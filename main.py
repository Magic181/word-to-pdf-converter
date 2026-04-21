from __future__ import annotations

import argparse
import ctypes
import gc
import json
import logging
import os
import queue
import subprocess
import sys
import threading
import time
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from collections.abc import Callable, Iterable
from dataclasses import asdict, dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore[import]
except ImportError:  # pragma: no cover - optional dependency
    DND_FILES = None
    TkinterDnD = None

try:
    import pythoncom  # type: ignore[import]
    import win32com.client  # type: ignore[import]
    from pywintypes import com_error  # type: ignore[import]
except ImportError:  # pragma: no cover - handled at runtime for missing dependency
    pythoncom = None
    win32com = None
    com_error = Exception


WORD_SUFFIXES = {".doc", ".docx"}
EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm", ".xlsb"}
POWERPOINT_SUFFIXES = {".ppt", ".pptx", ".pptm"}
SUPPORTED_INPUT_SUFFIXES = WORD_SUFFIXES | EXCEL_SUFFIXES | POWERPOINT_SUFFIXES
WD_EXPORT_FORMAT_PDF = 17
XL_TYPE_PDF = 0
PP_SAVE_AS_PDF = 32
PP_FIXED_FORMAT_TYPE_PDF = 2
WD_DO_NOT_SAVE_CHANGES = 0
WORD_EXIT_TIMEOUT_SECONDS = 5.0
BATCH_SESSION_RESTART_FAILURE_THRESHOLD = 2
DEFAULT_RETRY_ATTEMPTS = 1
DEFAULT_BATCH_WORKERS = 1
DEFAULT_WATCH_INTERVAL_SECONDS = 3.0
APP_CONFIG_PATH = Path(__file__).with_name("app_config.json")
APP_LOG_PATH = Path(__file__).with_name("office_to_pdf.log")


class WordConversionError(RuntimeError):
    def __init__(self, message: str, *, source: Path | None = None) -> None:
        self.source = source
        super().__init__(message)


@dataclass(slots=True)
class BatchResult:
    source: Path
    target: Path
    success: bool
    message: str


@dataclass(slots=True)
class ConversionOutcome:
    target: Path
    warning: str | None = None


@dataclass(slots=True)
class BatchConversionOutcome:
    results: list[BatchResult]
    warning: str | None = None
    duration_seconds: float = 0.0
    reused_single_session: bool = True
    session_restarts: int = 0
    parallel_workers_used: int = 1


@dataclass(slots=True)
class AppConfig:
    mode: str = "file"
    last_input: str = ""
    last_output: str = ""
    recursive: bool = False
    overwrite: bool = False
    failures_only: bool = False
    retry_attempts: int = DEFAULT_RETRY_ATTEMPTS
    batch_workers: int = DEFAULT_BATCH_WORKERS
    watch_enabled: bool = False
    watch_interval_seconds: float = DEFAULT_WATCH_INTERVAL_SECONDS


def setup_logging() -> logging.Logger:
    logger = logging.getLogger("office_to_pdf")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(threadName)s | %(message)s"
    )

    file_handler = logging.FileHandler(APP_LOG_PATH, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stderr)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    return logger


LOGGER = setup_logging()


def load_app_config() -> AppConfig:
    if not APP_CONFIG_PATH.exists():
        return AppConfig()

    try:
        raw = json.loads(APP_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        LOGGER.warning("Unable to load config from %s: %s", APP_CONFIG_PATH, exc)
        return AppConfig()

    config = AppConfig()
    for field_name in config.__dataclass_fields__:
        if field_name in raw:
            setattr(config, field_name, raw[field_name])

    config.mode = config.mode if config.mode in {"file", "directory"} else "file"
    config.retry_attempts = max(0, int(config.retry_attempts))
    config.batch_workers = max(1, int(config.batch_workers))
    config.watch_interval_seconds = max(1.0, float(config.watch_interval_seconds))
    return config


def save_app_config(config: AppConfig) -> None:
    try:
        APP_CONFIG_PATH.write_text(
            json.dumps(asdict(config), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except OSError as exc:
        LOGGER.warning("Unable to save config to %s: %s", APP_CONFIG_PATH, exc)


def is_word_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in WORD_SUFFIXES


def is_excel_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in EXCEL_SUFFIXES


def is_supported_input_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in SUPPORTED_INPUT_SUFFIXES


def detect_input_kind(path: Path) -> str | None:
    suffix = path.suffix.lower()
    if suffix in WORD_SUFFIXES:
        return "word"
    if suffix in EXCEL_SUFFIXES:
        return "excel"
    if suffix in POWERPOINT_SUFFIXES:
        return "powerpoint"
    return None


def supported_input_patterns() -> str:
    return " ".join(f"*{suffix}" for suffix in sorted(SUPPORTED_INPUT_SUFFIXES))


def supported_input_label() -> str:
    return ".doc/.docx/.xls/.xlsx/.xlsm/.xlsb/.ppt/.pptx/.pptm"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert Office documents to PDF on Windows."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        help=(
            "Path to a supported file or a directory containing supported files "
            f"({supported_input_label()})."
        ),
    )
    parser.add_argument(
        "-o",
        "--output",
        dest="output_path",
        help=(
            "Optional target PDF path for a single file, or target directory when "
            "the input is a directory."
        ),
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing PDF files.",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Recursively scan subdirectories when the input is a directory.",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch the desktop GUI.",
    )
    parser.add_argument(
        "--retry-attempts",
        type=int,
        default=DEFAULT_RETRY_ATTEMPTS,
        help="Retry a failed file this many extra times before marking it as failed.",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=DEFAULT_BATCH_WORKERS,
        help="Number of parallel batch workers. Values above 1 may open multiple Office instances.",
    )
    return parser


def ensure_dependencies() -> None:
    if pythoncom is None or win32com is None:
        raise WordConversionError(
            "Missing dependency: pywin32. Install it with `pip install -r requirements.txt`."
        )


def resolve_input_path(input_path: str) -> Path:
    source = Path(input_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"Input path does not exist: {source}")
    return source


def resolve_file_paths(input_path: str, output_path: str | None) -> tuple[Path, Path]:
    source = resolve_input_path(input_path)
    if not is_supported_input_file(source):
        raise ValueError(
            f"Input file must be one of: {supported_input_label()}."
        )

    if output_path:
        target = Path(output_path).expanduser().resolve()
    else:
        target = source.with_suffix(".pdf")

    if target.suffix.lower() != ".pdf":
        raise ValueError("Output file must use the .pdf extension.")

    return source, target


def collect_supported_files(directory: Path, recursive: bool = False) -> list[Path]:
    iterator: Iterable[Path]
    if recursive:
        iterator = directory.rglob("*")
    else:
        iterator = directory.iterdir()

    files = [path for path in iterator if is_supported_input_file(path)]
    return sorted(files, key=lambda path: str(path).lower())


def build_batch_target(source: Path, input_root: Path, output_root: Path | None) -> Path:
    if output_root is None:
        return source.with_suffix(".pdf")
    relative_path = source.relative_to(input_root)
    return (output_root / relative_path).with_suffix(".pdf")


def build_list_target(source: Path, output_root: Path | None) -> Path:
    if output_root is None:
        return source.with_suffix(".pdf")
    return (output_root / source.with_suffix(".pdf").name).resolve()


def _format_com_error(exc: Exception) -> str:
    text = str(exc).strip()
    if text:
        return text.replace("\r", " ").replace("\n", " ")
    return exc.__class__.__name__


def _get_word_process_id(word_app: object) -> int | None:
    try:
        hwnd = int(word_app.Hwnd)
    except Exception:
        return None

    process_id = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
    return int(process_id.value) or None


def _wait_for_process_exit(pid: int, timeout_seconds: float) -> bool:
    def is_process_running(process_id: int) -> bool:
        result = subprocess.run(
            ["tasklist", "/FI", f"PID eq {process_id}", "/FO", "CSV", "/NH"],
            check=False,
            capture_output=True,
            text=True,
        )
        line = result.stdout.strip().strip('"')
        if not line:
            return False
        return "No tasks are running" not in line and "INFO:" not in line

    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if not is_process_running(pid):
            return True
        time.sleep(0.1)
    return False


def _force_terminate_process(pid: int) -> None:
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/T", "/F"],
        check=False,
        capture_output=True,
        text=True,
    )


def _format_duration(seconds: float) -> str:
    if seconds < 1:
        return f"{seconds:.2f}s"
    if seconds < 60:
        return f"{seconds:.1f}s"
    minutes, remainder = divmod(seconds, 60)
    return f"{int(minutes)}m {remainder:.1f}s"


def _close_document(document: object | None) -> str | None:
    if document is None:
        return None

    try:
        try:
            document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        except TypeError:
            document.Close()
    except Exception as exc:
        return f"Document close warning: {_format_com_error(exc)}"
    return None


def _quit_office_application(
    app_object: object | None, pid: int | None, app_name: str
) -> str | None:
    if app_object is None:
        return None

    quit_error: str | None = None
    try:
        try:
            app_object.Quit(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        except TypeError:
            app_object.Quit()
    except Exception as exc:
        quit_error = f"{app_name} quit warning: {_format_com_error(exc)}"

    if pid is not None and not _wait_for_process_exit(pid, WORD_EXIT_TIMEOUT_SECONDS):
        _force_terminate_process(pid)
        if not _wait_for_process_exit(pid, 2.0):
            suffix = f"{app_name} process {pid} is still running after forced termination."
            quit_error = f"{quit_error} {suffix}".strip() if quit_error else suffix
        else:
            suffix = f"{app_name} process {pid} did not exit cleanly and was terminated."
            quit_error = f"{quit_error} {suffix}".strip() if quit_error else suffix

    return quit_error


class OfficeAutomationSession:
    app_name = "Office"
    prog_id = ""
    visible = False
    resource_label = "Document"
    cleanup_subject = "document"
    unavailable_message = "Office automation session is not available."
    export_failure_message = "Office export failed."

    def __init__(self) -> None:
        self.app = None
        self.pid: int | None = None
        self.cleanup_warnings: list[str] = []
        self._initialized = False

    def __enter__(self):
        ensure_dependencies()
        pythoncom.CoInitialize()
        self._initialized = True
        try:
            self.app = win32com.client.DispatchEx(self.prog_id)
            self._configure_application()
            self.pid = _get_word_process_id(self.app)
        except Exception:
            if self._initialized:
                pythoncom.CoUninitialize()
                self._initialized = False
            raise
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        quit_warning = _quit_office_application(self.app, self.pid, self.app_name)
        if quit_warning:
            self.cleanup_warnings.append(quit_warning)
        self.app = None
        self.pid = None
        gc.collect()
        if self._initialized:
            pythoncom.CoUninitialize()
            self._initialized = False

    def _configure_application(self) -> None:
        if self.app is not None:
            self.app.Visible = self.visible

    def _prepare_target(self, source: Path, target: Path, overwrite: bool) -> None:
        if target.exists() and not overwrite:
            raise FileExistsError(
                f"Target file already exists: {target}. Use --overwrite to replace it."
            )
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise WordConversionError(
                f"Unable to create output directory: {target.parent} ({exc})",
                source=source,
            ) from exc

    def _close_resource(self, resource: object | None) -> str | None:
        warning = _close_document(resource)
        if warning and self.resource_label != "Document":
            return warning.replace("Document", self.resource_label)
        return warning

    def _open_resource(self, source: Path) -> object:
        raise NotImplementedError

    def _export_resource(self, resource: object, target: Path) -> None:
        raise NotImplementedError

    def convert_file(
        self, source: Path, target: Path, overwrite: bool = False
    ) -> ConversionOutcome:
        self._prepare_target(source, target, overwrite)

        resource = None
        local_warnings: list[str] = []
        try:
            if self.app is None:
                raise WordConversionError(self.unavailable_message, source=source)

            resource = self._open_resource(source)
            self._export_resource(resource, target)
        except com_error as exc:
            details = _format_com_error(exc)
            raise WordConversionError(
                f"{self.export_failure_message} Details: {details}",
                source=source,
            ) from exc
        except Exception as exc:
            if isinstance(exc, WordConversionError):
                raise
            raise WordConversionError(
                f"Unexpected conversion failure for {source.name}: {exc}",
                source=source,
            ) from exc
        finally:
            close_warning = self._close_resource(resource)
            if close_warning:
                local_warnings.append(close_warning)

        if not target.exists():
            details = " ".join(local_warnings).strip()
            message = f"Export did not create the PDF file: {target}"
            if details:
                message = f"{message} Cleanup notes: {details}"
            raise WordConversionError(message, source=source)

        warning = None
        if local_warnings:
            warning = (
                f"Converted successfully, but {self.cleanup_subject} cleanup reported warnings. "
                f"{' '.join(local_warnings)}"
            )
        return ConversionOutcome(target=target, warning=warning)


class WordAutomationSession(OfficeAutomationSession):
    app_name = "Word"
    prog_id = "Word.Application"
    visible = False
    resource_label = "Document"
    cleanup_subject = "document"
    unavailable_message = "Word automation session is not available."
    export_failure_message = (
        "Word export failed. Make sure Microsoft Word is installed and the document can be opened normally."
    )

    def _configure_application(self) -> None:
        super()._configure_application()
        if self.app is not None:
            self.app.DisplayAlerts = 0

    def _open_resource(self, source: Path) -> object:
        return self.app.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )

    def _export_resource(self, resource: object, target: Path) -> None:
        resource.ExportAsFixedFormat(
            OutputFileName=str(target),
            ExportFormat=WD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )


class ExcelAutomationSession(OfficeAutomationSession):
    app_name = "Excel"
    prog_id = "Excel.Application"
    visible = False
    resource_label = "Workbook"
    cleanup_subject = "workbook"
    unavailable_message = "Excel automation session is not available."
    export_failure_message = (
        "Excel export failed. Make sure Microsoft Excel is installed and the document can be opened normally."
    )

    def _configure_application(self) -> None:
        super()._configure_application()
        if self.app is not None:
            self.app.DisplayAlerts = False

    def _open_resource(self, source: Path) -> object:
        return self.app.Workbooks.Open(str(source), ReadOnly=True)

    def _export_resource(self, resource: object, target: Path) -> None:
        resource.ExportAsFixedFormat(
            Type=XL_TYPE_PDF,
            Filename=str(target),
            OpenAfterPublish=False,
        )


class PowerPointAutomationSession(OfficeAutomationSession):
    app_name = "PowerPoint"
    prog_id = "PowerPoint.Application"
    visible = True
    resource_label = "Presentation"
    cleanup_subject = "presentation"
    unavailable_message = "PowerPoint automation session is not available."
    export_failure_message = (
        "PowerPoint export failed. Make sure Microsoft PowerPoint is installed and the document can be opened normally."
    )

    def _open_resource(self, source: Path) -> object:
        return self.app.Presentations.Open(
            str(source),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )

    def _export_resource(self, resource: object, target: Path) -> None:
        resource.ExportAsFixedFormat(
            str(target),
            PP_FIXED_FORMAT_TYPE_PDF,
            1,
            False,
            2,
            1,
            False,
            None,
        )


def _merge_outcome_cleanup_warnings(
    outcome: ConversionOutcome, session: OfficeAutomationSession
) -> ConversionOutcome:
    if not session.cleanup_warnings:
        return outcome

    cleanup_warning = (
        f"Converted successfully, but {session.app_name} cleanup reported warnings. "
        f"{' '.join(session.cleanup_warnings)}"
    )
    warning = " ".join(part for part in [outcome.warning, cleanup_warning] if part)
    return ConversionOutcome(target=outcome.target, warning=warning)


def convert_supported_file_to_pdf(
    source: Path, target: Path, overwrite: bool = False
) -> ConversionOutcome:
    kind = detect_input_kind(source)
    if kind is None:
        raise WordConversionError(
            f"Unsupported input format: {source.suffix}. Supported: {supported_input_label()}",
            source=source,
        )

    session_class = {
        "word": WordAutomationSession,
        "excel": ExcelAutomationSession,
        "powerpoint": PowerPointAutomationSession,
    }[kind]
    with session_class() as session:
        outcome = session.convert_file(source, target, overwrite=overwrite)
        return _merge_outcome_cleanup_warnings(outcome, session)


def convert_word_to_pdf(
    source: Path, target: Path, overwrite: bool = False
) -> ConversionOutcome:
    return convert_supported_file_to_pdf(source, target, overwrite=overwrite)


def _open_session_for_kind(kind: str):
    if kind == "word":
        return WordAutomationSession().__enter__()
    if kind == "excel":
        return ExcelAutomationSession().__enter__()
    if kind == "powerpoint":
        return PowerPointAutomationSession().__enter__()
    raise WordConversionError(f"Unsupported session kind: {kind}")


def _extract_session_pid(session: object) -> int | None:
    return getattr(session, "pid", None)


def _drain_session_warnings(session: object, collector: list[str]) -> None:
    warnings = list(getattr(session, "cleanup_warnings", []))
    if warnings:
        collector.extend(warnings)
        getattr(session, "cleanup_warnings").clear()

@dataclass(slots=True)
class BatchTask:
    index: int
    source: Path
    target: Path
    kind: str


class ProgressTracker:
    def __init__(
        self,
        total: int,
        callback: Callable[[int, int, BatchResult], None] | None,
    ) -> None:
        self.total = total
        self.callback = callback
        self._completed = 0
        self._lock = threading.Lock()

    def emit(self, result: BatchResult) -> None:
        if self.callback is None:
            return
        with self._lock:
            self._completed += 1
            completed = self._completed
        self.callback(completed, self.total, result)


def _normalize_supported_sources(sources: Iterable[Path]) -> list[Path]:
    files = [path.resolve() for path in sources if is_supported_input_file(path.resolve())]
    if not files:
        raise FileNotFoundError(
            f"No valid supported files were provided for batch conversion ({supported_input_label()})."
        )
    return files


def _build_batch_tasks(
    sources: list[Path],
    target_builder: Callable[[Path], Path],
) -> list[BatchTask]:
    tasks: list[BatchTask] = []
    for index, source in enumerate(sources):
        kind = detect_input_kind(source)
        if kind is None:
            continue
        tasks.append(
            BatchTask(
                index=index,
                source=source,
                target=target_builder(source),
                kind=kind,
            )
        )
    return tasks


def _restart_session(
    session: OfficeAutomationSession | None,
    kind: str,
    cleanup_warnings: list[str],
) -> OfficeAutomationSession:
    if session is not None:
        session.__exit__(None, None, None)
        _drain_session_warnings(session, cleanup_warnings)
    try:
        return _open_session_for_kind(kind)
    except Exception as exc:
        raise WordConversionError(f"Unable to start {kind} session: {exc}") from exc


def _convert_with_retries(
    session: OfficeAutomationSession,
    task: BatchTask,
    *,
    overwrite: bool,
    retry_attempts: int,
    cleanup_warnings: list[str],
) -> tuple[OfficeAutomationSession, ConversionOutcome, int]:
    restarts = 0
    active_session = session
    last_error: Exception | None = None
    total_attempts = retry_attempts + 1

    for attempt in range(1, total_attempts + 1):
        try:
            outcome = active_session.convert_file(
                task.source,
                task.target,
                overwrite=overwrite,
            )
            return active_session, outcome, restarts
        except Exception as exc:
            last_error = exc
            LOGGER.warning(
                "Conversion attempt %s/%s failed for %s: %s",
                attempt,
                total_attempts,
                task.source,
                exc,
            )
            if attempt >= total_attempts:
                break
            active_session = _restart_session(active_session, task.kind, cleanup_warnings)
            restarts += 1

    assert last_error is not None
    raise last_error


def _process_task_chunk(
    tasks: list[BatchTask],
    *,
    overwrite: bool,
    retry_attempts: int,
    progress_tracker: ProgressTracker,
) -> tuple[list[BatchTask], list[BatchResult], list[str], int]:
    if not tasks:
        return [], [], [], 0

    results: list[BatchResult] = []
    cleanup_warnings: list[str] = []
    session_restarts = 0
    consecutive_failures = 0
    session: OfficeAutomationSession | None = None
    current_kind: str | None = None

    try:
        for offset, task in enumerate(tasks):
            if session is None or current_kind != task.kind:
                session = _restart_session(session, task.kind, cleanup_warnings)
                current_kind = task.kind
                consecutive_failures = 0

            try:
                session, outcome, retry_restarts = _convert_with_retries(
                    session,
                    task,
                    overwrite=overwrite,
                    retry_attempts=retry_attempts,
                    cleanup_warnings=cleanup_warnings,
                )
                session_restarts += retry_restarts
                result = BatchResult(
                    source=task.source,
                    target=task.target,
                    success=True,
                    message=outcome.warning or "Converted successfully.",
                )
                consecutive_failures = 0
            except Exception as exc:
                result = BatchResult(
                    source=task.source,
                    target=task.target,
                    success=False,
                    message=str(exc),
                )
                consecutive_failures += 1

            results.append(result)
            progress_tracker.emit(result)

            if (
                consecutive_failures >= BATCH_SESSION_RESTART_FAILURE_THRESHOLD
                and offset < len(tasks) - 1
                and session is not None
            ):
                previous_pid = _extract_session_pid(session)
                session = _restart_session(session, task.kind, cleanup_warnings)
                session_restarts += 1
                consecutive_failures = 0
                LOGGER.warning(
                    "Restarted %s session after consecutive failures. Previous PID=%s",
                    task.kind,
                    previous_pid,
                )
    finally:
        if session is not None:
            session.__exit__(None, None, None)
            _drain_session_warnings(session, cleanup_warnings)

    return tasks, results, cleanup_warnings, session_restarts


def _split_tasks_for_parallelism(
    tasks: list[BatchTask], max_workers: int
) -> list[list[BatchTask]]:
    worker_count = max(1, min(max_workers, len(tasks)))
    if worker_count == 1:
        return [tasks]

    chunks: list[list[BatchTask]] = [[] for _ in range(worker_count)]
    for index, task in enumerate(tasks):
        chunks[index % worker_count].append(task)
    return [chunk for chunk in chunks if chunk]


def _run_batch_tasks(
    tasks: list[BatchTask],
    *,
    overwrite: bool,
    retry_attempts: int,
    max_workers: int,
    progress_callback: Callable[[int, int, BatchResult], None] | None = None,
) -> BatchConversionOutcome:
    if not tasks:
        raise FileNotFoundError("No supported files are available for batch conversion.")

    start_time = time.perf_counter()
    progress_tracker = ProgressTracker(len(tasks), progress_callback)
    cleanup_warnings: list[str] = []
    session_restarts = 0
    indexed_results: list[tuple[int, BatchResult]] = []
    chunks = _split_tasks_for_parallelism(tasks, max_workers)
    parallel_workers_used = min(max_workers, len(chunks))

    if parallel_workers_used == 1:
        chunk_tasks, chunk_results, chunk_warnings, chunk_restarts = _process_task_chunk(
            chunks[0],
            overwrite=overwrite,
            retry_attempts=retry_attempts,
            progress_tracker=progress_tracker,
        )
        cleanup_warnings.extend(chunk_warnings)
        session_restarts += chunk_restarts
        indexed_results.extend(
            (task.index, result) for task, result in zip(chunk_tasks, chunk_results)
        )
    else:
        with ThreadPoolExecutor(
            max_workers=parallel_workers_used,
            thread_name_prefix="office-batch",
        ) as executor:
            futures: list[Future[tuple[list[BatchTask], list[BatchResult], list[str], int]]] = [
                executor.submit(
                    _process_task_chunk,
                    chunk,
                    overwrite=overwrite,
                    retry_attempts=retry_attempts,
                    progress_tracker=progress_tracker,
                )
                for chunk in chunks
            ]
            for future in as_completed(futures):
                chunk_tasks, chunk_results, chunk_warnings, chunk_restarts = future.result()
                cleanup_warnings.extend(chunk_warnings)
                session_restarts += chunk_restarts
                indexed_results.extend(
                    (task.index, result)
                    for task, result in zip(chunk_tasks, chunk_results)
                )

    results = [result for _, result in sorted(indexed_results, key=lambda item: item[0])]
    warning = None
    if cleanup_warnings:
        warning = (
            "Batch finished, but Office cleanup reported warnings. "
            f"{' '.join(cleanup_warnings)}"
        )

    duration_seconds = time.perf_counter() - start_time
    return BatchConversionOutcome(
        results=results,
        warning=warning,
        duration_seconds=duration_seconds,
        reused_single_session=(parallel_workers_used == 1 and session_restarts == 0),
        session_restarts=session_restarts,
        parallel_workers_used=parallel_workers_used,
    )


def convert_directory_to_pdf(
    input_dir: Path,
    output_dir: Path | None = None,
    *,
    recursive: bool = False,
    overwrite: bool = False,
    retry_attempts: int = DEFAULT_RETRY_ATTEMPTS,
    max_workers: int = DEFAULT_BATCH_WORKERS,
    progress_callback: Callable[[int, int, BatchResult], None] | None = None,
) -> BatchConversionOutcome:
    if not input_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {input_dir}")

    files = collect_supported_files(input_dir, recursive=recursive)
    if not files:
        raise FileNotFoundError(
            f"No supported files found in directory: {input_dir} ({supported_input_label()})"
        )

    tasks = _build_batch_tasks(
        files,
        lambda source: build_batch_target(source, input_dir, output_dir),
    )
    return _run_batch_tasks(
        tasks,
        overwrite=overwrite,
        retry_attempts=retry_attempts,
        max_workers=max_workers,
        progress_callback=progress_callback,
    )


def convert_file_list_to_pdf(
    sources: list[Path],
    output_dir: Path | None = None,
    *,
    overwrite: bool = False,
    retry_attempts: int = DEFAULT_RETRY_ATTEMPTS,
    max_workers: int = DEFAULT_BATCH_WORKERS,
    progress_callback: Callable[[int, int, BatchResult], None] | None = None,
) -> BatchConversionOutcome:
    if not sources:
        raise FileNotFoundError("No files were provided for batch conversion.")

    files = _normalize_supported_sources(sources)
    tasks = _build_batch_tasks(
        files,
        lambda source: build_list_target(source, output_dir),
    )
    return _run_batch_tasks(
        tasks,
        overwrite=overwrite,
        retry_attempts=retry_attempts,
        max_workers=max_workers,
        progress_callback=progress_callback,
    )


def format_batch_summary(results: list[BatchResult]) -> str:
    success_count = sum(1 for result in results if result.success)
    failed_count = len(results) - success_count
    return f"Completed batch conversion. Success: {success_count}, Failed: {failed_count}"


def run_cli(args: argparse.Namespace) -> int:
    if not args.input_path:
        raise ValueError("input_path is required unless --gui is used.")

    source = resolve_input_path(args.input_path)

    if source.is_file():
        file_source, target = resolve_file_paths(args.input_path, args.output_path)
        outcome = convert_supported_file_to_pdf(
            file_source,
            target,
            overwrite=args.overwrite,
        )
        print(f"Converted successfully: {outcome.target}")
        if outcome.warning:
            print(f"Warning: {outcome.warning}", file=sys.stderr)
        return 0

    if source.is_dir():
        output_dir = (
            Path(args.output_path).expanduser().resolve() if args.output_path else None
        )
        batch_outcome = convert_directory_to_pdf(
            source,
            output_dir,
            recursive=args.recursive,
            overwrite=args.overwrite,
            retry_attempts=max(0, args.retry_attempts),
            max_workers=max(1, args.workers),
        )
        for result in batch_outcome.results:
            status = "OK" if result.success else "FAILED"
            print(f"[{status}] {result.source} -> {result.target}")
            if not result.success:
                print(f"        {result.message}")
            elif result.message != "Converted successfully.":
                print(f"        {result.message}")

        print(format_batch_summary(batch_outcome.results))
        if batch_outcome.parallel_workers_used > 1:
            session_note = (
                f"Batch used {batch_outcome.parallel_workers_used} parallel Office worker(s) "
                f"with {batch_outcome.session_restarts} session restart(s)."
            )
        elif batch_outcome.reused_single_session:
            session_note = "Batch used a single reusable Office session."
        else:
            session_note = (
                "Batch reused Office sessions with "
                f"{batch_outcome.session_restarts} automatic restart(s)."
            )
        print(f"Info: {session_note}")
        print(f"Info: Batch duration: {_format_duration(batch_outcome.duration_seconds)}")
        if batch_outcome.warning:
            print(f"Warning: {batch_outcome.warning}", file=sys.stderr)
        return 0 if all(result.success for result in batch_outcome.results) else 1

    raise ValueError(f"Unsupported input path: {source}")


class WordToPdfApp:
    def __init__(self) -> None:
        self.config = load_app_config()
        self.root = TkinterDnD.Tk() if TkinterDnD is not None else tk.Tk()
        self.root.title("Office to PDF Converter")
        self.root.geometry("1380x920")
        self.root.minsize(1240, 860)
        self.root.configure(bg="#eef3f1")

        self.mode_var = tk.StringVar(value=self.config.mode)
        self.input_var = tk.StringVar(value=self.config.last_input)
        self.output_var = tk.StringVar(value=self.config.last_output)
        self.recursive_var = tk.BooleanVar(value=self.config.recursive)
        self.overwrite_var = tk.BooleanVar(value=self.config.overwrite)
        self.failures_only_var = tk.BooleanVar(value=self.config.failures_only)
        self.retry_attempts_var = tk.IntVar(value=self.config.retry_attempts)
        self.batch_workers_var = tk.IntVar(value=self.config.batch_workers)
        self.watch_enabled_var = tk.BooleanVar(value=self.config.watch_enabled)
        self.drop_hint_var = tk.StringVar(
            value=(
                "Drag a Word/Excel/PowerPoint file or folder onto the window. "
                "You can also drop multiple supported files for a one-off batch."
                if TkinterDnD is not None
                else "Install tkinterdnd2 to enable drag and drop."
            )
        )
        self.status_var = tk.StringVar(
            value="Ready. Select an Office file or a directory to begin."
        )
        self.summary_var = tk.StringVar(value="No results yet.")
        self.progress_var = tk.DoubleVar(value=0.0)
        self.preview_title_var = tk.StringVar(
            value="Select a result row to inspect source and output details."
        )
        self.preview_source_var = tk.StringVar(value="-")
        self.preview_output_var = tk.StringVar(value="-")
        self.preview_details_var = tk.StringVar(value="-")
        self.worker_thread: threading.Thread | None = None
        self.event_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.result_rows: list[dict[str, str]] = []
        self.sort_column = "status"
        self.sort_descending = False
        self._results_refresh_job: str | None = None
        self._results_should_scroll_end = False
        self._results_prefer_last_selection = True
        self._pending_selection_key: tuple[str, str, str] | None = None
        self.dropped_files: list[Path] = []
        self._auto_output_path: Path | None = None
        self.watch_snapshot: dict[Path, tuple[int, int]] = {}
        self.drop_hint_default_bg = "#f6f8fb"
        self.drop_hint_active_bg = "#e7f0fb"
        self.drop_hint_default_border = self.colors["border"] if hasattr(self, "colors") else "#d8dee6"

        self._apply_styles()
        self._build_ui()
        self._build_context_menu()
        self._register_drop_targets()
        self._toggle_mode()
        self._sync_output_with_input(force=False)
        self._update_preview_panel(None)
        self.root.protocol("WM_DELETE_WINDOW", self._handle_close)
        self.root.after(80, self._poll_events)
        self.root.after(
            int(max(1.0, self.config.watch_interval_seconds) * 1000),
            self._poll_folder_watch,
        )

    def _apply_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            try:
                style.theme_use("xpnative")
            except tk.TclError:
                try:
                    style.theme_use("default")
                except tk.TclError:
                    pass

        self.colors = {
            "bg": "#f3f5f7",
            "panel": "#f7f8fa",
            "card": "#ffffff",
            "hero": "#e9edf2",
            "hero_accent": "#c9d2dc",
            "text": "#17212b",
            "muted": "#66717d",
            "border": "#d8dee6",
            "input": "#fdfefe",
            "primary": "#244a73",
            "primary_active": "#1d3a59",
            "secondary": "#eef2f6",
            "secondary_active": "#e2e8ef",
            "drop_active_border": "#6f8fb3",
        }

        style.configure(".", font=("Segoe UI", 10))
        style.configure("App.TFrame", background=self.colors["bg"])
        style.configure("Card.TFrame", background=self.colors["card"])
        style.configure(
            "Card.TLabelframe",
            background=self.colors["card"],
            bordercolor=self.colors["border"],
            relief="solid",
            borderwidth=1,
            padding=0,
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=self.colors["card"],
            foreground=self.colors["text"],
            font=("Segoe UI Semibold", 10),
        )
        style.configure(
            "Body.TLabel",
            background=self.colors["card"],
            foreground=self.colors["text"],
        )
        style.configure(
            "Muted.TLabel",
            background=self.colors["card"],
            foreground=self.colors["muted"],
        )
        style.configure(
            "Primary.TButton",
            font=("Segoe UI Semibold", 10),
            padding=(16, 10),
            foreground="#ffffff",
            background=self.colors["primary"],
            borderwidth=1,
            focusthickness=1,
            focuscolor=self.colors["primary"],
            relief="flat",
        )
        style.map(
            "Primary.TButton",
            background=[
                ("disabled", "#94b9ae"),
                ("active", self.colors["primary_active"]),
                ("pressed", self.colors["primary_active"]),
            ],
            foreground=[
                ("disabled", "#f3f7f5"),
                ("active", "#ffffff"),
                ("pressed", "#ffffff"),
            ],
        )
        style.configure(
            "Secondary.TButton",
            padding=(14, 10),
            foreground=self.colors["text"],
            background=self.colors["secondary"],
            borderwidth=1,
            focusthickness=1,
            focuscolor=self.colors["secondary"],
            relief="flat",
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("active", self.colors["secondary_active"]),
                ("pressed", self.colors["secondary_active"]),
            ],
            foreground=[
                ("active", self.colors["text"]),
                ("pressed", self.colors["text"]),
            ],
        )
        style.configure(
            "App.TRadiobutton",
            background=self.colors["card"],
            foreground=self.colors["text"],
            font=("Segoe UI", 10),
        )
        style.map(
            "App.TRadiobutton",
            background=[("active", self.colors["card"])],
        )
        style.configure(
            "App.TCheckbutton",
            background=self.colors["card"],
            foreground=self.colors["text"],
            font=("Segoe UI", 10),
        )
        style.map(
            "App.TCheckbutton",
            background=[("active", self.colors["card"])],
        )
        style.configure(
            "App.Horizontal.TProgressbar",
            troughcolor="#dfe5ec",
            background=self.colors["primary"],
            bordercolor="#dfe5ec",
            lightcolor=self.colors["primary"],
            darkcolor=self.colors["primary"],
            thickness=10,
        )
        style.configure(
            "Results.Treeview",
            background="#f7f9fb",
            fieldbackground="#f7f9fb",
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            rowheight=34,
            relief="flat",
            font=("Segoe UI", 9),
        )
        style.map(
            "Results.Treeview",
            background=[("selected", "#dbe7f3")],
            foreground=[("selected", self.colors["text"])],
        )
        style.configure(
            "Results.Treeview.Heading",
            background="#eef2f6",
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            relief="flat",
            font=("Segoe UI Semibold", 9),
            padding=(8, 8),
        )

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=28, style="App.TFrame")
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)

        hero = tk.Frame(
            container,
            bg=self.colors["hero"],
            padx=28,
            pady=22,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
        )
        hero.pack(fill="x")

        hero_top = tk.Frame(hero, bg=self.colors["hero"])
        hero_top.pack(fill="x")
        tk.Label(
            hero_top,
            text="Office to PDF Converter",
            font=("Segoe UI Semibold", 22),
            fg=self.colors["text"],
            bg=self.colors["hero"],
        ).pack(anchor="w")
        tk.Label(
            hero_top,
            text="A desktop utility for reliable Office-to-PDF conversion in office workflows.",
            font=("Segoe UI", 10),
            fg=self.colors["muted"],
            bg=self.colors["hero"],
        ).pack(anchor="w", pady=(6, 0))

        meta_row = tk.Frame(hero, bg=self.colors["hero"])
        meta_row.pack(anchor="w", pady=(14, 0))
        tk.Label(
            meta_row,
            text="Windows + Microsoft Office required",
            font=("Segoe UI", 9),
            fg=self.colors["muted"],
            bg=self.colors["hero"],
        ).pack(side="left")
        tk.Label(
            meta_row,
            text="  |  ",
            font=("Segoe UI", 9),
            fg=self.colors["hero_accent"],
            bg=self.colors["hero"],
        ).pack(side="left")
        tk.Label(
            meta_row,
            text="Single-file and batch conversion",
            font=("Segoe UI", 9),
            fg=self.colors["muted"],
            bg=self.colors["hero"],
        ).pack(side="left")
        tk.Label(
            meta_row,
            text="  |  ",
            font=("Segoe UI", 9),
            fg=self.colors["hero_accent"],
            bg=self.colors["hero"],
        ).pack(side="left")
        tk.Label(
            meta_row,
            text=(
                "Drag files or folders into the window"
                if TkinterDnD is not None
                else "Install tkinterdnd2 to enable drag and drop"
            ),
            font=("Segoe UI", 9),
            fg=self.colors["muted"],
            bg=self.colors["hero"],
        ).pack(side="left")

        self.drop_hint_label = tk.Label(
            hero,
            textvariable=self.drop_hint_var,
            font=("Segoe UI", 10),
            fg=self.colors["text"],
            bg=self.drop_hint_default_bg,
            padx=14,
            pady=10,
            anchor="w",
            justify="left",
            highlightthickness=1,
            highlightbackground=self.colors["border"],
        )
        self.drop_hint_label.pack(fill="x", pady=(16, 0))
        self.drop_hint_default_border = self.colors["border"]

        body = ttk.Frame(container, style="App.TFrame")
        body.pack(fill="both", expand=True, pady=(22, 0))
        body.columnconfigure(0, weight=6, minsize=520)
        body.columnconfigure(1, weight=8, minsize=780)
        body.rowconfigure(1, weight=1)

        self.left_column = ttk.Frame(body, style="App.TFrame")
        self.left_column.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 20))
        self.left_column.columnconfigure(0, weight=1)

        self.right_column = ttk.Frame(body, style="App.TFrame")
        self.right_column.grid(row=0, column=1, rowspan=2, sticky="nsew")
        self.right_column.columnconfigure(0, weight=1)
        self.right_column.rowconfigure(2, weight=1)

        mode_frame = ttk.LabelFrame(self.left_column, text="Mode", padding=20, style="Card.TLabelframe")
        mode_frame.pack(fill="x")
        mode_frame.columnconfigure(0, weight=1)
        mode_frame.columnconfigure(1, weight=1)

        ttk.Radiobutton(
            mode_frame,
            text="Single file",
            value="file",
            variable=self.mode_var,
            command=self._toggle_mode,
            style="App.TRadiobutton",
        ).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            mode_frame,
            text="Directory batch",
            value="directory",
            variable=self.mode_var,
            command=self._toggle_mode,
            style="App.TRadiobutton",
        ).grid(row=0, column=1, sticky="w", padx=(22, 0))

        mode_hint = ttk.Label(
            mode_frame,
            text="Use batch mode to convert every supported Office file inside a folder.",
            style="Muted.TLabel",
        )
        mode_hint.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        mode_hint.configure(wraplength=420, justify="left")

        path_frame = ttk.LabelFrame(self.left_column, text="Paths", padding=20, style="Card.TLabelframe")
        path_frame.pack(fill="x", pady=(16, 0))
        path_frame.columnconfigure(1, weight=1)
        path_frame.columnconfigure(2, minsize=126)

        ttk.Label(path_frame, text="Input", style="Body.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        self.input_entry = ttk.Entry(path_frame, textvariable=self.input_var, width=44)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=(12, 10), pady=(0, 8))
        ttk.Button(
            path_frame,
            text="Browse",
            command=self._browse_input,
            style="Secondary.TButton",
        ).grid(
            row=0, column=2, sticky="ew", pady=(0, 8)
        )

        self.output_label = ttk.Label(path_frame, text="Output", style="Body.TLabel")
        self.output_label.grid(row=1, column=0, sticky="w", pady=(0, 8))
        self.output_entry = ttk.Entry(path_frame, textvariable=self.output_var, width=44)
        self.output_entry.grid(row=1, column=1, sticky="ew", padx=(12, 10), pady=(0, 8))
        ttk.Button(
            path_frame,
            text="Browse",
            command=self._browse_output,
            style="Secondary.TButton",
        ).grid(
            row=1, column=2, sticky="ew", pady=(0, 8)
        )

        ttk.Label(
            path_frame,
            text="Leave output empty to export beside the source file or folder.",
            style="Muted.TLabel",
        ).grid(row=2, column=0, columnspan=3, sticky="ew", pady=(2, 0))
        path_frame.grid_slaves(row=2, column=0)[0].configure(wraplength=420, justify="left")

        self.pending_frame = ttk.LabelFrame(
            self.left_column,
            text="Dropped Files",
            padding=20,
            style="Card.TLabelframe",
        )
        self.pending_frame.columnconfigure(0, weight=1)

        ttk.Label(
            self.pending_frame,
            text=(
                "Files dropped for one-off batch conversion "
                "(.doc/.docx/.xls/.xlsx/.xlsm/.xlsb/.ppt/.pptx/.pptm)"
            ),
            style="Muted.TLabel",
        ).grid(row=0, column=0, sticky="w")

        list_frame = ttk.Frame(self.pending_frame, style="Card.TFrame")
        list_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.pending_listbox = tk.Listbox(
            list_frame,
            height=6,
            activestyle="none",
            font=("Segoe UI", 9),
            bg="#f7f9fb",
            fg=self.colors["text"],
            selectbackground="#dbe7f3",
            selectforeground=self.colors["text"],
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            relief="flat",
        )
        self.pending_listbox.grid(row=0, column=0, sticky="nsew")

        pending_scroll = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.pending_listbox.yview
        )
        pending_scroll.grid(row=0, column=1, sticky="ns")
        self.pending_listbox.configure(yscrollcommand=pending_scroll.set)

        pending_actions = ttk.Frame(self.pending_frame, style="Card.TFrame")
        pending_actions.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        pending_actions.columnconfigure(0, weight=1)
        pending_actions.columnconfigure(1, weight=1)

        self.remove_pending_button = tk.Button(
            pending_actions,
            text="Remove Selected",
            command=self._remove_selected_dropped_files,
            font=("Segoe UI", 9),
            fg=self.colors["text"],
            bg=self.colors["secondary"],
            activeforeground=self.colors["text"],
            activebackground=self.colors["secondary_active"],
            relief="flat",
            bd=0,
            padx=12,
            pady=9,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=self.colors["border"],
        )
        self.remove_pending_button.grid(row=0, column=0, sticky="ew")

        self.clear_pending_button = tk.Button(
            pending_actions,
            text="Clear All",
            command=self._clear_dropped_files,
            font=("Segoe UI", 9),
            fg=self.colors["text"],
            bg=self.colors["secondary"],
            activeforeground=self.colors["text"],
            activebackground=self.colors["secondary_active"],
            relief="flat",
            bd=0,
            padx=12,
            pady=9,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=self.colors["border"],
        )
        self.clear_pending_button.grid(row=0, column=1, sticky="ew", padx=(12, 0))

        options_frame = ttk.LabelFrame(self.left_column, text="Options", padding=20, style="Card.TLabelframe")
        options_frame.pack(fill="x", pady=(16, 0))
        options_frame.columnconfigure(0, weight=1)
        options_frame.columnconfigure(1, weight=0, minsize=92)
        options_frame.columnconfigure(2, weight=1)
        options_frame.columnconfigure(3, weight=0, minsize=92)

        self.recursive_check = ttk.Checkbutton(
            options_frame,
            text="Include subdirectories",
            variable=self.recursive_var,
            style="App.TCheckbutton",
        )
        self.recursive_check.grid(row=0, column=0, sticky="w")
        self.overwrite_check = ttk.Checkbutton(
            options_frame,
            text="Overwrite existing PDF files",
            variable=self.overwrite_var,
            style="App.TCheckbutton",
        )
        self.overwrite_check.grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.watch_checkbox = ttk.Checkbutton(
            options_frame,
            text="Watch input folder",
            variable=self.watch_enabled_var,
            style="App.TCheckbutton",
        )
        self.watch_checkbox.grid(row=2, column=0, sticky="w", pady=(10, 0))

        ttk.Label(options_frame, text="Retry attempts", style="Body.TLabel").grid(
            row=3, column=0, sticky="w", pady=(16, 6)
        )
        self.retry_spinbox = ttk.Spinbox(
            options_frame,
            from_=0,
            to=5,
            textvariable=self.retry_attempts_var,
            width=6,
        )
        self.retry_spinbox.grid(row=3, column=1, sticky="w", pady=(16, 6), padx=(18, 0))

        ttk.Label(options_frame, text="Batch workers", style="Body.TLabel").grid(
            row=3, column=2, sticky="w", pady=(16, 6), padx=(18, 0)
        )
        self.workers_spinbox = ttk.Spinbox(
            options_frame,
            from_=1,
            to=4,
            textvariable=self.batch_workers_var,
            width=6,
        )
        self.workers_spinbox.grid(row=3, column=3, sticky="w", pady=(16, 6), padx=(12, 0))

        ttk.Label(
            options_frame,
            text="Retries rebuild the Office session. Parallel workers can improve large-batch throughput.",
            style="Muted.TLabel",
        ).grid(row=4, column=0, columnspan=4, sticky="ew", pady=(6, 0))
        options_frame.grid_slaves(row=4, column=0)[0].configure(wraplength=430, justify="left")

        action_frame = ttk.Frame(self.left_column, padding=(0, 18, 0, 0), style="App.TFrame")
        action_frame.pack(fill="x")
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)

        self.start_button = tk.Button(
            action_frame,
            text="Start Conversion",
            command=self._start_conversion,
            font=("Segoe UI Semibold", 10),
            fg="#ffffff",
            bg=self.colors["primary"],
            activeforeground="#ffffff",
            activebackground=self.colors["primary_active"],
            relief="flat",
            bd=0,
            padx=16,
            pady=11,
            cursor="hand2",
        )
        self.start_button.grid(row=0, column=0, sticky="ew")

        self.clear_button = tk.Button(
            action_frame,
            text="Clear Results",
            command=self._clear_log,
            font=("Segoe UI", 10),
            fg=self.colors["text"],
            bg=self.colors["secondary"],
            activeforeground=self.colors["text"],
            activebackground=self.colors["secondary_active"],
            relief="flat",
            bd=0,
            padx=16,
            pady=11,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=self.colors["border"],
        )
        self.clear_button.grid(row=0, column=1, sticky="ew", padx=(12, 0))
        self._refresh_pending_file_list()

        progress_frame = ttk.LabelFrame(self.right_column, text="Progress", padding=20, style="Card.TLabelframe")
        progress_frame.grid(row=0, column=0, sticky="ew")

        self.progress = ttk.Progressbar(
            progress_frame,
            mode="determinate",
            maximum=100,
            variable=self.progress_var,
            style="App.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x")

        self.status_label = ttk.Label(
            progress_frame,
            textvariable=self.status_var,
            style="Muted.TLabel",
        )
        self.status_label.pack(anchor="w", pady=(10, 0))

        tk.Label(
            progress_frame,
            text="Tip: batch mode preserves the relative folder structure in the output directory.",
            font=("Segoe UI", 9),
            fg=self.colors["muted"],
            bg=self.colors["card"],
            wraplength=620,
            justify="left",
        ).pack(anchor="w", pady=(12, 0))

        preview_frame = ttk.LabelFrame(
            self.right_column,
            text="Preview",
            padding=20,
            style="Card.TLabelframe",
        )
        preview_frame.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        preview_frame.columnconfigure(0, weight=0, minsize=72)
        preview_frame.columnconfigure(1, weight=1)

        ttk.Label(
            preview_frame,
            textvariable=self.preview_title_var,
            style="Body.TLabel",
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(preview_frame, text="Source", style="Body.TLabel").grid(
            row=1, column=0, sticky="nw", pady=(12, 0)
        )
        ttk.Label(
            preview_frame,
            textvariable=self.preview_source_var,
            style="Muted.TLabel",
            wraplength=700,
            justify="left",
        ).grid(row=1, column=1, sticky="w", pady=(12, 0))
        ttk.Label(preview_frame, text="Output", style="Body.TLabel").grid(
            row=2, column=0, sticky="nw", pady=(10, 0)
        )
        ttk.Label(
            preview_frame,
            textvariable=self.preview_output_var,
            style="Muted.TLabel",
            wraplength=700,
            justify="left",
        ).grid(row=2, column=1, sticky="w", pady=(10, 0))
        ttk.Label(preview_frame, text="Details", style="Body.TLabel").grid(
            row=3, column=0, sticky="nw", pady=(10, 0)
        )
        ttk.Label(
            preview_frame,
            textvariable=self.preview_details_var,
            style="Muted.TLabel",
            wraplength=700,
            justify="left",
        ).grid(row=3, column=1, sticky="w", pady=(10, 0))

        preview_actions = ttk.Frame(preview_frame, style="Card.TFrame")
        preview_actions.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        preview_actions.columnconfigure(0, weight=1)
        preview_actions.columnconfigure(1, weight=1)
        preview_actions.columnconfigure(2, weight=1)

        self.open_source_button = ttk.Button(
            preview_actions,
            text="Open Source",
            command=self._open_selected_source,
            style="Secondary.TButton",
        )
        self.open_source_button.grid(row=0, column=0, sticky="ew")
        self.open_output_button = ttk.Button(
            preview_actions,
            text="Open Output",
            command=self._open_selected_output_file,
            style="Secondary.TButton",
        )
        self.open_output_button.grid(row=0, column=1, sticky="ew", padx=(12, 12))
        self.open_folder_button = ttk.Button(
            preview_actions,
            text="Open Output Folder",
            command=self._open_selected_output_folder,
            style="Secondary.TButton",
        )
        self.open_folder_button.grid(row=0, column=2, sticky="ew")

        log_frame = ttk.LabelFrame(self.right_column, text="Results", padding=20, style="Card.TLabelframe")
        log_frame.grid(row=2, column=0, sticky="nsew", pady=(16, 0))
        log_frame.rowconfigure(1, weight=1)
        log_frame.columnconfigure(0, weight=1)

        results_toolbar = ttk.Frame(log_frame, style="Card.TFrame")
        results_toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        results_toolbar.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            results_toolbar,
            text="Only show failures",
            variable=self.failures_only_var,
            command=self._refresh_results_table,
            style="App.TCheckbutton",
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            results_toolbar,
            textvariable=self.summary_var,
            style="Muted.TLabel",
        ).grid(row=0, column=1, sticky="e")

        columns = ("status", "source", "output", "details")
        self.results_table = ttk.Treeview(
            log_frame,
            columns=columns,
            show="headings",
            style="Results.Treeview",
        )
        self.results_table.heading(
            "status", text="Status", command=lambda: self._sort_results_by("status")
        )
        self.results_table.heading(
            "source", text="Source", command=lambda: self._sort_results_by("source")
        )
        self.results_table.heading(
            "output", text="Output", command=lambda: self._sort_results_by("output")
        )
        self.results_table.heading(
            "details", text="Details", command=lambda: self._sort_results_by("details")
        )
        self.results_table.column("status", width=130, minwidth=110, anchor="center", stretch=False)
        self.results_table.column("source", width=250, minwidth=220, anchor="w")
        self.results_table.column("output", width=250, minwidth=220, anchor="w")
        self.results_table.column("details", width=360, minwidth=280, anchor="w")
        self.results_table.grid(row=1, column=0, sticky="nsew")
        self.results_table.tag_configure("ok", background="#edf7f2", foreground="#1d5d43")
        self.results_table.tag_configure("failed", background="#fbefef", foreground="#8a2f2f")
        self.results_table.tag_configure("error", background="#fff4e8", foreground="#8c4b1f")
        self.results_table.tag_configure("info", background="#f6f8fb", foreground=self.colors["text"])
        self.results_table.bind("<<TreeviewSelect>>", self._handle_result_selection)
        self.results_table.bind("<Double-1>", self._handle_row_double_click)
        self.results_table.bind("<Button-3>", self._show_context_menu)

        table_scroll_y = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.results_table.yview
        )
        table_scroll_y.grid(row=1, column=1, sticky="ns")
        self.results_table.configure(yscrollcommand=table_scroll_y.set)

        table_scroll_x = ttk.Scrollbar(
            log_frame, orient="horizontal", command=self.results_table.xview
        )
        table_scroll_x.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.results_table.configure(xscrollcommand=table_scroll_x.set)

    def _build_context_menu(self) -> None:
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(
            label="Open Output Folder",
            command=self._open_selected_output_folder,
        )
        self.context_menu.add_separator()
        self.context_menu.add_command(
            label="Copy Source Path",
            command=lambda: self._copy_selected_path("source"),
        )
        self.context_menu.add_command(
            label="Copy Output Path",
            command=lambda: self._copy_selected_path("output"),
        )
        self.context_menu.add_command(
            label="Copy Details",
            command=lambda: self._copy_selected_path("details"),
        )

    def _register_drop_targets(self) -> None:
        if DND_FILES is None:
            return

        targets = [
            (self.root, self._handle_input_drop),
            (self.input_entry, self._handle_input_drop),
            (self.output_entry, self._handle_output_drop),
        ]
        for widget, handler in targets:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", handler)
            widget.dnd_bind("<<DropEnter>>", self._handle_drop_enter)
            widget.dnd_bind("<<DropLeave>>", self._handle_drop_leave)

    def _handle_drop_enter(self, event: object) -> str:
        self.drop_hint_label.configure(
            bg=self.drop_hint_active_bg,
            highlightbackground=self.colors["drop_active_border"],
        )
        self.status_var.set("Drop here to use the selected file or folder.")
        return "break"

    def _handle_drop_leave(self, event: object) -> str:
        self.drop_hint_label.configure(
            bg=self.drop_hint_default_bg,
            highlightbackground=self.drop_hint_default_border,
        )
        return "break"

    def _parse_drop_paths(self, data: str) -> list[Path]:
        raw_items = self.root.tk.splitlist(data)
        return [Path(item).expanduser().resolve() for item in raw_items]

    def _show_drop_error(self, message: str) -> str:
        self._handle_drop_leave(None)
        self.status_var.set(message)
        messagebox.showerror("Unsupported Drop", message)
        return "break"

    def _current_output_path(self) -> Path | None:
        output_text = self.output_var.get().strip()
        if not output_text:
            return None
        try:
            return Path(output_text).expanduser().resolve()
        except OSError:
            return None

    def _sync_output_with_input(self, force: bool = False) -> None:
        if self.mode_var.get() != "file" or self.dropped_files:
            return

        input_text = self.input_var.get().strip()
        if not input_text:
            if force:
                self.output_var.set("")
                self._auto_output_path = None
            return

        source = Path(input_text).expanduser()
        if source.suffix.lower() not in SUPPORTED_INPUT_SUFFIXES:
            return

        suggested_output = source.with_suffix(".pdf").resolve()
        current_output = self._current_output_path()
        should_replace = (
            force
            or current_output is None
            or (self._auto_output_path is not None and current_output == self._auto_output_path)
        )
        if should_replace:
            self.output_var.set(str(suggested_output))
            self._auto_output_path = suggested_output
        elif current_output is not None:
            self._auto_output_path = None

    def _clear_dropped_files(self) -> None:
        self.dropped_files = []
        self._refresh_pending_file_list()
        self.drop_hint_var.set(
            "Drag a Word/Excel/PowerPoint file or folder onto the window. "
            "You can also drop multiple supported files for a one-off batch."
            if TkinterDnD is not None
            else "Install tkinterdnd2 to enable drag and drop."
        )
        if self.mode_var.get() == "directory" and not self.input_var.get().strip():
            self.mode_var.set("file")
            self._toggle_mode()
        self._persist_config()

    def _refresh_pending_file_list(self) -> None:
        self.pending_listbox.delete(0, "end")
        for path in self.dropped_files:
            self.pending_listbox.insert("end", path.name)

        if self.dropped_files:
            if not self.pending_frame.winfo_manager():
                self.pending_frame.pack(fill="x", pady=(16, 0))
            self.input_var.set(f"{len(self.dropped_files)} dropped files")
            self.drop_hint_var.set(
                f"Using {len(self.dropped_files)} dropped files for a one-off batch conversion."
            )
        else:
            if self.pending_frame.winfo_manager():
                self.pending_frame.pack_forget()
            if self.input_var.get().endswith("dropped files"):
                self.input_var.set("")

    def _remove_selected_dropped_files(self) -> None:
        selected = list(self.pending_listbox.curselection())
        if not selected:
            self.status_var.set("Select one or more dropped files to remove.")
            return

        for index in reversed(selected):
            del self.dropped_files[index]

        if len(self.dropped_files) == 1:
            remaining = self.dropped_files[0]
            self.dropped_files = []
            self.mode_var.set("file")
            self.input_var.set(str(remaining))
            self._sync_output_with_input(force=True)
            self.drop_hint_var.set(
                "Single-file mode is ready. Drag another Office file here to replace it."
            )
            self.status_var.set(f"Kept one file and switched to single-file mode: {remaining.name}")
            self._toggle_mode()
            self._refresh_pending_file_list()
            return

        if not self.dropped_files:
            self._clear_dropped_files()
            self.status_var.set("Cleared the dropped-file batch list.")
            return

        self._refresh_pending_file_list()
        self.status_var.set(
            f"Removed selected files. {len(self.dropped_files)} dropped files remain."
        )
        self._persist_config()

    def _set_batch_file_drop(self, files: list[Path]) -> None:
        self.dropped_files = files
        self.mode_var.set("directory")
        self._auto_output_path = None
        self._toggle_mode()
        first_parent = files[0].parent
        if not self.output_var.get().strip():
            self.output_var.set(str(first_parent))
        self.drop_hint_var.set(
            f"Using {len(files)} dropped files for a one-off batch conversion."
        )
        self._refresh_pending_file_list()
        self.status_var.set(
            f"Prepared a dropped-file batch with {len(files)} files."
        )
        self._handle_drop_leave(None)
        self._persist_config()

    def _handle_input_drop(self, event: object) -> str:
        data = getattr(event, "data", "")
        paths = self._parse_drop_paths(data)
        if not paths:
            self.status_var.set("No valid item was dropped.")
            return "break"

        if len(paths) == 1 and paths[0].is_dir():
            path = paths[0]
            self.dropped_files = []
            self._refresh_pending_file_list()
            self.mode_var.set("directory")
            self.input_var.set(str(path))
            self._auto_output_path = None
            self._toggle_mode()
            self.drop_hint_var.set(
                "Directory batch mode is ready. Drag a folder here anytime to replace it."
            )
            self.status_var.set(f"Input directory selected by drag and drop: {path}")
            self._handle_drop_leave(event)
            self._update_preview_panel(None)
            self._persist_config()
            return "break"

        supported_files = [path for path in paths if is_supported_input_file(path)]
        invalid_paths = [path for path in paths if not is_supported_input_file(path) and not path.is_dir()]
        dropped_directories = [path for path in paths if path.is_dir()]

        if invalid_paths:
            names = ", ".join(path.name for path in invalid_paths[:3])
            suffix = " ..." if len(invalid_paths) > 3 else ""
            return self._show_drop_error(
                f"Only supported files ({supported_input_label()}) or folders can be dropped as input. Unsupported item(s): {names}{suffix}"
            )

        if dropped_directories and len(paths) > 1:
            return self._show_drop_error(
                "Please drop either one folder or one/multiple supported files, not a mixed selection."
            )

        if len(supported_files) > 1:
            self._set_batch_file_drop(supported_files)
            return "break"

        if len(supported_files) == 1:
            path = supported_files[0]
            self.dropped_files = []
            self._refresh_pending_file_list()
            self.mode_var.set("file")
            self.input_var.set(str(path))
            self._toggle_mode()
            self._sync_output_with_input(force=True)
            self.drop_hint_var.set(
                "Single-file mode is ready. Drag another Office file here to replace it."
            )
            self.status_var.set(f"Input file selected by drag and drop: {path.name}")
            self._handle_drop_leave(event)
            self._update_preview_panel(None)
            self._persist_config()
            return "break"

        return self._show_drop_error(
            f"Only supported files ({supported_input_label()}) or directories can be dropped as input."
        )

    def _handle_output_drop(self, event: object) -> str:
        data = getattr(event, "data", "")
        paths = self._parse_drop_paths(data)
        if not paths:
            self.status_var.set("No valid item was dropped.")
            return "break"

        path = paths[0]
        if path.is_dir():
            self.output_var.set(str(path))
            self.status_var.set(f"Output directory selected by drag and drop: {path}")
            self._handle_drop_leave(event)
            self._update_preview_panel(None)
            self._persist_config()
            return "break"

        if path.suffix.lower() == ".pdf":
            self.output_var.set(str(path))
            self.status_var.set(f"Output PDF selected by drag and drop: {path.name}")
            self._handle_drop_leave(event)
            self._update_preview_panel(None)
            self._persist_config()
            return "break"

        return self._show_drop_error(
            "Drop a PDF file or a directory into the output field."
        )

    def _toggle_mode(self) -> None:
        directory_mode = self.mode_var.get() == "directory"
        self.output_label.configure(
            text="Output directory" if directory_mode else "Output PDF"
        )
        if directory_mode:
            self._auto_output_path = None
            self.recursive_check.state(["!disabled"])
            self.watch_checkbox.state(["!disabled"])
        else:
            self.recursive_var.set(False)
            self.recursive_check.state(["disabled"])
            self.watch_enabled_var.set(False)
            self.watch_checkbox.state(["disabled"])
            self.watch_snapshot = {}
            self._sync_output_with_input(force=False)
        self._update_preview_panel(None)

    def _browse_input(self) -> None:
        if self.mode_var.get() == "directory":
            selected = filedialog.askdirectory(title="Select input directory")
        else:
            selected = filedialog.askopenfilename(
                title="Select input file",
                filetypes=[
                    ("Supported documents", supported_input_patterns()),
                    ("Word documents", "*.doc *.docx"),
                    ("Excel documents", "*.xls *.xlsx *.xlsm *.xlsb"),
                    ("PowerPoint documents", "*.ppt *.pptx *.pptm"),
                    ("All files", "*.*"),
                ],
            )
        if selected:
            self._clear_dropped_files()
            self.input_var.set(selected)
            self._sync_output_with_input(force=True)
            self._update_preview_panel(None)
            self._persist_config()

    def _browse_output(self) -> None:
        if self.mode_var.get() == "directory":
            selected = filedialog.askdirectory(title="Select output directory")
        else:
            selected = filedialog.asksaveasfilename(
                title="Save PDF as",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
            )
        if selected:
            self.output_var.set(selected)
            self._auto_output_path = None
            self._update_preview_panel(None)
            self._persist_config()

    def _append_result(
        self,
        status: str,
        source: str,
        output: str,
        details: str,
        *,
        prefer_selection: bool = True,
    ) -> None:
        log_level = logging.WARNING if status.upper() in {"FAILED", "ERROR"} else logging.INFO
        LOGGER.log(log_level, "%s | %s -> %s | %s", status, source, output, details)
        if prefer_selection and source != "-" and output != "-":
            self._pending_selection_key = (source, output, details)
        self.result_rows.append(
            {
                "status": status,
                "source": source,
                "output": output,
                "details": details,
            }
        )
        self._schedule_results_refresh(
            scroll_to_end=True,
            prefer_last_selection=prefer_selection,
        )

    def _clear_log(self) -> None:
        self.result_rows.clear()
        self._pending_selection_key = None
        self._refresh_results_table()
        self._update_preview_panel(None)

    def _status_tag(self, status: str) -> str:
        return {
            "OK": "ok",
            "FAILED": "failed",
            "ERROR": "error",
            "INFO": "info",
        }.get(status.upper(), "info")

    def _sort_value(self, row: dict[str, str], column: str) -> tuple[int, str]:
        status_rank = {
            "ERROR": 0,
            "FAILED": 1,
            "INFO": 2,
            "OK": 3,
        }
        if column == "status":
            return (status_rank.get(row["status"].upper(), 99), row["status"].lower())
        return (0, row[column].lower())

    def _filtered_rows(self) -> list[dict[str, str]]:
        if not self.failures_only_var.get():
            return list(self.result_rows)
        return [
            row
            for row in self.result_rows
            if row["status"].upper() in {"FAILED", "ERROR"}
        ]

    def _update_summary_text(self, shown_rows: list[dict[str, str]]) -> None:
        total = len(self.result_rows)
        shown = len(shown_rows)
        failed = sum(
            1 for row in self.result_rows if row["status"].upper() in {"FAILED", "ERROR"}
        )
        self.summary_var.set(
            f"Showing {shown} of {total} items  |  Failures: {failed}"
        )

    def _schedule_results_refresh(
        self,
        *,
        scroll_to_end: bool = False,
        prefer_last_selection: bool = True,
    ) -> None:
        self._results_should_scroll_end = self._results_should_scroll_end or scroll_to_end
        self._results_prefer_last_selection = (
            self._results_prefer_last_selection and prefer_last_selection
        )
        if self._results_refresh_job is None:
            self._results_refresh_job = self.root.after_idle(self._flush_results_refresh)

    def _flush_results_refresh(self) -> None:
        self._results_refresh_job = None
        scroll_to_end = self._results_should_scroll_end
        prefer_last_selection = self._results_prefer_last_selection
        self._results_should_scroll_end = False
        self._results_prefer_last_selection = True
        self._refresh_results_table(
            scroll_to_end=scroll_to_end,
            prefer_last_selection=prefer_last_selection,
        )

    def _refresh_results_table(
        self,
        scroll_to_end: bool = False,
        prefer_last_selection: bool = True,
    ) -> None:
        current_values = self._selected_row_values()
        selection_key = self._pending_selection_key
        if selection_key is None and current_values is not None:
            selection_key = (current_values[1], current_values[2], current_values[3])
        for item in self.results_table.get_children():
            self.results_table.delete(item)

        rows = self._filtered_rows()
        rows.sort(
            key=lambda row: self._sort_value(row, self.sort_column),
            reverse=self.sort_descending,
        )

        last_item_id = None
        for row in rows:
            item_id = self.results_table.insert(
                "",
                "end",
                values=(
                    self._status_symbol(row["status"]),
                    row["source"],
                    row["output"],
                    row["details"],
                ),
                tags=(self._status_tag(row["status"]),),
            )
            last_item_id = item_id

            if selection_key is not None and (
                row["source"],
                row["output"],
                row["details"],
            ) == selection_key:
                self.results_table.selection_set(item_id)
                self.results_table.focus(item_id)

        if scroll_to_end and last_item_id is not None:
            self.results_table.see(last_item_id)
            if prefer_last_selection and not self.results_table.selection():
                self.results_table.selection_set(last_item_id)
                self.results_table.focus(last_item_id)

        self._refresh_heading_labels()
        self._update_summary_text(rows)
        self._update_preview_panel(self._selected_row_values())
        self._pending_selection_key = None

    def _sort_results_by(self, column: str) -> None:
        if self.sort_column == column:
            self.sort_descending = not self.sort_descending
        else:
            self.sort_column = column
            self.sort_descending = False
        self._refresh_results_table()

    def _selected_item_id(self) -> str | None:
        selection = self.results_table.selection()
        if selection:
            return selection[0]
        focused = self.results_table.focus()
        return focused or None

    def _copy_text(self, text: str, success_message: str) -> None:
        if not text or text == "-":
            self.status_var.set("No path is available for the selected row.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update_idletasks()
        self.status_var.set(success_message)

    def _copy_selected_path(self, field: str) -> None:
        values = self._selected_row_values()
        if values is None:
            self.status_var.set("Select a result row first.")
            return

        status, source, output, details = values
        mapping = {
            "source": (source, "Source path copied."),
            "output": (output, "Output path copied."),
            "details": (details, "Details copied."),
        }
        text, success_message = mapping[field]
        self._copy_text(text, success_message)

    def _open_selected_output_folder(self) -> None:
        values = self._effective_preview_values()
        if values is None:
            self.status_var.set("Select a result row or choose an output path first.")
            return

        _, _, output, _ = values
        if not output or output == "-":
            self.status_var.set("No output path is available right now.")
            return

        output_path = Path(output)
        folder = output_path.parent if output_path.suffix else output_path
        if not folder.exists():
            self.status_var.set("Output folder does not exist yet.")
            return

        try:
            os.startfile(str(folder))
            self.status_var.set(f"Opened output folder: {folder}")
        except OSError as exc:
            self.status_var.set(f"Unable to open output folder: {exc}")

    def _handle_row_double_click(self, event: tk.Event[tk.Misc]) -> None:
        item_id = self.results_table.identify_row(event.y)
        if item_id:
            self.results_table.selection_set(item_id)
            self.results_table.focus(item_id)
            self._open_selected_output_folder()

    def _show_context_menu(self, event: tk.Event[tk.Misc]) -> None:
        item_id = self.results_table.identify_row(event.y)
        if item_id:
            self.results_table.selection_set(item_id)
            self.results_table.focus(item_id)

        has_selection = self._selected_item_id() is not None
        state = "normal" if has_selection else "disabled"
        for index in (0, 2, 3, 4):
            self.context_menu.entryconfigure(index, state=state)
        self.context_menu.tk_popup(event.x_root, event.y_root)
        self.context_menu.grab_release()

    def _set_running_state(self, running: bool) -> None:
        if running:
            self.start_button.configure(
                state="disabled",
                bg="#8aa0ba",
                disabledforeground="#eef3f8",
                cursor="arrow",
            )
            self.status_var.set("Converting... please keep Word available in the background.")
        else:
            self.start_button.configure(
                state="normal",
                bg=self.colors["primary"],
                fg="#ffffff",
                activebackground=self.colors["primary_active"],
                activeforeground="#ffffff",
                cursor="hand2",
            )

    def _start_conversion(self) -> None:
        if self.worker_thread is not None and self.worker_thread.is_alive():
            messagebox.showinfo("In Progress", "A conversion task is already running.")
            return

        input_text = self.input_var.get().strip()
        output_text = self.output_var.get().strip()
        mode = self.mode_var.get()

        if not input_text and not self.dropped_files:
            messagebox.showerror("Missing Input", "Please select an input file or directory.")
            return

        self.retry_attempts_var.set(max(0, self.retry_attempts_var.get()))
        self.batch_workers_var.set(max(1, self.batch_workers_var.get()))
        self._persist_config()
        self.progress_var.set(0)
        self.status_var.set("Starting conversion...")
        self._append_result(
            "INFO",
            "-",
            "-",
            "Starting conversion task.",
            prefer_selection=False,
        )
        self._set_running_state(True)

        self.worker_thread = threading.Thread(
            target=self._worker_run,
            args=(mode, input_text, output_text),
            daemon=True,
        )
        self.worker_thread.start()

    def _worker_run(self, mode: str, input_text: str, output_text: str) -> None:
        try:
            if mode == "file":
                source, target = resolve_file_paths(
                    input_text, output_text or None
                )
                outcome = convert_word_to_pdf(
                    source, target, overwrite=self.overwrite_var.get()
                )
                self.event_queue.put(("file_success", (source, outcome)))
                return

            if self.dropped_files:
                output_dir = Path(output_text).expanduser().resolve() if output_text else None

                def on_progress(index: int, total: int, result: BatchResult) -> None:
                    self.event_queue.put(("batch_progress", (index, total, result)))

                batch_outcome = convert_file_list_to_pdf(
                    self.dropped_files,
                    output_dir,
                    overwrite=self.overwrite_var.get(),
                    retry_attempts=max(0, self.retry_attempts_var.get()),
                    max_workers=max(1, self.batch_workers_var.get()),
                    progress_callback=on_progress,
                )
                self.event_queue.put(("batch_complete", batch_outcome))
                return

            input_dir = resolve_input_path(input_text)
            if not input_dir.is_dir():
                raise NotADirectoryError(f"Input path is not a directory: {input_dir}")

            output_dir = Path(output_text).expanduser().resolve() if output_text else None

            def on_progress(index: int, total: int, result: BatchResult) -> None:
                self.event_queue.put(("batch_progress", (index, total, result)))

            batch_outcome = convert_directory_to_pdf(
                input_dir,
                output_dir,
                recursive=self.recursive_var.get(),
                overwrite=self.overwrite_var.get(),
                retry_attempts=max(0, self.retry_attempts_var.get()),
                max_workers=max(1, self.batch_workers_var.get()),
                progress_callback=on_progress,
            )
            self.event_queue.put(("batch_complete", batch_outcome))
        except Exception as exc:
            self.event_queue.put(("error", str(exc)))

    def _poll_events(self) -> None:
        try:
            while True:
                event_type, payload = self.event_queue.get_nowait()
                if event_type == "file_success":
                    source, outcome = payload
                    self.progress_var.set(100)
                    self.status_var.set("Single-file conversion completed.")
                    self._append_result(
                        "OK",
                        str(source),
                        str(outcome.target),
                        outcome.warning or "Single-file conversion completed.",
                    )
                    self._set_running_state(False)
                elif event_type == "batch_progress":
                    index, total, result = payload
                    percentage = (index / total) * 100
                    status = "OK" if result.success else "FAILED"
                    self.status_var.set(f"Processing {index}/{total}...")
                    self.progress_var.set(percentage)
                    self._append_result(
                        status,
                        str(result.source),
                        str(result.target),
                        result.message,
                    )
                elif event_type == "batch_complete":
                    batch_outcome = payload
                    self.progress_var.set(100)
                    summary = format_batch_summary(batch_outcome.results)
                    self.status_var.set(summary)
                    self._append_result(
                        "INFO",
                        "-",
                        "-",
                        summary,
                        prefer_selection=False,
                    )
                    if batch_outcome.parallel_workers_used > 1:
                        session_note = (
                            f"This batch used {batch_outcome.parallel_workers_used} parallel Office workers "
                            f"and restarted sessions {batch_outcome.session_restarts} time(s)."
                        )
                    elif batch_outcome.reused_single_session:
                        session_note = "This batch reused a single Office session."
                    else:
                        session_note = (
                            "This batch automatically restarted the Office session "
                            f"{batch_outcome.session_restarts} time(s) after failures."
                        )
                    self._append_result(
                        "INFO",
                        "-",
                        "-",
                        session_note,
                        prefer_selection=False,
                    )
                    self._append_result(
                        "INFO",
                        "-",
                        "-",
                        f"Batch duration: {_format_duration(batch_outcome.duration_seconds)}",
                        prefer_selection=False,
                    )
                    if batch_outcome.warning:
                        self._append_result(
                            "INFO",
                            "-",
                            "-",
                            batch_outcome.warning,
                            prefer_selection=False,
                        )
                    self._set_running_state(False)
                elif event_type == "error":
                    self.status_var.set("Conversion failed.")
                    self._append_result(
                        "ERROR",
                        "-",
                        "-",
                        str(payload),
                        prefer_selection=False,
                    )
                    messagebox.showerror("Conversion Error", str(payload))
                    self._set_running_state(False)
        except queue.Empty:
            pass
        finally:
            self.root.after(80, self._poll_events)

    def _status_symbol(self, status: str) -> str:
        symbol_map = {
            "OK": "OK Success",
            "FAILED": "X Failed",
            "ERROR": "! Error",
            "INFO": "i Info",
        }
        return symbol_map.get(status.upper(), status)

    def _refresh_heading_labels(self) -> None:
        labels = {
            "status": "Status",
            "source": "Source",
            "output": "Output",
            "details": "Details",
        }
        arrow = " v" if self.sort_descending else " ^"
        for column, label in labels.items():
            heading = label + arrow if column == self.sort_column else label
            self.results_table.heading(
                column,
                text=heading,
                command=lambda col=column: self._sort_results_by(col),
            )

    def _selected_row_values(self) -> tuple[str, str, str, str] | None:
        item_id = self._selected_item_id()
        if not item_id:
            return None
        values = self.results_table.item(item_id, "values")
        if len(values) != 4:
            return None
        status, source, output, details = (str(value) for value in values)
        for prefix in ("OK ", "X ", "! ", "i "):
            if status.startswith(prefix):
                status = status[len(prefix) :]
                break
        return status, source, output, details

    def _result_value_to_path(self, value: str) -> Path | None:
        if not value or value == "-":
            return None
        return Path(value)

    def _preview_fallback_values(self) -> tuple[str, str, str, str] | None:
        input_text = self.input_var.get().strip()
        output_text = self.output_var.get().strip()
        if not input_text and not output_text:
            return None
        details = (
            "Current input/output selection. Run a conversion or select a result row for more details."
        )
        return ("Current", input_text or "-", output_text or "-", details)

    def _set_preview_button_states(
        self,
        source_path: Path | None,
        output_path: Path | None,
    ) -> None:
        self.open_source_button.configure(
            state="normal" if source_path is not None else "disabled"
        )
        self.open_output_button.configure(
            state="normal" if output_path is not None else "disabled"
        )
        self.open_folder_button.configure(
            state="normal" if output_path is not None else "disabled"
        )

    def _update_preview_panel(self, values: tuple[str, str, str, str] | None) -> None:
        if values is None:
            values = self._preview_fallback_values()
            if values is None:
                self.preview_title_var.set(
                    "Select a result row to inspect source and output details."
                )
                self.preview_source_var.set("-")
                self.preview_output_var.set("-")
                self.preview_details_var.set("-")
                self._set_preview_button_states(None, None)
                return

        status, source, output, details = values
        self.preview_title_var.set(f"{status} item preview")
        self.preview_source_var.set(source)
        self.preview_output_var.set(output)
        self.preview_details_var.set(details)
        source_path = self._result_value_to_path(source)
        output_path = self._result_value_to_path(output)
        self._set_preview_button_states(source_path, output_path)

    def _handle_result_selection(self, event: tk.Event[tk.Misc] | None = None) -> None:
        self._update_preview_panel(self._selected_row_values())

    def _effective_preview_values(self) -> tuple[str, str, str, str] | None:
        return self._selected_row_values() or self._preview_fallback_values()

    def _open_path(self, path: Path, label: str) -> None:
        if not path.exists():
            self.status_var.set(f"{label} does not exist.")
            return
        try:
            os.startfile(str(path))
            self.status_var.set(f"Opened {label.lower()}: {path}")
        except OSError as exc:
            self.status_var.set(f"Unable to open {label.lower()}: {exc}")

    def _open_selected_source(self) -> None:
        values = self._effective_preview_values()
        if values is None:
            self.status_var.set("Select a result row or choose an input file first.")
            return
        _, source, _, _ = values
        source_path = self._result_value_to_path(source)
        if source_path is None:
            self.status_var.set("No source path is available right now.")
            return
        self._open_path(source_path, "Source file")

    def _open_selected_output_file(self) -> None:
        values = self._effective_preview_values()
        if values is None:
            self.status_var.set("Select a result row or choose an output path first.")
            return
        _, _, output, _ = values
        output_path = self._result_value_to_path(output)
        if output_path is None:
            self.status_var.set("No output file is available right now.")
            return
        self._open_path(output_path, "Output file")

    def _current_config(self) -> AppConfig:
        return AppConfig(
            mode=self.mode_var.get(),
            last_input=self.input_var.get().strip(),
            last_output=self.output_var.get().strip(),
            recursive=bool(self.recursive_var.get()),
            overwrite=bool(self.overwrite_var.get()),
            failures_only=bool(self.failures_only_var.get()),
            retry_attempts=max(0, int(self.retry_attempts_var.get())),
            batch_workers=max(1, int(self.batch_workers_var.get())),
            watch_enabled=bool(self.watch_enabled_var.get()),
            watch_interval_seconds=max(1.0, self.config.watch_interval_seconds),
        )

    def _persist_config(self) -> None:
        self.config = self._current_config()
        save_app_config(self.config)

    def _handle_close(self) -> None:
        self._persist_config()
        self.root.destroy()

    def _watch_input_directory(self) -> Path | None:
        if self.mode_var.get() != "directory" or self.dropped_files:
            return None
        input_text = self.input_var.get().strip()
        if not input_text:
            return None
        path = Path(input_text).expanduser()
        if not path.exists() or not path.is_dir():
            return None
        return path.resolve()

    def _snapshot_supported_files(
        self, directory: Path
    ) -> dict[Path, tuple[int, int]]:
        snapshot: dict[Path, tuple[int, int]] = {}
        for path in collect_supported_files(
            directory,
            recursive=bool(self.recursive_var.get()),
        ):
            try:
                stat = path.stat()
            except OSError:
                continue
            snapshot[path] = (stat.st_mtime_ns, stat.st_size)
        return snapshot

    def _poll_folder_watch(self) -> None:
        try:
            watch_dir = self._watch_input_directory()
            if not self.watch_enabled_var.get() or watch_dir is None:
                self.watch_snapshot = {}
                return

            snapshot = self._snapshot_supported_files(watch_dir)
            if not self.watch_snapshot:
                self.watch_snapshot = snapshot
                return

            added = [path for path in snapshot if path not in self.watch_snapshot]
            removed = [path for path in self.watch_snapshot if path not in snapshot]
            changed = [
                path
                for path, info in snapshot.items()
                if path in self.watch_snapshot and self.watch_snapshot[path] != info
            ]
            if added or removed or changed:
                details = []
                if added:
                    details.append(f"added {len(added)}")
                if changed:
                    details.append(f"updated {len(changed)}")
                if removed:
                    details.append(f"removed {len(removed)}")
                summary = (
                    f"Watch detected changes in {watch_dir.name}: {', '.join(details)}."
                )
                self.status_var.set(summary)
                self._append_result("INFO", str(watch_dir), "-", summary)
                LOGGER.info(summary)
            self.watch_snapshot = snapshot
        except Exception as exc:
            LOGGER.warning("Folder watch failed: %s", exc)
        finally:
            self.root.after(
                int(max(1.0, self.config.watch_interval_seconds) * 1000),
                self._poll_folder_watch,
            )

    def run(self) -> None:
        self.root.mainloop()


def launch_gui() -> int:
    app = WordToPdfApp()
    app.run()
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.gui or not args.input_path:
            return launch_gui()
        return run_cli(args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
