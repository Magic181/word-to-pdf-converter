from __future__ import annotations

import argparse
import gc
import os
import queue
import sys
import threading
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pythoncom  # type: ignore[import]
    import win32com.client  # type: ignore[import]
    from pywintypes import com_error  # type: ignore[import]
except ImportError:  # pragma: no cover - handled at runtime for missing dependency
    pythoncom = None
    win32com = None
    com_error = Exception


WORD_SUFFIXES = {".doc", ".docx"}
WD_EXPORT_FORMAT_PDF = 17
WD_DO_NOT_SAVE_CHANGES = 0


@dataclass(slots=True)
class BatchResult:
    source: Path
    target: Path
    success: bool
    message: str


def is_word_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in WORD_SUFFIXES


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert Word documents (.doc/.docx) to PDF on Windows."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        help="Path to a Word file or a directory containing Word files.",
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
    return parser


def ensure_dependencies() -> None:
    if pythoncom is None or win32com is None:
        raise RuntimeError(
            "Missing dependency: pywin32. Install it with `pip install -r requirements.txt`."
        )


def resolve_input_path(input_path: str) -> Path:
    source = Path(input_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"Input path does not exist: {source}")
    return source


def resolve_file_paths(input_path: str, output_path: str | None) -> tuple[Path, Path]:
    source = resolve_input_path(input_path)
    if not is_word_file(source):
        raise ValueError("Input file must be a .doc or .docx document.")

    if output_path:
        target = Path(output_path).expanduser().resolve()
    else:
        target = source.with_suffix(".pdf")

    if target.suffix.lower() != ".pdf":
        raise ValueError("Output file must use the .pdf extension.")

    return source, target


def collect_word_files(directory: Path, recursive: bool = False) -> list[Path]:
    iterator: Iterable[Path]
    if recursive:
        iterator = directory.rglob("*")
    else:
        iterator = directory.iterdir()

    files = [path for path in iterator if is_word_file(path)]
    return sorted(files, key=lambda path: str(path).lower())


def build_batch_target(source: Path, input_root: Path, output_root: Path | None) -> Path:
    if output_root is None:
        return source.with_suffix(".pdf")
    relative_path = source.relative_to(input_root)
    return (output_root / relative_path).with_suffix(".pdf")


def convert_word_to_pdf(source: Path, target: Path, overwrite: bool = False) -> Path:
    ensure_dependencies()

    if target.exists() and not overwrite:
        raise FileExistsError(
            f"Target file already exists: {target}. Use --overwrite to replace it."
        )

    target.parent.mkdir(parents=True, exist_ok=True)
    word = None
    document = None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        document = word.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        document.ExportAsFixedFormat(
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
    except com_error as exc:
        raise RuntimeError(
            "Word export failed. Make sure Microsoft Word is installed and the document can be opened normally."
        ) from exc
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            finally:
                document = None
        if word is not None:
            try:
                word.Quit()
            finally:
                word = None
        gc.collect()
        pythoncom.CoUninitialize()

    if not target.exists():
        raise RuntimeError(f"Export did not create the PDF file: {target}")

    return target


def convert_directory_to_pdf(
    input_dir: Path,
    output_dir: Path | None = None,
    *,
    recursive: bool = False,
    overwrite: bool = False,
    progress_callback: Callable[[int, int, BatchResult], None] | None = None,
) -> list[BatchResult]:
    if not input_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {input_dir}")

    files = collect_word_files(input_dir, recursive=recursive)
    if not files:
        raise FileNotFoundError(f"No Word files found in directory: {input_dir}")

    results: list[BatchResult] = []
    total = len(files)
    for index, source in enumerate(files, start=1):
        target = build_batch_target(source, input_dir, output_dir)
        try:
            convert_word_to_pdf(source, target, overwrite=overwrite)
            result = BatchResult(
                source=source,
                target=target,
                success=True,
                message="Converted successfully.",
            )
        except Exception as exc:
            result = BatchResult(
                source=source,
                target=target,
                success=False,
                message=str(exc),
            )

        results.append(result)
        if progress_callback is not None:
            progress_callback(index, total, result)

    return results


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
        result = convert_word_to_pdf(file_source, target, overwrite=args.overwrite)
        print(f"Converted successfully: {result}")
        return 0

    if source.is_dir():
        output_dir = (
            Path(args.output_path).expanduser().resolve() if args.output_path else None
        )
        results = convert_directory_to_pdf(
            source,
            output_dir,
            recursive=args.recursive,
            overwrite=args.overwrite,
        )
        for result in results:
            status = "OK" if result.success else "FAILED"
            print(f"[{status}] {result.source} -> {result.target}")
            if not result.success:
                print(f"        {result.message}")

        print(format_batch_summary(results))
        return 0 if all(result.success for result in results) else 1

    raise ValueError(f"Unsupported input path: {source}")


class WordToPdfApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Word to PDF Converter")
        self.root.geometry("1140x760")
        self.root.minsize(1000, 700)
        self.root.configure(bg="#eef3f1")

        self.mode_var = tk.StringVar(value="file")
        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=False)
        self.overwrite_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(
            value="Ready. Select a Word file or a directory to begin."
        )
        self.progress_var = tk.DoubleVar(value=0.0)
        self.worker_thread: threading.Thread | None = None
        self.event_queue: queue.Queue[tuple[str, object]] = queue.Queue()

        self._apply_styles()
        self._build_ui()
        self._build_context_menu()
        self._toggle_mode()
        self.root.after(150, self._poll_events)

    def _apply_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
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
            borderwidth=0,
            focusthickness=0,
        )
        style.map(
            "Primary.TButton",
            background=[
                ("disabled", "#94b9ae"),
                ("active", self.colors["primary_active"]),
            ],
            foreground=[("disabled", "#f3f7f5")],
        )
        style.configure(
            "Secondary.TButton",
            padding=(14, 10),
            foreground=self.colors["text"],
            background=self.colors["secondary"],
            borderwidth=0,
            focusthickness=0,
        )
        style.map(
            "Secondary.TButton",
            background=[("active", self.colors["secondary_active"])],
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
            rowheight=28,
            relief="flat",
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
            text="Word to PDF Converter",
            font=("Segoe UI Semibold", 22),
            fg=self.colors["text"],
            bg=self.colors["hero"],
        ).pack(anchor="w")
        tk.Label(
            hero_top,
            text="A desktop utility for reliable Word-to-PDF conversion in office workflows.",
            font=("Segoe UI", 10),
            fg=self.colors["muted"],
            bg=self.colors["hero"],
        ).pack(anchor="w", pady=(6, 0))

        meta_row = tk.Frame(hero, bg=self.colors["hero"])
        meta_row.pack(anchor="w", pady=(14, 0))
        tk.Label(
            meta_row,
            text="Windows + Microsoft Word required",
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

        body = ttk.Frame(container, style="App.TFrame")
        body.pack(fill="both", expand=True, pady=(22, 0))
        body.columnconfigure(0, weight=11)
        body.columnconfigure(1, weight=8)
        body.rowconfigure(1, weight=1)

        left_column = ttk.Frame(body, style="App.TFrame")
        left_column.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 18))
        left_column.columnconfigure(0, weight=1)

        right_column = ttk.Frame(body, style="App.TFrame")
        right_column.grid(row=0, column=1, rowspan=2, sticky="nsew")
        right_column.columnconfigure(0, weight=1)
        right_column.rowconfigure(1, weight=1)

        mode_frame = ttk.LabelFrame(left_column, text="Mode", padding=20, style="Card.TLabelframe")
        mode_frame.pack(fill="x")

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
            text="Use batch mode when you want to convert every Word file inside a folder.",
            style="Muted.TLabel",
        )
        mode_hint.grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))

        path_frame = ttk.LabelFrame(left_column, text="Paths", padding=20, style="Card.TLabelframe")
        path_frame.pack(fill="x", pady=(16, 0))
        path_frame.columnconfigure(1, weight=1)

        ttk.Label(path_frame, text="Input", style="Body.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        self.input_entry = ttk.Entry(path_frame, textvariable=self.input_var)
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
        self.output_entry = ttk.Entry(path_frame, textvariable=self.output_var)
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
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(2, 0))

        options_frame = ttk.LabelFrame(left_column, text="Options", padding=20, style="Card.TLabelframe")
        options_frame.pack(fill="x", pady=(16, 0))

        self.recursive_check = ttk.Checkbutton(
            options_frame,
            text="Include subdirectories",
            variable=self.recursive_var,
            style="App.TCheckbutton",
        )
        self.recursive_check.grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(
            options_frame,
            text="Overwrite existing PDF files",
            variable=self.overwrite_var,
            style="App.TCheckbutton",
        ).grid(row=0, column=1, sticky="w", padx=(22, 0))

        action_frame = ttk.Frame(left_column, padding=(0, 18, 0, 0), style="App.TFrame")
        action_frame.pack(fill="x")

        self.start_button = ttk.Button(
            action_frame,
            text="Start Conversion",
            command=self._start_conversion,
            style="Primary.TButton",
        )
        self.start_button.pack(side="left")

        ttk.Button(
            action_frame,
            text="Clear Log",
            command=self._clear_log,
            style="Secondary.TButton",
        ).pack(
            side="left", padx=(10, 0)
        )

        progress_frame = ttk.LabelFrame(right_column, text="Progress", padding=20, style="Card.TLabelframe")
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
            wraplength=360,
            justify="left",
        ).pack(anchor="w", pady=(12, 0))

        log_frame = ttk.LabelFrame(right_column, text="Results", padding=20, style="Card.TLabelframe")
        log_frame.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        columns = ("status", "source", "output", "details")
        self.results_table = ttk.Treeview(
            log_frame,
            columns=columns,
            show="headings",
            style="Results.Treeview",
        )
        self.results_table.heading("status", text="Status")
        self.results_table.heading("source", text="Source")
        self.results_table.heading("output", text="Output")
        self.results_table.heading("details", text="Details")
        self.results_table.column("status", width=90, minwidth=80, anchor="center")
        self.results_table.column("source", width=210, minwidth=160, anchor="w")
        self.results_table.column("output", width=210, minwidth=160, anchor="w")
        self.results_table.column("details", width=260, minwidth=220, anchor="w")
        self.results_table.grid(row=0, column=0, sticky="nsew")
        self.results_table.tag_configure("ok", background="#edf7f2", foreground="#1d5d43")
        self.results_table.tag_configure("failed", background="#fbefef", foreground="#8a2f2f")
        self.results_table.tag_configure("error", background="#fff4e8", foreground="#8c4b1f")
        self.results_table.tag_configure("info", background="#f6f8fb", foreground=self.colors["text"])
        self.results_table.bind("<Double-1>", self._handle_row_double_click)
        self.results_table.bind("<Button-3>", self._show_context_menu)

        table_scroll_y = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.results_table.yview
        )
        table_scroll_y.grid(row=0, column=1, sticky="ns")
        self.results_table.configure(yscrollcommand=table_scroll_y.set)

        table_scroll_x = ttk.Scrollbar(
            log_frame, orient="horizontal", command=self.results_table.xview
        )
        table_scroll_x.grid(row=1, column=0, sticky="ew", pady=(10, 0))
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

    def _toggle_mode(self) -> None:
        directory_mode = self.mode_var.get() == "directory"
        self.output_label.configure(
            text="Output directory" if directory_mode else "Output PDF"
        )
        if directory_mode:
            self.recursive_check.state(["!disabled"])
        else:
            self.recursive_var.set(False)
            self.recursive_check.state(["disabled"])

    def _browse_input(self) -> None:
        if self.mode_var.get() == "directory":
            selected = filedialog.askdirectory(title="Select input directory")
        else:
            selected = filedialog.askopenfilename(
                title="Select Word file",
                filetypes=[("Word documents", "*.doc *.docx"), ("All files", "*.*")],
            )
        if selected:
            self.input_var.set(selected)

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

    def _append_result(
        self,
        status: str,
        source: str,
        output: str,
        details: str,
    ) -> None:
        normalized_status = status.upper()
        tag = {
            "OK": "ok",
            "FAILED": "failed",
            "ERROR": "error",
            "INFO": "info",
        }.get(normalized_status, "info")
        item_id = self.results_table.insert(
            "",
            "end",
            values=(status, source, output, details),
            tags=(tag,),
        )
        self.results_table.see(item_id)

    def _clear_log(self) -> None:
        for item in self.results_table.get_children():
            self.results_table.delete(item)

    def _selected_item_id(self) -> str | None:
        selection = self.results_table.selection()
        if selection:
            return selection[0]
        focused = self.results_table.focus()
        return focused or None

    def _selected_row_values(self) -> tuple[str, str, str, str] | None:
        item_id = self._selected_item_id()
        if not item_id:
            return None
        values = self.results_table.item(item_id, "values")
        if len(values) != 4:
            return None
        return tuple(str(value) for value in values)

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
        values = self._selected_row_values()
        if values is None:
            self.status_var.set("Select a result row first.")
            return

        _, _, output, _ = values
        if not output or output == "-":
            self.status_var.set("No output path is available for the selected row.")
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
            self.start_button.state(["disabled"])
            self.status_var.set("Converting... please keep Word available in the background.")
        else:
            self.start_button.state(["!disabled"])

    def _start_conversion(self) -> None:
        if self.worker_thread is not None and self.worker_thread.is_alive():
            messagebox.showinfo("In Progress", "A conversion task is already running.")
            return

        input_text = self.input_var.get().strip()
        output_text = self.output_var.get().strip()
        mode = self.mode_var.get()

        if not input_text:
            messagebox.showerror("Missing Input", "Please select an input file or directory.")
            return

        self.progress_var.set(0)
        self.status_var.set("Starting conversion...")
        self._append_result("INFO", "-", "-", "Starting conversion task.")
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
                result = convert_word_to_pdf(
                    source, target, overwrite=self.overwrite_var.get()
                )
                self.event_queue.put(("file_success", (source, result)))
                return

            input_dir = resolve_input_path(input_text)
            if not input_dir.is_dir():
                raise NotADirectoryError(f"Input path is not a directory: {input_dir}")

            output_dir = Path(output_text).expanduser().resolve() if output_text else None

            def on_progress(index: int, total: int, result: BatchResult) -> None:
                self.event_queue.put(("batch_progress", (index, total, result)))

            results = convert_directory_to_pdf(
                input_dir,
                output_dir,
                recursive=self.recursive_var.get(),
                overwrite=self.overwrite_var.get(),
                progress_callback=on_progress,
            )
            self.event_queue.put(("batch_complete", results))
        except Exception as exc:
            self.event_queue.put(("error", str(exc)))

    def _poll_events(self) -> None:
        try:
            while True:
                event_type, payload = self.event_queue.get_nowait()
                if event_type == "file_success":
                    source, result = payload
                    self.progress_var.set(100)
                    self.status_var.set("Single-file conversion completed.")
                    self._append_result(
                        "OK",
                        str(source),
                        str(result),
                        "Single-file conversion completed.",
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
                    results = payload
                    self.progress_var.set(100)
                    summary = format_batch_summary(results)
                    self.status_var.set(summary)
                    self._append_result("INFO", "-", "-", summary)
                    self._set_running_state(False)
                elif event_type == "error":
                    self.status_var.set("Conversion failed.")
                    self._append_result("ERROR", "-", "-", str(payload))
                    messagebox.showerror("Conversion Error", str(payload))
                    self._set_running_state(False)
        except queue.Empty:
            pass
        finally:
            self.root.after(150, self._poll_events)

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
