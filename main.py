from __future__ import annotations

import argparse
import gc
import sys
from pathlib import Path

try:
    import pythoncom  # type: ignore[import]
    import win32com.client  # type: ignore[import]
    from pywintypes import com_error  # type: ignore[import]
except ImportError:  # pragma: no cover - handled at runtime for missing dependency
    pythoncom = None
    win32com = None
    com_error = Exception


WD_EXPORT_FORMAT_PDF = 17
WD_DO_NOT_SAVE_CHANGES = 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert a Word document (.doc/.docx) to PDF on Windows."
    )
    parser.add_argument("input_path", help="Path to the source Word document.")
    parser.add_argument(
        "-o",
        "--output",
        dest="output_path",
        help="Optional target PDF path. Defaults to the same name beside the source file.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite the target PDF if it already exists.",
    )
    return parser


def resolve_paths(input_path: str, output_path: str | None) -> tuple[Path, Path]:
    source = Path(input_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"Input file does not exist: {source}")
    if source.suffix.lower() not in {".doc", ".docx"}:
        raise ValueError("Input file must be a .doc or .docx document.")

    if output_path:
        target = Path(output_path).expanduser().resolve()
    else:
        target = source.with_suffix(".pdf")

    if target.suffix.lower() != ".pdf":
        raise ValueError("Output file must use the .pdf extension.")

    return source, target


def ensure_dependencies() -> None:
    if pythoncom is None or win32com is None:
        raise RuntimeError(
            "Missing dependency: pywin32. Install it with `pip install -r requirements.txt`."
        )


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


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        source, target = resolve_paths(args.input_path, args.output_path)
        result = convert_word_to_pdf(source, target, overwrite=args.overwrite)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Converted successfully: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
