# Office to PDF Converter

A Windows desktop tool for converting Word, Excel, and PowerPoint files to PDF with Microsoft Office automation.

## Features

- Convert Word files: `.doc`, `.docx`
- Convert Excel files: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`
- Convert PowerPoint files: `.ppt`, `.pptx`, `.pptm`
- Use the CLI or the built-in desktop GUI
- Convert one file, a whole directory, or a dragged-in file list
- Reuse Office sessions for batch work to reduce startup overhead
- Retry failed files automatically and rebuild the Office session when needed
- Run batch jobs with configurable parallel workers
- Drag and drop files, folders, and output targets in the GUI
- Inspect results in a sortable table with preview details and quick-open actions
- Watch an input folder in the GUI and surface file changes in the results list

## Requirements

- Windows
- Python 3.10 or newer
- Microsoft Office installed locally

## Installation

```bash
pip install -r requirements.txt
```

Or install the project as a command-line tool:

```bash
pip install .
```

## CLI Usage

Convert a single file:

```bash
python main.py "D:\docs\example.docx"
```

Choose a custom output path:

```bash
python main.py "D:\docs\budget.xlsx" -o "D:\output\budget.pdf"
```

Run a recursive batch with retries and parallel workers:

```bash
python main.py "D:\docs\office-files" -o "D:\docs\pdf-output" --recursive --overwrite --retry-attempts 1 --workers 2
```

Launch the GUI:

```bash
python main.py --gui
```

Or simply:

```bash
python main.py
```

## GUI Highlights

- Drag one Office file to switch into single-file mode
- Drag a folder to switch into directory batch mode
- Drag multiple supported files to create a one-off batch queue
- Tune retry count and batch workers from the left-side options panel
- Watch the selected input folder for added, changed, or removed Office files
- Review each result in the preview panel and open the source, output file, or output folder

## Notes

- This tool is designed for local desktop use, not for headless servers.
- Conversion quality depends on Microsoft Office being able to open the source file normally.
- Batch jobs may open multiple Office instances when parallel workers are greater than `1`.
- Runtime settings are stored locally in `app_config.json`.
- Runtime logs are written to `office_to_pdf.log`.

## Project Structure

```text
.
├── main.py
├── pyproject.toml
├── requirements.txt
└── README.md
```

## License

MIT. See [LICENSE](LICENSE).
