# Office to PDF Converter

A lightweight Windows tool for converting Word/Excel/PowerPoint documents to PDF with high layout fidelity.

This project uses `pywin32` to automate a locally installed Microsoft Word instance and export `.doc` or `.docx` files to PDF. It now supports both command-line usage and a desktop GUI, including batch conversion for entire directories.

## Features

- Convert Word files (`.doc`, `.docx`) to PDF
- Convert Excel files (`.xls`, `.xlsx`, `.xlsm`, `.xlsb`) to PDF
- Convert PowerPoint files (`.ppt`, `.pptx`, `.pptm`) to PDF
- Launch a desktop GUI for easier everyday use
- Support drag and drop for input files, input folders, and output targets in the GUI
- Batch-convert all Word files in a directory
- Optionally scan subdirectories recursively
- Use the source filename by default and export beside the original document
- Support custom output paths
- Support overwriting an existing PDF
- Keep formatting by relying on Microsoft Word's native export

## Requirements

- Windows
- Python 3.10 or newer
- Microsoft Word installed locally

## Installation

Install dependencies directly:

```bash
pip install -r requirements.txt
```

Or install the project as a command-line tool:

```bash
pip install .
```

## Usage

Run the script directly:

```bash
python main.py "D:\docs\example.docx"
```

Run the installed command:

```bash
wp-transform "D:\docs\example.docx"
```

Launch the GUI:

```bash
python main.py --gui
```

Or simply run without arguments:

```bash
python main.py
```

Specify an output PDF path:

```bash
wp-transform "D:\docs\example.docx" -o "D:\output\example.pdf"
```

Overwrite an existing PDF:

```bash
wp-transform "D:\docs\example.docx" --overwrite
```

Batch-convert a directory:

```bash
wp-transform "D:\docs\word-files" -o "D:\docs\pdf-output" --overwrite
```

Batch-convert a directory recursively:

```bash
wp-transform "D:\docs\word-files" -o "D:\docs\pdf-output" --recursive --overwrite
```

When the input is a directory:

- `input_path` is treated as the source folder
- `-o/--output` is treated as the target folder
- the relative folder structure is preserved in the output directory

## Example Output

```text
Converted successfully: D:\docs\example.pdf
```

## Notes

- If Microsoft Word is not installed, conversion will fail.
- This tool is designed for local desktop use, not for headless servers.
- The conversion quality depends on Word being able to open the source document normally.
- In batch mode, files that fail to convert are reported individually in the log or terminal output.

## Project Structure

```text
.
├── main.py
├── pyproject.toml
├── requirements.txt
└── README.md
```

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
