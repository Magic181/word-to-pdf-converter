# Word to PDF Converter

A lightweight Windows CLI tool for converting Word documents to PDF with high layout fidelity.

This project uses `pywin32` to automate a locally installed Microsoft Word instance and export `.doc` or `.docx` files to PDF. For personal desktop use on Windows, this is usually one of the most reliable ways to preserve pagination, fonts, headers, footers, and overall document layout.

## Features

- Convert `.doc` and `.docx` files to PDF
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

Specify an output PDF path:

```bash
wp-transform "D:\docs\example.docx" -o "D:\output\example.pdf"
```

Overwrite an existing PDF:

```bash
wp-transform "D:\docs\example.docx" --overwrite
```

## Example Output

```text
Converted successfully: D:\docs\example.pdf
```

## Notes

- If Microsoft Word is not installed, conversion will fail.
- This tool is designed for local desktop use, not for headless servers.
- The conversion quality depends on Word being able to open the source document normally.

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
