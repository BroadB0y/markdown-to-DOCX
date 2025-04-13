# Markdown to DOCX Converter

A Python script that converts Markdown files to Microsoft Word (.docx) format while preserving headings, paragraphs, lists, and tables.

## Features

- Supports headings (h1-h6)
- Converts paragraphs, bullet lists, and numbered lists
- Handles tables with proper formatting
- Preserves basic text formatting from Markdown
- Automatic output filename generation (same as input but with .docx extension)

## Requirements

- Python 3.x
- Required packages:
  - `markdown`
  - `python-docx`
  - `beautifulsoup4`

## Installation

1. Install the required packages:
   ```bash
   pip install markdown python-docx beautifulsoup4
   ```

## Usage

Run the script and provide the input Markdown filename when prompted:
```bash
python main.py
```

The script will:
1. Ask for input filename (e.g., `document.md`)
2. Generate a Word document with the same base name (e.g., `document.docx`)
3. Save it in the same directory.
