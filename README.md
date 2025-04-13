# Markdown to DOCX Converter

This Python script converts Markdown files to Microsoft Word (DOCX) format while preserving basic formatting elements like headings, lists, and tables.

## Requirements

- Python 3.x
- Required Python packages:
  - `markdown`
  - `python-docx`
  - `beautifulsoup4`

## Installation

1. Install the required packages using pip:

```bash
pip install markdown python-docx beautifulsoup4
```

2. Clone or download this repository

## Usage

1. Prepare your Markdown file (default input filename is `input.md`)
2. Run the script:

```bash
python main.py
```

3. The script will generate a Word document named `output.docx`

### Custom File Names

You can modify the input and output file names by editing these lines in the script:

```python
input_md = "input.md"    # Your Markdown file
output_docx = "output.docx"  # Output DOCX filename
```

## Supported Markdown Elements

- Headers (`#`, `##`, `###`)
- Paragraphs
- Unordered lists (`-` or `*`)
- Ordered lists (numbered items)
- Tables
- Basic text formatting (bold, italic, etc.)


