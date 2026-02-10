

# Multi-format Word Counter (Report Generator)

A small Tkinter desktop app that scans a folder, extracts text from common document/code formats, counts occurrences of chosen words, and writes a `report*.txt` table to an output folder (without overwriting existing reports).

NOTE: This project is too basic, open to further development and feedback

Implementation lives in [src/extract_text_pdf.py](src/extract_text_pdf.py) (entry point: [`main`](src/extract_text_pdf.py)).

## Features

- **GUI workflow** (tabs for Folders / Formats / Words / Run)
- **Counts whole-word matches** using word boundaries (regex `\b...\b`)
- **Case-sensitive or case-insensitive** matching
- **Optional recursive scan** (include subfolders)
- **Progress + cancel** (cancels between files; PDF extraction checks between pages)
- **No overwrites**: reports are saved as `report.txt`, `report(1).txt`, `report(2).txt`, ...

Core logic: [`generate_word_report_for_files`](src/extract_text_pdf.py)

## Supported formats

Built-in handling for:

- Plain text-like: `.txt`, `.md`, `.rst`, `.log`, `.csv`, `.tsv`, `.json`, `.ndjson`, `.jsonl`, `.xml`, `.xhtml`, `.yaml`, `.yml`, `.toml`, `.ini`, `.cfg`, `.env`, `.properties`, many code extensions, `.srt`, `.ipynb`, etc. (see `PLAIN_TEXT_EXTS` in [src/extract_text_pdf.py](src/extract_text_pdf.py))
- Container formats (optional dependencies):
  - `.pdf` via `pypdf`
  - `.docx` via `python-docx`
  - `.rtf` via `striprtf`
- HTML parsing is improved if `beautifulsoup4` is installed; otherwise raw HTML is scanned.

## Installation

Create and activate a virtual environment (recommended), then install dependencies.

```sh
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
```

Install from your requirements file:

```sh
python -m pip install -r requirements.txt
```

Optional format support (install only what you need):

```sh
python -m pip install pypdf python-docx striprtf beautifulsoup4
```

## Run

Launch the GUI:

```sh
python -m src.extract_text_pdf
```

Or:

```sh
python src/extract_text_pdf.py
```

## Usage

1. **Folders tab**: pick the folder to scan and where to save the report.
2. **Formats tab**: choose which extensions to include (and optionally add “Other extensions”).
3. **Words tab**: enter words separated by spaces and/or commas, choose case sensitivity and recursion.
4. **Run tab**: click **Generate report**.

The output report is a plain text table with one row per file, per-word counts, and a total column. If extraction fails for a file, the report includes an error note for that row.

## Notes / behavior

- Matching is **whole-word** based. For a word `w`, the pattern is effectively `\b<w>\b` (escaped).
- “Cancel” stops after the current file (and between pages for PDFs).
- Only selected and supported formats are scanned; unsupported extensions are ignored even if entered.

## Development

- Main module: [src/extract_text_pdf.py](src/extract_text_pdf.py)
- Test file present: [extract_text_pdf.spec](extract_text_pdf.spec)

