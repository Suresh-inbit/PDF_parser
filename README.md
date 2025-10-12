
# ProposalExtractorGemini

A small utility to extract structured fields from PDF proposal documents using the Google Generative AI (Gemini) SDK and update an Excel sheet.

## Repository contents

- `main.py` - Primary script containing `ProposalExtractorGemini` and a CLI-style runner.
- `requirements.txt` - Python dependencies.
- `proposals/` - Folder with proposal PDFs organized by TPN.
- `proposals_sheet.xlsx` - Example Excel sheets used by the script.

## What it does

The script reads an Excel file and for each row with a TPN No. finds the corresponding PDF in `proposals/`. It uploads the PDF to the Gemini SDK, asks the model to extract a fixed set of fields, and writes the extracted values back to columns L..AB in the Excel sheet.

## Requirements

- Python 3.10+ (3.13 recommended)
- A working Google Gemini API key / access to Google Generative AI
- The Python packages listed in `requirements.txt` (pandas, openpyxl, etc.)

## Installation

1. Clone the repository

```bash
https://github.com/Suresh-inbit/PDF_parser.git
```

2. Create and activate a virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate    # Linux / macOS
# .venv\Scripts\activate   # Windows (PowerShell)
```

3. Install dependencies

```bash
pip install -r requirements.txt
```

If a package from `requirements.txt` is unavailable on PyPI for your Python version, install compatible versions manually (see Troubleshooting below).

## Configuration

Edit `main.py` and set your API key in the `API_KEY` variable near the bottom of the file, or modify the script to read from an environment variable. Example (bash):

```bash
export GEMINI_API_KEY="your_api_key_here"
```


The script expects the Excel sheet to have headers on row 5 (so `pandas.read_excel(..., header=4)`). Data rows start at Excel row 6.

## Usage

Basic run (after configuring API key and paths in `main.py`):

```bash
python main.py
```

What the script does:

- Loads the Excel file configured in `main.py`.
- Scans `proposals/` for PDFs (expects filenames like `ProposalID_<TPN>_finalproposal.pdf`).
- Uploads each PDF to Gemini and runs the extraction prompt.
- Writes results into columns L..AB of a new Excel file (does not overwrite the original by default, set OUPUT_FILE to EXCEL_FILE for writing to same file).

To run the extractor on a single file for debugging, you can import and call the class from a Python shell or small script.

## Troubleshooting

- Timeout / 504 errors: the script now includes upload + generate retries with exponential backoff, but network or API-side rate limits can still cause failures. Increase `max_retries` in `extract_from_pdf` or add longer backoff.
- `ModuleNotFoundError: No module named 'google.genai'` - ensure you installed the correct client package for the Gemini SDK (package name and availability may change). If the official package isn't available, check provider docs for the correct install instructions or use an alternative client.
- Excel headers mismatch: open your Excel and ensure the header row is row 5. If headers are elsewhere, change `header=4` in `pandas.read_excel` calls.



## License / Credits
Developed with AI assistance (Google Gemini SDK).  
