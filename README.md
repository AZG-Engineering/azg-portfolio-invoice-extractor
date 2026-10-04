# Invoice PDF to spreadsheet extractor

Reads every PDF invoice in a folder and writes one spreadsheet row per invoice:
Vendor, Invoice Date, Invoice Number and Total, plus the Source File and a Review
column.

It never guesses. A field it can't read with certainty is left blank, the row is
highlighted, and the Review column says why ("Total not found", "No text layer
(scanned image): enter by hand"). A person then only has to look at the flagged rows.

**Demo built with made-up sample data.** Every vendor, customer, address, invoice
number and amount in `samples\` is invented.

Built by AZG Engineering.

## What's in the folder

| Path | What it is |
| --- | --- |
| `extract.py` | The extractor. One command, one spreadsheet. |
| `vendors.toml` | The vendor layouts it knows how to read. Edit this to add a vendor. |
| `samples\` | 10 made-up invoice PDFs: 3 vendor layouts, 8 clean, 2 deliberate problem files. |
| `generate_samples.py` | Rebuilds `samples\` exactly. |
| `tests\test_extract.py` | The tests: expected values for every sample, both problem files flagged. |
| `tools\make_screenshots.py` | Rebuilds the three images in `screenshots\`. |
| `tools\check_metadata.py` | Shows the author fields of every spreadsheet, PDF and image here. |
| `tools\showcase.py`, `tools\fonts\` | The AZG showcase kit (the screenshot template) and the Inter font it uses. |
| `screenshots\` | Three 1600x1200 PNG images of the demo. |
| `output\` | Where the spreadsheet is written. Not kept in git; it is rebuilt on every run. |

## How to run it

Windows, Python 3.11 or newer. Everything installs into this folder only.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe extract.py samples output\invoices.xlsx
```

It prints:

```
10 files, 8 OK, 2 need review
  REVIEW  northwind-10528.pdf: Total not found
  REVIEW  scan-0007.pdf: No text layer (scanned image): enter by hand
Saved output\invoices.xlsx
```

Point it at any other folder of PDFs the same way:
`extract.py <pdf_folder> <output.xlsx>`. Close the spreadsheet in Excel before
running it again, or it can't be overwritten.

### The two problem files

- `northwind-10528.pdf` stops at "Continued on next page", so there is no total. The
  vendor, date and number are filled in; Total is blank and flagged.
- `scan-0007.pdf` is a picture of an invoice with no text in it, like a scan. Every
  field is blank and the row is flagged. This tool does not do OCR.

### Tests

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m pytest -v
```

## How to change it

**Add a vendor.** Copy a `[[vendor]]` block in `vendors.toml` and change:

- `name`: what goes in the Vendor column;
- `match`: text that only this vendor's invoices contain;
- `invoice_number`, `invoice_date`, `total`: a pattern for each, anchored on the label
  printed next to the value, with one pair of brackets around the part to keep;
- `date_format`: how the date is written (examples are in the file).

To see the text the extractor sees in a PDF, and so what to anchor on:

```powershell
.\.venv\Scripts\python.exe -c "import extract, pathlib; print(extract.read_text(pathlib.Path('samples/contoso-0318.pdf')))"
```

Then add the new invoice's expected values to `EXPECTED` in `tests\test_extract.py`
and run the tests.

**The never-guess rules** are in `parse_invoice` and `single_match` in `extract.py`:

- no text in the PDF: every field blank;
- no vendor layout matches, or more than one does: every field blank;
- a pattern finds nothing: that field blank ("... not found");
- a pattern finds two different values: that field blank ("... has more than one
  candidate"), with both values shown;
- a date or amount that doesn't parse: that field blank ("... unreadable").

**Add a column.** Add the field to `Row`, read it in `parse_invoice`, and add it to
`HEADERS`, `COLUMN_WIDTHS` and the row written in `write_xlsx`.

**Change the look of the spreadsheet** (widths, number formats, row heights):
`write_xlsx`. The colours are the constants at the top of `extract.py` and come from
the AZG build style guide: a navy header band, one thin rule under each row, and
amber with dark-amber text for rows that need review.

**Change the sample invoices:** edit `INVOICES` in `generate_samples.py`, then

```powershell
.\.venv\Scripts\python.exe generate_samples.py
```

## Limits

- Text PDFs only. Scanned invoices are flagged, not read.
- One invoice per PDF file.
- A vendor must have a layout in `vendors.toml`; unknown vendors are flagged.
- Amounts are read as printed, with a dot for decimals and commas for thousands.

## Screenshots

`tools\make_screenshots.py` rebuilds the three images. It needs Windows with Excel
installed: Excel prints the spreadsheet to PDF in the background, and nothing is saved
back. The images are composites, not screen grabs, so no window, account name or file
path appears in them.

The frame (navy strip, caption, white card, footer) comes from the AZG showcase kit,
`tools\showcase.py`, which follows the AZG build style guide. Don't edit that copy:
change the kit in `Products\showcase-kit\` and sync it. A 400 px wide copy of each
image is written to `screenshots\_work\` to check that the headline still reads at
thumbnail size.

## Dependencies and licences

Nothing is copied into this repository from a third party: no templates, icons or
code. The packages below are installed by pip into `.venv`.

Needed to run the extractor (`requirements.txt`):

| Package | Version | Licence |
| --- | --- | --- |
| pdfplumber | 0.11.10 | MIT |
| openpyxl | 3.1.5 | MIT |
| pdfminer.six (via pdfplumber) | 20260107 | MIT |
| Pillow (via pdfplumber) | 12.3.0 | MIT-CMU |
| pypdfium2 (via pdfplumber) | 5.13.0 | BSD-3-Clause / Apache-2.0, plus PDFium's own licences |
| charset-normalizer (via pdfminer.six) | 3.5.2 | MIT |
| cryptography (via pdfminer.six) | 50.0.2 | Apache-2.0 OR BSD-3-Clause |
| cffi (via cryptography) | 2.1.1 | MIT-0 |
| pycparser (via cffi) | 3.0 | BSD-3-Clause |
| et_xmlfile (via openpyxl) | 2.0.0 | MIT |

Needed only for samples, tests and screenshots (`requirements-dev.txt`):

| Package | Version | Licence | Used for |
| --- | --- | --- | --- |
| reportlab | 5.0.1 | BSD | drawing the sample PDFs |
| pytest | 9.1.1 | MIT | tests |
| iniconfig, pluggy (via pytest) | 2.3.0, 1.6.0 | MIT | tests |
| packaging (via pytest) | 26.3 | Apache-2.0 OR BSD-2-Clause | tests |
| Pygments (via pytest) | 2.21.0 | BSD-2-Clause | tests |
| colorama (via pytest) | 0.4.6 | BSD | tests |
| pywin32 | 312 | PSF | screenshots: driving Excel |
| PyMuPDF | 1.28.2 | **AGPL-3.0, or a paid Artifex licence** | screenshots: PDF to image |

PyMuPDF is the one to know about. It is used only by `tools\make_screenshots.py`;
`extract.py` does not import it and runs without it. If
this tool is ever sold or handed over, leave PyMuPDF out or swap the screenshot step
to pypdfium2, which is already installed.

Kept in this repository, for the screenshots only:

| Item | Version | Licence | Where |
| --- | --- | --- | --- |
| Inter font (Regular, SemiBold) | 4.1 | SIL Open Font License 1.1 | `tools\fonts\`, with the licence text in `tools\fonts\OFL.txt` |
| AZG showcase kit | 1.2.0 | AZG Engineering's own | `tools\showcase.py`, an exact copy of `Products\showcase-kit\showcase.py` |

No system font is drawn into the images.

## Licence
Copyright © 2026 Azimuth Group LLC. All rights reserved. Published as a work sample; no licence to reuse it is granted. The Inter font in `tools\fonts\` is under the SIL Open Font License 1.1 (`tools\fonts\OFL.txt`).
