"""Rebuild the three showcase images in screenshots/ (PNG, exactly 1600x1200).

Usage (Windows, with Excel installed):
    python tools\\make_screenshots.py

The images are composites on the AZG showcase template (tools\\showcase.py), not
screen grabs, so no window, account name or file path can appear in them:
  - invoice pages are drawn from the sample PDFs with PyMuPDF;
  - spreadsheet content is printed to PDF by Excel itself, then drawn with PyMuPDF;
  - the run summary is the extractor's real output, drawn as text.

A 400 px wide copy of each image is written to screenshots\\_work\\ so the
headline can be checked at thumbnail size.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pymupdf
from PIL import Image, ImageChops, ImageDraw

import showcase as kit  # tools\showcase.py: a copy of the AZG showcase kit

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"
OUTPUT_XLSX = ROOT / "output" / "invoices.xlsx"
SHOTS = ROOT / "screenshots"
WORK = SHOTS / "_work"

FEATURED = "contoso-0331.pdf"
FEATURED_FIELDS = ("CONTOSO FREIGHT", "CF-2026-0331", "27 Aug 2026", "USD 2,485.30")
THUMBNAILS = ("contoso-0318.pdf", "fabrikam-00771.pdf", "northwind-10528.pdf", "scan-0007.pdf")


def autocrop(image: Image.Image, pad: int = 6) -> Image.Image:
    image = image.convert("RGB")
    box = ImageChops.difference(image, Image.new("RGB", image.size, "white")).getbbox()
    if box is None:
        return image
    return image.crop((max(box[0] - pad, 0), max(box[1] - pad, 0),
                       min(box[2] + pad, image.width), min(box[3] + pad, image.height)))


def pdf_page_image(pdf_path: Path, dpi: int = 200) -> Image.Image:
    with pymupdf.open(pdf_path) as document:
        pixmap = document[0].get_pixmap(dpi=dpi, alpha=False)
        return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)


def excel_range_image(xlsx: Path, address: str, name: str, keep_rows: set[int] | None = None) -> Image.Image:
    """Have Excel print a range to PDF, then draw that PDF as an image."""
    import win32com.client as win32

    WORK.mkdir(parents=True, exist_ok=True)
    pdf_path = WORK / f"{name}.pdf"
    excel = win32.DispatchEx("Excel.Application")  # a private instance, never the user's open Excel
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(xlsx), 0, True)  # no link updates, read-only
        sheet = workbook.Worksheets(1)
        if keep_rows is not None:
            for row in range(2, sheet.UsedRange.Rows.Count + 1):
                if row not in keep_rows:
                    sheet.Rows(row).Hidden = True
        setup = sheet.PageSetup
        setup.Orientation = 2  # landscape
        setup.Zoom = False
        setup.FitToPagesWide = 1
        setup.FitToPagesTall = 1
        setup.PrintHeadings = False
        setup.PrintGridlines = False
        sheet.Range(address).ExportAsFixedFormat(0, str(pdf_path), 0, False, False)
        workbook.Close(False)  # nothing is saved back
    finally:
        excel.Quit()
    return autocrop(pdf_page_image(pdf_path, dpi=400))


def featured_invoice_image() -> Image.Image:
    """The featured invoice, cropped to the printed part, with the four fields boxed in teal."""
    dpi = 200
    scale = dpi / 72
    pdf_path = SAMPLES / FEATURED
    page_image = pdf_page_image(pdf_path, dpi).convert("RGBA")
    overlay = Image.new("RGBA", page_image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    teal = tuple(int(kit.TEAL[i:i + 2], 16) for i in (1, 3, 5))
    lowest = 0.0
    with pymupdf.open(pdf_path) as document:
        page = document[0]
        for text in FEATURED_FIELDS:
            hits = page.search_for(text)
            if len(hits) != 1:
                raise SystemExit(f"Expected one '{text}' on {FEATURED}, found {len(hits)}")
            rect = hits[0]
            lowest = max(lowest, rect.y1)
            box = ((rect.x0 - 5) * scale, (rect.y0 - 3) * scale, (rect.x1 + 5) * scale, (rect.y1 + 3) * scale)
            draw.rounded_rectangle(box, radius=10, fill=teal + (40,), outline=teal + (255,), width=5)
    marked = Image.alpha_composite(page_image, overlay).convert("RGB")
    return marked.crop((0, 0, marked.width, round((lowest + 30) * scale)))


def finish(card: Image.Image, caption: str, file_name: str) -> None:
    path = kit.compose(card, caption, SHOTS / file_name)
    small = kit.thumbnail(path, WORK)
    print(f"  {path.name}  {kit.WIDTH}x{kit.HEIGHT}   (thumbnail: _work\\{small.name})")


def shot_invoice_and_row() -> None:
    """Before: the invoice, fields boxed. After: its one row in the output sheet."""
    row_number = 1 + sorted(p.name.lower() for p in SAMPLES.glob("*.pdf")).index(FEATURED) + 1
    row_strip = excel_range_image(OUTPUT_XLSX, "A1:E11", "row", keep_rows={row_number})
    invoice = featured_invoice_image()

    # Give the row strip the full card width, and the invoice the height that is left.
    left, top, right, bottom = kit.inner(kit.new_card())
    strip_height = row_strip.height * (right - left) / row_strip.width
    usable = bottom - top - kit.ARROW_GAP - 2 * kit.LABEL_ROW
    card, scales = kit.before_after(invoice, row_strip, orientation="vertical", share=(usable - strip_height) / usable,
                                    before_note="a PDF invoice, as it arrives", after_note="one clean spreadsheet row")
    print(f"  invoice drawn at {scales[0]:.2f}x, row at {scales[1]:.2f}x")
    finish(card, "Invoice PDF in, spreadsheet row out", "1-invoice-to-row.png")


def shot_output_sheet(summary_line: str) -> None:
    """The headline result, then the whole output sheet with its two flagged rows."""
    sheet = excel_range_image(OUTPUT_XLSX, "A1:F11", "sheet")
    card = kit.new_card()
    left, top, right, bottom = kit.inner(card)
    note = "Blank cells are blank on purpose. The Review column says why, so a person only checks those rows."
    scaled, _ = kit.scale_to_fit(sheet, right - left, 10_000)
    headline_height, note_height = 110, 40
    spare = (bottom - top) - headline_height - scaled.height - note_height
    gap = max(24, spare // 4)

    y = top + gap // 2
    kit.headline(card, summary_line.replace(", ", "  ·  "), card.width // 2, y + headline_height // 2, size=84)
    y += headline_height + gap
    box, scale = kit.place(card, sheet, (left, y, right, y + scaled.height), valign="top")
    kit.text(ImageDraw.Draw(card), (card.width // 2, box[3] + gap + note_height // 2), note, 26, "regular", kit.SLATE, anchor="mm")
    print(f"  sheet drawn at {scale:.2f}x")
    finish(card, "Unreadable invoices flagged, never guessed", "2-output-sheet-flagged-rows.png")


def page_thumbnail(file_name: str, width: int) -> Image.Image:
    """The printed top part of a sample invoice page, with a thin frame."""
    page = pdf_page_image(SAMPLES / file_name, dpi=150)
    page = page.crop((0, 0, page.width, round(page.height * 0.75)))
    page, _ = kit.scale_to_fit(page, width - 2, 10_000)
    framed = Image.new("RGB", (page.width + 2, page.height + 2), kit.LINE)
    framed.paste(page, (1, 1))
    return framed


def shot_one_command(command: str, lines: list[str]) -> None:
    """A few of the sample invoices on top, the extractor's real output below."""
    card = kit.new_card()
    draw = ImageDraw.Draw(card)
    left, top, right, bottom = kit.inner(card)
    width = right - left

    gap = 24
    thumb_width = (width - 3 * gap) // 4
    thumbs = [page_thumbnail(name, thumb_width) for name in THUMBNAILS]
    label_height = 44

    pad, step = 48, 56
    details = lines[1:]
    panel_height = pad + 48 + 24 + 108 + step * len(details) + pad - 10
    block = thumbs[0].height + label_height + 28 + panel_height
    y = top + max(0, (bottom - top - block) // 2)

    for index, (name, thumb) in enumerate(zip(THUMBNAILS, thumbs)):
        x = left + index * (thumb_width + gap)
        card.paste(thumb, (x, y))
        kit.text(draw, (x + thumb_width // 2, y + thumb.height + label_height // 2), name, 22, "regular", kit.SLATE, anchor="mm")
    y += thumbs[0].height + label_height + 28

    draw.rounded_rectangle((left, y, right, y + panel_height), radius=18, fill=kit.INK)
    x, y = left + pad, y + pad
    kit.text(draw, (x, y), "> " + command, 32, "regular", kit.LINE)
    y += 48 + 24
    kit.text(draw, (x, y), lines[0], 84, "semibold", kit.WHITE)
    y += 108
    for line in details:
        review = "REVIEW" in line
        kit.text(draw, (x + (30 if review else 0), y), " ".join(line.split()), 34, "regular",
                 kit.AMBER_FILL if review else kit.LINE)
        y += step
    widest = max(kit.text_width(lines[0], 84, "semibold"), *(kit.text_width(" ".join(l.split()), 34) + 30 for l in details))
    if x + widest > right - pad:
        raise SystemExit("The run output is too wide for the picture")
    finish(card, "One command reads the whole folder", "3-one-command-summary.png")


def main() -> int:
    SHOTS.mkdir(exist_ok=True)
    command = ["extract.py", "samples", "output\\invoices.xlsx"]
    run = subprocess.run([sys.executable, *command], cwd=ROOT, capture_output=True, text=True, check=True)
    lines = run.stdout.rstrip().splitlines()
    print(f"showcase kit {kit.__version__}. Extractor said: {lines[0]}")
    print("Writing screenshots:")
    shot_invoice_and_row()
    shot_output_sheet(lines[0])
    shot_one_command("python " + " ".join(command), lines)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
