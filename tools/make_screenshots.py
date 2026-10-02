"""Rebuild the three portfolio screenshots in screenshots/ (PNG, exactly 1600x1200).

Usage (Windows, with Excel installed):
    python tools\\make_screenshots.py

The images are composites, not screen grabs, so no window, account name or file
path can appear in them:
  - invoice pages are drawn from the sample PDFs with PyMuPDF;
  - spreadsheet content is printed to PDF by Excel itself, then drawn with PyMuPDF;
  - the terminal picture is the extractor's real output, drawn as text.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pymupdf
from PIL import Image, ImageChops, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"
OUTPUT_XLSX = ROOT / "output" / "invoices.xlsx"
SHOTS = ROOT / "screenshots"
WORK = SHOTS / "_work"
FONTS = Path(r"C:\Windows\Fonts")

WIDTH, HEIGHT = 1600, 1200
STRIP_HEIGHT = 104
MARGIN = 40
CONTENT_BOX = (MARGIN, STRIP_HEIGHT + 28, WIDTH - MARGIN, HEIGHT - 56)
BACKGROUND = "#F4F6F8"
STRIP = "#1F3A5F"
ACCENT = (245, 166, 35)
FOOTNOTE = "Demo \u00b7 sample data"

FEATURED = "contoso-0331.pdf"
FEATURED_FIELDS = ("CONTOSO FREIGHT", "CF-2026-0331", "27 Aug 2026", "USD 2,485.30")
THUMBNAILS = ("contoso-0318.pdf", "fabrikam-00771.pdf", "northwind-10528.pdf", "scan-0007.pdf")


def font(file_name: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONTS / file_name), size)


def content_size() -> tuple[int, int]:
    left, top, right, bottom = CONTENT_BOX
    return right - left, bottom - top


def fit(image: Image.Image, max_width: int, max_height: int) -> Image.Image:
    scale = min(max_width / image.width, max_height / image.height)
    size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
    return image.resize(size, Image.LANCZOS)


def compose(content: Image.Image, caption: str, out_path: Path) -> None:
    """Title strip on top, the content scaled to fill the rest, footnote bottom right."""
    canvas = Image.new("RGB", (WIDTH, HEIGHT), BACKGROUND)
    draw = ImageDraw.Draw(canvas)
    draw.rectangle((0, 0, WIDTH, STRIP_HEIGHT), fill=STRIP)
    draw.text((MARGIN + 8, STRIP_HEIGHT // 2), caption, font=font("seguisb.ttf", 50), fill="white", anchor="lm")

    left, top, right, bottom = CONTENT_BOX
    scaled = fit(content.convert("RGB"), right - left, bottom - top)
    x = left + (right - left - scaled.width) // 2
    y = top + (bottom - top - scaled.height) // 2
    canvas.paste(scaled, (x, y))

    draw.text((WIDTH - MARGIN, HEIGHT - 28), FOOTNOTE, font=font("segoeui.ttf", 22), fill="#8A94A0", anchor="rm")
    assert canvas.size == (WIDTH, HEIGHT)
    canvas.save(out_path, "PNG")
    print(f"  {out_path.name}  {canvas.width}x{canvas.height}")


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


def excel_range_image(xlsx: Path, address: str, name: str, keep_rows: set[int] | None = None,
                      headings: bool = False) -> Image.Image:
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
        setup.PrintHeadings = headings
        setup.PrintGridlines = False
        sheet.Range(address).ExportAsFixedFormat(0, str(pdf_path), 0, False, False)
        workbook.Close(False)  # nothing is saved back
    finally:
        excel.Quit()
    return autocrop(pdf_page_image(pdf_path, dpi=400))


def featured_invoice_image() -> Image.Image:
    """The featured invoice, cropped to the printed part, with the four fields boxed."""
    dpi = 200
    scale = dpi / 72
    pdf_path = SAMPLES / FEATURED
    page_image = pdf_page_image(pdf_path, dpi).convert("RGBA")
    overlay = Image.new("RGBA", page_image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
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
            draw.rounded_rectangle(box, radius=10, fill=ACCENT + (36,), outline=ACCENT + (255,), width=5)
    marked = Image.alpha_composite(page_image, overlay).convert("RGB")
    cropped = marked.crop((0, 0, marked.width, round((lowest + 34) * scale)))
    framed = Image.new("RGB", (cropped.width + 4, cropped.height + 4), "#C5CCD3")
    framed.paste(cropped, (2, 2))
    return framed


def shot_invoice_and_row() -> None:
    row_number = 1 + sorted(p.name.lower() for p in SAMPLES.glob("*.pdf")).index(FEATURED) + 1
    row_strip = excel_range_image(OUTPUT_XLSX, "A1:E11", "row", keep_rows={row_number})
    invoice = featured_invoice_image()

    width, height = content_size()
    content = Image.new("RGB", (width, height), BACKGROUND)
    arrow_height, gap = 64, 18
    strip = fit(row_strip, width, 190)
    invoice = fit(invoice, width, height - strip.height - arrow_height - 2 * gap)

    y = (height - invoice.height - strip.height - arrow_height - 2 * gap) // 2
    content.paste(invoice, ((width - invoice.width) // 2, y))
    y += invoice.height + gap
    draw = ImageDraw.Draw(content)
    middle = width // 2
    draw.rectangle((middle - 14, y, middle + 14, y + 30), fill=ACCENT)
    draw.polygon(((middle - 44, y + 30), (middle + 44, y + 30), (middle, y + arrow_height)), fill=ACCENT)
    y += arrow_height + gap
    content.paste(strip, ((width - strip.width) // 2, y))
    compose(content, "Invoice PDF in, spreadsheet row out", SHOTS / "1-invoice-to-row.png")


def shot_output_sheet(summary_line: str) -> None:
    sheet = excel_range_image(OUTPUT_XLSX, "A1:F11", "sheet", headings=True)
    width, height = content_size()
    content = Image.new("RGB", (width, height), BACKGROUND)
    sheet = fit(sheet, width, height - 260)
    callout_font = font("seguisb.ttf", 84)
    gap = 70
    block = sheet.height + gap + 100
    y = (height - block) // 2
    content.paste(sheet, ((width - sheet.width) // 2, y))
    draw = ImageDraw.Draw(content)
    draw.text((width // 2, y + sheet.height + gap + 50), summary_line.replace(", ", "  \u00b7  "),
              font=callout_font, fill=STRIP, anchor="mm")
    compose(content, "Unreadable invoices flagged, never guessed", SHOTS / "2-output-sheet-flagged-rows.png")


def page_thumbnail(file_name: str, width: int) -> Image.Image:
    """The printed top part of a sample invoice page, with a thin frame."""
    page = pdf_page_image(SAMPLES / file_name, dpi=120)
    page = page.crop((0, 0, page.width, round(page.height * 0.62)))
    page = fit(page, width - 4, 10_000)
    framed = Image.new("RGB", (page.width + 4, page.height + 4), "#C5CCD3")
    framed.paste(page, (2, 2))
    return framed


def shot_terminal(command: str, lines: list[str]) -> None:
    """A few of the sample invoices on top, the extractor's real output below."""
    width, height = content_size()
    content = Image.new("RGB", (width, height), BACKGROUND)
    draw = ImageDraw.Draw(content)

    gap = 24
    thumb_width = (width - 3 * gap) // 4
    thumbs = [page_thumbnail(name, thumb_width) for name in THUMBNAILS]
    label_font = font("segoeui.ttf", 24)
    label_height = 44

    mono = font("consola.ttf", 36)
    mono_bold = font("consolab.ttf", 36)
    pad, step = 44, 60
    terminal_height = pad + step + 14 + step * len(lines) + pad - 20

    y = (height - thumbs[0].height - label_height - gap - terminal_height) // 2
    for index, (name, thumb) in enumerate(zip(THUMBNAILS, thumbs)):
        x = index * (thumb_width + gap)
        content.paste(thumb, (x, y))
        draw.text((x + thumb_width // 2, y + thumb.height + label_height // 2), name,
                  font=label_font, fill="#4A5563", anchor="mm")
    y += thumbs[0].height + label_height + gap

    draw.rounded_rectangle((0, y, width - 1, y + terminal_height), radius=22, fill="#1E2430")
    x, y = pad, y + pad
    draw.text((x, y), "> " + command, font=mono, fill="#9FB3C8")
    y += step + 14
    for index, line in enumerate(lines):
        if index == 0:
            draw.text((x, y), line, font=mono_bold, fill="#FFFFFF")
        elif "REVIEW" in line:
            draw.text((x, y), line, font=mono, fill="#F5C26B")
        else:
            draw.text((x, y), line, font=mono, fill="#C9D4E0")
        y += step
    longest = max(draw.textlength(text, font=mono_bold) for text in ["> " + command, *lines])
    if x + longest > width - 30:
        raise SystemExit("Terminal text is too wide for the picture")
    compose(content, "One command reads the whole folder", SHOTS / "3-one-command-summary.png")


def main() -> int:
    SHOTS.mkdir(exist_ok=True)
    command = ["extract.py", "samples", "output\\invoices.xlsx"]
    run = subprocess.run([sys.executable, *command], cwd=ROOT, capture_output=True, text=True, check=True)
    lines = run.stdout.rstrip().splitlines()
    print("Extractor said:", lines[0])
    print("Writing screenshots:")
    shot_invoice_and_row()
    shot_output_sheet(lines[0])
    shot_terminal("python " + " ".join(command), lines)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
