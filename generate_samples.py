"""Generate the made-up sample invoices used by the demo and the tests.

Usage:
    python generate_samples.py [output_folder]      (default: samples)

Every name, address, number and amount here is invented. The set is fixed, so
running this again produces the same invoices:

  - 3 vendor layouts (Northwind Supply Co., Contoso Freight, Fabrikam Office Goods)
  - 8 clean invoices
  - 2 deliberate problem files:
      northwind-10528.pdf  the totals block is missing ("Continued on next page")
      scan-0007.pdf        a picture of an invoice with no text layer, like a scan
"""

from __future__ import annotations

import io
import random
import sys
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

AUTHOR = "AZG Engineering"
SAMPLE_NOTE = "SAMPLE - made-up data for demonstration only"

PAGE_W, PAGE_H = letter
LEFT, RIGHT = 54, PAGE_W - 54
CENT = Decimal("0.01")


@dataclass(frozen=True)
class Invoice:
    file_name: str
    layout: str  # northwind | contoso | fabrikam
    number: str
    issued: date
    # northwind / fabrikam: (description, quantity, unit price)
    # contoso:              (service, details, charge)
    lines: tuple[tuple, ...]
    omit_total: bool = False
    as_scan: bool = False


INVOICES: tuple[Invoice, ...] = (
    Invoice("northwind-10482.pdf", "northwind", "NW-10482", date(2026, 8, 3), (
        ("Copy paper, 10-ream case", 4, "42.50"),
        ("Packing tape, 6-pack", 3, "11.20"),
        ("Shipping labels, box of 500", 2, "18.75"),
    )),
    Invoice("northwind-10497.pdf", "northwind", "NW-10497", date(2026, 8, 19), (
        ("Corrugated boxes 12x12x8, bundle of 25", 6, "31.40"),
        ("Bubble wrap roll, 24 in", 2, "27.95"),
    )),
    Invoice("northwind-10513.pdf", "northwind", "NW-10513", date(2026, 9, 8), (
        ("Stretch film, 4-roll case", 12, "64.00"),
        ("Box cutter, safety", 12, "6.85"),
        ("Pallet labels, roll", 3, "22.10"),
        ("Nitrile gloves, box", 10, "9.40"),
    )),
    Invoice("northwind-10528.pdf", "northwind", "NW-10528", date(2026, 9, 22), (
        ("Copy paper, 10-ream case", 8, "42.50"),
        ("Toner-safe envelopes, box", 4, "15.30"),
    ), omit_total=True),
    Invoice("contoso-0318.pdf", "contoso", "CF-2026-0318", date(2026, 8, 11), (
        ("LTL freight, 2 pallets", "Zone 3, 412 mi", "685.00"),
        ("Fuel surcharge", "8.5% of linehaul", "58.23"),
        ("Liftgate delivery", "Flat fee", "75.00"),
    )),
    Invoice("contoso-0331.pdf", "contoso", "CF-2026-0331", date(2026, 8, 27), (
        ("Full truckload, 53 ft dry van", "Zone 5, 1,140 mi", "2180.00"),
        ("Fuel surcharge", "8.5% of linehaul", "185.30"),
        ("Detention, 2 hours", "After 2 free hours", "120.00"),
    )),
    Invoice("contoso-0346.pdf", "contoso", "CF-2026-0346", date(2026, 9, 15), (
        ("LTL freight, 1 pallet", "Zone 2, 188 mi", "312.00"),
        ("Fuel surcharge", "8.5% of linehaul", "26.52"),
        ("Residential delivery", "Flat fee", "45.00"),
    )),
    Invoice("scan-0007.pdf", "contoso", "CF-2026-0352", date(2026, 9, 24), (
        ("LTL freight, 3 pallets", "Zone 4, 655 mi", "940.00"),
        ("Fuel surcharge", "8.5% of linehaul", "79.90"),
    ), as_scan=True),
    Invoice("fabrikam-00771.pdf", "fabrikam", "FOG-00771", date(2026, 8, 6), (
        ("Desk organizer, mesh", 6, "14.90"),
        ("Whiteboard markers, 12-pack", 4, "9.75"),
        ("Toner cartridge, black", 2, "68.00"),
    )),
    Invoice("fabrikam-00784.pdf", "fabrikam", "FOG-00784", date(2026, 9, 2), (
        ("Ergonomic chair mat", 3, "54.25"),
        ("Monitor riser, bamboo", 5, "32.80"),
        ("Cable sleeves, 10 ft", 8, "7.15"),
    )),
)

TAX_RATE = {"northwind": Decimal("0.06"), "fabrikam": Decimal("0.055"), "contoso": Decimal("0")}


def money(value: Decimal) -> str:
    return f"{value:,.2f}"


def amounts(inv: Invoice) -> tuple[Decimal, Decimal, Decimal]:
    """Return (subtotal, tax, total) for an invoice."""
    if inv.layout == "contoso":
        subtotal = sum((Decimal(charge) for _, _, charge in inv.lines), Decimal("0"))
    else:
        subtotal = sum((Decimal(qty) * Decimal(price) for _, qty, price in inv.lines), Decimal("0"))
    subtotal = subtotal.quantize(CENT)
    tax = (subtotal * TAX_RATE[inv.layout]).quantize(CENT, rounding=ROUND_HALF_UP)
    return subtotal, tax, subtotal + tax


def new_canvas(target, title: str) -> canvas.Canvas:
    # invariant=1 keeps the file identical from run to run (fixed dates and IDs).
    c = canvas.Canvas(target, pagesize=letter, invariant=1)
    c.setAuthor(AUTHOR)
    c.setCreator("AZG Engineering sample generator")
    c.setTitle(title)
    c.setSubject("Made-up sample invoice for a demo")
    return c


def footer(c: canvas.Canvas) -> None:
    c.setFillColor(colors.HexColor("#8A94A0"))
    c.setFont("Helvetica", 8)
    c.drawCentredString(PAGE_W / 2, 40, SAMPLE_NOTE)


def bill_to(c: canvas.Canvas, x: float, y: float, heading: str) -> None:
    c.setFillColor(colors.HexColor("#555555"))
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x, y, heading)
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    for i, line in enumerate(("Example Client Co.", "100 Example Street", "Sampleville, ZZ 00000"), start=1):
        c.drawString(x, y - 14 * i, line)


def draw_northwind(c: canvas.Canvas, inv: Invoice) -> None:
    navy = colors.HexColor("#1F3A5F")
    c.setFillColor(navy)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(LEFT, 730, "Northwind Supply Co.")
    c.setFillColor(colors.HexColor("#555555"))
    c.setFont("Helvetica", 9)
    c.drawString(LEFT, 715, "400 Sample Road, Suite 12")
    c.drawString(LEFT, 703, "Anytown, ZZ 00000")
    c.setFillColor(colors.HexColor("#9AA5B1"))
    c.setFont("Helvetica-Bold", 26)
    c.drawRightString(RIGHT, 724, "INVOICE")

    bill_to(c, LEFT, 668, "Bill To:")

    meta = (
        ("Invoice No:", inv.number),
        ("Invoice Date:", inv.issued.strftime("%m/%d/%Y")),
        ("Due Date:", (inv.issued + timedelta(days=30)).strftime("%m/%d/%Y")),
        ("PO Number:", "PO-" + inv.number[-4:]),
    )
    c.setFillColor(colors.black)
    y = 668
    for label, value in meta:
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(RIGHT - 80, y, label)
        c.setFont("Helvetica", 10)
        c.drawRightString(RIGHT, y, value)
        y -= 14

    top = 590
    c.setFillColor(navy)
    c.rect(LEFT, top - 6, RIGHT - LEFT, 20, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(LEFT + 6, top, "Description")
    c.drawRightString(370, top, "Qty")
    c.drawRightString(460, top, "Unit Price")
    c.drawRightString(RIGHT - 6, top, "Amount")

    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    y = top - 22
    for description, qty, price in inv.lines:
        c.drawString(LEFT + 6, y, description)
        c.drawRightString(370, y, str(qty))
        c.drawRightString(460, y, "$" + money(Decimal(price)))
        c.drawRightString(RIGHT - 6, y, "$" + money(Decimal(qty) * Decimal(price)))
        y -= 18
    c.setStrokeColor(colors.HexColor("#C5CCD3"))
    c.line(LEFT, y + 8, RIGHT, y + 8)

    y -= 12
    if inv.omit_total:
        c.setFont("Helvetica-Oblique", 10)
        c.setFillColor(colors.HexColor("#555555"))
        c.drawRightString(RIGHT - 6, y, "Continued on next page")
        return

    subtotal, tax, total = amounts(inv)
    for label, value in (("Subtotal", subtotal), ("Sales Tax (6.0%)", tax)):
        c.drawRightString(460, y, label)
        c.drawRightString(RIGHT - 6, y, "$" + money(value))
        y -= 16
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(navy)
    c.drawRightString(460, y - 4, "TOTAL DUE")
    c.drawRightString(RIGHT - 6, y - 4, "$" + money(total))
    c.setFillColor(colors.HexColor("#555555"))
    c.setFont("Helvetica", 9)
    c.drawString(LEFT, y - 40, "Thank you for your business. Payment terms: net 30 days.")


def draw_contoso(c: canvas.Canvas, inv: Invoice) -> None:
    teal = colors.HexColor("#0B5563")
    c.setFillColor(teal)
    c.rect(0, PAGE_H - 84, PAGE_W, 84, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(LEFT, PAGE_H - 52, "CONTOSO FREIGHT")
    c.setFont("Helvetica", 12)
    c.drawRightString(RIGHT, PAGE_H - 52, "Freight Invoice")

    bill_to(c, LEFT, 672, "BILLED TO")

    box_left = 350
    c.setFillColor(colors.HexColor("#EEF4F5"))
    c.rect(box_left, 610, RIGHT - box_left, 78, stroke=0, fill=1)
    meta = (
        ("Invoice #", inv.number),
        ("Date", f"{inv.issued.day} {inv.issued.strftime('%b %Y')}"),
        ("Terms", "Net 30"),
        ("Shipment ref", "SH-4" + inv.number[-4:]),
    )
    c.setFillColor(colors.black)
    y = 672
    for label, value in meta:
        c.setFont("Helvetica", 10)
        c.drawString(box_left + 10, y, label)
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(RIGHT - 10, y, value)
        y -= 16

    delivered = inv.issued - timedelta(days=2)
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.HexColor("#555555"))
    c.drawString(LEFT, 590, f"Delivered {delivered.day} {delivered.strftime('%b %Y')}, signed for at dock.")

    top = 560
    c.setFont("Helvetica-Bold", 8)
    c.drawString(LEFT, top, "SERVICE")
    c.drawString(270, top, "DETAILS")
    c.drawRightString(RIGHT, top, "CHARGE (USD)")
    c.setStrokeColor(teal)
    c.setLineWidth(1.2)
    c.line(LEFT, top - 6, RIGHT, top - 6)

    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    c.setStrokeColor(colors.HexColor("#C5CCD3"))
    c.setLineWidth(0.5)
    y = top - 24
    for service, details, charge in inv.lines:
        c.drawString(LEFT, y, service)
        c.drawString(270, y, details)
        c.drawRightString(RIGHT, y, money(Decimal(charge)))
        c.line(LEFT, y - 7, RIGHT, y - 7)
        y -= 22

    _, _, total = amounts(inv)
    y -= 14
    c.setFillColor(teal)
    c.rect(box_left, y - 12, RIGHT - box_left, 32, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(box_left + 10, y, "Amount Due")
    c.drawRightString(RIGHT - 10, y, "USD " + money(total))


def draw_fabrikam(c: canvas.Canvas, inv: Invoice) -> None:
    plum = colors.HexColor("#5B2A6E")
    c.setFillColor(colors.HexColor("#333333"))
    c.setFont("Helvetica", 30)
    c.drawString(LEFT, 722, "Invoice")
    c.setFillColor(plum)
    c.setFont("Helvetica-Bold", 14)
    c.drawRightString(RIGHT, 734, "Fabrikam Office Goods")
    c.setFillColor(colors.HexColor("#555555"))
    c.setFont("Helvetica", 9)
    c.drawRightString(RIGHT, 720, "22 Placeholder Avenue")
    c.drawRightString(RIGHT, 708, "Testburg, ZZ 00000")

    c.setStrokeColor(plum)
    c.setLineWidth(1.5)
    c.line(LEFT, 690, RIGHT, 690)

    columns = (
        (LEFT, "Invoice number", inv.number),
        (LEFT + 180, "Date of issue", inv.issued.isoformat()),
        (LEFT + 360, "Customer ref", "C-2" + inv.number[-3:]),
    )
    for x, label, value in columns:
        c.setFillColor(colors.HexColor("#777777"))
        c.setFont("Helvetica", 8)
        c.drawString(x, 672, label)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x, 656, value)

    bill_to(c, LEFT, 622, "Sold to")

    top = 540
    c.setFillColor(colors.HexColor("#555555"))
    c.setFont("Helvetica-Bold", 9)
    c.drawString(LEFT + 6, top, "Item")
    c.drawRightString(380, top, "Qty")
    c.drawRightString(460, top, "Price")
    c.drawRightString(RIGHT - 6, top, "Line total")

    c.setFont("Helvetica", 10)
    y = top - 22
    for i, (description, qty, price) in enumerate(inv.lines):
        if i % 2 == 0:
            c.setFillColor(colors.HexColor("#F3EEF5"))
            c.rect(LEFT, y - 6, RIGHT - LEFT, 20, stroke=0, fill=1)
        c.setFillColor(colors.black)
        c.drawString(LEFT + 6, y, description)
        c.drawRightString(380, y, str(qty))
        c.drawRightString(460, y, money(Decimal(price)))
        c.drawRightString(RIGHT - 6, y, money(Decimal(qty) * Decimal(price)))
        y -= 20

    subtotal, tax, total = amounts(inv)
    y -= 10
    for label, value in (("Net amount", subtotal), ("Tax (5.5%)", tax)):
        c.drawRightString(460, y, label)
        c.drawRightString(RIGHT - 6, y, "$" + money(value))
        y -= 16
    c.setStrokeColor(plum)
    c.setLineWidth(1)
    c.line(360, y + 8, RIGHT, y + 8)
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(460, y - 8, "Grand Total")
    c.drawRightString(RIGHT - 6, y - 8, "$" + money(total))


LAYOUTS = {"northwind": draw_northwind, "contoso": draw_contoso, "fabrikam": draw_fabrikam}


def render_pdf(inv: Invoice) -> bytes:
    buffer = io.BytesIO()
    c = new_canvas(buffer, f"Invoice {inv.number} (demo sample)")
    LAYOUTS[inv.layout](c, inv)
    footer(c)
    c.showPage()
    c.save()
    return buffer.getvalue()


def as_scanned_pdf(pdf_bytes: bytes) -> bytes:
    """Turn a PDF page into a picture-only PDF, the way a scanner would."""
    import pdfplumber
    from PIL import Image

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0].to_image(resolution=130).original.convert("L")

    # A little skew, grey paper and some specks, so it looks like a scan.
    page = page.rotate(0.8, resample=Image.BICUBIC, fillcolor=255)
    page = page.point(lambda v: int(v * 0.90 + 14))
    rng = random.Random(7)
    pixels = page.load()
    for _ in range(1800):
        x, y = rng.randrange(page.width), rng.randrange(page.height)
        pixels[x, y] = rng.randint(120, 200)

    buffer = io.BytesIO()
    c = new_canvas(buffer, "Scanned invoice (demo sample)")
    c.drawImage(ImageReader(page), 0, 0, width=PAGE_W, height=PAGE_H)
    c.showPage()
    c.save()
    return buffer.getvalue()


def generate(out_dir: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for inv in INVOICES:
        data = render_pdf(inv)
        if inv.as_scan:
            data = as_scanned_pdf(data)
        path = out_dir / inv.file_name
        path.write_bytes(data)
        written.append(path)
    return written


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    out_dir = Path(args[0]) if args else Path(__file__).parent / "samples"
    written = generate(out_dir)
    print(f"Wrote {len(written)} sample invoices to {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
