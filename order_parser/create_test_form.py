#!/usr/bin/env python3
"""
Creates a sample order form image for testing the parser.
Run this script to generate test_order_form.png in the input/ folder.
"""

from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
OUTPUT_PATH = SCRIPT_DIR / "input" / "test_order_form.png"


def create_test_order_form():
    """Generate a simple order form image with tabular data."""

    # Create image
    width, height = 800, 600
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)

    # Use default font (no external font file needed)
    try:
        font_large = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 24)
        font_medium = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 16)
        font_small = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 14)
    except OSError:
        # Fallback to default font
        font_large = ImageFont.load_default()
        font_medium = ImageFont.load_default()
        font_small = ImageFont.load_default()

    # Header
    draw.text((50, 30), "ACME SUPPLIES CO.", fill="black", font=font_large)
    draw.text((50, 65), "Purchase Order #PO-2024-001", fill="gray", font=font_medium)
    draw.text((50, 90), "Date: January 15, 2024", fill="gray", font=font_small)

    # Draw horizontal line
    draw.line([(50, 120), (750, 120)], fill="black", width=2)

    # Table header
    y_start = 140
    row_height = 35
    columns = [
        (50, "Qty"),
        (120, "SKU"),
        (250, "Description"),
        (500, "Size"),
        (600, "Unit Price")
    ]

    # Header row
    for x, label in columns:
        draw.text((x, y_start), label, fill="black", font=font_medium)

    # Header underline
    draw.line([(50, y_start + 25), (750, y_start + 25)], fill="black", width=1)

    # Data rows
    data = [
        ("25", "TEE-BLU-001", "Cotton T-Shirt Blue", "Large", "$15.99"),
        ("50", "TEE-RED-002", "Cotton T-Shirt Red", "Medium", "$15.99"),
        ("10", "HOOD-BLK-003", "Pullover Hoodie Black", "XL", "$45.00"),
        ("100", "CAP-WHT-004", "Baseball Cap White", "One Size", "$12.50"),
        ("30", "POLO-NAV-005", "Polo Shirt Navy", "Small", "$28.00"),
    ]

    y = y_start + 40
    for qty, sku, desc, size, price in data:
        draw.text((50, y), qty, fill="black", font=font_small)
        draw.text((120, y), sku, fill="black", font=font_small)
        draw.text((250, y), desc, fill="black", font=font_small)
        draw.text((500, y), size, fill="black", font=font_small)
        draw.text((600, y), price, fill="black", font=font_small)
        y += row_height

    # Footer line
    draw.line([(50, y + 10), (750, y + 10)], fill="black", width=1)

    # Total
    draw.text((500, y + 25), "Subtotal:", fill="black", font=font_medium)
    draw.text((600, y + 25), "$2,847.50", fill="black", font=font_medium)

    # Save
    OUTPUT_PATH.parent.mkdir(exist_ok=True)
    image.save(OUTPUT_PATH, "PNG")
    print(f"Test order form created: {OUTPUT_PATH}")


if __name__ == "__main__":
    create_test_order_form()
