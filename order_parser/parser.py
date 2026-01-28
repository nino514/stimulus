#!/usr/bin/env python3
"""
Order Form Parser
Extracts structured data from order forms (PDFs and images) using Claude Vision API.
Outputs results to orders_export.csv
"""

import os
import sys
import csv
import json
import base64
import shutil
import time
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from PIL import Image
import anthropic

# Conditional import for PDF support
try:
    from pdf2image import convert_from_path
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

# Configuration
SCRIPT_DIR = Path(__file__).parent
INPUT_DIR = SCRIPT_DIR / "input"
PROCESSED_DIR = SCRIPT_DIR / "processed"
FAILED_DIR = SCRIPT_DIR / "failed"
OUTPUT_FILE = SCRIPT_DIR / "orders_export.csv"

SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif"}
SUPPORTED_PDF_EXTENSIONS = {".pdf"}
PDF_DPI = 300
MAX_RETRIES = 3
RETRY_DELAY_BASE = 2  # seconds

# CSV column headers
CSV_HEADERS = [
    "source_file",
    "page_number",
    "quantity",
    "sku",
    "description",
    "unit_price",
    "size",
    "confidence"
]

EXTRACTION_PROMPT = """You are an order form data extraction assistant. Analyze this image of an order form and extract all order line items into structured data.

For each line item in the order table, extract:
- quantity: The number of units ordered (integer)
- sku: The product code, item number, or SKU
- description: The product name or description
- unit_price: The price per unit (numeric, without currency symbols)
- size: The size if specified (e.g., S, M, L, XL, or dimensions), otherwise null

Rules:
1. Skip header rows - only extract actual data rows
2. If a field is unclear or missing, use null
3. If the image is blurry or partially unreadable, extract what you can and note issues
4. Return ONLY valid JSON, no markdown formatting or code blocks

Return a JSON object in this exact format:
{
  "line_items": [
    {"quantity": 10, "sku": "ABC-123", "description": "Product Name", "unit_price": 19.99, "size": "Large"}
  ],
  "confidence": "high",
  "notes": "Any issues or observations about the document"
}

If no order table is found, return:
{"line_items": [], "confidence": "low", "notes": "No order table detected"}"""


def load_api_key() -> str:
    """Load the Anthropic API key from environment."""
    load_dotenv(SCRIPT_DIR / ".env")
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: ANTHROPIC_API_KEY not found.")
        print("Please create a .env file with: ANTHROPIC_API_KEY=your-key-here")
        sys.exit(1)
    return api_key


def ensure_directories():
    """Create necessary directories if they don't exist."""
    INPUT_DIR.mkdir(exist_ok=True)
    PROCESSED_DIR.mkdir(exist_ok=True)
    FAILED_DIR.mkdir(exist_ok=True)


def get_input_files() -> list[Path]:
    """Get all supported files from the input directory."""
    files = []
    all_extensions = SUPPORTED_IMAGE_EXTENSIONS | SUPPORTED_PDF_EXTENSIONS

    for file_path in INPUT_DIR.iterdir():
        if file_path.is_file() and not file_path.name.startswith("."):
            if file_path.suffix.lower() in all_extensions:
                files.append(file_path)

    return sorted(files)


def image_to_base64(image_path: Path) -> tuple[str, str]:
    """Convert an image file to base64 and return with media type."""
    suffix = image_path.suffix.lower()
    media_type_map = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".gif": "image/gif"
    }
    media_type = media_type_map.get(suffix, "image/png")

    with open(image_path, "rb") as f:
        image_data = base64.standard_b64encode(f.read()).decode("utf-8")

    return image_data, media_type


def pil_image_to_base64(image: Image.Image) -> tuple[str, str]:
    """Convert a PIL Image to base64 PNG."""
    import io
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    image_data = base64.standard_b64encode(buffer.getvalue()).decode("utf-8")
    return image_data, "image/png"


def pdf_to_images(pdf_path: Path) -> list[Image.Image]:
    """Convert a PDF to a list of PIL Images (one per page)."""
    if not PDF_SUPPORT:
        raise RuntimeError(
            "PDF support requires pdf2image and poppler. "
            "Install poppler: brew install poppler (macOS) or apt-get install poppler-utils (Linux)"
        )
    return convert_from_path(pdf_path, dpi=PDF_DPI)


def call_vision_api(client: anthropic.Anthropic, image_data: str, media_type: str) -> Optional[dict]:
    """Send image to Claude Vision API and return parsed response."""
    for attempt in range(MAX_RETRIES):
        try:
            response = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=4096,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": media_type,
                                    "data": image_data
                                }
                            },
                            {
                                "type": "text",
                                "text": EXTRACTION_PROMPT
                            }
                        ]
                    }
                ]
            )

            # Extract text response
            response_text = response.content[0].text

            # Clean up response if it contains markdown code blocks
            if response_text.startswith("```"):
                lines = response_text.split("\n")
                # Remove first and last lines (code block markers)
                lines = [l for l in lines if not l.startswith("```")]
                response_text = "\n".join(lines)

            # Parse JSON
            return json.loads(response_text)

        except anthropic.RateLimitError:
            if attempt < MAX_RETRIES - 1:
                delay = RETRY_DELAY_BASE * (2 ** attempt)
                print(f"  Rate limited, retrying in {delay}s...")
                time.sleep(delay)
            else:
                raise
        except json.JSONDecodeError as e:
            print(f"  Warning: Failed to parse API response as JSON: {e}")
            print(f"  Response was: {response_text[:200]}...")
            return None
        except Exception as e:
            if attempt < MAX_RETRIES - 1:
                delay = RETRY_DELAY_BASE * (2 ** attempt)
                print(f"  Error: {e}, retrying in {delay}s...")
                time.sleep(delay)
            else:
                raise

    return None


def process_image(client: anthropic.Anthropic, image_data: str, media_type: str,
                  source_file: str, page_number: int) -> list[dict]:
    """Process a single image and return extracted rows."""
    result = call_vision_api(client, image_data, media_type)

    if not result:
        return []

    rows = []
    confidence = result.get("confidence", "low")
    line_items = result.get("line_items", [])

    for item in line_items:
        # Clean up price (remove currency symbols if present)
        unit_price = item.get("unit_price")
        if isinstance(unit_price, str):
            unit_price = unit_price.replace("$", "").replace("€", "").replace("£", "").strip()
            try:
                unit_price = float(unit_price)
            except ValueError:
                unit_price = None

        row = {
            "source_file": source_file,
            "page_number": page_number,
            "quantity": item.get("quantity"),
            "sku": item.get("sku"),
            "description": item.get("description"),
            "unit_price": unit_price,
            "size": item.get("size"),
            "confidence": confidence
        }
        rows.append(row)

    if result.get("notes"):
        print(f"  Note: {result.get('notes')}")

    return rows


def process_file(client: anthropic.Anthropic, file_path: Path) -> tuple[list[dict], bool]:
    """Process a single file (PDF or image) and return extracted rows."""
    print(f"Processing: {file_path.name}")

    rows = []
    suffix = file_path.suffix.lower()

    try:
        if suffix in SUPPORTED_PDF_EXTENSIONS:
            # PDF: convert to images and process each page
            images = pdf_to_images(file_path)
            print(f"  PDF has {len(images)} page(s)")

            for page_num, image in enumerate(images, start=1):
                print(f"  Processing page {page_num}...")
                image_data, media_type = pil_image_to_base64(image)
                page_rows = process_image(client, image_data, media_type,
                                         file_path.name, page_num)
                rows.extend(page_rows)
                print(f"    Extracted {len(page_rows)} line items")

        elif suffix in SUPPORTED_IMAGE_EXTENSIONS:
            # Image: process directly
            image_data, media_type = image_to_base64(file_path)
            rows = process_image(client, image_data, media_type, file_path.name, 1)
            print(f"  Extracted {len(rows)} line items")

        else:
            print(f"  Skipping unsupported file type: {suffix}")
            return [], False

        return rows, True

    except Exception as e:
        print(f"  ERROR: {e}")
        return [], False


def write_csv(rows: list[dict], append: bool = False):
    """Write rows to the output CSV file."""
    mode = "a" if append else "w"
    file_exists = OUTPUT_FILE.exists() and append

    with open(OUTPUT_FILE, mode, newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)

        if not file_exists:
            writer.writeheader()

        writer.writerows(rows)


def move_file(file_path: Path, destination_dir: Path):
    """Move a file to the destination directory."""
    dest_path = destination_dir / file_path.name

    # Handle name conflicts
    counter = 1
    while dest_path.exists():
        stem = file_path.stem
        suffix = file_path.suffix
        dest_path = destination_dir / f"{stem}_{counter}{suffix}"
        counter += 1

    shutil.move(str(file_path), str(dest_path))


def main():
    """Main entry point."""
    print("=" * 50)
    print("ORDER FORM PARSER")
    print("=" * 50)
    print()

    # Setup
    api_key = load_api_key()
    client = anthropic.Anthropic(api_key=api_key)
    ensure_directories()

    # Get files to process
    files = get_input_files()

    if not files:
        print(f"No files found in {INPUT_DIR}/")
        print("Place PDF or image files in the input/ folder and run again.")
        return

    print(f"Found {len(files)} file(s) to process")
    print()

    # Process files
    all_rows = []
    success_count = 0
    fail_count = 0

    for file_path in files:
        rows, success = process_file(client, file_path)

        if success:
            all_rows.extend(rows)
            move_file(file_path, PROCESSED_DIR)
            success_count += 1
        else:
            move_file(file_path, FAILED_DIR)
            fail_count += 1

        print()

    # Write output
    if all_rows:
        write_csv(all_rows, append=False)
        print(f"Output written to: {OUTPUT_FILE}")
    else:
        print("No data extracted.")

    # Summary
    print()
    print("=" * 50)
    print("SUMMARY")
    print("=" * 50)
    print(f"Files processed successfully: {success_count}")
    print(f"Files failed: {fail_count}")
    print(f"Total line items extracted: {len(all_rows)}")

    if fail_count > 0:
        print(f"\nFailed files moved to: {FAILED_DIR}/")

    if success_count > 0:
        print(f"Processed files moved to: {PROCESSED_DIR}/")


if __name__ == "__main__":
    main()
