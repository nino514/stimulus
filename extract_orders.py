#!/usr/bin/env python3
"""
Extract order data from Stimulus Athletic Excel invoice files.

Usage:
    python extract_orders.py <input.xlsx> <output.csv>

Example:
    python extract_orders.py sample_orders/invoice.xlsx orders_extracted.csv
"""

import sys
import csv
from pathlib import Path
from openpyxl import load_workbook


def find_cell_value(sheet, search_text):
    """Find a cell containing search_text and return the value from the next non-empty cell in that row."""
    search_upper = search_text.upper()
    for row in sheet.iter_rows():
        for i, cell in enumerate(row):
            if cell.value and search_upper in str(cell.value).upper():
                # Look for next non-empty cell in this row
                for next_cell in row[i + 1:]:
                    if next_cell.value and str(next_cell.value).strip():
                        return str(next_cell.value).strip()
    return ""


def find_personalization_sections(sheet):
    """Find all rows that contain 'PERSONALIZATION' header."""
    sections = []
    for row_idx, row in enumerate(sheet.iter_rows(min_row=1), start=1):
        for cell in row:
            if cell.value and "PERSONALIZATION" in str(cell.value).upper():
                sections.append(row_idx)
                break
    return sections


def find_product_description_above(sheet, start_row):
    """Look backwards from start_row to find the nearest 'PRODUCTS INCLUDED' value."""
    for row_idx in range(start_row - 1, 0, -1):
        for cell in sheet[row_idx]:
            if cell.value and "PRODUCTS INCLUDED" in str(cell.value).upper():
                # Get value from next non-empty cell in this row
                row_cells = list(sheet[row_idx])
                cell_idx = row_cells.index(cell)
                for next_cell in row_cells[cell_idx + 1:]:
                    if next_cell.value and str(next_cell.value).strip():
                        return str(next_cell.value).strip()
    return ""


def parse_header_row(sheet, header_row_idx):
    """Parse the header row to find column indices for each field."""
    column_map = {}
    header_mappings = {
        "GENDER": "gender",
        "JERSEY": "jersey_size",
        "SHORTS": "shorts_size",
        "SOCKS": "socks_size",
        "NAME": "personalization_name",
        "INITIALS OR NUMBER": "item_number",
        "INITIALS/NUMBER": "item_number",
        "NUMBER": "item_number",
        "QUANTITY": "quantity",
        "QTY": "quantity",
        "UNIT PRICE": "unit_price",
        "TOTAL PRICE": "total_price",
    }

    for cell in sheet[header_row_idx]:
        if cell.value:
            cell_text = str(cell.value).upper().strip()
            for header_text, field_name in header_mappings.items():
                if header_text in cell_text:
                    # Don't overwrite if already found (prioritize exact matches)
                    if field_name not in column_map:
                        column_map[field_name] = cell.column - 1  # 0-indexed
                    break

    return column_map


def clean_price(value):
    """Remove $ from price and return as string."""
    if value is None:
        return ""
    text = str(value).strip()
    if text.startswith("$"):
        text = text[1:]
    return text


def clean_value(value):
    """Clean a cell value for CSV output."""
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in ("none", "n/a", "na", "-"):
        return ""
    return text


def is_empty_row(row_values):
    """Check if all values in a row are empty."""
    return all(v is None or str(v).strip() == "" for v in row_values)


def is_skip_row(row_values):
    """Check if this row should be skipped (totals, subtotals)."""
    row_text = " ".join(str(v) for v in row_values if v).upper()
    skip_keywords = ["TOTAL QTY", "SUBTOTAL", "TOTAL PRICE"]
    return any(kw in row_text for kw in skip_keywords)


def extract_personalization_data(sheet, section_row, column_map):
    """Extract all data rows from a personalization section."""
    rows = []
    header_row = section_row + 1  # Header is right after PERSONALIZATION
    data_start = header_row + 1   # Data starts after header

    max_col = max(column_map.values()) + 1 if column_map else 15

    for row_idx in range(data_start, sheet.max_row + 1):
        row_cells = [sheet.cell(row=row_idx, column=c + 1).value for c in range(max_col + 5)]

        # Stop if empty row
        if is_empty_row(row_cells):
            break

        # Stop if we hit next product section
        row_text = " ".join(str(v) for v in row_cells if v).upper()
        if "PRODUCT INFORMATION" in row_text:
            break

        # Skip total/subtotal rows
        if is_skip_row(row_cells):
            continue

        # Skip if this looks like a header row (contains GENDER text)
        if "GENDER" in row_text and "SIZE" in row_text:
            continue

        # Extract data using column map
        row_data = {}
        for field_name, col_idx in column_map.items():
            value = sheet.cell(row=row_idx, column=col_idx + 1).value
            if field_name in ("unit_price", "total_price"):
                row_data[field_name] = clean_price(value)
            else:
                row_data[field_name] = clean_value(value)

        # Only add row if it has some identifying data
        if row_data.get("gender") or row_data.get("personalization_name") or row_data.get("item_number"):
            rows.append(row_data)

    return rows


def extract_orders(input_path, output_path):
    """Main extraction function."""
    print(f"Loading: {input_path}")
    workbook = load_workbook(input_path, data_only=True)
    sheet = workbook.active

    # Step 1: Extract order metadata
    order_id = find_cell_value(sheet, "ORDER ID")
    customer_name = find_cell_value(sheet, "RECEIVER'S NAME")
    customer_email = find_cell_value(sheet, "EMAIL ADDRESS")

    print(f"Order ID: {order_id}")
    print(f"Customer: {customer_name}")
    print(f"Email: {customer_email}")

    # Step 2: Find all personalization sections
    personalization_rows = find_personalization_sections(sheet)
    print(f"Found {len(personalization_rows)} personalization section(s)")

    # Step 3: Extract data from each section
    all_records = []

    for section_row in personalization_rows:
        # Find product description above this section
        product_desc = find_product_description_above(sheet, section_row)
        print(f"  Section at row {section_row}: {product_desc[:50]}...")

        # Parse header row to get column positions
        header_row = section_row + 1
        column_map = parse_header_row(sheet, header_row)

        if not column_map:
            print(f"  Warning: Could not find column headers at row {header_row}")
            continue

        # Extract personalization rows
        rows = extract_personalization_data(sheet, section_row, column_map)
        print(f"  Extracted {len(rows)} personalization row(s)")

        # Add order metadata to each row
        for row in rows:
            record = {
                "order_id": order_id,
                "customer_name": customer_name,
                "customer_email": customer_email,
                "product_description": product_desc,
                "personalization_name": row.get("personalization_name", ""),
                "gender": row.get("gender", ""),
                "jersey_size": row.get("jersey_size", ""),
                "shorts_size": row.get("shorts_size", ""),
                "socks_size": row.get("socks_size", ""),
                "item_number": row.get("item_number", ""),
                "quantity": row.get("quantity", ""),
                "unit_price": row.get("unit_price", ""),
                "total_price": row.get("total_price", ""),
            }
            all_records.append(record)

    # Step 4: Write CSV output
    fieldnames = [
        "order_id", "customer_name", "customer_email", "product_description",
        "personalization_name", "gender", "jersey_size", "shorts_size",
        "socks_size", "item_number", "quantity", "unit_price", "total_price"
    ]

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_records)

    print(f"\nWrote {len(all_records)} records to: {output_path}")
    return all_records


def main():
    if len(sys.argv) != 3:
        print("Usage: python extract_orders.py <input.xlsx> <output.csv>")
        print("Example: python extract_orders.py invoice.xlsx orders.csv")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        sys.exit(1)

    extract_orders(input_path, output_path)


if __name__ == "__main__":
    main()
