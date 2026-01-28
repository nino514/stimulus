#!/usr/bin/env python3
"""
Extract order data from Stimulus Athletic Excel invoice files.

Usage:
    python extract_orders.py                         # Batch mode: process all .xlsx in current folder
    python extract_orders.py <input.xlsx> <output.csv>  # Single file mode

Examples:
    python extract_orders.py                         # Creates all_orders_extracted.csv
    python extract_orders.py invoice.xlsx orders.csv # Process single file
"""

import sys
import csv
from pathlib import Path
from openpyxl import load_workbook


# CSV column order
FIELDNAMES = [
    "order_id", "date_of_order", "customer_name", "customer_email",
    "phone", "billing_address", "delivery_address", "product_description",
    "personalization_name", "gender", "jersey_size", "shorts_size",
    "socks_size", "item_number", "quantity", "unit_price", "total_price"
]


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
    """
    Parse the header row to find column indices for each field.
    Handles merged 'SIZE' header with sub-headers (JERSEY, SHORTS, SOCKS) in row below.
    """
    column_map = {}

    # Mappings for main header row
    header_mappings = {
        "GENDER": "gender",
        "NAME": "personalization_name",
        "INITIALS OR NUMBER": "item_number",
        "INITIALS/NUMBER": "item_number",
        "NUMBER": "item_number",
        "QUANTITY": "quantity",
        "QTY": "quantity",
        "UNIT PRICE": "unit_price",
        "TOTAL PRICE": "total_price",
    }

    # Mappings for sub-header row (under SIZE)
    size_sub_mappings = {
        "JERSEY": "jersey_size",
        "SHORTS": "shorts_size",
        "SOCKS": "socks_size",
    }

    has_size_header = False

    # First pass: scan main header row
    for cell in sheet[header_row_idx]:
        if cell.value:
            cell_text = str(cell.value).upper().strip()

            # Check if this is a merged SIZE header
            if "SIZE" in cell_text and cell_text not in ("JERSEY SIZE", "SHORTS SIZE", "SOCKS SIZE"):
                has_size_header = True
                continue

            # Check against main mappings
            for header_text, field_name in header_mappings.items():
                if header_text in cell_text:
                    if field_name not in column_map:
                        column_map[field_name] = cell.column - 1  # 0-indexed
                    break

            # Also check for size columns in main header (in case they're not merged)
            for header_text, field_name in size_sub_mappings.items():
                if header_text in cell_text:
                    if field_name not in column_map:
                        column_map[field_name] = cell.column - 1
                    break

    # Second pass: if SIZE header found, scan sub-header row below
    if has_size_header:
        sub_header_row_idx = header_row_idx + 1
        for cell in sheet[sub_header_row_idx]:
            if cell.value:
                cell_text = str(cell.value).upper().strip()
                for header_text, field_name in size_sub_mappings.items():
                    if header_text in cell_text:
                        if field_name not in column_map:
                            column_map[field_name] = cell.column - 1
                        break

    return column_map, has_size_header


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


def extract_personalization_data(sheet, section_row, column_map, has_sub_header):
    """Extract all data rows from a personalization section."""
    rows = []
    header_row = section_row + 1  # Header is right after PERSONALIZATION

    # If there's a sub-header row (for SIZE columns), data starts 2 rows after header
    if has_sub_header:
        data_start = header_row + 2
    else:
        data_start = header_row + 1

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
        if "GENDER" in row_text and ("SIZE" in row_text or "NAME" in row_text):
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


def extract_from_file(input_path, verbose=True):
    """Extract all records from a single Excel file. Returns list of records."""
    if verbose:
        print(f"\nProcessing: {input_path.name}")
        print("-" * 50)

    workbook = load_workbook(input_path, data_only=True)
    sheet = workbook.active

    # Step 1: Extract order metadata
    order_id = find_cell_value(sheet, "ORDER ID")
    date_of_order = find_cell_value(sheet, "DATE OF ORDER")
    customer_name = find_cell_value(sheet, "RECEIVER'S NAME")
    customer_email = find_cell_value(sheet, "EMAIL ADDRESS")
    phone = find_cell_value(sheet, "TELEPHONE")
    billing_address = find_cell_value(sheet, "BILLING ADDRESS")
    delivery_address = find_cell_value(sheet, "DELIVERY ADDRESS")

    if verbose:
        print(f"  Order ID: {order_id}")
        print(f"  Customer: {customer_name}")

    # Step 2: Find all personalization sections
    personalization_rows = find_personalization_sections(sheet)

    # Step 3: Extract data from each section
    all_records = []

    for section_row in personalization_rows:
        # Find product description above this section
        product_desc = find_product_description_above(sheet, section_row)

        # Parse header row to get column positions
        header_row = section_row + 1
        column_map, has_sub_header = parse_header_row(sheet, header_row)

        if not column_map:
            if verbose:
                print(f"  Warning: Could not find column headers at row {header_row}")
            continue

        # Extract personalization rows
        rows = extract_personalization_data(sheet, section_row, column_map, has_sub_header)

        # Add order metadata to each row
        for row in rows:
            record = {
                "order_id": order_id,
                "date_of_order": date_of_order,
                "customer_name": customer_name,
                "customer_email": customer_email,
                "phone": phone,
                "billing_address": billing_address,
                "delivery_address": delivery_address,
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

    if verbose:
        print(f"  Extracted: {len(all_records)} rows")

    return all_records


def write_csv(records, output_path):
    """Write records to CSV file."""
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)


def process_single_file(input_path, output_path):
    """Process a single Excel file and write to CSV."""
    records = extract_from_file(input_path, verbose=True)
    write_csv(records, output_path)
    print(f"\nWrote {len(records)} records to: {output_path}")


def process_batch():
    """Process all .xlsx files in current directory and combine into one CSV."""
    current_dir = Path(".")
    xlsx_files = sorted(current_dir.glob("*.xlsx"))

    if not xlsx_files:
        print("No .xlsx files found in current directory.")
        print("Place your Excel order files here and run again.")
        sys.exit(1)

    print("=" * 60)
    print("BATCH PROCESSING MODE")
    print("=" * 60)
    print(f"Found {len(xlsx_files)} Excel file(s) to process:\n")

    for f in xlsx_files:
        print(f"  - {f.name}")

    all_records = []
    files_processed = 0
    files_failed = []

    for xlsx_file in xlsx_files:
        try:
            records = extract_from_file(xlsx_file, verbose=True)
            all_records.extend(records)
            files_processed += 1
        except Exception as e:
            print(f"  ERROR: {e}")
            files_failed.append((xlsx_file.name, str(e)))

    # Write combined CSV
    output_path = Path("all_orders_extracted.csv")
    write_csv(all_records, output_path)

    # Print summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"  Total files processed: {files_processed}")
    if files_failed:
        print(f"  Files with errors: {len(files_failed)}")
        for name, error in files_failed:
            print(f"    - {name}: {error}")
    print(f"  Total rows extracted: {len(all_records)}")
    print(f"  Output saved to: {output_path.absolute()}")
    print("=" * 60)


def main():
    if len(sys.argv) == 1:
        # No arguments: batch mode
        process_batch()
    elif len(sys.argv) == 3:
        # Two arguments: single file mode
        input_path = Path(sys.argv[1])
        output_path = Path(sys.argv[2])

        if not input_path.exists():
            print(f"Error: Input file not found: {input_path}")
            sys.exit(1)

        process_single_file(input_path, output_path)
    else:
        print("Usage:")
        print("  python extract_orders.py                           # Batch: process all .xlsx files")
        print("  python extract_orders.py <input.xlsx> <output.csv> # Single file mode")
        print()
        print("Examples:")
        print("  python extract_orders.py                           # Creates all_orders_extracted.csv")
        print("  python extract_orders.py invoice.xlsx orders.csv   # Process one file")
        sys.exit(1)


if __name__ == "__main__":
    main()
