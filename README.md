# Stimulus Order Extractor

Extract order data from Stimulus Athletic Excel invoice files into clean CSV format.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Batch Mode (Recommended)

Process all `.xlsx` files in the current folder and combine into one CSV:

```bash
python extract_orders.py
```

This will:
1. Find all `.xlsx` files in the current directory
2. Extract data from each file
3. Combine all data into `all_orders_extracted.csv`
4. Print a summary of files processed and rows extracted

### Single File Mode

Process a specific file:

```bash
python extract_orders.py <input.xlsx> <output.csv>
```

Example:
```bash
python extract_orders.py WILLMAR-CARDINALS_ORDER-FORM.xlsx willmar_orders.csv
```

## Output CSV Columns

| Column | Description |
|--------|-------------|
| order_id | Order identifier from header |
| date_of_order | Order date |
| customer_name | Receiver's name |
| customer_email | Customer email address |
| phone | Customer phone number |
| billing_address | Billing address |
| delivery_address | Shipping address |
| product_description | Products included in this kit |
| personalization_name | Name on the jersey |
| gender | Men/Women/Youth |
| jersey_size | Jersey size |
| shorts_size | Shorts size |
| socks_size | Socks size (if applicable) |
| item_number | Player number or initials |
| quantity | Quantity ordered |
| unit_price | Price per item |
| total_price | Line total |

## Example Output

```
============================================================
BATCH PROCESSING MODE
============================================================
Found 3 Excel file(s) to process:

  - ORDER-FORM-1.xlsx
  - ORDER-FORM-2.xlsx
  - ORDER-FORM-3.xlsx

Processing: ORDER-FORM-1.xlsx
--------------------------------------------------
  Order ID: WILLMAR CARDINALS (DECEMBER 2025)
  Customer: JEFF WINTER
  Extracted: 28 rows
...

============================================================
SUMMARY
============================================================
  Total files processed: 3
  Total rows extracted: 127
  Output saved to: D:\stimulus_parser\all_orders_extracted.csv
============================================================
```
