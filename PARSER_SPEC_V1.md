# Excel Order Parser - V1 Specification (MVP)

## Overview

Simple Python script to extract order data from Stimulus Athletic Excel invoices.

**Input:** Single `.xlsx` file
**Output:** Single `.csv` file

---

## Output CSV Columns

```
order_id, customer_name, customer_email, product_description, personalization_name, gender, jersey_size, shorts_size, socks_size, item_number, quantity, unit_price, total_price
```

---

## Extraction Logic

### Step 1: Extract Order Metadata (scan rows top-to-bottom)

| Look for cell containing | Extract value from | Store as |
|--------------------------|-------------------|----------|
| `"ORDER ID"` | Same row, next non-empty cell | `order_id` |
| `"RECEIVER'S NAME"` | Same row, next non-empty cell | `customer_name` |
| `"EMAIL ADDRESS"` | Same row, next non-empty cell | `customer_email` |

### Step 2: Find Product Sections

For each row containing `"PRODUCTS INCLUDED"`:
- Extract the value in the adjacent cell → `product_description`
- Note the row number

### Step 3: Find Personalization Sections

For each row containing `"PERSONALIZATION"`:
- The next row is the **header row** (GENDER | SIZE... ) → skip it
- Read data rows starting from header + 1
- Stop when you hit:
  - Empty row (all cells blank), OR
  - Row containing `"PRODUCT INFORMATION"`, OR
  - Row containing `"TOTAL QTY"` or `"SUBTOTAL"`

### Step 4: Extract Personalization Data Rows

For each data row in a personalization section, extract by **column position** (based on header):

| Column | Maps to |
|--------|---------|
| GENDER | `gender` |
| JERSEY (under SIZE) | `jersey_size` |
| SHORTS (under SIZE) | `shorts_size` |
| SOCKS (under SIZE) | `socks_size` |
| NAME | `personalization_name` |
| INITIALS OR NUMBER | `item_number` |
| QUANTITY | `quantity` |
| UNIT PRICE | `unit_price` |
| TOTAL PRICE | `total_price` |

### Step 5: Link Products to Personalization

Each personalization section belongs to the **most recent** `"PRODUCTS INCLUDED"` found above it.

---

## Edge Cases (V1)

| Scenario | Handling |
|----------|----------|
| Empty cell | Write blank to CSV (not "None") |
| Row contains "TOTAL QTY" or "SUBTOTAL" | Skip row |
| Header row (contains "GENDER") | Skip row |
| Price formatting (`$95.00`) | Strip `$` and write as number |

---

## File Structure

```
stimulus/
├── extract_orders.py    # Main script (single file)
├── requirements.txt     # Just: openpyxl
└── sample_orders/       # Place Excel files here
    └── input.xlsx
```

---

## Usage

```bash
python extract_orders.py sample_orders/input.xlsx output.csv
```

---

## NOT in V1 (Future)

- Multiple file batch processing
- Fuzzy header matching
- Validation reports
- Error recovery
- Multi-sheet support
