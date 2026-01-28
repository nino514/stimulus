# Excel Order Parser - Technical Specification

## Overview

This document specifies the design for a robust Python parser that extracts order data from Stimulus Athletic's custom Excel invoice files and outputs normalized CSV data.

**Input:** Excel files (`.xlsx`) with custom invoice layout
**Output:** Clean CSV with one row per personalization item

---

## 1. Document Structure Model

Based on analysis of the sample file, the Excel document follows this hierarchy:

```
┌─────────────────────────────────────────────────────────────┐
│ HEADER (Logo, Company Info)                                 │
├─────────────────────────────────────────────────────────────┤
│ ORDER METADATA                                              │
│   ├── Order ID                                              │
│   ├── Date of Order                                         │
│   ├── Delivery Date (Estimated)                             │
│   └── Order Type                                            │
├─────────────────────────────────────────────────────────────┤
│ CUSTOMER INFORMATION                                        │
│   ├── Billing Address                                       │
│   ├── Delivery Address                                      │
│   ├── Receiver's Name                                       │
│   ├── Telephone No.                                         │
│   └── Email Address                                         │
├─────────────────────────────────────────────────────────────┤
│ DESIGN INFORMATION                                          │
│   ├── Design Request ID                                     │
│   ├── Dropbox Link                                          │
│   └── Design Notes                                          │
├─────────────────────────────────────────────────────────────┤
│ ORDER SUMMARY (table)                                       │
│   └── [Item #, Products, Quantity, Unit Price, Total]       │
├─────────────────────────────────────────────────────────────┤
│ PRODUCT SECTION 1                                           │
│   ├── Product Info (images, description, printing type)     │
│   └── PERSONALIZATION TABLE                                 │
│         └── [Gender, Sizes, Name, Number, Qty, Prices...]   │
├─────────────────────────────────────────────────────────────┤
│ PRODUCT SECTION 2                                           │
│   ├── Product Info                                          │
│   └── PERSONALIZATION TABLE                                 │
├─────────────────────────────────────────────────────────────┤
│ PRODUCT SECTION N...                                        │
└─────────────────────────────────────────────────────────────┘
```

---

## 2. Section Identification Strategy

### 2.1 Section Markers

Each major section is identified by scanning for **marker text** in cells. The parser uses a prioritized list of markers with fuzzy matching:

| Section Type | Primary Markers | Fallback Markers |
|--------------|-----------------|------------------|
| Order Metadata | `"ORDER ID"` | `"ORDER_ID"`, `"ORDERID"` |
| Customer Info | `"CUSTOMER INFORMATION"` | `"CUSTOMER INFO"`, `"BILLING ADDRESS"` |
| Design Info | `"DESIGN INFORMATION"` | `"DESIGN REQUEST ID"` |
| Order Summary | `"ORDER SUMMARY"` | `"ITEM"` + `"PRODUCTS"` + `"QUANTITY"` in same row |
| Product Section | `"PRODUCT INFORMATION"` | `"PRODUCTS INCLUDED"` |
| Personalization | `"PERSONALIZATION"` | `"GENDER"` + `"SIZE"` header row |

### 2.2 Section Detection Algorithm

```
ALGORITHM: detect_sections(sheet)

INPUT: openpyxl worksheet object
OUTPUT: List of Section objects with type, start_row, end_row

1. SCAN all rows (1 to max_row):
   a. For each cell in row:
      - Normalize text: uppercase, strip whitespace, remove special chars
      - Check against marker dictionary
      - If match found, record (row_number, section_type, confidence_score)

2. SORT detected markers by row_number

3. DETERMINE section boundaries:
   a. For each detected marker:
      - start_row = marker row
      - end_row = next marker row - 1 (or last data row if final section)

4. VALIDATE section sequence:
   a. Expected order: METADATA → CUSTOMER → DESIGN → SUMMARY → PRODUCTS
   b. Flag warnings if sections appear out of order or are missing

5. RETURN list of Section(type, start_row, end_row, confidence)
```

### 2.3 Product Section Identification (Detail)

Product sections require special handling because they repeat and contain nested personalization tables.

```
ALGORITHM: identify_product_sections(sheet, start_row)

1. FIND all rows containing "PRODUCT INFORMATION" (case-insensitive)
   - Store as product_header_rows[]

2. FIND all rows containing "PERSONALIZATION"
   - Store as personalization_header_rows[]

3. PAIR each product header with its personalization section:
   FOR each product_header in product_header_rows:
       - Find the NEXT personalization_header that comes AFTER this product_header
       - product.info_start = product_header row
       - product.info_end = personalization_header row - 1
       - product.personalization_start = personalization_header row
       - product.personalization_end = next product_header row - 1
                                       (or end of data if last product)

4. VALIDATE pairing:
   - Each product should have exactly one personalization section
   - Warn if personalization found without preceding product header

5. RETURN list of ProductSection objects
```

**Visual representation of product boundary detection:**

```
Row 45: ████ PRODUCT INFORMATION ████  ← product_1.info_start
Row 46: [Product images and details]
Row 47: PRODUCTS INCLUDED: RED GAME JERSEY...
Row 48: PRINTING TYPE: SUBLIMATION
...
Row 55: ████ PERSONALIZATION ████      ← product_1.personalization_start
Row 56: GENDER | JERSEY | SHORTS | ... ← header row (skip)
Row 57: MEN    | Small  | Small  | ... ← data row 1
Row 58: MEN    | Small  | Small  | ... ← data row 2
...
Row 82: MEN    | Large  | Large  | ... ← data row N
Row 83: [empty or TOTAL QTY row]       ← product_1.personalization_end
Row 84: ████ PRODUCT INFORMATION ████  ← product_2.info_start (next product)
```

---

## 3. Personalization Row Extraction

### 3.1 Header Row Detection

The personalization table header row contains column names. Detection strategy:

```
ALGORITHM: find_personalization_header(sheet, section_start_row)

1. Starting from section_start_row, scan forward (max 5 rows)

2. For each row, check if it contains MULTIPLE expected header terms:
   - Required: "GENDER" or "NAME"
   - Supporting: "SIZE", "JERSEY", "SHORTS", "SOCKS", "QUANTITY", "PRICE"

3. SCORING:
   - +2 points for each required term found
   - +1 point for each supporting term found
   - Row with highest score (minimum 3) = header row

4. If no valid header found, WARN and attempt column inference from data patterns

5. RETURN header_row_number, column_mapping
```

### 3.2 Column Position Mapping

Once header row is found, map column positions to field names:

```
ALGORITHM: map_columns(header_row)

INPUT: Row containing header cells
OUTPUT: Dictionary {column_index: field_name}

1. READ all non-empty cells in header row

2. For each cell, normalize and match against known headers:

   HEADER_MAPPINGS = {
       # Gender
       "GENDER": "gender",

       # Sizes (may be sub-headers under "SIZE")
       "JERSEY": "jersey_size",
       "SHORTS": "shorts_size",
       "SOCKS": "socks_size",
       "SIZE": None,  # Parent header, skip

       # Identification
       "NAME": "personalization_name",
       "INITIALS OR NUMBER": "item_number",
       "INITIALS/NUMBER": "item_number",
       "NUMBER": "item_number",
       "INITIALS": "item_number",

       # Pricing
       "QUANTITY": "quantity",
       "QTY": "quantity",
       "UNIT PRICE": "unit_price",
       "PRICE": "unit_price",
       "TOTAL PRICE": "total_price",
       "TOTAL": "total_price",
   }

3. Handle merged header cells:
   - "SIZE" often spans multiple columns (JERSEY, SHORTS, SOCKS)
   - If "SIZE" found as merged cell, look at row below for sub-headers

4. VALIDATE required columns present:
   - Required: gender, personalization_name, quantity
   - Optional: all size columns, item_number, prices

5. RETURN column_map = {0: "gender", 1: "jersey_size", 2: "shorts_size", ...}
```

### 3.3 Data Row Extraction

```
ALGORITHM: extract_personalization_rows(sheet, start_row, end_row, column_map)

INPUT:
  - sheet: worksheet object
  - start_row: first data row (after header)
  - end_row: last row of this personalization section
  - column_map: {column_index: field_name}

OUTPUT: List of PersonalizationItem objects

1. FOR each row from start_row to end_row:

   a. SKIP if row is empty (all cells None or whitespace)

   b. SKIP if row is a summary row:
      - Contains "TOTAL QTY" or "SUBTOTAL"
      - Contains only numbers in price columns (no gender/name)

   c. EXTRACT cell values using column_map:
      item = {}
      for col_index, field_name in column_map:
          cell_value = sheet.cell(row, col_index).value
          item[field_name] = clean_value(cell_value, field_name)

   d. VALIDATE item has minimum required fields:
      - Must have: gender OR name OR item_number (at least one identifier)
      - If validation fails, log warning and skip row

   e. APPEND item to results

2. RETURN list of PersonalizationItem
```

### 3.4 Data Cleaning Functions

```python
def clean_value(raw_value, field_name):
    """Clean and normalize extracted cell values."""

    if raw_value is None:
        return None

    # Convert to string and strip
    value = str(raw_value).strip()

    if value == "" or value.lower() in ("n/a", "na", "-", "none"):
        return None

    # Field-specific cleaning
    if field_name == "gender":
        return normalize_gender(value)  # "MEN" → "Men", "M" → "Men"

    elif field_name in ("jersey_size", "shorts_size", "socks_size"):
        return normalize_size(value)  # "MEDIUM" → "Medium", "M" → "Medium"

    elif field_name in ("unit_price", "total_price"):
        return parse_currency(value)  # "$95.00" → 95.00

    elif field_name == "quantity":
        return parse_number(value)  # "1" → 1, "1.00" → 1

    elif field_name == "item_number":
        return str(value)  # Keep as string (could be "12" or "AB")

    else:
        return value
```

---

## 4. Column Header Mapping (Complete Reference)

### 4.1 Input → Output Field Mapping

| Excel Header (variations) | Output CSV Column | Data Type | Required |
|---------------------------|-------------------|-----------|----------|
| `ORDER ID`, `Order ID` | `order_id` | string | Yes |
| `RECEIVER'S NAME`, `Receiver's Name` | `customer_name` | string | Yes |
| `EMAIL ADDRESS`, `Email Address` | `customer_email` | string | Yes |
| `PRODUCTS INCLUDED`, `Products` | `product_description` | string | Yes |
| `NAME` | `personalization_name` | string | Yes |
| `GENDER`, `Gender` | `gender` | string | No |
| `JERSEY`, `Jersey Size` | `jersey_size` | string | No |
| `SHORTS`, `Shorts Size` | `shorts_size` | string | No |
| `SOCKS`, `Socks Size` | `socks_size` | string | No |
| `INITIALS OR NUMBER`, `Number` | `item_number` | string | No |
| `UNIT PRICE`, `Price` | `unit_price` | decimal | No |
| `QUANTITY`, `Qty` | `quantity` | integer | Yes |
| `TOTAL PRICE`, `Total` | `total_price` | decimal | No |

### 4.2 Header Normalization Rules

```python
HEADER_ALIASES = {
    # Order metadata
    "order_id": ["ORDER ID", "ORDER_ID", "ORDERID", "ORDER #", "ORDER NO"],
    "date_of_order": ["DATE OF ORDER", "ORDER DATE", "DATE"],

    # Customer info
    "customer_name": ["RECEIVER'S NAME", "RECEIVERS NAME", "CUSTOMER NAME", "NAME", "RECIPIENT"],
    "customer_email": ["EMAIL ADDRESS", "EMAIL", "E-MAIL"],
    "billing_address": ["BILLING ADDRESS", "BILL TO"],
    "delivery_address": ["DELIVERY ADDRESS", "SHIP TO", "SHIPPING ADDRESS"],
    "phone": ["TELEPHONE NO", "TELEPHONE", "PHONE", "TEL"],

    # Product info
    "product_description": ["PRODUCTS INCLUDED", "PRODUCTS", "PRODUCT", "DESCRIPTION", "ITEM"],
    "printing_type": ["PRINTING TYPE", "PRINT TYPE"],
    "socks_code": ["SOCKS CODE", "SOCK CODE"],

    # Personalization
    "gender": ["GENDER", "SEX"],
    "jersey_size": ["JERSEY", "JERSEY SIZE"],
    "shorts_size": ["SHORTS", "SHORTS SIZE", "SHORT"],
    "socks_size": ["SOCKS", "SOCKS SIZE", "SOCK SIZE"],
    "personalization_name": ["NAME", "PLAYER NAME", "TEAM NAME"],
    "item_number": ["INITIALS OR NUMBER", "INITIALS/NUMBER", "NUMBER", "NO", "#", "INITIALS"],

    # Pricing
    "quantity": ["QUANTITY", "QTY", "COUNT"],
    "unit_price": ["UNIT PRICE", "PRICE", "EACH"],
    "total_price": ["TOTAL PRICE", "TOTAL", "LINE TOTAL", "AMOUNT"],
}

def normalize_header(raw_header):
    """Convert raw header text to canonical field name."""
    normalized = raw_header.upper().strip()

    for field_name, aliases in HEADER_ALIASES.items():
        if normalized in aliases:
            return field_name

    # Fuzzy match for typos (Levenshtein distance ≤ 2)
    for field_name, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            if levenshtein_distance(normalized, alias) <= 2:
                return field_name

    return None  # Unknown header
```

---

## 5. Edge Case Handling

### 5.1 Missing or Empty Values

| Scenario | Detection | Handling |
|----------|-----------|----------|
| Empty cell | `cell.value is None` | Set field to `None` in output |
| Whitespace only | `str(value).strip() == ""` | Set field to `None` |
| "N/A" or "-" | Match against skip list | Set field to `None` |
| Missing required field | Field is `None` after extraction | Log warning, include row with `None` |
| Entire row empty | All cells `None` | Skip row silently |

### 5.2 Missing Size Columns

Some products may not have all size columns (e.g., socks not included):

```python
def handle_missing_size_columns(column_map, product_info):
    """
    Check if size columns are expected based on product description.
    """
    product_desc = product_info.get("products_included", "").upper()

    expected_sizes = {
        "jersey_size": "JERSEY" in product_desc,
        "shorts_size": "SHORTS" in product_desc,
        "socks_size": "SOCKS" in product_desc,
    }

    for size_field, expected in expected_sizes.items():
        if expected and size_field not in column_map.values():
            log_warning(f"Expected {size_field} column but not found in headers")
        elif not expected and size_field in column_map.values():
            # Size column present but product doesn't include it - OK, might be empty
            pass
```

### 5.3 Multi-Row Headers (Merged Cells)

The "SIZE" header spans multiple columns with sub-headers below:

```
Row 55: | GENDER |     SIZE      | NAME | ...
Row 56: |        | JERSEY|SHORTS|SOCKS |      | ...  (sub-headers)
Row 57: | MEN    | Small | Small| M    | ...  (data)
```

**Detection and handling:**

```python
def detect_merged_header(sheet, header_row):
    """
    Detect if header row has merged cells requiring sub-header row.
    """
    merged_ranges = sheet.merged_cells.ranges

    for merge_range in merged_ranges:
        if merge_range.min_row == header_row:
            # Found merged cell in header row
            parent_value = sheet.cell(header_row, merge_range.min_col).value

            if parent_value and "SIZE" in str(parent_value).upper():
                # Check row below for sub-headers
                sub_header_row = header_row + 1
                sub_headers = {}

                for col in range(merge_range.min_col, merge_range.max_col + 1):
                    sub_val = sheet.cell(sub_header_row, col).value
                    if sub_val:
                        sub_headers[col] = normalize_header(str(sub_val))

                return sub_header_row, sub_headers

    return None, {}
```

### 5.4 Summary/Total Rows Mixed with Data

Personalization tables often end with summary rows:

```
Row 80: | MEN | Large | Large | | WILLMAR | 26 | 1 | $95.00 | $95.00 |  ← data
Row 81: |     |       |       | |         |    |   | TOTAL QTY | 26.00 | ← summary (skip)
Row 82: |     |       |       | |         |    |   | SUBTOTAL  | $2,375 | ← summary (skip)
```

**Detection:**

```python
def is_summary_row(row_data, column_map):
    """
    Detect if row is a summary/total row rather than data.
    """
    # Check for summary keywords
    summary_keywords = ["TOTAL", "SUBTOTAL", "SUM", "GRAND TOTAL"]

    row_text = " ".join(str(v) for v in row_data.values() if v).upper()

    if any(keyword in row_text for keyword in summary_keywords):
        return True

    # Check if row lacks required identifier fields
    has_gender = row_data.get("gender") is not None
    has_name = row_data.get("personalization_name") is not None
    has_number = row_data.get("item_number") is not None

    if not (has_gender or has_name or has_number):
        # Row has no identifiers - likely a summary row
        return True

    return False
```

### 5.5 Page Breaks / Continuation Tables

When personalization data spans pages, the header row may repeat:

```
Page 1, Row 82: | MEN | Large | WILLMAR | 13 | ...
--- PAGE BREAK ---
Page 2, Row 1:  | GENDER | JERSEY | SHORTS | NAME | ...  ← repeated header
Page 2, Row 2:  | MEN    | Medium | Medium | WILLMAR | 14 | ...
```

**Handling:**

```python
def is_repeated_header(row_data, expected_headers):
    """
    Detect if a data row is actually a repeated header row.
    """
    row_values = [str(v).upper().strip() for v in row_data.values() if v]
    header_values = [h.upper() for h in expected_headers]

    # If >50% of row values match header names, it's a repeated header
    matches = sum(1 for v in row_values if v in header_values)

    return matches > len(row_values) * 0.5
```

### 5.6 Data Type Variations

| Field | Expected Variations | Normalization |
|-------|---------------------|---------------|
| Gender | `"MEN"`, `"Men"`, `"M"`, `"MALE"`, `"WOMEN"`, `"W"`, `"FEMALE"` | Map to `"Men"` or `"Women"` |
| Sizes | `"Small"`, `"S"`, `"SM"`, `"SMALL"`, `"small"` | Map to `"Small"`, `"Medium"`, `"Large"`, `"XS"`, `"XL"`, etc. |
| Prices | `"$95.00"`, `"95.00"`, `"95"`, `"$95"` | Parse to `Decimal("95.00")` |
| Quantity | `"1"`, `"1.00"`, `1`, `1.0` | Parse to `int(1)` |

```python
SIZE_NORMALIZATION = {
    # Extra Small
    "XS": "XS", "EXTRA SMALL": "XS", "XSMALL": "XS",
    # Small
    "S": "Small", "SM": "Small", "SMALL": "Small",
    # Medium
    "M": "Medium", "MED": "Medium", "MEDIUM": "Medium",
    # Large
    "L": "Large", "LG": "Large", "LARGE": "Large",
    # Extra Large
    "XL": "XL", "EXTRA LARGE": "XL", "XLARGE": "XL",
    # 2XL, 3XL, etc.
    "XXL": "2XL", "2XL": "2XL", "XXXL": "3XL", "3XL": "3XL",
}

GENDER_NORMALIZATION = {
    "MEN": "Men", "M": "Men", "MALE": "Men", "MENS": "Men", "MAN": "Men",
    "WOMEN": "Women", "W": "Women", "FEMALE": "Women", "WOMENS": "Women", "WOMAN": "Women",
    "YOUTH": "Youth", "Y": "Youth", "KIDS": "Youth", "CHILD": "Youth",
}
```

### 5.7 Error Recovery Strategy

```
ALGORITHM: extract_with_recovery(sheet, section)

1. TRY normal extraction pipeline

2. IF extraction fails or returns 0 rows:
   a. Log detailed error with row numbers
   b. Attempt FALLBACK strategies:

   FALLBACK 1: Relaxed header matching
      - Use fuzzy matching with higher tolerance
      - Accept partial header matches

   FALLBACK 2: Position-based extraction
      - If previous files had consistent column positions,
        use those positions as hints

   FALLBACK 3: Pattern-based row detection
      - Look for rows matching pattern: [text, size, size, size, text, number, number, currency, currency]
      - Extract without relying on headers

3. IF all fallbacks fail:
   a. Mark section as "extraction_failed"
   b. Include raw row data in error report
   c. Continue processing remaining sections

4. RETURN extracted data + warnings list
```

---

## 6. Output Specification

### 6.1 CSV Output Format

**Filename:** `{original_filename}_extracted.csv`

**Columns (in order):**

| # | Column Name | Type | Example |
|---|-------------|------|---------|
| 1 | `order_id` | string | `"WILLMAR CARDINALS (DECEMBER 2025)"` |
| 2 | `customer_name` | string | `"JEFF WINTER"` |
| 3 | `customer_email` | string | `"WINTER_J@WILLMAR.K12.MN.US"` |
| 4 | `product_description` | string | `"RED GAME JERSEY + WHITE GAME JERSEY + GRAY GAME SHORTS"` |
| 5 | `personalization_name` | string | `"WILLMAR"` |
| 6 | `gender` | string | `"Men"` |
| 7 | `jersey_size` | string | `"Medium"` |
| 8 | `shorts_size` | string | `"Medium"` |
| 9 | `socks_size` | string | `"M"` or `null` |
| 10 | `item_number` | string | `"14"` |
| 11 | `unit_price` | decimal | `95.00` |
| 12 | `quantity` | integer | `1` |
| 13 | `total_price` | decimal | `95.00` |

### 6.2 Sample Output

```csv
order_id,customer_name,customer_email,product_description,personalization_name,gender,jersey_size,shorts_size,socks_size,item_number,unit_price,quantity,total_price
"WILLMAR CARDINALS (DECEMBER 2025)","JEFF WINTER","WINTER_J@WILLMAR.K12.MN.US","RED GAME JERSEY + WHITE GAME JERSEY + GRAY GAME SHORTS","WILLMAR","Men","Small","Small",,"2",95.00,1,95.00
"WILLMAR CARDINALS (DECEMBER 2025)","JEFF WINTER","WINTER_J@WILLMAR.K12.MN.US","RED GAME JERSEY + WHITE GAME JERSEY + GRAY GAME SHORTS","WILLMAR","Men","Small","Small",,"3",95.00,1,95.00
...
"WILLMAR CARDINALS (DECEMBER 2025)","JEFF WINTER","WINTER_J@WILLMAR.K12.MN.US","PINK GK JERSEY + PINK GAME SHORTS + PINK SOCKS","WILLMAR","Men","L","L","M","1",69.00,1,69.00
"WILLMAR CARDINALS (DECEMBER 2025)","JEFF WINTER","WINTER_J@WILLMAR.K12.MN.US","BLUE GK JERSEY + BLUE GAME SHORTS + BLUE SOCKS","WILLMAR","Men","L","L","M","1",69.00,1,69.00
```

### 6.3 Validation Report Output

In addition to CSV, generate a validation report:

**Filename:** `{original_filename}_validation.json`

```json
{
  "file_name": "willmar_cardinals_dec2025.xlsx",
  "extraction_timestamp": "2026-01-28T10:30:00Z",
  "status": "success_with_warnings",

  "summary": {
    "total_rows_extracted": 28,
    "products_found": 3,
    "order_id": "WILLMAR CARDINALS (DECEMBER 2025)",
    "customer_name": "JEFF WINTER"
  },

  "validation_checks": {
    "quantity_sum_matches": {
      "passed": true,
      "expected": 28,
      "actual": 28
    },
    "total_price_matches": {
      "passed": true,
      "expected": 2513.00,
      "actual": 2513.00
    },
    "required_fields_present": {
      "passed": true,
      "missing_fields": []
    }
  },

  "warnings": [
    {
      "type": "missing_value",
      "row": 15,
      "field": "socks_size",
      "message": "Socks size empty but product includes socks"
    }
  ],

  "errors": []
}
```

---

## 7. Implementation Architecture

### 7.1 Module Structure

```
stimulus/
├── excel_parser/
│   ├── __init__.py
│   ├── main.py              # CLI entry point
│   ├── parser.py            # Main ExcelOrderParser class
│   ├── section_detector.py  # Section boundary detection
│   ├── extractors/
│   │   ├── __init__.py
│   │   ├── metadata.py      # Order metadata extraction
│   │   ├── customer.py      # Customer info extraction
│   │   ├── product.py       # Product section extraction
│   │   └── personalization.py  # Personalization table extraction
│   ├── normalizers.py       # Data cleaning and normalization
│   ├── validators.py        # Validation and cross-checking
│   └── output.py            # CSV and report generation
├── tests/
│   ├── test_parser.py
│   ├── test_section_detector.py
│   ├── test_extractors.py
│   └── fixtures/
│       └── sample_order.xlsx
└── requirements.txt
```

### 7.2 Main Parser Class

```python
class ExcelOrderParser:
    """
    Main parser class orchestrating the extraction pipeline.
    """

    def __init__(self, config: Optional[ParserConfig] = None):
        self.config = config or ParserConfig()
        self.section_detector = SectionDetector()
        self.validators = [
            QuantitySumValidator(),
            PriceTotalValidator(),
            RequiredFieldsValidator(),
        ]
        self.warnings: List[Warning] = []
        self.errors: List[Error] = []

    def parse(self, file_path: Path) -> ParseResult:
        """
        Main entry point for parsing an Excel file.

        Returns ParseResult containing extracted data and validation info.
        """
        workbook = load_workbook(file_path, data_only=True)
        sheet = workbook.active

        # Pass 1: Detect document structure
        sections = self.section_detector.detect(sheet)

        # Pass 2: Extract data from each section
        order_metadata = self._extract_metadata(sheet, sections)
        customer_info = self._extract_customer(sheet, sections)
        products = self._extract_products(sheet, sections)

        # Pass 3: Validate and cross-check
        validation_results = self._validate(order_metadata, customer_info, products)

        # Build output records
        records = self._build_output_records(order_metadata, customer_info, products)

        return ParseResult(
            records=records,
            validation=validation_results,
            warnings=self.warnings,
            errors=self.errors
        )

    def to_csv(self, result: ParseResult, output_path: Path) -> None:
        """Write extracted records to CSV file."""
        ...

    def to_validation_report(self, result: ParseResult, output_path: Path) -> None:
        """Write validation report to JSON file."""
        ...
```

### 7.3 Dependencies

```
# requirements.txt
openpyxl>=3.1.0      # Excel file reading
python-Levenshtein>=0.20.0  # Fuzzy string matching (optional)
pydantic>=2.0.0      # Data validation models (optional)
```

---

## 8. Testing Strategy

### 8.1 Unit Tests

| Component | Test Cases |
|-----------|------------|
| `SectionDetector` | Finds all sections, handles missing sections, handles out-of-order sections |
| `HeaderMapper` | Maps all known aliases, handles unknown headers, handles merged cells |
| `PersonalizationExtractor` | Extracts all rows, skips empty rows, skips summary rows, handles repeated headers |
| `Normalizers` | Gender normalization, size normalization, price parsing, quantity parsing |
| `Validators` | Quantity sum check, price total check, required fields check |

### 8.2 Integration Tests

| Scenario | Input | Expected Output |
|----------|-------|-----------------|
| Normal file | Sample order with 3 products | 28 rows extracted, all validations pass |
| Missing socks column | Order without socks | Rows extracted with null socks_size |
| Repeated header mid-table | Table spanning pages | All data rows extracted, header rows skipped |
| Malformed prices | `"$95"`, `"95.00"`, `"95"` mixed | All parsed to `95.00` |

---

## 9. Future Considerations

### 9.1 Potential Enhancements (Not in Initial Scope)

1. **Multiple file batch processing** - Process entire folders
2. **Database direct insert** - Skip CSV, insert directly to DB
3. **Template learning** - Auto-detect column positions from new file formats
4. **GUI interface** - Drag-and-drop file processing
5. **Excel output** - Output to formatted Excel in addition to CSV

### 9.2 Known Limitations

1. **Images not extracted** - Product images are ignored
2. **Single sheet only** - Assumes all data on first sheet
3. **English only** - Header matching assumes English text
4. **Manual template changes** - Major layout changes require code updates

---

## Approval Checklist

Before implementation, please confirm:

- [ ] Section identification strategy is correct
- [ ] Column mappings cover all expected headers
- [ ] Edge case handling is sufficient
- [ ] Output format meets database requirements
- [ ] Validation checks are appropriate

**Questions for clarification:**

1. Are there other header variations I should account for?
2. Should the output include any additional calculated fields?
3. What should happen if validation fails - stop processing or continue with warnings?
4. Do you need support for multiple sheets in a single workbook?
