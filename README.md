# Stimulus Order Extractor

Extract order data from Stimulus Athletic Excel invoice files into clean CSV format.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python extract_orders.py <input.xlsx> <output.csv>
```

### Example

```bash
python extract_orders.py WILLMAR-CARDINALS_ORDER-FORM_122125.xlsx orders.csv
```

## Output CSV Columns

| Column | Description |
|--------|-------------|
| order_id | Order identifier from header |
| customer_name | Receiver's name |
| customer_email | Customer email address |
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
