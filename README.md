# Packaging Sales & Inventory Automation

A realistic portfolio project that automates sales reporting for a small packaging-material distributor.

## What it does
- Reads order, product, and customer data from `source_sales_data.xlsx`
- Cleans inconsistent customer names/cities and standardizes text fields
- Detects duplicate order IDs and flags invalid dates, quantities, prices, and payment statuses
- Calculates gross sales, discounts, net sales, amounts received, and outstanding dues
- Creates monthly sales, product performance, customer performance, payment analysis, and low-stock reports
- Formats the Excel report with filters, frozen headers, currency formatting, and attention highlights

## Main libraries
- `pandas` — reading, cleaning, transforming, grouping, and summarizing spreadsheet data
- `openpyxl` — Excel workbook formatting and styling

The project also uses Python's built-in `pathlib` for file paths. No other third-party library is required.

## Folder structure
```text
packaging_sales_automation/
├── automation.py
├── source_sales_data.xlsx
├── requirements.txt
├── README.md
└── reports/
    └── packaging_sales_report.xlsx  (created after running)
```

## Setup (Windows / VS Code)
1. Install Python 3.10 or newer.
2. Open this project folder in VS Code.
3. Open Terminal → New Terminal.
4. Run:
   ```bash
   py -m pip install -r requirements.txt
   ```
5. Run:
   ```bash
   py automation.py
   ```
6. Open `reports/packaging_sales_report.xlsx`.

If `py` does not work, try `python -m pip install -r requirements.txt` and `python automation.py`.

## Output worksheets
- **Executive Summary** — key metrics and a note about estimated profit
- **Monthly Sales** — monthly order count, sales, discounts, receipts, and dues
- **Product Performance** — sales and estimated gross profit by product
- **Customer Performance** — customer sales and outstanding dues
- **Payment Analysis** — paid/pending/overdue breakdown
- **Low Stock Alerts** — products at or below their reorder threshold
- **Cleaned Orders** — cleaned order-level data with validation flags

## Important
The workbook contains synthetic demo data, not real customer data. Estimated gross profit uses the product unit costs in the Products sheet and excludes overheads, taxes, returns, and other adjustments. This is a portfolio demonstration, not accounting software.

## Portfolio / GitHub description
**Packaging Sales & Inventory Automation | Python, pandas, openpyxl**

Built an Excel automation workflow for a packaging-material distributor. The project cleans sales data, validates records, calculates net sales and outstanding receivables, summarizes customer/product performance, flags low inventory, and exports a professionally formatted multi-sheet Excel report.
