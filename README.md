# Gift Shop Sales Analysis — Excel Project

A data analysis project built in Excel: cleaning raw data, merging three
related tables with formulas, and summarizing results on a dashboard with charts.

## Files

| File | Description |
|---|---|
| `gift_shop_sales_analysis.xlsx` | The Excel project (raw data, cleaned data, summary, dashboard) |
| `customers.csv` | Raw customer data (204 records) |
| `orders.csv` | Raw order data (1,000 records) |
| `products.csv` | Raw product catalog (70 records) |

## What's inside the workbook

- **README** — project overview and notes on every cleaning step
- **Raw_Customers / Raw_Orders / Raw_Products** — original, untouched data
- **Cleaned_Data** — all three tables merged into one record set using `INDEX`/`MATCH`,
  with cleaned phone numbers, real date values, and calculated columns
- **Summary** — revenue totals by category and occasion, built with `SUMIFS`/`COUNTIFS`
- **Dashboard** — 4 KPI cards plus a bar chart (revenue by category) and a pie chart
  (revenue share by occasion)

## Data cleaning performed

1. **Phone numbers** — raw values had inconsistent lengths (9, 10, or 12 digits, with or
   without a country code). Standardized every number to `+91-XXXXXXXXXX`.
2. **City names** — removed extra leading/trailing whitespace with `TRIM()`.
3. **Dates** — raw dates were stored as text (`DD-MM-YYYY`). Converted them into real Excel
   date values so they sort correctly and support date math (e.g. delivery time in days).
4. **Duplicates** — checked `Order_ID` and `Customer_ID` for duplicates with `COUNTIF`
   (none found).
5. **Merging** — combined customer, product, and order data into a single table using
   `INDEX`/`MATCH` lookups instead of copy-pasting.

## Skills demonstrated

Excel formulas (`INDEX`/`MATCH`, `SUMIFS`, `COUNTIFS`, `TRIM`, `DATE`, `RIGHT`/`MID`/`LEFT`),
data cleaning, table structuring, and native Excel charts (bar + pie).

<img width="512" height="209" alt="image" src="https://github.com/user-attachments/assets/41f68643-0326-48ac-bd4a-e9e5f6ce0ec6" />

