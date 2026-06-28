# Retail Orders — Data Cleaning & Validation Project

## Project Overview

This project cleans, validates, and prepares an order-level retail sales dataset exported from multiple internal systems. The raw data contained inconsistent formatting, date anomalies, duplicate records, missing values, invalid discounts, and order status issues. The output is an analysis-ready dataset with full documentation and pivot-based summary reports for business review.

---


## Dataset Summary

| Metric | Value |
|---|---|
| Raw input records | 932 |
| Exact duplicate rows removed | 20 |
| Duplicate order IDs flagged for review | 12 |
| Final clean records | **900** |
| Columns in cleaned output | 34 |

---

## Tasks Completed

### Task 2 — Clean Text Fields
Standardized 10 text columns: `customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`.

Techniques used: TRIM, PROPER, SUBSTITUTE, Find & Replace.

Issues resolved: leading/trailing spaces, double internal spaces, mixed case variants (e.g. `COMPLETED`, `completed`, `  Completed `), typo variants (e.g. `Small  Business`, `Standard  Class`).

### Task 3 — Clean and Validate Dates
Standardized `order_date` and `ship_date` from 5 different formats to `YYYY-MM-DD`.

New column created: `shipping_delay_days` = ship_date − order_date.

Flags applied: `INVALID_SHIP_BEFORE_ORDER` for 20 records where ship date preceded order date.

### Task 4 — Handle Duplicates
- 20 exact duplicate rows identified and removed (keep first occurrence)
- 12 order IDs with conflicting data identified and flagged with `DUPLICATE_FLAGGED_FOR_REVIEW`
- All flagged records preserved in `data_quality_report.xlsx` → Sheet: Duplicate Details

### Task 5 — Apply Business Rules
All 9 business rules applied. See `cleaning_log.md` for full decision detail.

### Task 6 — Create Calculated Columns
8 new columns added to `cleaned_orders.xlsx`:
`cleaned_discount`, `calculated_sales`, `calculated_profit`, `profit_margin`, `shipping_delay_days`, `order_month`, `order_year`, `data_quality_flag`

### Task 8 — Pivot Summary Report
7 pivot sheets created in `pivot_summary.xlsx` covering region, category, ship mode, segment, issue analysis, monthly trend, and refunded orders.

---

## Output Files

### cleaned_orders.xlsx
The primary deliverable. Contains all 900 records after deduplication, with the following added columns:

| New Column | Description |
|---|---|
| `order_date_clean` | Standardized order date (YYYY-MM-DD) |
| `ship_date_clean` | Standardized ship date (YYYY-MM-DD) |
| `shipping_delay_days` | Days between order and ship date |
| `cleaned_discount` | Validated discount (negatives/high values kept for flagging) |
| `calculated_sales` | Quantity × Unit Price × (1 − cleaned_discount) |
| `calculated_profit` | calculated_sales − cost |
| `profit_margin` | calculated_profit ÷ calculated_sales |
| `order_month` | Month extracted from order_date_clean |
| `order_year` | Year extracted from order_date_clean |
| `discount_flag` | Flags invalid or missing discount values |
| `shipping_date_flag` | Flags ship-before-order records |
| `duplicate_flag` | Flags conflicting duplicate order IDs |
| `data_quality_flag` | Master flag: CLEAN / WARNING / INVALID |

### data_quality_report.xlsx
Three sheets:
- **Summary** — counts of every issue category found
- **Duplicate Details** — all rows with duplicate order IDs listed for manual review
- **Invalid Discounts** — all 30 records with negative or above-50% discounts

### pivot_summary.xlsx
Seven sheets:

| Sheet | Analysis |
|---|---|
| 1_Region | Sales & profit by region (completed orders, sorted by sales) |
| 2_Category | Sales & profit by category and sub-category |
| 3_ShipMode | Order count and avg delay by ship mode |
| 4_Segment | Profit margin by customer segment |
| 5_Issues_by_Region | Cancelled / returned / failed orders by region |
| 6_Monthly_Trend | Monthly sales trend (chronological) |
| 7_Refunded | Returned orders separately summarized by region |

---

## Data Quality Flag Summary

| Flag | Records | Meaning |
|---|---|---|
| CLEAN | 790 | No issues detected |
| WARNING | 60 | Minor fill applied (region or ship_mode set to Unknown) |
| INVALID | 50 | Hard error — invalid discount or ship-before-order |

> **Note:** INVALID records are retained in `cleaned_orders.xlsx` with flags visible. They are excluded from completed-sales pivot summaries via the `order_status` filter.

---

## How to Use the Cleaned Data

1. Open `cleaned_orders.xlsx` for row-level analysis
2. Filter `data_quality_flag = "CLEAN"` for fully trusted records only
3. Filter `order_status = "Completed"` for sales performance analysis
4. Filter `order_status = "Returned"` for returns/refund analysis
5. Use `calculated_sales` and `calculated_profit` columns — not the original `sales` and `profit` columns — for any financial calculations
6. Refer to `data_quality_report.xlsx` for a full audit of issues found

---

## Tools Used

- Microsoft Excel (TRIM, PROPER, SUBSTITUTE, DATEVALUE, MONTH, YEAR, COUNTIF, IF, ISNUMBER, ISBLANK, PivotTable)
- Python / pandas / openpyxl (for automated pipeline validation and output generation)

---
