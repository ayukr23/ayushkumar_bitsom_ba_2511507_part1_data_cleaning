# Cleaning Log — Retail Orders Dataset

**File:** raw_orders.xlsx  
**Cleaned Output:** cleaned_orders.xlsx  
**Total Raw Records:** 932  
**Total Clean Records:** 900  
**Log Format:** Decision → Rationale → Records Affected

---

## Task 2 — Text Field Cleaning

### Decision 2.1 — Remove Leading and Trailing Spaces
**Columns affected:** `customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`  
**Method:** TRIM() formula applied to all 10 columns  
**Rationale:** Trailing/leading spaces cause COUNTIF, VLOOKUP, and pivot grouping to treat identical values as distinct. For example, `"Paid "` and `"Paid"` would appear as two separate payment statuses in a pivot.  
**Records affected:** All 932 rows processed; spacing issues present across all 10 columns

---

### Decision 2.2 — Remove Internal Double Spaces
**Columns affected:** `segment`, `ship_mode`  
**Method:** SUBSTITUTE(value, "  ", " ") combined with TRIM()  
**Examples found:**
- `"Small  Business"` → `"Small Business"`
- `"Standard  Class"` → `"Standard Class"`
- `"Home  Office"` → `"Home Office"`

**Rationale:** TRIM() only removes leading/trailing spaces, not internal double spaces. SUBSTITUTE targets the internal occurrence directly.  
**Records affected:** ~18 records

---

### Decision 2.3 — Standardize to Title Case
**Columns affected:** All 10 text columns  
**Method:** PROPER() formula  
**Examples resolved:**

| Raw Value | Cleaned Value |
|---|---|
| `COMPLETED` | `Completed` |
| `completed` | `Completed` |
| `  Completed ` | `Completed` |
| `CANCELLED` | `Cancelled` |
| `cancelled` | `Cancelled` |
| `PENDING` | `Pending` |
| `failed` | `Failed` |
| `SMALL BUSINESS` | `Small Business` |
| `HOME OFFICE` | `Home Office` |
| `STANDARD CLASS` | `Standard Class` |
| `FIRST CLASS` | `First Class` |
| `  North ` | `North` |
| `NORTH` | `North` |
| `WEST` | `West` |
| `EAST` | `East` |
| `Corporate ` | `Corporate` |
| `Paid ` | `Paid` |

**Rationale:** Consistent title case ensures all pivot groupings, filters, and COUNTIF formulas work correctly. Mixed case creates phantom duplicate categories.  
**Records affected:** Variants found across all 10 columns

---

## Task 3 — Date Cleaning and Validation

### Decision 3.1 — Standardize Date Formats to YYYY-MM-DD
**Columns affected:** `order_date`, `ship_date`  
**Problem:** Five different date formats found in the same columns:

| Format Example | Format Code |
|---|---|
| `21 Jul 2024` | DD MMM YYYY |
| `07/27/2024` | MM/DD/YYYY |
| `2024-05-24` | YYYY-MM-DD |
| `28-11-2024` | DD-MM-YYYY |
| `05 Sep 2024` | DD MMM YYYY |

**Method:** DATEVALUE() for text-stored dates; Text to Columns (Data → Text to Columns → Date: DMY) for hyphen-separated formats; Flash Fill for pattern-based conversion  
**Output columns created:** `order_date_clean`, `ship_date_clean` (both formatted YYYY-MM-DD)  
**Rationale:** A single consistent date format is required for MONTH(), YEAR(), and date arithmetic to work reliably. Mixing formats causes Excel to misread dates or treat them as text.  
**Records affected:** All 932 rows reprocessed

---

### Decision 3.2 — Create `shipping_delay_days` Column
**Formula:** `= ship_date_clean − order_date_clean`  
**Output:** Integer number of days between order and ship date  
**Rationale:** Task requirement. Also used to identify invalid shipping records (negative values).  
**Records affected:** All 900 records in cleaned output  
**Note:** Negative values indicate ship date before order date — these are separately flagged (see Decision 3.3)

---

### Decision 3.3 — Flag Ship Date Before Order Date
**Column created:** `shipping_date_flag`  
**Flag value:** `INVALID_SHIP_BEFORE_ORDER`  
**Condition:** `shipping_delay_days < 0`  
**Records flagged:** 20  
**Rationale:** A ship date earlier than the order date is a logical impossibility and indicates a data entry error. These records are flagged as INVALID in `data_quality_flag` and excluded from completed-sales analysis.  
**Action taken:** Records retained with flag. Not deleted. Require investigation by the data entry or logistics team.

---

## Task 4 — Duplicate Handling

### Decision 4.1 — Remove Exact Duplicate Rows
**Method:** Data → Remove Duplicates → Select All Columns → OK  
**Records removed:** 20  
**Rationale:** Rows that are identical across every column represent the same transaction entered twice. Keeping them would inflate sales totals and order counts. First occurrence is kept; subsequent identical rows are deleted.  
**Documentation:** Count of 20 recorded in `data_quality_report.xlsx` → Summary sheet

---

### Decision 4.2 — Flag Conflicting Duplicate Order IDs
**Method:** COUNTIF($A$2:A2, A2) > 1 to identify 2nd+ occurrences of the same order_id  
**Column created:** `duplicate_flag`  
**Flag value:** `DUPLICATE_FLAGGED_FOR_REVIEW`  
**Duplicate order IDs found:** 12 unique order IDs with conflicting data  
**Records flagged:** 12 (the non-first occurrences)  
**Rationale:** These rows share an order_id but differ in at least one field (e.g. different sales amount, different status). This suggests a system sync error or a manual re-entry with updated values. They cannot be silently deleted — the correct version must be verified against the source system.  
**Action taken:** First occurrence of each order_id retained in cleaned data. Conflicting duplicates moved to `data_quality_report.xlsx` → Duplicate Details sheet for human review.

---

## Task 5 — Business Rules

### Decision 5.1 — Missing `region` → Fill as "Unknown"
**Rule:** Missing region values filled with the string `Unknown`  
**Method:** `=IF(OR(G2="", ISBLANK(G2)), "Unknown", G2)`  
**Records affected:** 25  
**Rationale:** Per business rule. Region is required for geographic analysis. Blank values would be excluded from regional pivots, causing undercounting. "Unknown" is a named bucket that keeps the record in analysis while making the gap visible.  
**Flag in quality report:** These records receive `data_quality_flag = WARNING`

---

### Decision 5.2 — Missing `ship_mode` → Fill as "Unknown"
**Rule:** Missing ship_mode values filled with the string `Unknown`  
**Method:** `=IF(OR(M2="", ISBLANK(M2)), "Unknown", M2)`  
**Records affected:** 20  
**Rationale:** Same rationale as Decision 5.1. Ship mode is required for logistics analysis (Pivot 3). Blanks excluded from pivot groupings.  
**Flag in quality report:** These records receive `data_quality_flag = WARNING`

---

### Decision 5.3 — Missing `discount` → Treat as 0 (if sales fields valid)
**Rule:** If discount is blank AND quantity and unit_price are both valid numbers, set cleaned_discount = 0  
**Method:**
```
=IF(AND(ISBLANK(P2), ISNUMBER(N2), ISNUMBER(O2)), 0, "")
```
**Records affected:** 18  
**Rationale:** Per business rule. A missing discount most likely means no discount was applied, not that data is corrupt. This is only safe to assume when the other sales fields are present and valid. Records are flagged as `DISCOUNT_MISSING_SET_TO_0` so the assumption is traceable.  
**Flag applied:** `discount_flag = DISCOUNT_MISSING_SET_TO_0`  
**Data quality flag:** WARNING

---

### Decision 5.4 — Negative Discount → Flag as Invalid
**Rule:** Any discount value below 0 is invalid  
**Records affected:** 15  
**Examples found:** `-0.19`, `-0.23`, `-0.09`, `-0.15`, `-0.16`, `-0.14`, `-0.06`, `-0.1`, `-0.13`  
**Rationale:** A negative discount implies the customer is being charged more than the list price. This has no valid business meaning in this dataset context and is treated as a data entry error (e.g. minus sign entered by mistake).  
**Flag applied:** `discount_flag = INVALID_NEGATIVE_DISCOUNT`  
**Data quality flag:** INVALID  
**Action taken:** Value retained as-is in `cleaned_discount`. Record flagged. Excluded from completed-sales pivot summaries.

---

### Decision 5.5 — Discount Above 50% → Flag as Invalid
**Rule:** Discount above 0.5 (50%) is outside the allowed range  
**Records affected:** 15  
**Examples found:** `0.55`, `0.65`, `"70%"`, `"85%"`  
**Pre-processing:** Text values `"70%"` and `"85%"` converted to `0.7` and `0.85` via Find & Replace before flagging  
**Rationale:** Discounts of 55%, 65%, 70%, 85% are implausible for a retail operation. These likely represent data entry errors (e.g. entering 70 instead of 0.07) or system export errors. Business threshold set at 50% maximum.  
**Flag applied:** `discount_flag = INVALID_DISCOUNT_ABOVE_50PCT`  
**Data quality flag:** INVALID  
**Action taken:** Value retained. Record flagged. Excluded from completed-sales pivot summaries.

---

### Decision 5.6 — Cancelled Orders Excluded from Sales Summary
**Rule:** Cancelled orders must not contribute to completed sales summary  
**Method:** `order_status` filter applied in all sales-related pivots (Pivot 1, 2, 4, 6) — filter set to `Completed` only  
**Records with Cancelled status:** Present in dataset  
**Rationale:** Cancelled orders represent transactions that did not complete. Including them would overstate revenue figures used in business review.  
**Action taken:** Records retained in `cleaned_orders.xlsx` with original status. Excluded at pivot level via filter. Not deleted.

---

### Decision 5.7 — Failed Payment Orders Excluded from Sales Summary
**Rule:** Failed payment orders must not contribute to completed sales summary  
**Method:** Same `order_status = Completed` filter as Decision 5.6  
**Rationale:** Failed payments mean revenue was not collected. Including them in sales totals misrepresents actual realized revenue.  
**Action taken:** Retained in full dataset. Excluded at pivot level. Visible in Pivot 5 (Issues by Region).

---

### Decision 5.8 — Returned/Refunded Orders Separately Summarized
**Rule:** Returned orders must be summarized separately, not mixed into completed sales  
**Method:** Dedicated pivot sheet `7_Refunded` in `pivot_summary.xlsx`, filtered to `order_status = Returned`  
**Rationale:** Returns represent revenue reversal. They are commercially significant but must not offset or be combined with forward sales. Separate summary allows leadership to see both the sales achievement and the returns exposure independently.  
**Action taken:** Pivot 7 created exclusively for returned orders, broken down by region.

---

### Decision 5.9 — Ship Date Before Order Date → Flag as Invalid Shipping Record
**Rule:** Ship date earlier than order date must be flagged as invalid  
**Already documented in:** Decision 3.3  
**Column:** `shipping_date_flag = INVALID_SHIP_BEFORE_ORDER`  
**Records affected:** 20  
**Data quality flag:** INVALID

---

## Task 6 — Calculated Columns

### Decision 6.1 — Use `calculated_sales` Instead of Raw `sales`
**Formula:** `= quantity × unit_price × (1 − cleaned_discount)`  
**Rationale:** The raw `sales` column was sourced from multiple internal systems with inconsistent calculation logic. Re-deriving sales from its component fields ensures consistency and catches records where the original sales figure did not match the formula. All pivot summaries use `calculated_sales`.

---

### Decision 6.2 — Use `calculated_profit` Instead of Raw `profit`
**Formula:** `= calculated_sales − cost`  
**Rationale:** Profit is derived from `calculated_sales` (not raw sales) to maintain internal consistency. Using raw profit with calculated sales would create a mixed-source figure that could misrepresent margins.

---

### Decision 6.3 — Divide-by-Zero Guard on `profit_margin`
**Formula:** `=IF(calculated_sales <> 0, calculated_profit / calculated_sales, 0)`  
**Rationale:** Prevents #DIV/0! errors for any records where calculated_sales resolves to zero (e.g. 100% discount). Margin is set to 0 in these cases.

---

### Decision 6.4 — `data_quality_flag` Logic and Priority
**Three-tier system:**

| Flag | Condition | Priority |
|---|---|---|
| INVALID | Negative discount OR discount > 50% OR ship before order | Highest |
| WARNING | Region = Unknown OR ship_mode = Unknown OR discount was missing | Medium |
| CLEAN | None of the above | Default |

**Rationale:** INVALID takes priority over WARNING. A record with both an invalid discount and a filled region is still INVALID, not WARNING. This ensures the master flag always reflects the worst issue present.  
**Formula:**
```
=IF(OR(discount_flag="INVALID_NEGATIVE_DISCOUNT",
       discount_flag="INVALID_DISCOUNT_ABOVE_50PCT",
       shipping_date_flag="INVALID_SHIP_BEFORE_ORDER"),
   "INVALID",
   IF(OR(region="Unknown",
         ship_mode="Unknown",
         discount_flag="DISCOUNT_MISSING_SET_TO_0"),
      "WARNING",
      "CLEAN"))
```

---

## Final Record Count Reconciliation

| Stage | Records | Notes |
|---|---|---|
| Raw input | 932 | raw_orders.xlsx |
| After removing exact duplicates | 912 | 20 removed |
| After flagging conflicting order IDs | 900 | 12 flagged, removed from main dataset |
| **Final cleaned output** | **900** | cleaned_orders.xlsx |

### Data Quality Distribution (900 records)

| Flag | Count | % of Total |
|---|---|---|
| CLEAN | 790 | 87.8% |
| WARNING | 60 | 6.7% |
| INVALID | 50 | 5.6% |

---

## Issues Requiring Follow-Up (Action Items for Data Team)

| # | Issue | Records | Recommended Action |
|---|---|---|---|
| 1 | Conflicting duplicate order IDs | 12 | Cross-check against source system; determine correct version |
| 2 | Invalid negative discounts | 15 | Verify with sales team — likely data entry error (missing decimal or sign) |
| 3 | Discounts above 50% | 15 | Verify with pricing team — check if special promotions apply |
| 4 | Ship date before order date | 20 | Verify with logistics team — possible date entry swap |
| 5 | Missing region (filled Unknown) | 25 | Enrich from customer master data or order address |
| 6 | Missing ship mode (filled Unknown) | 20 | Retrieve from shipping/logistics system |