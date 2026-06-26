# Cleaning Log — Orders Dataset

**Source file:** `cleaned_orders.xlsx` (sheet `raw_orders`, 932 records, 21 columns)
**Output file:** `cleaned_orders.xlsx` (sheet `cleaned_orders`, 912 records, 38 columns)
**Tools used:** Python (pandas, openpyxl) for parsing/standardization logic; Excel formulas (via openpyxl) for all derived/calculated columns so the workbook stays dynamic; LibreOffice headless recalculation to verify zero formula errors.

The raw `raw_orders` sheet was never modified — it is preserved unchanged inside `cleaned_orders.xlsx` for audit purposes, per the project's "do not overwrite raw data" rule.

---

## 1. Issues Found (Discovery Summary)

| # | Issue | Detail |
|---|-------|--------|
| 1 | Inconsistent text casing | `segment`, `region`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`, `customer_name` contained UPPER, lower, and Title case versions of the same value (e.g. `EAST`, `East`, `east`) |
| 2 | Double/internal spacing | Values such as `"Aarav  Sharma"`, `"Small  Business"`, `"Office  Supplies"`, `"Standard  Class"` had double spaces |
| 3 | Missing `region` | 26 raw records (25 after exact-dup removal) had a blank region |
| 4 | Missing `ship_mode` | 22 raw records (21 after exact-dup removal) had a blank ship_mode |
| 5 | Missing `discount` | 18 records had no discount value at all |
| 6 | Discount stored inconsistently | Mostly decimals (`0.2`), but 8 records used a percentage string (`"70%"`, `"85%"`) |
| 7 | Negative discounts | 15 records had a negative discount (e.g. `-0.19`) — not business-valid |
| 8 | Discount above allowed range | 15 records had discount > 30% (`0.55`, `0.65`, `70%`, `85%`) |
| 9 | Mixed date formats | `order_date`/`ship_date` arrived in 4 different formats: `D Mon YYYY`, `MM/DD/YYYY`, `DD-MM-YYYY`, `YYYY-MM-DD` |
| 10 | Ship date before order date | 22 raw records (21 after dedup) had a ship_date earlier than the order_date |
| 11 | Exact duplicate rows | 20 rows were 100% identical copies (every column matched another row) |
| 12 | Conflicting duplicate order_ids | 12 order_ids appeared twice with differing dates, sales, status, or other fields (24 records involved) |
| 13 | Sales/profit formula mismatches | 20 records (with otherwise valid discount) had `sales` and `profit` values that did not equal `quantity × unit_price × (1 − discount)` and `sales − cost` respectively |
| 14 | Logically inconsistent status | 2 records were marked `order_status = Completed` while `payment_status = Failed` — not logically possible |

No missing `order_date`/`ship_date` values and no unparseable/non-date text values were found in this dataset.

---

## 2. Cleaning Actions Performed

### Text fields (`customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`)
- Trimmed leading/trailing whitespace (equivalent to Excel `TRIM`).
- Collapsed multiple internal spaces to a single space (equivalent to nested `SUBSTITUTE(text," ","")` style cleanup, or Excel's `TRIM` after replacing double-spaces).
- Standardized casing to Title Case (equivalent to Excel `PROPER()`), e.g. `EAST`→`East`, `small  business`→`Small Business`.
- Checked for stray special characters (none found beyond normal letters/spaces/hyphens — no further substitution needed).
- Verified `category`/`sub_category` had no true spelling variants beyond casing/spacing (no synonym mapping was required after normalization).

### Dates (`order_date`, `ship_date`)
- Detected 4 source formats and parsed each value with the matching format (`%d %b %Y`, `%m/%d/%Y`, `%d-%m-%Y`, `%Y-%m-%d`).
- Disambiguated `MM/DD/YYYY` vs `DD/MM/YYYY` by confirming the first slash-component never exceeds 12 across the whole column, consistent with US-style month-first input.
- Converted all dates to a single standard format (`YYYY-MM-DD`) in the cleaned sheet.
- Flagged records where `ship_date < order_date` as `date_sequence_flag = "Invalid - Ship Before Order"`.
- Calculated `shipping_delay_days = ship_date − order_date` as a live Excel formula.

### Discount
- Parsed percentage-string values (`"70%"`) into decimal form (`0.70`).
- Classified every value into `discount_status`: `Valid` (0%–30%), `Missing`, `Negative - Invalid`, `Above Allowed Range - Invalid`.
- Original raw text (`discount_original`) and parsed numeric (`discount_value`) are both retained for audit; `cleaned_discount` is a formula that uses the parsed value only when `Valid`, otherwise defaults to 0.

### Duplicates
- Step 1: Removed rows that were **100% identical** across all 21 source columns (20 rows), keeping the first occurrence. These are logged on the `removed_exact_duplicates` sheet.
- Step 2: For order_ids that still repeated after Step 1 (12 order_ids / 24 rows), **no rows were deleted**. Each is flagged `duplicate_status = "Conflicting Duplicate - Flagged for Review"` so a human can decide which version is correct.

### Calculated columns (all implemented as live Excel formulas referencing the cleaned columns)
- `cleaned_discount`, `calculated_sales`, `calculated_profit`, `profit_margin`, `shipping_delay_days`, `date_sequence_flag`, `order_month`, `order_year`, `sales_mismatch_flag`, `profit_mismatch_flag`, `order_status_issue_flag`, `data_quality_flag`.
- `data_quality_flag` rolls every check into one of three states:
  - **Invalid** — invalid/out-of-range discount, ship-before-order, conflicting duplicate, sales/profit mismatch, or Completed+Failed status inconsistency.
  - **Warning** — region/ship_mode was missing and filled, or discount was missing and defaulted to 0 (none of which are "wrong", just incomplete source data).
  - **Clean** — none of the above.

---

## 3. Business Rules Applied

| Rule Area | Action Taken |
|---|---|
| Missing `region` | Filled as `"Unknown"`; `region_was_missing = Yes` flag set; counted in Warning records |
| Missing `ship_mode` | Filled as `"Unknown"`; `ship_mode_was_missing = Yes` flag set; counted in Warning records |
| Missing `discount` | Verified `quantity`, `unit_price`, `sales`, `cost`, `profit` were all present and valid for every one of the 18 affected records, then defaulted `cleaned_discount` to 0 |
| Negative discount | Flagged `discount_status = "Negative - Invalid"`; `cleaned_discount` forced to 0 for calculation purposes; original value preserved for audit |
| Discount above allowed range (>30%) | Flagged `discount_status = "Above Allowed Range - Invalid"`; `cleaned_discount` forced to 0; original value preserved |
| Cancelled orders | Excluded from the completed-sales pivot summaries (`order_status != "Completed"`) |
| Failed payments | Excluded from the completed-sales pivot summaries (`payment_status != "Paid"`) |
| Refunded orders | Summarized **separately** by region in `pivot_summary.xlsx` (`NonCompleted_by_Region` sheet) rather than folded into completed sales |
| Ship date before order date | Flagged `date_sequence_flag = "Invalid - Ship Before Order"`; record kept, not deleted |

**"Completed sales" definition (assumption):** a record counts toward the completed-sales summaries only if `order_status = "Completed"` **and** `payment_status = "Paid"`. This was chosen because the rules state both Cancelled orders and Failed payments must be excluded — applying both filters together is the strictest, safest reading. `Returned` orders are excluded from completed sales by definition (`order_status != "Completed"`), even when payment was originally `Paid`, since the sale was effectively reversed.

---

## 4. Assumptions Made

1. **Allowed discount range:** 0%–30%. This was inferred from the distribution of valid values in the data (0%, 5%, 10%, 15%, 20%, 25% all occur frequently and form a clean step pattern); 30% was chosen as a safety ceiling slightly above the highest "normal" value (25%) to avoid being overly strict, while values such as 55%, 65%, 70%, 85% sit far outside this cluster and were treated as data-entry errors.
2. **Date format resolution:** Since `MM/DD/YYYY` and `DD/MM/YYYY` look identical when day ≤ 12, the format was inferred by scanning the *entire* column: the first component never exceeded 12, while the second component did — confirming a consistent **month-first (US-style)** convention for all `/`-separated dates.
3. **Invalid/missing discount → 0 for calculations:** Rather than guessing a "true" discount, invalid or missing discounts are treated as 0% in `calculated_sales`/`calculated_profit` so the formulas remain conservative and auditable. The original flag is preserved so these records are never silently treated as fully "clean."
4. **"Completed sales" = Completed status AND Paid payment** (see above) — the rules named two separate exclusion criteria (Cancelled, Failed) which combine into this filter.
5. **Customer identity:** `customer_id` does not map 1:1 to `customer_name` in this dataset (the same name appears under dozens of different IDs), so `customer_id` was treated as an order-level reference field only, not a stable customer key, and was left untouched.

---

## 5. Records Removed

- **20 rows removed** — exact full-row duplicates (kept the first occurrence of each). Listed in full on the `removed_exact_duplicates` sheet of `cleaned_orders.xlsx`.
- No other rows were deleted. Total: 932 raw → 912 cleaned.

## 6. Records Flagged (not removed)

- **24 records** — conflicting duplicate order_ids (12 groups), flagged for manual review.
- **30 records** — invalid discount (15 negative, 15 above allowed range).
- **18 records** — missing discount, defaulted to 0.
- **25 records** — missing region, filled as Unknown.
- **21 records** — missing ship_mode, filled as Unknown.
- **21 records** — ship_date earlier than order_date.
- **20 records** — sales/profit calculation mismatch despite a valid discount.
- **2 records** — Completed order status with a Failed payment status (logically inconsistent).
- **112 records total flagged `Invalid`**, **44 flagged `Warning`**, **756 flagged `Clean`** (`data_quality_flag` — some records trip more than one rule, so these category counts overlap with the bullets above; see `data_quality_report.xlsx` for the full breakdown).

## 7. Limitations of This Cleaning Process

- For the 30 invalid-discount and 24 conflicting-duplicate records, the *true* intended value cannot be recovered from the data alone — they are flagged for a human reviewer rather than auto-corrected.
- The 20 genuine sales/profit mismatches (valid discount, but reported sales/profit don't match the formula) could not be explained by any single rule; they may reflect manual overrides, rounding from a different source system, or entry errors upstream — they are flagged, not silently corrected.
- "Similar category names written differently" beyond case/spacing (e.g. true synonyms like "Tech" vs "Technology") were not present in this dataset, so no synonym-mapping table was needed — but the cleaning script would need to be extended with an explicit mapping dictionary if a future data load introduced true spelling variants.
- Date format disambiguation relies on a column-wide pattern (no day value > 12 for `/`-separated dates); if a future load contains a genuinely ambiguous single record, it would be parsed as month-first by default and should be spot-checked.
- `customer_id` inconsistency (many IDs per name) was noted but not resolved, since no rule required customer-identity reconciliation.
