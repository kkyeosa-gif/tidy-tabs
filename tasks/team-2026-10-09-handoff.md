# Handoff for posts dated 2026-10-09

Status: titles and openings (hook-writer step) done, templates built and recalculated in LibreOffice Calc 24.2 (Linux) only. Nothing was opened in Excel or Google Sheets, so `tested_in` must say LibreOffice Calc only. Sample data is fictional (say so once per post). Post Example tables must show the verified values below.

Replacement note: the researcher's topics 1, 2 and 5 from 2026-10-08 duplicated existing posts and were dropped (see tasks/briefs.md and tasks/topics-seed.md).

Hook-writer title candidates were weighed against the style guide: each title carries "Excel" and "Google Sheets", names the task and the function people search for, and has no colon or dash.

---

## 1. quantity-price-tiers-lookup-excel-google-sheets
- Title: How to Look Up Wholesale Price Tiers by Quantity in Excel and Google Sheets
- Opening: Use an approximate-match lookup on a sorted tier table: `=VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE)` returns the unit price for the last tier whose minimum quantity is less than or equal to the order quantity. `LOOKUP` and `INDEX` with `MATCH(...,1)` give the same price.
- Template: `templates/tidy-tabs-quantity-price-tiers.xlsx`. Sheets: Tiers, Orders, Unsorted demo, How to use.
- Tiers: minimum quantity 1 / 12 / 36 / 72 at $12.00 / $10.50 / $9.25 / $8.00.
- Formulas (Orders row 2):
  - `=VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE)`
  - `=LOOKUP(C2,Tiers!$A$2:$A$5,Tiers!$B$2:$B$5)`
  - `=INDEX(Tiers!$B$2:$B$5,MATCH(C2,Tiers!$A$2:$A$5,1))`
  - `=C2*D2` (order total)
  - Tiers!F3: `=IFERROR(VLOOKUP(F2,$A$2:$B$5,2,TRUE),"Below first tier")`
- Verified values (Example table):

| Quantity | Unit price | Order total |
|---|---|---|
| 1 | $12.00 | $12.00 |
| 11 | $12.00 | $132.00 |
| 12 | $10.50 | $126.00 |
| 35 | $10.50 | $367.50 |
| 36 | $9.25 | $333.00 |
| 71 | $9.25 | $656.75 |
| 72 | $8.00 | $576.00 |
| 500 | $8.00 | $4,000.00 |

- Notes: all three methods agree on all 8 rows. Quantity 0 gives #N/A with plain VLOOKUP and "Below first tier" with IFERROR. Unsorted demo tab (tiers in order 36, 1, 72, 12): LibreOffice returned #N/A for quantity 5 and 20, $12.00 for 40 (correct $9.25) and $10.50 for 80 (correct $8.00). Say only "wrong price or #N/A"; do not claim Excel or Sheets behave the same. Existing post `vlookup-price-list` covers exact match (FALSE); link to it. Microsoft/Google VLOOKUP help pages are the sources for Excel/Sheets behavior.

---

## 2. weighted-average-cost-sumproduct-excel-google-sheets
- Title: How to Calculate Weighted Average Cost With SUMPRODUCT in Excel and Google Sheets
- Opening: Divide the total money spent on an item by the total quantity bought: `=SUMPRODUCT((Purchases!$B$2:$B$40=A2)*Purchases!$C$2:$C$40*Purchases!$D$2:$D$40)/B2`, where B2 is that item's total quantity. This weighted average counts a 40 lb purchase more than a 10 lb one, unlike a plain average of the prices.
- Template: `templates/tidy-tabs-weighted-average-cost.xlsx`. Sheets: Purchases, Summary, How to use.
- Formulas:
  - Purchases E2: `=C2*D2`
  - Summary B2: `=SUMIFS(Purchases!$C$2:$C$40,Purchases!$B$2:$B$40,A2)`
  - Summary C2: `=SUMIFS(Purchases!$E$2:$E$40,Purchases!$B$2:$B$40,A2)`
  - Summary D2: `=C2/B2`
  - Summary E2: `=SUMPRODUCT((Purchases!$B$2:$B$40=A2)*Purchases!$C$2:$C$40*Purchases!$D$2:$D$40)/B2`
  - Summary F2: `=AVERAGEIFS(Purchases!$D$2:$D$40,Purchases!$B$2:$B$40,A2)`
- Purchases (fictional): soy wax 10 lb at $4.00 (07/06/2026), 40 lb at $3.50 (08/03/2026), 25 lb at $3.80 (09/07/2026); cotton wicks 200 at $0.12, 500 at $0.09, 300 at $0.10; glass jars 48 at $1.85, 96 at $1.60, 24 at $2.10.
- Verified values (Example table):

| Item | Total quantity | Total spent | Weighted average | Simple average |
|---|---|---|---|---|
| Soy wax (lb) | 75 | $275.00 | $3.6667 | $3.7667 |
| Cotton wicks | 1,000 | $99.00 | $0.0990 | $0.1033 |
| Glass jars | 168 | $292.80 | $1.7429 | $1.8500 |

- Notes: SUMPRODUCT and SUMIFS methods return identical values. Short hand check in the post: soy wax first two buys, 10 lb x $4.00 + 40 lb x $3.50 = $180.00 over 50 lb = $3.60 (simple average $3.75); the full table above uses all three buys. It is an average purchase cost, not an inventory valuation method or tax advice; say so. Show costs to 4 decimals for wicks, otherwise the difference disappears.

---

## 3. last-order-date-per-customer-maxifs-excel-google-sheets
- Title: How to Find the Last Order Date for Each Customer With MAXIFS in Excel and Google Sheets
- Opening: Use `=MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A2)` to return the most recent order date for the customer in A2. Then pull that day's amount with `=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2,Orders!$A$2:$A$200,B2)`.
- Template: `templates/tidy-tabs-last-order-date-maxifs.xlsx`. Sheets: Orders, Summary, How to use.
- Formulas (Summary row 2; stored in the file as `_xlfn.MAXIFS`, which LibreOffice 24.2 calculated correctly):
  - B2: `=MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A2)`
  - C2: `=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2,Orders!$A$2:$A$200,B2)`
  - D2: `=$H$2-B2` (H2 report date = 10/09/2026)
  - E2: `=COUNTIFS(Orders!$B$2:$B$200,A2)`
- Verified values (Example table):

| Customer | Last order date | Last order amount | Days since (as of 10/09/2026) | Orders |
|---|---|---|---|---|
| Harbor Coffee | 09/17/2026 | $865.00 | 22 | 3 |
| Maple Street Bakery | 09/30/2026 | $428.40 | 9 | 3 |
| Oak & Ember Candles | 09/09/2026 | $510.75 | 30 | 2 |
| Lakeview Florist | 09/24/2026 | $242.00 | 15 | 2 |

- Notes: two orders from one customer on the same day make the SUMIFS amount the sum of both (documented in the template). A customer with no orders returns 0 in LibreOffice (shows as a 1899/1900 date in Excel formatting terms; do not describe Excel's exact display as tested). MAXIFS exists in Excel 2019 and Microsoft 365 and in Google Sheets, not Excel 2016 or earlier (cite the Microsoft MAXIFS page; not tested in Excel). Text dates are ignored.

---

## 4. count-orders-by-month-countifs-excel-google-sheets
- Title: How to Count Orders by Month With COUNTIFS in Excel and Google Sheets
- Opening: Count the orders in a month with two date tests: `=COUNTIFS(Orders!$A$2:$A$200,">="&A2,Orders!$A$2:$A$200,"<"&EDATE(A2,1))`, where A2 is the first day of the month. Swap `COUNTIFS` for `SUMIFS` and add the amount column to total the revenue.
- Template: `templates/tidy-tabs-count-orders-by-month.xlsx`. Sheets: Orders, Monthly, How to use.
- Formulas (Monthly row 2):
  - B2: `=COUNTIFS(Orders!$A$2:$A$200,">="&A2,Orders!$A$2:$A$200,"<"&EDATE(A2,1))`
  - C2: `=SUMIFS(Orders!$C$2:$C$200,Orders!$A$2:$A$200,">="&A2,Orders!$A$2:$A$200,"<"&EDATE(A2,1))`
  - Totals: B6 `=SUM(B2:B5)`, B7 `=COUNT(Orders!A2:A200)`, C7 `=SUM(Orders!C2:C200)`
- Verified values (Example table):

| Month start | Orders | Revenue |
|---|---|---|
| 07/01/2026 | 2 | $365.50 |
| 08/01/2026 | 3 | $549.25 |
| 09/01/2026 | 3 | $693.75 |
| 10/01/2026 | 2 | $395.50 |
| Total | 10 | $2,004.00 |

- Notes: 08/31/2026 lands in August and 09/01/2026 in September (edge rows in the sample). The "before the first of next month" test is explained as the reason it also handles dates with a time, but that was not tested; do not say it was. Rows 6 and 7 on the Monthly tab agree (10 and $2,004.00). Text dates are skipped by the criteria.

---

## 5. average-order-value-by-channel-averageifs-excel-google-sheets
- Title: How to Find Average Order Value by Sales Channel With AVERAGEIFS in Excel and Google Sheets
- Opening: Use `=AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2)` to get the average order amount for the channel named in A2, such as Etsy, Shopify or Farmers market. Add `=COUNTIFS(Orders!$B$2:$B$200,A2)` next to it to show how many orders each average is based on.
- Template: `templates/tidy-tabs-average-order-value-by-channel.xlsx`. Sheets: Orders, By channel, How to use.
- Formulas (By channel row 2):
  - B2: `=COUNTIFS(Orders!$B$2:$B$200,A2)`
  - C2: `=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2)`
  - D2: `=AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2)`
  - E2: `=C2/B2` (check, equals D2)
  - D5 (all channels): `=AVERAGE(Orders!C2:C200)`
- Verified values (Example table):

| Channel | Orders | Revenue | Average order value |
|---|---|---|---|
| Etsy | 3 | $115.50 | $38.50 |
| Shopify | 3 | $219.50 | $73.17 |
| Farmers market | 4 | $116.00 | $29.00 |
| All channels | 10 | $451.00 | $45.10 |

- Notes: Shopify average is 73.1666..., show $73.17. The all-channel $45.10 is the average of the 10 orders, not the average of the three channel averages (which would be $46.89; arithmetic only, not a template cell). A channel with no orders makes AVERAGEIFS return #DIV/0! (documented in the template, not reproduced in a test). Keep channel names spelled identically.

---

## Files changed
- `scripts/build-templates.py` (5 new functions, registered in `__main__`)
- 5 new templates in `templates/` (listed above); no existing template binary was changed (rebuild timestamp-only diffs were reverted with `git checkout -- templates`)
- `tasks/briefs.md`, `tasks/topics-seed.md` (replacement note, 3 dropped entries removed or struck)
- Nothing committed.
