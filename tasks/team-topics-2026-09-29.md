# Team topics — 2026-09-29 queue (5 new post topics)

Handed to script-writers. Titles, slugs, and hooks are in English per the
style guide. Templates are built and formula-checked in LibreOffice Calc
only — see `tasks/briefs.md` (2026-09-28 section) for the exact commands
and results. Do not write or imply "tested in Excel/Google Sheets"
anywhere; use the `Works in` line below as the post's `tested_in`.

These are genuinely new topics: checked against every `title:` in
`posts/ready/*.md` and `posts/published/*.md`, and against
`tasks/topics-seed.md` (fully used up). `tasks/search-stats.md` does not
exist yet (Search Console not connected), so these were picked on the
3-gate test alone, not on search data.

---

## 1. How to Create a Dependent Drop-Down List in Excel and Google Sheets

- **Slug:** `dependent-dropdown-list-excel-google-sheets`
- **First paragraph (hook):** Make a second dropdown follow the first by naming
  each list after its category and pointing the second cell's data validation
  at `=INDIRECT(A2)`; picking "Candles" in column A then limits column B to
  only candle subcategories.
- **`> Works in:`** LibreOffice Calc 24.2.7 (Linux). Google Sheets steps from
  Google Docs Editors Help (not tested here); Excel steps from Microsoft
  Support (not tested here).
- **Template:** `templates/tidy-tabs-dependent-dropdown-list.xlsx` (new;
  built by `dependent_dropdown()` in `scripts/build-templates.py`). Two
  tabs: **Order form** (Category dropdown + dependent Subcategory dropdown)
  and **Lists** (one column per category, each column a named range with the
  same name as its header).
- **What was verified:** Clicking the actual dropdown UI can't be reproduced
  headlessly, so the underlying mechanism was checked instead: a test
  formula `=COUNTA(INDIRECT(A2))` was placed next to each of the 4 sample
  categories. After forcing recalculation
  (`soffice --headless --convert-to xlsx`), the formula returned the exact
  item count of each category's named range (Candles→3, Soap→2, Jewelry→4,
  Cards→2), confirming INDIRECT resolves the right named range for every
  category. Script-writer: say plainly that the dropdown-click itself
  wasn't screen-tested, only the INDIRECT/named-range mechanism behind it.

---

## 2. How to Build a Client Contact List Template in Google Sheets

- **Slug:** `client-contact-list-template-google-sheets`
- **First paragraph (hook):** Track when you last talked to a client and flag
  who's due for a follow-up with one formula: `=IF(TODAY()>=E2,"Follow up
  now","OK")`, where E2 is Last contact plus your follow-up interval in days.
- **`> Works in:`** LibreOffice Calc 24.2.7 (Linux), recalculated 09/28/2026.
  Google Sheets steps from Google Docs Editors Help (not tested here).
- **Template:** `templates/tidy-tabs-client-contact-list.xlsx` (new; built by
  `client_contact_list()`). One tab, **Clients**: Client, Company, Last
  contact, Follow up every (days), Next follow-up, Status, Notes. Status
  cells turn red via conditional formatting when overdue.
- **What was verified:** Forced recalculation on 09/28/2026 and read back
  Status for all 4 sample rows: Dana Ruiz (last contact 09/01 + 21 days =
  09/22, overdue) → "Follow up now"; Emma Lee (09/20 + 14 = 10/04, not yet)
  → "OK"; Sam Patel (08/15 + 30 = 09/14, overdue) → "Follow up now"; Jamie
  Chen (09/25 + 7 = 10/02, not yet) → "OK". All 4 matched hand-calculated
  expectations. Note for the post: this uses `TODAY()`, so results change
  depending on the day the reader opens the file — say that explicitly,
  the way the craft inventory and project tracker posts already do.

---

## 3. How to Build a Packing Slip Template in Google Sheets

- **Slug:** `packing-slip-template-google-sheets`
- **First paragraph (hook):** Build a one-page packing slip with an Order #,
  a Ship to address, and an items table that totals itself with
  `=SUM(D11:D14)` under Qty packed, then set the page to US Letter,
  portrait, fit to one page.
- **`> Works in:`** LibreOffice Calc 24.2.7 (Linux). Google Sheets print
  steps from Google Docs Editors Help (not tested here).
- **Template:** `templates/tidy-tabs-packing-slip-template.xlsx` (new; built
  by `packing_slip()`). One printable tab (**Packing slip**) plus a
  **How to use** notes tab.
- **What was verified:** Total packed formula recalculated to `9`
  (2+3+1+3), matching the sample quantities. Print check:
  `soffice --headless --convert-to pdf` on the full workbook produced a
  2-page PDF (page 2 is the non-printed "How to use" notes tab, same as
  every other template in this repo), so a second file was made with only
  the Packing slip tab and converted the same way — that one PDF is exactly
  1 page, 612x792 pt (US Letter), portrait, per `pdfinfo`. Post should note
  readers only print the Packing slip tab.

---

## 4. How to Add a Running Balance Column in Excel and Google Sheets

- **Slug:** `running-balance-column-excel-google-sheets`
- **First paragraph (hook):** Keep a running balance by adding each row's
  money in and subtracting money out from the balance above it:
  `=E2+C3-D3` copied down, starting from an opening balance you type
  directly into the first row.
- **`> Works in:`** LibreOffice Calc 24.2.7 (Linux). Excel/Google Sheets
  steps from Microsoft Support / Google Docs Editors Help (not tested here).
- **Template:** `templates/tidy-tabs-running-balance-cash-log.xlsx` (new;
  built by `running_balance()`). One tab, **Cash log**: Date, Description,
  Money in, Money out, Balance — sample data is a craft-fair/farmers-market
  cash box.
- **What was verified:** Forced recalculation and read back the Balance
  column: 250.00 (opening) → 430.00 (+180 market sales) → 384.50 (-45.50
  supplies) → 480.70 (+96.20 Etsy payout) → 445.70 (-35.00 booth fee).
  Every step matched hand arithmetic exactly.

---

## 5. How to Highlight Duplicate Values in Excel and Google Sheets

- **Slug:** `highlight-duplicate-values-excel-google-sheets`
- **First paragraph (hook):** Select the column, open conditional formatting,
  and use the custom formula `=COUNTIF($B$2:$B$16,B2)>1` to highlight every
  value that repeats, without deleting or moving any rows.
- **`> Works in:`** LibreOffice Calc 24.2.7 (Linux). Google Sheets/Excel
  conditional-formatting steps from Google Docs Editors Help / Microsoft
  Support (not tested here).
- **Template:** `templates/tidy-tabs-highlight-duplicates-sample.xlsx` (new;
  built by `highlight_duplicates()`). One tab, **Orders**: Order #, Customer
  email, and a `Duplicate?` formula column, with the matching conditional
  formatting rule already applied to column B.
- **What was verified:** Forced recalculation and read back column C for
  all 15 sample rows: the 4 repeated emails (priya.nair@example.com x3,
  sam.patel@example.com x3, dana.ruiz@example.com x2,
  ana.gomez@example.com x2 — back-to-back, an accidental double entry) all
  came back "Duplicate"; the other 7 unique emails came back blank. Matches
  what was expected before the run.
- Distinct from the already-published "Remove Duplicate Rows" post:
  this one flags repeats in place with conditional formatting and keeps
  every row (useful when a repeat is a legitimate second order from the
  same customer), instead of deleting rows.

---

## Notes for whoever writes these up

- `scripts/build-templates.py` was extended with 5 new functions
  (`dependent_dropdown`, `client_contact_list`, `packing_slip`,
  `running_balance`, `highlight_duplicates`) and re-run; all 5 new `.xlsx`
  files are already in `templates/`.
- Environment note: this session found `libreoffice-calc` was not installed
  (only `libreoffice-core`/`libreoffice-common` were present), which made
  every `soffice --convert-to` call silently fail with "source file could
  not be loaded". It was installed with `apt-get install -y
  libreoffice-calc` before any of the verification above. If a future
  session hits the same "source file could not be loaded" error, check
  `dpkg -l | grep libreoffice-calc` first.
