# Team topics — 2026-10-06 queue (5 new post topics)

Posts dated 2026-10-06 PT. No tasks/search-stats.md exists (Search Console not
connected): no search data, topics chosen on the 3-gate test only. Each was checked
against every title/slug in posts/ready and posts/published and against
topics-seed.md, team-topics-2026-09-29.md and team-topics-2026-10-02.md. No tax,
insurance, or legal judgment in any of them. Build each template in
scripts/build-templates.py and verify by recalculating in LibreOffice Calc
(`soffice --headless --convert-to xlsx`, read back with openpyxl data_only). Only claim
LibreOffice in `tested_in`. Titles and hooks go to hook-writer.

Formula mix is new to the blog: PMT, percent change, RANK, TEXT, TEXTJOIN.

Ranked best first.

1. Loan payment calculator (PMT)
- Slug: `loan-payment-calculator-pmt-excel-google-sheets`
- Brief: Monthly payment on an equipment or van loan with `=PMT(Rate/12,Years*12,-Amount)`, plus total paid and total interest.
- Target keyword: loan payment calculator excel PMT
- Template: tidy-tabs-loan-payment-calculator.xlsx (inputs: $15,000.00, 6.50%, 5 years; payment, total paid, total interest; 3 side-by-side scenarios)
- Verify: hand-calc payment with P*r/(1-(1+r)^-n) (expected $293.49 for the sample), compare to recalculated PMT; check sign handling (negative amount gives positive payment).
- Competition: Microsoft PMT help page, Investopedia, bank calculators; twist on small-business equipment purchase with a downloadable comparison of three terms. Numbers are illustrations only, no lending advice.

2. Month-over-month percent change
- Slug: `percent-change-month-over-month-excel-google-sheets`
- Brief: Sales growth with `=IF(B2=0,"",(C2-B2)/B2)` formatted as %, plus a fix for the common "divide by the wrong month" mistake.
- Target keyword: percent change formula excel
- Template: tidy-tabs-percent-change.xlsx (12 months of fictional Etsy shop revenue, change in $ and in %, negative change shown in red)
- Verify: recalculate; hand-check Jan $2,400.00 to Feb $2,760.00 equals 15.0%; confirm the zero-prior-month row stays blank.
- Competition: Microsoft and ExcelJet percent-change pages; twist on monthly shop revenue with the zero-month case.

3. Rank top customers
- Slug: `rank-top-customers-excel-google-sheets`
- Brief: Rank customers by yearly total with `=RANK(B2,$B$2:$B$11,0)`, and break ties with `+COUNTIF($B$2:B2,B2)-1`.
- Target keyword: rank values in excel
- Template: tidy-tabs-rank-top-customers.xlsx (10 fictional clients with totals, rank column, tie-broken rank column, 2 deliberate ties)
- Verify: recalculate and compare to a hand-sorted list; confirm tied totals get the same RANK and unique ranks after the tie-break.
- Competition: Microsoft RANK help, ExcelJet; twist on freelancer client revenue and ties. Avoid repeating the SUMIFS steps from the Etsy and aging posts; totals are typed in.

4. Invoice number generator
- Slug: `invoice-number-generator-text-excel-google-sheets`
- Brief: Auto-numbered IDs like INV-2026-0007 with `="INV-"&TEXT(B2,"yyyy")&"-"&TEXT(ROWS($A$2:A2),"0000")`.
- Target keyword: automatic invoice number excel
- Template: tidy-tabs-invoice-number-generator.xlsx (invoice dates, generated numbers, 12 fictional rows)
- Verify: recalculate and read back as text; check zero padding, and that the year follows the date in column B. TEXT format codes checked in LibreOffice only (locale-dependent in other apps, say so).
- Competition: low; mostly forum answers and VBA solutions, so a formula-only version works. Links to the existing invoice template post without repeating it.

5. Combine address columns (TEXTJOIN)
- Slug: `combine-address-columns-textjoin-excel-google-sheets`
- Brief: Join street, unit, city and state into one line, skipping blank cells, with `=TEXTJOIN(", ",TRUE,A2:D2)`.
- Target keyword: combine columns excel skip blanks textjoin
- Template: tidy-tabs-combine-address-columns.xlsx (8 fictional addresses, some without a unit, joined column, ampersand version for comparison)
- Verify: recalculate and compare against the ampersand version, which leaves stray ", ," for blank cells; check TRUE vs FALSE on the ignore-empty argument.
- Competition: Microsoft TEXTJOIN help, ExcelJet; twist on mailing labels. It is the reverse of the published split city/state/ZIP post, so link to it.
