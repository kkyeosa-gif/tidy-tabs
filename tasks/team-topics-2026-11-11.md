# Team topics - 2026-11-11 queue (5 new post topics)

Posts dated 2026-11-11. tasks/search-stats.md does not exist: no data (no search volume or competition numbers; none are claimed). Topics were chosen on the 3-gate test only (buildable and checkable by us in LibreOffice Calc, a download that gives a reason to visit, no tax, wage-law or legal judgment).

Duplicate check done against posts/ready, posts/published, tasks/topics-seed.md and tasks/team-topics-*.md (grep of file names, seed titles and briefs for truncate, word boundary, zip valid, top 3, large, streak, month name, month number, char(10), line break, last 5, average of last, nth row, datevalue). Nothing matched as a topic. Closest existing posts and how they differ:
- Topic 1 vs count-listing-characters-tags (counts length only), count-words-in-cell and extract-text-between-parentheses: this one cuts a title to a character limit at the last whole word.
- Topic 2 vs a-keep-leading-zeros-zip-codes-excel (storing ZIPs), a-split-city-state-zip and shipping-zone-from-zip (lookup by ZIP): this one checks the shape of each ZIP (5 digits or 5+4) and says what is wrong. It does not check that a ZIP exists.
- Topic 3 vs a-split-city-state-zip (comma separated) and combine-address-columns (joins): this one splits one cell with line breaks (CHAR(10)) into columns.
- Topic 4 vs seasonal-sales-index-by-month and year-over-year-sales-by-month (month numbers already typed): this one turns month names typed as text (Jan, March, Sept.) into month numbers.
- Topic 5 vs rolling-3-month-average-sales (fixed table, 3 months) and three-biggest-unpaid-invoices: this one averages the last N entries of a log that keeps growing, with N in a cell.

Dropped before choosing: sum of the top 3 values (too close to three-biggest-unpaid-invoices), longest streak of days with sales (needs array formulas, hard to explain), text dates like Nov 5, 2026 to real dates (DATEVALUE depends on locale, we cannot check Excel or Sheets), sum every Nth row, daily amount still needed for a goal (monthly-sales-goal-progress exists).

Function families: LEFT/FIND/SUBSTITUTE/LEN (1), LEN/ISNUMBER/MID/IF (2), FIND/MID/SUBSTITUTE/REPT/TRIM with CHAR(10) (3), MATCH/LEFT/TRIM/LOWER (4), AVERAGE/INDEX/COUNT/MIN (5). No TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, dynamic arrays.

How to verify in LibreOffice (all five): build the .xlsx with openpyxl, recalculate headless (soffice --headless --convert-to xlsx), read the values back and compare with an independent Python model; also type the edge cases in the Test lines. Excel and Google Sheets are not run by us; the posts say LibreOffice Calc only unless someone actually tests there.

Ranked best first:

1. How to Shorten a Product Title to a Character Limit Without Cutting a Word in Excel and Google Sheets
   - Slug: 2026-11-11-team-01-shorten-title-to-character-limit-whole-word-left-find-substitute-excel-google-sheets
   - Search phrase: "truncate text at word boundary excel", "shorten text to 140 characters without cutting words", "cut off text at last space google sheets", "limit characters keep whole words".
   - Template: tidy-tabs-shorten-title-whole-word.xlsx (Titles, How to use). B1 character limit (140, a setting; the platform's real limit is the reader's to check), header row 3, rows 4 to 203 (12 sample titles, fictional): A Title | B Length | C Shortened title | D Length after | E Status.
     - C4 `=IF(LEN(TRIM(A4))<=$B$1,TRIM(A4),LEFT(TRIM(A4),FIND("~",SUBSTITUTE(LEFT(TRIM(A4),$B$1+1)," ","~",LEN(LEFT(TRIM(A4),$B$1+1))-LEN(SUBSTITUTE(LEFT(TRIM(A4),$B$1+1)," ",""))))-1))`; E4 says OK, Shortened or One long word.
     - Expected: a 150 character title cut at the last space before character 141; a title of exactly 140 stays; a 141 character title is cut.
   - Test: length 139, 140, 141, 142, a space at position 141, a title with no spaces (cannot cut, flagged), double spaces, limit 20 and 1, text limit, blank rows, 200 rows, random titles against a Python model.
   - Who ranks: Exceljet and Stack Overflow generic truncate answers. Angle: Etsy-style listing titles with the limit in one cell and a before and after length.

2. How to Check ZIP Codes for the Right Format (5 Digits or ZIP+4) in Excel and Google Sheets
   - Slug: 2026-11-11-team-02-check-zip-code-format-5-digit-zip-plus-4-len-isnumber-excel-google-sheets
   - Search phrase: "validate zip code format excel", "check 5 digit zip code google sheets", "flag zip codes missing leading zero", "zip+4 formula excel".
   - Template: tidy-tabs-check-zip-format.xlsx (ZIPs, How to use). Header row 3, rows 4 to 203 (15 sample ZIPs, fictional customers): A Customer | B ZIP (text) | C ZIP cleaned | D Digits | E Result | F Fix suggestion; summary H3:I8.
     - C4 `=TRIM(B4)`; E4 `=IF(C4="","",IF(AND(LEN(C4)=5,ISNUMBER(--C4)),"OK",IF(AND(LEN(C4)=10,MID(C4,6,1)="-",ISNUMBER(--LEFT(C4,5)),ISNUMBER(--RIGHT(C4,4))),"OK (ZIP+4)",IF(AND(LEN(C4)=4,ISNUMBER(--C4)),"Missing leading zero","Check"))))`; F4 `=IF(E4="Missing leading zero","0"&C4,"")`.
     - Expected: 02134 OK; 2134 Missing leading zero with fix 02134; 90210-1234 OK (ZIP+4); 9021 Missing leading zero; 1234567 Check; 12O45 (letter O) Check.
   - Test: 4, 5, 6, 9 and 10 characters, letter O for zero, decimal 2134.0, numbers stored as numbers, spaces, hyphen in the wrong place, blank rows, 200 rows, random strings against a Python model.
   - Who ranks: generic regex answers (not available in plain formulas), a few Excel forum threads. Angle: format check only; the post says plainly it does not confirm a ZIP exists and links the USPS lookup for that.

3. How to Split an Address With Line Breaks Into Separate Columns in Excel and Google Sheets
   - Slug: 2026-11-11-team-03-split-address-with-line-breaks-into-columns-char-10-mid-excel-google-sheets
   - Search phrase: "split cell by line break excel formula", "separate text on new line google sheets", "char(10) split into columns", "address in one cell with line breaks".
   - Template: tidy-tabs-split-line-breaks.xlsx (Addresses, How to use). Header row 3, rows 4 to 203 (10 sample ship-to blocks, fictional, 2 to 4 lines): A Pasted address | B Lines | C Line 1 | D Line 2 | E Line 3 | F Line 4.
     - B4 `=IF(TRIM(A4)="","",LEN(A4)-LEN(SUBSTITUTE(A4,CHAR(10),""))+1)`; C4 `=IF($B4="","",TRIM(MID(SUBSTITUTE(A4,CHAR(10),REPT(" ",200)),(COLUMN()-COLUMN($C4))*200+1,200)))` with the column offset written as a number per column; blanks for missing lines.
     - Expected: a 3 line block gives name, street, "Austin, TX 78701" and a blank fourth column.
   - Test: 1, 2, 3, 4 and 5 lines (5 flagged), trailing line break, empty line in the middle, lines longer than 200 characters (flag), spaces around lines, blank rows, 200 rows, Python model. The sample file stores line breaks as real CHAR(10) inside cells with wrap text on.
   - Who ranks: Exceljet, Ablebits (often using TEXTSPLIT or Text to Columns). Angle: formula version that works without TEXTSPLIT, for addresses pasted from order emails.

4. How to Turn Month Names Into Month Numbers in Excel and Google Sheets
   - Slug: 2026-11-11-team-04-month-name-to-month-number-match-left-excel-google-sheets
   - Search phrase: "convert month name to number excel", "jan feb mar to 1 2 3 google sheets", "month text to month number formula", "sort months in order not alphabetical".
   - Template: tidy-tabs-month-name-to-number.xlsx (Sales, Months, How to use). Sales: header row 3, rows 4 to 203 (14 sample rows, fictional): A Month (as typed) | B Sales | C Month no. | D Month (clean); Months: list of the 12 names with 3-letter keys used by MATCH; summary table by month number with SUMIFS.
     - C4 `=IF(TRIM(A4)="","",IFERROR(MATCH(LOWER(LEFT(TRIM(A4),3)),Months!$B$4:$B$15,0),"Check"))` with the keys jan to dec in lower case; D4 `=IF(ISNUMBER(C4),INDEX(Months!$A$4:$A$15,C4),"")`.
     - Expected: Jan, january, JANUARY and "  Jan " all give 1; Sept. and September give 9; "Marc" gives 3 only because it starts with mar (stated as a limit); "Ju" and "Foo" give Check.
   - Test: all 12 names in full, 3-letter and upper case, Sept, May, text with a trailing period, numbers typed (Check), blank rows, 200 rows, summary totals against Python.
   - Who ranks: Exceljet and Microsoft help for MONTH and DATEVALUE tricks (locale dependent). Angle: a plain MATCH list that behaves the same in any locale, plus the totals by month number.

5. How to Average the Last 5 Entries in a Growing List in Excel and Google Sheets
   - Slug: 2026-11-11-team-05-average-last-n-entries-growing-list-index-count-excel-google-sheets
   - Search phrase: "average last 5 values in column excel", "average of last n entries google sheets", "moving average of latest entries as list grows", "average last 10 non blank cells".
   - Template: tidy-tabs-average-last-n.xlsx (Log, How to use). B1 how many latest entries (5, a setting), header row 3, rows 4 to 203 (14 sample weekly sales entries, fictional, typed top to bottom with blank rows below): A Date | B Amount | C Average of last N (on every row); summary E3:F8 (Entries, Latest row, Average of last N, Average of all).
     - Last row F5 `=IF(COUNT(B4:B203)=0,"",LOOKUP(2,1/ISNUMBER(B4:B203),ROW(B4:B203)))`; average F6 `=IF(F5="","",AVERAGE(INDEX(B4:B203,MAX(1,F5-3-$B$1+1)):INDEX(B4:B203,F5-3)))`; per-row C4 `=IF(ISNUMBER(B4),AVERAGE(INDEX($B$4:$B$203,MAX(1,ROW()-3-$B$1+1)):B4),"")`.
     - Expected: with 14 entries and N=5 the average of entries 10 to 14; N larger than the number of entries averages all of them.
   - Test: 0, 1, 4, 5, 6 and 200 entries, N = 1, 5, 200, 0, negative, 2.5 and text (Check), a text or blank cell inside the range, blank rows between entries, random lists against a Python model.
   - Who ranks: Exceljet and Microsoft AVERAGE help, OFFSET-based forum answers. Angle: non-volatile INDEX range instead of OFFSET, N in a cell, shown next to the average of all entries. It describes a number; it does not forecast.
