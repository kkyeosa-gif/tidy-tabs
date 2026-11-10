# Handoff 2026-11-11: five topics researched, hooks written, templates built and tested, heroes made (LibreOffice Calc 24.2.7.2 only), for the script-writers

Stage 1 of the content team (researcher, hook-writer, builder, designer) was done by one agent. Nothing is committed or pushed; no post file exists yet in posts/ready. Files: tasks/team-topics-2026-11-11.md, tasks/hooks-2026-11-11.md, scripts/build-templates.py (`shorten_title_whole_word`, `check_zip_format`, `split_line_breaks`, `month_name_to_number`, `average_last_n`), scripts/make-heroes.py (`heroes_2026_11_11`, run with `--only-2026-11-11`), five new files in templates/, five heroes in images/, results in tasks/briefs.md ("2026-11-11 builder results"). tasks/topics-seed.md already carried the new titles when this run started.

Exact version: **LibreOffice 24.2.7.2 420(Build:2)**. Nothing was opened in Excel or Google Sheets and no post may claim it. Totals: 73 edited workbooks compared (28,672 comparisons, 0 mismatches against independent Python models) plus the 5 unedited samples checked by hand; no error value in any recalculated workbook. All five files exported to PDF: every page US Letter, no error values in the PDF text. Bugs found and fixed are in tasks/briefs.md (a text N such as "abc" gave #VALUE! in the average file through AND; fixed with a nested IF).

Repo note: only the five new builder functions were run, so no older template was rewritten; `git status` shows the two scripts, the topics, hooks, handoff, briefs and seed files, five new .xlsx files and five image folders. No older file and no __pycache__ changed.

Post file names: `posts/ready/2026-11-11-team-0N-<slug>.md`. Sample data is fictional; say so once per post. Colors: yellow cells are inputs, gray cells are formulas. The `> Works in` callout and `tested_in` must say LibreOffice Calc 24.2.7.2 only. Rules for all five: no TODAY, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE, TEXTJOIN, dynamic arrays. The ZIP post must not say a ZIP exists or is right for an address (shape only; link the USPS ZIP Code lookup for that). The title post must not state any marketplace's real limit (140 is a setting). The average post describes numbers; it does not forecast.

| # | Slug | Template | tested_in counts |
|---|---|---|---|
| 1 | 2026-11-11-team-01-shorten-title-to-character-limit-whole-word-left-find-substitute-excel-google-sheets | tidy-tabs-shorten-title-whole-word.xlsx | 19 workbooks, 6,973 comparisons |
| 2 | 2026-11-11-team-02-check-zip-code-format-5-digit-zip-plus-4-len-isnumber-excel-google-sheets | tidy-tabs-check-zip-format.xlsx | 7 workbooks, 5,069 comparisons |
| 3 | 2026-11-11-team-03-split-address-with-line-breaks-into-columns-char-10-mid-excel-google-sheets | tidy-tabs-split-line-breaks.xlsx | 7 workbooks, 7,378 comparisons |
| 4 | 2026-11-11-team-04-month-name-to-month-number-match-left-excel-google-sheets | tidy-tabs-month-name-to-number.xlsx | 7 workbooks, 2,421 comparisons |
| 5 | 2026-11-11-team-05-average-last-n-entries-growing-list-index-count-excel-google-sheets | tidy-tabs-average-last-n.xlsx | 33 workbooks, 6,831 comparisons |

Hero files (1600x900 PNG, real LibreOffice PDF render of the template with its sample data; caption "PDF export from LibreOffice Calc 24.2"; alt text must be written per post, 8 to 15 words, not the title):
1. images/2026-11-11-team-01-shorten-title-to-character-limit-whole-word-left-find-substitute-excel-google-sheets/shorten-title-whole-word-template.png (caption: "The Titles tab with a limit of 140: titles over the limit are cut at the last space, one long word is flagged. PDF export from LibreOffice Calc 24.2.")
2. images/2026-11-11-team-02-check-zip-code-format-5-digit-zip-plus-4-len-isnumber-excel-google-sheets/check-zip-code-format-template.png (caption: "Fifteen sample ZIPs with the result and the one-zero fix suggestion. PDF export from LibreOffice Calc 24.2.")
3. images/2026-11-11-team-03-split-address-with-line-breaks-into-columns-char-10-mid-excel-google-sheets/split-address-line-breaks-template.png (caption: "Ten pasted ship-to blocks split into Line 1 to Line 4; the 5 line block is flagged. PDF export from LibreOffice Calc 24.2.")
4. images/2026-11-11-team-04-month-name-to-month-number-match-left-excel-google-sheets/month-name-to-month-number-template.png (caption: "Month names as typed, the month number and clean name, and sales totaled by month number. PDF export from LibreOffice Calc 24.2.")
5. images/2026-11-11-team-05-average-last-n-entries-growing-list-index-count-excel-google-sheets/average-last-5-entries-template.png (caption: "Fourteen weekly sales with the average of the last 5 on every row and the summary at the right. PDF export from LibreOffice Calc 24.2.")
The AI photo (`image_prompts`, `image_alt`) is still open for every post: one scene each, 2 or 3 specific objects, no readable text (ideas: title = a listing box lid with a printed card and a measuring tape; ZIP = a stack of address-labeled envelopes with one label torn; line breaks = a shipping label sheet beside a printed order email; months = a wall calendar page and a ledger book; average = a row of weekly receipts clipped to a board).

Every workbook has a "How to use" tab. Titles and openings: tasks/hooks-2026-11-11.md. Deviations from tasks/team-topics-2026-11-11.md: topic and slug unchanged; the built formulas and cell maps below differ in the ways listed per template and in tasks/briefs.md.

---
## 1. shorten-title-to-character-limit-whole-word-left-find-substitute-excel-google-sheets
Title: How to Shorten a Product Title to a Character Limit Without Cutting a Word in Excel and Google Sheets
Template: templates/tidy-tabs-shorten-title-whole-word.xlsx (tabs: Titles, How to use)

Layout
- Titles: A1 label, B1 character limit (input, 140), note C1, header row 3, data rows 4:203 (12 sample titles). A Title (input, wrap text), B Length, C Shortened title, D Length after, E Status (formulas). Summary G3:H9: titles listed, already within the limit, shortened at a space, one long word, rows showing Check, longest length after, CHECK.

Key formulas (row 4; `t` = TRIM(A4), `cut` = LEFT(t,$B$1+1))
- B4 `=IF(t="","",LEN(t))`
- C4 `=IF(t="","",IF(NOT(ISNUMBER($B$1)),"Check",IF(OR($B$1<1,$B$1<>INT($B$1)),"Check",IF(ISNUMBER(FIND("~",t)),"Check",IF(LEN(t)<=$B$1,t,IFERROR(LEFT(t,FIND("~",SUBSTITUTE(cut," ","~",LEN(cut)-LEN(SUBSTITUTE(cut," ",""))))-1),LEFT(t,$B$1)))))))` (the post teaches the core `=LEFT(TRIM(A4),FIND("~",SUBSTITUTE(LEFT(TRIM(A4),$B$1+1)," ","~",LEN(LEFT(TRIM(A4),$B$1+1))-LEN(SUBSTITUTE(LEFT(TRIM(A4),$B$1+1)," ",""))))-1)`)
- D4 `=IF(OR(E4="",E4="Check"),"",LEN(C4))`; E4 OK, Shortened, One long word or Check
- H9 `=IF(AND(ISNUMBER(B1),H7=0,H3=H4+H5+H6,COUNTIF(D4:D203,">"&B1)=0),"OK","Check")`

Sample values (fictional titles; limit 140)
| Title (start) | Length | After | Status |
|---|---|---|---|
| Handmade Soy Candle Lavender Vanilla 8 oz Gift Set... | 150 | 138 | Shortened |
| Oatmeal Goat Milk Soap Bar | 26 | 26 | OK |
| Personalized Cutting Board ... Custom Charcuterie | 140 | 140 | OK |
| same with "Names" | 141 | 129 | Shortened |
| the 140 title plus " Tray" (a space at position 141) | 145 | 140 | Shortened |
| Brass Hoop Earrings Small Gold Huggie ... Set | 158 | 137 | Shortened |
| Letterpress Birthday Card Set of 6 ... | 122 | 122 | OK |
| Ceramic Mug 12 oz ... | 127 | 127 | OK |
| Macrame Plant Hanger ... | 55 | 55 | OK |
| Hand  Poured  Soy ... (doubled spaces) | 146 | 138 | Shortened |
| HandmadeSoyCandle... (no spaces, 149 characters) | 149 | 140 | One long word |
| "  Leather Key Fob Keychain Personalized Initial  " | 45 | 45 | OK |
Summary H3:H9: 12, 6, 5, 1, 0, 140, OK.

Pitfalls the post must state
- The +1 in `$B$1+1` is the point: a space at position 141 means the word ending at 140 is whole and stays (the 145 character row shows it). A 141 character title is cut before its last word.
- TRIM first: spaces at the ends and doubled spaces inside are removed and counted as one before the length is compared.
- A title with no space in its first limit+1 characters cannot be cut at a space; the file keeps the first limit characters and says One long word. SUBSTITUTE with instance 0 would fail, which is why IFERROR is there.
- A ~ anywhere in a title returns Check (the formula uses ~ as a marker). A limit that is blank, text, 0, negative or has a decimal makes every row Check.
- The limit is a setting; check your marketplace for its real number. Not tested: Excel, Google Sheets, non-breaking spaces (TRIM does not remove CHAR(160)).

---
## 2. check-zip-code-format-5-digit-zip-plus-4-len-isnumber-excel-google-sheets
Title: How to Check ZIP Codes for the Right Format (5 Digits or ZIP+4) in Excel and Google Sheets
Template: templates/tidy-tabs-check-zip-format.xlsx (tabs: ZIPs, How to use)

Layout
- ZIPs: note A1, header row 3, data rows 4:203 (15 sample customers). A Customer, B ZIP (text format) are inputs; C ZIP cleaned, D Length, E Result, F Fix suggestion. Summary H3:I9: ZIPs listed, OK 5 digits, OK ZIP+4, Missing leading zero, Check, Blank ZIP, CHECK.

Key formulas (row 4)
- C4 `=TRIM(B4)`; D4 `=IF(C4="","",LEN(C4))`
- E4 `=IF(C4="",IF(TRIM(A4)="","","Blank ZIP"),IF(AND(LEN(C4)=5,SUMPRODUCT(--ISNUMBER(FIND(MID(C4,{1,2,3,4,5},1),"0123456789")))=5),"OK",IF(AND(LEN(C4)=10,MID(C4,6,1)="-",SUMPRODUCT(--ISNUMBER(FIND(MID(C4,{1,2,3,4,5,7,8,9,10},1),"0123456789")))=9),"OK (ZIP+4)",IF(AND(LEN(C4)=4,SUMPRODUCT(...{1,2,3,4}...)=4),"Missing leading zero",IF(AND(LEN(C4)=9,MID(C4,5,1)="-",SUMPRODUCT(...{1,2,3,4,6,7,8,9}...)=8),"Missing leading zero","Check")))))`
- F4 `=IF(E4="Missing leading zero","0"&C4,"")`
- I9 `=IF(I3=I4+I5+I6+I7+I8,"OK","Check")`
- Deviation from the topic: the topic's `ISNUMBER(--C4)` was replaced by the per-character FIND test because `--` accepts "1,234", "$1234", "+1234" and "1E100" as numbers (reasoned, not run). The posts may still show `ISNUMBER(--C4)` as a shorter idea only if they say it accepts those.

Sample values (fictional customers)
| ZIP typed | Result | Fix |
|---|---|---|
| 02134 | OK | |
| 2134 | Missing leading zero | 02134 |
| 90210-1234 | OK (ZIP+4) | |
| 9021 | Missing leading zero | 09021 |
| 1234567 | Check | |
| 12O45 (letter O) | Check | |
| " 07030 " | OK | |
| 60614 | OK | |
| 10001-0001 | OK (ZIP+4) | |
| 2134-5678 | Missing leading zero | 02134-5678 |
| 8101 | Missing leading zero | 08101 |
| (blank, a customer is listed) | Blank ZIP | |
| 331 | Check | |
| 30301-12 | Check | |
| 55401 | OK | |
Summary I3:I9: 15, 4, 2, 4, 4, 1, OK.

Pitfalls the post must state
- Format only: 00000 passes; no ZIP is looked up. Link the USPS ZIP Code lookup for existence.
- The fix adds one 0 and only suggests; the original may have lost more than one zero or be a different ZIP.
- A ZIP stored as a number (2134) is already damaged: it shows Missing leading zero. Format column B as Plain text before pasting so zeros are kept (link the existing keep-leading-zeros post).
- Letter O for zero, a letter after digits, a comma, $, an exponent such as 1E100 and full width digits all show Check. Not tested: Excel, Google Sheets (SUMPRODUCT with an array constant inside MID), non-breaking spaces and tabs.

---
## 3. split-address-with-line-breaks-into-columns-char-10-mid-excel-google-sheets
Title: How to Split an Address With Line Breaks Into Separate Columns in Excel and Google Sheets
Template: templates/tidy-tabs-split-line-breaks.xlsx (tabs: Addresses, How to use)

Layout
- Addresses: note A1, header row 3, data rows 4:203 (10 sample ship-to blocks with real line breaks, wrap text on). A Pasted address (input); B Line breaks + 1, C Line 1, D Line 2, E Line 3, F Line 4, G Status (formulas). Summary I3:J6: addresses listed, four lines or fewer, more than 4 lines, CHECK.

Key formulas (row 4; `L` = LEN($A4), `S` = SUBSTITUTE($A4,CHAR(10),REPT(" ",L)))
- B4 `=IF(TRIM(A4)="","",LEN(A4)-LEN(SUBSTITUTE(A4,CHAR(10),""))+1)`
- C4 `=IF($B4="","",TRIM(CLEAN(MID(S,0*L+1,L))))`; D4 uses `1*L+1`, E4 `2*L+1`, F4 `3*L+1` (the offset is a typed number per column)
- G4 `=IF($B4="","",IF(TRIM(CLEAN(MID(S,4*L+1,L*L)))<>"","More than 4 lines","OK"))`
- J6 `=IF(J3=J4+J5,"OK","Check")`
- Deviation from the topic: the pad and window are LEN(A4) wide, not 200, so there is no 200 character limit (a fixed 200 pad breaks as soon as the lines before one add up to more than 200; reasoned, not run). CLEAN was added for CHAR(13).

Sample values (fictional)
| Lines | Line 1 | Line 2 | Line 3 | Line 4 | Status |
|---|---|---|---|---|---|
| 3 | Maria Lopez | 418 Elm Street | Austin, TX 78701 | | OK |
| 4 | Harbor Dental | Attn: Front Desk | 92 Water Street | Portland, ME 04101 | OK |
| 3 | Dev Patel | 77 Birch Lane Apt 3B | Boston, MA 02134 | | OK |
| 4 | Lakeside Yoga | 1200 Shore Drive | Suite 4 | Chicago, IL 60614 | OK |
| 3 | Jordan Reed | 5 Cedar Court | Denver, CO 80202 | | OK |
| 3 | "  Pine & Co. Realty  " becomes Pine & Co. Realty | 300 Main Street | Hartford, CT 06103 | | OK |
| 4 | Ana Souza | (empty line) | 9 Oak Road | Miami, FL 33101 | OK |
| 5 | Sam Whitaker | Unit 12, Maple Plaza | 88 Market Street | Attn: Receiving Dock B | More than 4 lines (Seattle, WA 98101 is the 5th) |
| 4 (trailing break) | Priya Shah | 14 River Bend Drive | Newark, NJ 07102 | | OK |
| 3 | Cedar Print Shop | 2210 Hilltop Road | Asheville, NC 28801 | | OK |
Summary J3:J6: 10, 9, 1, OK.

Pitfalls the post must state
- Each CHAR(10) becomes as many spaces as the cell has characters, so every line starts in its own window; the post can show the one-line version of the idea, not the 200 pad.
- Line breaks + 1 counts a trailing break as one more line. An empty line in the middle leaves its column blank and keeps the later lines in their columns.
- TRIM also squeezes double spaces inside a line. CLEAN removes CHAR(13) (CRLF text) and other control characters.
- 5 or more lines: only 4 columns are built; Status says More than 4 lines (copy a column and use 4*, 5*...).
- In Excel a line break is typed with Alt+Enter; in Google Sheets with Ctrl+Enter (Windows). These are menu/shortcut facts from Microsoft and Google help, not tested by us; link the help pages. Not tested: Excel, Google Sheets, cells longer than about 6,000 characters.

---
## 4. month-name-to-month-number-match-left-excel-google-sheets
Title: How to Turn Month Names Into Month Numbers in Excel and Google Sheets
Template: templates/tidy-tabs-month-name-to-number.xlsx (tabs: Sales, Months, How to use)

Layout
- Sales: note A1, header row 3, data rows 4:203 (14 sample rows). A Month (as typed), B Sales (inputs); C Month no., D Month (clean) (formulas). Table F3:I15 (F month no. 1 to 12, G name, H entries, I sales). F16 Rows showing Check (H16, I16), F17 Months 1 to 12 (H17, I17), F18 All sales (I18), F19 Amounts that are not numbers (H19), F20 CHECK (I20).
- Months: A4:A15 names, B4:B15 keys jan to dec (lower case), C4:C15 numbers 1 to 12.

Key formulas (row 4)
- C4 `=IF(TRIM(A4)="","",IF(SUMPRODUCT(--ISNUMBER(FIND({"*","?","~"},A4)))>0,"Check",IFERROR(MATCH(LOWER(LEFT(TRIM(A4),3)),Months!$B$4:$B$15,0),"Check")))` (the post teaches the core `=MATCH(LOWER(LEFT(TRIM(A4),3)),Months!$B$4:$B$15,0)`)
- D4 `=IF(ISNUMBER(C4),INDEX(Months!$A$4:$A$15,C4),"")`
- H4 `=COUNTIF($C$4:$C$203,F4)`; I4 `=SUMIFS($B$4:$B$203,$C$4:$C$203,F4)`
- I20 `=IF(AND(H19=0,ROUND(I17+I16-I18,2)=0),"OK","Check")`

Sample values (fictional)
| Typed | Sales | Month no. | Clean |
|---|---|---|---|
| Jan | $1,250.00 | 1 | January |
| february | $980.50 | 2 | February |
| MARCH | $1,410.00 | 3 | March |
| "  Apr " | $1,175.25 | 4 | April |
| May | $1,620.00 | 5 | May |
| Jun. | $1,890.75 | 6 | June |
| July | $1,735.00 | 7 | July |
| Sept. | $1,302.40 | 9 | September |
| October | $1,544.00 | 10 | October |
| Nov | $2,210.80 | 11 | November |
| Dec | $2,895.00 | 12 | December |
| Marc | $310.00 | 3 | March |
| Ju | $150.00 | Check | |
| Foo | $90.00 | Check | |
Table: August 0 entries $0.00; March 2 entries $1,720.00; Rows showing Check 2 ($240.00); months 1 to 12: 12 entries, $18,323.70; all sales $18,563.70; not numbers 0; CHECK OK.

Pitfalls the post must state
- Only the first three letters are read: Marc, Marketing, Mayo, Janitor and "May 2026" all give a month. Ju and Ma give Check (fewer than three useful letters or no match).
- A number typed in the month column (1, 12) or a real date shows Check; this tab reads names. For month numbers already typed, link the existing seasonal-sales-index-by-month post.
- * ? ~ in a cell return Check (MATCH reads them as wildcards).
- A sale on a row with no month name, or an amount that is text, makes CHECK show Check.
- For another language, change the keys on the Months tab. Not tested: Excel, Google Sheets (SUMPRODUCT over FIND with an array constant), a non-US language setting. The locale independence of MATCH is a property of the formula (no date parsing); the post must not say we checked it in another locale.

---
## 5. average-last-n-entries-growing-list-index-count-excel-google-sheets
Title: How to Average the Last 5 Entries in a Growing List in Excel and Google Sheets
Template: templates/tidy-tabs-average-last-n.xlsx (tabs: Log, How to use)

Layout
- Log: A1 label, B1 how many latest entries (input, 5), note C1, header row 3, data rows 4:203 (14 sample weekly sales, 08/03/2026 to 11/02/2026, typed top to bottom, empty rows below). A Date, B Amount (inputs); C Average of last N (formula on every row). Summary E3:F10 (E3 header "Summary", F3 "Value"): F4 entries, F5 row of the latest entry, F6 average of last N, F7 average of all entries, F8 numbers in the last-N rows, F9 cells in B that are not numbers, F10 CHECK.

Key formulas
- F5 `=IF(COUNT(B4:B203)=0,"",LOOKUP(2,1/ISNUMBER(B4:B203),ROW(B4:B203)))`
- F6 `=IF(F5="","",IF(NOT(ISNUMBER($B$1)),"Check",IF(OR($B$1<1,$B$1<>INT($B$1)),"Check",AVERAGE(INDEX(B4:B203,MAX(1,F5-ROW($B$3)-$B$1+1)):INDEX(B4:B203,F5-ROW($B$3))))))` (the post teaches `=AVERAGE(INDEX(B4:B203,MAX(1,F5-3-$B$1+1)):INDEX(B4:B203,F5-3))`)
- C4 `=IF(ISNUMBER(B4),IF(ISNUMBER($B$1),IF(AND($B$1>=1,$B$1=INT($B$1)),AVERAGE(INDEX($B$4:$B$203,MAX(1,ROW(B4)-ROW($B$3)-$B$1+1)):B4),"Check"),"Check"),"")`
- F7 `=IF(F4=0,"",AVERAGE(B4:B203))`; F10 `=IF(AND(<N is a whole number of 1 or more>,F9=0),"OK","Check")`

Sample values (fictional weekly sales)
| Date | Amount | Average of last 5 |
|---|---|---|
| 08/03/2026 | $1,180.00 | $1,180.00 |
| 08/10/2026 | $1,245.50 | $1,212.75 |
| 08/17/2026 | $1,310.25 | $1,245.25 |
| 08/24/2026 | $1,190.00 | $1,231.44 |
| 08/31/2026 | $1,422.75 | $1,269.70 |
| 09/07/2026 | $1,365.00 | $1,306.70 |
| 09/14/2026 | $1,288.40 | $1,315.28 |
| 09/21/2026 | $1,510.00 | $1,355.23 |
| 09/28/2026 | $1,475.60 | $1,412.35 |
| 10/05/2026 | $1,620.00 | $1,451.80 |
| 10/12/2026 | $1,398.25 | $1,458.45 |
| 10/19/2026 | $1,702.50 | $1,541.27 |
| 10/26/2026 | $1,655.00 | $1,570.27 |
| 11/02/2026 | $1,788.90 | $1,632.93 |
Summary F4:F10: 14, 17, $1,632.93 (entries 10 to 14), $1,439.44, 5, 0, OK. With N larger than 14 the average of last N equals the average of all.

Pitfalls the post must state
- The window is the last N ROWS ending at the latest number, not the last N numbers. A blank or text cell inside it is skipped by AVERAGE, so it can hold fewer than N numbers (F8 shows how many).
- Early rows average what exists so far; the first row shows its own amount.
- N must be a whole number of 1 or more; blank, text (even "5" typed as text), 0, negative or 2.5 shows Check on every row and in F6. A text cell in column B makes CHECK show Check.
- INDEX(...):INDEX(...) is not volatile, unlike OFFSET, which recalculates on every change (Microsoft documents OFFSET as volatile; link it). The 3 in the post is the header row offset; the file uses ROW($B$3) so it survives inserted rows above.
- This describes numbers already typed. It does not forecast. Not tested: Excel, Google Sheets (LOOKUP(2,1/ISNUMBER(range),ROW(range)) and the INDEX range are the open questions).
