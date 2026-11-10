---
title: How to Turn Month Names Into Month Numbers in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 7 workbooks and 2,421 cell comparisons. Checked: all 12 month names as full, 3 letter, upper case, lower case, with a period and with spaces, Sept and Sept., Ju, Foo, Marc, Marketing, Mayo, Janitor, May 2026, names with * ? or ~, numbers typed in the month column, blank rows, text and negative sales, a sale on a row with no month, and 5 random sets of 200 rows. Excel and Google Sheets were not used. Not run: a non-US language setting, more rows than the template holds.
search_description: Turn Jan, february or Sept. into 1, 2 and 9 with =MATCH(LOWER(LEFT(TRIM(A4),3)),Months!$B$4:$B$15,0), then total sales by month. Free template.
image_prompts: A wall calendar page with a few handwritten circles, a closed linen ledger book and a pencil on a wooden craft studio desk, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Wall calendar page with circled days beside a closed ledger book and pencil
threads: Sales export says Jan, february, MARCH and "  Apr "?\n=MATCH(LOWER(LEFT(TRIM(A4),3)),Months!$B$4:$B$15,0) gives 1, 2, 3 and 4 from a list of jan to dec.\nOnly the first 3 letters are read, so Marc also gives 3 (made-up sales).
---

Turn a month name into its number with `=MATCH(LOWER(LEFT(TRIM(A4),3)),Months!$B$4:$B$15,0)`, where Months!B4:B15 lists jan to dec in lower case. Jan, january, JANUARY and "  Jan " all give 1, and Sept. and September give 9.

Only the first three letters are read, so Marc also gives 3. The template returns Check for names it cannot match, such as Ju.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (7 workbooks, 2,421 comparisons). The formulas use only IF, MATCH, LOWER, LEFT, TRIM, INDEX and SUMIFS, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Sales table with month names as typed, a month number column and sales totaled by month number](images/2026-11-11-team-04-month-name-to-month-number-match-left-excel-google-sheets/month-name-to-month-number-template.png) *Month names as typed, the month number and clean name, and sales totaled by month number. PDF export from LibreOffice Calc 24.2.*

The sales rows in this post are made up. They say nothing about any real shop.

## Steps

A sales export often has months typed by hand, so one column holds Jan, february, MARCH and Sept. side by side. A number is easier to sort and total than a name.

1. On a tab named Months, list the keys in B4:B15: jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec, all lower case. Put the month names in A4:A15 and the numbers 1 to 12 in C4:C15.
2. Paste your export with the month as typed in column A and the sales in column B.
3. In C4, type `=MATCH(LOWER(LEFT(TRIM(A4),3)),Months!$B$4:$B$15,0)`. TRIM removes stray spaces, LEFT keeps three letters, LOWER makes the case match, and MATCH with 0 finds the exact key.
4. In D4, type `=INDEX(Months!$A$4:$A$15,C4)` to show the clean month name.
5. Copy C4:D4 down to the last row.
6. To total by month, put 1 to 12 in F4:F15 and use `=SUMIFS($B$4:$B$203,$C$4:$C$203,F4)` beside each number.

The template wraps step 3 so a blank row stays blank and a name with no match shows Check instead of an error.

Help pages: Microsoft's [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) and Google's [MATCH](https://support.google.com/docs/answer/3093378). They describe the function, not our test.

## Example

Sample rows (fictional):

| Typed | Sales | Month no. | Clean |
|---|---|---|---|
| Jan | $1,250.00 | 1 | January |
| february | $980.50 | 2 | February |
| MARCH | $1,410.00 | 3 | March |
| "  Apr " | $1,175.25 | 4 | April |
| Jun. | $1,890.75 | 6 | June |
| Sept. | $1,302.40 | 9 | September |
| Marc | $310.00 | 3 | March |
| Ju | $150.00 | Check | |
| Foo | $90.00 | Check | |

The full sample has 14 rows. March shows 2 entries and $1,720.00, August shows 0 entries and $0.00, and months 1 to 12 add up to $18,323.70. Two rows show Check, worth $240.00, so all sales come to $18,563.70. With the Check rows counted, the template's CHECK cell shows OK.

## Troubleshooting

### Marc, Marketing or Mayo give a month
Only the first three letters are read. Marc gives March, Mayo gives May, Janitor gives January and "May 2026" gives May. If your column can hold other words, scan the Month (clean) column for surprises. Ju and Ma show Check because they have fewer than three letters to match.

### A number or a date shows Check
This tab reads names. A 1 or a 12 typed in the month column, or a real date, shows Check. For month numbers you already have, see [our seasonal index post](https://tidytabs.blogspot.com/2026/10/how-to-find-your-busiest-and-slowest.html).

### A name with * ? or ~ shows Check
MATCH reads those characters as wildcards, so the template stops those cells before MATCH runs.

### CHECK shows Check
Either an amount in column B is text, or a sale sits on a row with no month name. Fix the cell and the count of amounts that are not numbers goes back to 0.

### Another language
Change the keys on the Months tab, for example ene, feb, mar for Spanish. MATCH compares text and does not parse dates. We did not run the template under a non-US language setting.

## Template

[Download the .xlsx](templates/tidy-tabs-month-name-to-number.xlsx)

The Sales tab has room for 200 rows (rows 4 to 203), 14 fictional samples, the month number and clean name columns, a table of sales by month number and a CHECK cell. The Months tab holds the keys. Yellow cells are inputs and gray cells are formulas. A How to use tab explains each column.

To open it in Google Sheets, choose **File > Import > Upload**, then pick the file. The formulas were checked in LibreOffice Calc 24.2.7.2 only.
