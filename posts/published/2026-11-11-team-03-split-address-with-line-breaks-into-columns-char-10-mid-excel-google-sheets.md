---
threads_url: https://www.threads.com/@tin_ylab/post/DfnZwXtlJJX
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-split-address-with-line-breaks.html
title: How to Split an Address With Line Breaks Into Separate Columns in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with independent Python models over 7 workbooks and 7,378 comparisons. Checked: 25 fixed cells with 1 to 6 lines, a trailing break, empty lines in the middle, a leading break, CRLF breaks, spaces around lines, a tab, spaces only, a 250 character line, a 500 character fifth line, and 6 random sets of 200 rows. Excel and Google Sheets were not used. Not run: Excel, Google Sheets, a non-US language setting, more than 200 rows, cells longer than about 6,000 characters.
search_description: Split an address that has line breaks in one cell into separate columns with SUBSTITUTE, CHAR(10), REPT and MID, no TEXTSPLIT needed. Free template.
image_prompts: A sheet of blank shipping labels beside a printed order confirmation page with a few lines highlighted in yellow and a roll of packing tape on a wooden packing table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Sheet of blank shipping labels next to a highlighted order printout and tape
threads: One cell, four address lines, no commas to split on.\nSwap each CHAR(10) for a pile of spaces with SUBSTITUTE and REPT, then MID out each line. Line 1 starts at 0*LEN(A4)+1.\nThe sample ship-to blocks are made up.
---
Pull line 1 out of a cell that holds line breaks with `=TRIM(CLEAN(MID(SUBSTITUTE(A4,CHAR(10),REPT(" ",LEN(A4))),0*LEN(A4)+1,LEN(A4))))`, then change the 0 to 1, 2 and 3 for lines 2 to 4. A 3 line sample block gives the name, the street and Austin, TX 78701, with column 4 left blank (made-up address).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (7 workbooks, 7,378 comparisons). The formulas use only IF, TRIM, CLEAN, MID, SUBSTITUTE, REPT, CHAR and LEN, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Ten pasted ship-to blocks split into Line 1 to Line 4 with the five line block flagged](images/2026-11-11-team-03-split-address-with-line-breaks-into-columns-char-10-mid-excel-google-sheets/split-address-line-breaks-template.png) *Ten pasted ship-to blocks split into Line 1 to Line 4; the 5 line block is flagged. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. Put each pasted address in column **A**, with headers in row 3 and the first address in **A4**. A line break inside a cell is the character CHAR(10), and the formulas below find it.

1. In **B4**, count the lines: `=LEN(A4)-LEN(SUBSTITUTE(A4,CHAR(10),""))+1`. This is the number of line breaks plus 1.
2. In **C4**, type the line 1 formula:

   `=TRIM(CLEAN(MID(SUBSTITUTE(A4,CHAR(10),REPT(" ",LEN(A4))),0*LEN(A4)+1,LEN(A4))))`

3. In **D4**, **E4** and **F4**, use the same formula and change the typed 0 to 1, 2 and 3.
4. Fill **B4:F4** down to your last address.

How it works: SUBSTITUTE replaces each line break with as many spaces as the whole cell has characters. Every line then starts in its own window of that width. MID cuts out window number 0, 1, 2 or 3, and TRIM removes the padding spaces.

An empty line in the middle leaves its own column blank and keeps the later lines in their columns. A trailing line break counts as one more line in column **B**.

TRIM also squeezes double spaces inside a line to one. CLEAN removes CHAR(13), which can appear when text has Windows style line endings (the test set included CRLF breaks).

The template wraps these formulas so an empty row stays blank, and it adds a Status column that flags a block with more than 4 lines. The core formulas above are what do the work.

Help pages: Microsoft's [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332), [MID function](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028) and [CLEAN function](https://support.microsoft.com/en-us/office/clean-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41). They describe the functions, not our test.

## Example

Sample data is fictional: 10 ship-to blocks pasted from emails, each with real line breaks in one cell.

| Lines | Line 1 | Line 2 | Line 3 | Line 4 | Status |
|---|---|---|---|---|---|
| 3 | Maria Lopez | 418 Elm Street | Austin, TX 78701 | | OK |
| 4 | Harbor Dental | Attn: Front Desk | 92 Water Street | Portland, ME 04101 | OK |
| 3 | Dev Patel | 77 Birch Lane Apt 3B | Boston, MA 02134 | | OK |
| 4 | Lakeside Yoga | 1200 Shore Drive | Suite 4 | Chicago, IL 60614 | OK |
| 3 | Jordan Reed | 5 Cedar Court | Denver, CO 80202 | | OK |
| 3 | Pine & Co. Realty | 300 Main Street | Hartford, CT 06103 | | OK |
| 4 | Ana Souza | (empty line) | 9 Oak Road | Miami, FL 33101 | OK |
| 5 | Sam Whitaker | Unit 12, Maple Plaza | 88 Market Street | Attn: Receiving Dock B | More than 4 lines |
| 4 | Priya Shah | 14 River Bend Drive | Newark, NJ 07102 | | OK |
| 3 | Cedar Print Shop | 2210 Hilltop Road | Asheville, NC 28801 | | OK |

Row 6 was typed as `  Pine & Co. Realty  ` with spaces at both ends, and TRIM removed them. Ana Souza's block has an empty second line, so Line 2 is blank and the street and city stay in Line 3 and Line 4. Priya Shah's block ends with a line break, so column **B** shows 4 for a 3 line address.

Sam Whitaker's block has 5 lines. The fifth line, Seattle, WA 98101, has no column, so Status says More than 4 lines.

The summary block shows 10 addresses listed, 9 with four lines or fewer, 1 with more than 4 lines, and CHECK OK.

## Troubleshooting

### Status says More than 4 lines
Only four columns are built. Copy the Line 4 column and change its typed 3 to 4 for a fifth line, 5 for a sixth, and so on. The Status test in the template looks at what is left after line 4, so widen it the same way.

### Everything lands in Line 1
The cell has no real line break. Text pasted from some sources uses a different separator, such as a comma or a pipe. Use that character in place of CHAR(10).

### A line shows an extra gap in the middle
TRIM squeezes double spaces inside a line to one space. That is intended, but it also changes text where two spaces were meant.

### A row with an address shows nothing
The formulas treat a cell as empty when TRIM leaves nothing, so a cell with only spaces counts as empty.

### Cells with very long text
The padded text is about as long as the cell length times its line count. Cells longer than about 6,000 characters were not tested, and the padded text could pass Excel's 32,767 character limit.

## Template

[Download the .xlsx](templates/tidy-tabs-split-line-breaks.xlsx)

The file has two tabs. **Addresses** has 200 slots from row 4 to row 203 and 10 fictional sample ship-to blocks with wrap text on. Column **A** holds the pasted address. Columns **B** to **G** hold the line count, Line 1 to Line 4 and the Status. A summary block in **I3:J6** counts addresses listed, four lines or fewer, more than 4 lines, and CHECK. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The file was exported to PDF in LibreOffice and every page is US Letter.

To open it in Google Sheets, go to **File > Import > Upload** and select the file. The file was checked in LibreOffice Calc only.
