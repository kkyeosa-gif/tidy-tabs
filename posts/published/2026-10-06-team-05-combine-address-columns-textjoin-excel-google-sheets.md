---
threads_url: https://www.threads.com/@tin_ylab/post/DeMouvLG1zX
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-combine-address-columns-with.html
title: How to Combine Address Columns With TEXTJOIN in Excel and Google Sheets
labels: excel-formulas, mailing-labels, data-cleanup
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps follow Microsoft Support (neither app was tested here).
search_description: Join street, unit, city and state into one address line with TEXTJOIN and skip blank cells. Compare with the & join. Free .xlsx with 5 sample rows.
image_prompts: A stack of blank white mailing envelopes beside a sheet of unprinted adhesive address labels and a roll of postage stamps on a wooden table, a rubber stamp and a pen, soft window light, no screens, no readable writing, no logos
image_alt: Plain envelopes, blank adhesive labels, and a stamp roll arranged on a table
threads: 2450 Harbor Road, , Portland, ME is not an address.\nThat double comma is what the & join leaves when Unit is empty.\n=TEXTJOIN(", ",TRUE,A2:D2) skips the blank cell, and the TRUE is the whole trick.
---
Type `=TEXTJOIN(", ",TRUE,A2:D2)` to join street, unit, city and state into one line. The TRUE skips empty cells, so a row with no unit does not leave a stray comma.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and read back. The Excel and Google Sheets steps use the same function and link to Microsoft Support. Neither Excel nor Google Sheets was opened for this post.

![Address sheet with street, unit, city, state and ZIP columns and two joined address columns](images/2026-10-06-team-05-combine-address-columns-textjoin-excel-google-sheets/combine-address-columns-textjoin-template.png) *LibreOffice Calc 24.2 PDF export of the Addresses sheet with sample data (fake).*

## Steps

TEXTJOIN takes three things: a separator, a TRUE or FALSE for "ignore empty cells", and the range to join. One formula replaces a chain of `&` joins.

1. Put one part of the address in each column. The template uses **A** Street, **B** Unit, **C** City, **D** State and **E** ZIP, with headers in row 1.
2. Format column **E** as Text before you type ZIP codes. That keeps the leading zero in codes such as 02108.
3. In **F2**, type `=TEXTJOIN(", ",TRUE,A2:D2)` and press Enter. The result is one line with commas between the parts that have a value.
4. In **G2**, type `=TEXTJOIN(", ",TRUE,A2:D2)&" "&E2`. This adds the ZIP after a space, so the state and ZIP are not split by a comma.
5. Double-click the fill handle on **F2** and **G2**, or drag it down, to copy the formulas to the other rows.
6. To see why TEXTJOIN helps, type `=A2&", "&B2&", "&C2&", "&D2` in **H2** and fill it down. This is the older way to join.
7. To use the result on labels or in a mail merge, copy column F or G and choose **Paste special > Values only**, so the text stays when you delete the source columns.

### In Excel

The formula works as typed. Microsoft documents it on its [TEXTJOIN function](https://support.microsoft.com/en-us/office/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c) page, which lists the versions that include it: Excel 2019, Excel 2021 and Microsoft 365. Excel 2016 and earlier do not have TEXTJOIN. That version limit comes from Microsoft's page, not from a test here.

### In Google Sheets

Google Sheets also has TEXTJOIN with the same three arguments, so the formula should work as typed. This was not tested for this post, and the Google help page for it could not be verified, so no link is given.

## Example

Sample data below is fictional. The first five columns are typed in. The last three are formulas.

| Street | Unit | City | State | ZIP |
|---|---|---|---|---|
| 118 Elm Street | Apt 4B | Boston | MA | 02108 |
| 2450 Harbor Road | | Portland | ME | 04101 |
| 77 Mill Lane | Suite 210 | Hoboken | NJ | 07030 |
| 910 Cedar Avenue | | Providence | RI | 02903 |
| 36 Orchard Way | Unit 2 | Albany | NY | 12207 |

The TEXTJOIN formulas in F and G give:

| TEXTJOIN (skips blanks) | With ZIP |
|---|---|
| 118 Elm Street, Apt 4B, Boston, MA | 118 Elm Street, Apt 4B, Boston, MA 02108 |
| 2450 Harbor Road, Portland, ME | 2450 Harbor Road, Portland, ME 04101 |
| 77 Mill Lane, Suite 210, Hoboken, NJ | 77 Mill Lane, Suite 210, Hoboken, NJ 07030 |
| 910 Cedar Avenue, Providence, RI | 910 Cedar Avenue, Providence, RI 02903 |
| 36 Orchard Way, Unit 2, Albany, NY | 36 Orchard Way, Unit 2, Albany, NY 12207 |

Rows 3 and 5 have no unit, and TEXTJOIN simply skips it. The ampersand version in column H handles those two rows differently. For 2450 Harbor Road it returns `2450 Harbor Road, , Portland, ME`, with a double comma where the unit would be.

The ZIP column stays text, so 02108 and 02903 keep their leading zeros in the joined results.

## Troubleshooting

### The joined address has a double comma
You are probably joining with `&`, as in column H. Each `&` and each `", "` is added whether or not the cell has a value. In the template, the two rows without a unit show the double comma. Switch to `=TEXTJOIN(", ",TRUE,A2:D2)`.

### The ZIP code lost its leading zero
A ZIP typed into a General or Number cell turns 02108 into 2108 before any formula sees it. Format column E as Text first and retype the codes. The template's ZIP column is already Text, and the With ZIP results keep 02108 and 02903 intact. The [split city, state and ZIP post](https://tidytabs.blogspot.com/2026/09/how-to-split-city-state-and-zip-into.html) covers the reverse job, taking one line apart.

### Excel shows #NAME?
The likely cause is an older Excel. Microsoft's TEXTJOIN page lists Excel 2019, Excel 2021 and Microsoft 365, so Excel 2016 and earlier have no such function. This was not reproduced here, since only LibreOffice was used. In an older version, use the `&` join and delete the extra commas by hand.

### A blank unit still leaves a gap
Check the Unit cell for a space. A cell with a single space is not empty, so TEXTJOIN with TRUE treats it as text and keeps it. This was not reproduced in the template, whose empty cells are truly empty. Select the cell and press Delete.

## Template

[Download the .xlsx](templates/tidy-tabs-combine-address-columns.xlsx)

The file has two tabs. **Addresses** holds the five sample rows, the TEXTJOIN column, the version with ZIP, and the ampersand column for comparison. **How to use** has short fill-in notes, including the Excel version note.

The addresses are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
