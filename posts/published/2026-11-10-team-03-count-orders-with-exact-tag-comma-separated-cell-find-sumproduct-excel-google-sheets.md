---
threads_url: https://www.threads.com/@tin_ylab/post/DflJGNliU3Y
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-count-orders-with-exact-tag-in.html
title: How to Count Orders With an Exact Tag in a Comma Separated Cell in Excel and Google Sheets
labels: excel-formulas, small-business, etsy-sellers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 26 workbooks and 1,115 comparisons. Checked: the sample, gift against gifted, a gift and gift wrap, a tag with a comma at either end, spaces before and after commas, capitals, the same tag twice in one order, empty tags (,,), blank rows, a tag typed with a comma (shows Check), tags containing * ? and ~ (orders and amounts compared; the COUNTIF column was not), 200 orders, and 14 random sets of up to 200 rows with text and blank amounts. Excel and Google Sheets were not used. Not run: tab characters in a tag cell, non-breaking spaces, wildcard characters in the COUNTIF column, a non-US language setting, more than 200 orders.
search_description: Count orders that carry an exact tag in a comma separated cell with FIND and SUMPRODUCT, so gift does not count gifted or gift wrap. Free template.
image_prompts: Small gift boxes tied with twine and a stack of paper order slips with colored sticky tabs on a craft studio table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Gift boxes tied with twine next to paper order slips with colored tabs
threads: COUNTIF with *gift* says 7 orders. Only 4 are tagged gift.\nPut a comma around the cleaned tags and use FIND inside SUMPRODUCT, so gifted and gift wrap stop counting.\nSample orders are made up.
---
Put a comma at both ends of the cleaned tags, then count a tag with `=SUMPRODUCT(--ISNUMBER(FIND(","&LOWER(TRIM(F4))&",",$D$4:$D$203)))`. In the sample, gift finds 4 orders, while `=COUNTIF(B4:B203,"*gift*")` finds 7 because it also counts gifted and gift wrap.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (26 workbooks, 1,115 comparisons). The formulas use only IF, TRIM, LOWER, SUBSTITUTE, FIND, ISNUMBER, SUMPRODUCT and COUNTIF, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Orders with typed tags next to a table of exact tag counts versus partial match counts](images/2026-11-10-team-03-count-orders-with-exact-tag-comma-separated-cell-find-sumproduct-excel-google-sheets/count-exact-tag-template.png) *PDF export of the template in LibreOffice Calc 24.2.7.2, sample data is fictional*

## Steps

The steps are the same in Excel and Google Sheets. Put the order number in column **A**, the tags in column **B** (separated by commas) and the amount in column **C**, with headers in row 3 and orders from row 4. Type the tags you want to count in **F4:F13**.

1. In **D4**, type `=","&SUBSTITUTE(SUBSTITUTE(LOWER(TRIM(B4)),", ",",")," ,",",")&","`. This lowercases the tags, removes spaces next to commas and adds a comma at both ends. The entry `gift , personalized` becomes `,gift,personalized,`.
2. Fill **D4** down to your last order.
3. In **G4**, type `=SUMPRODUCT(--ISNUMBER(FIND(","&LOWER(TRIM(F4))&",",$D$4:$D$203)))`. It looks for the tag with a comma on each side, so `gift` finds `,gift,` but not `,gifted,` or `,gift wrap,`.
4. For the dollars on those orders, type this in **H4**:

   `=SUMPRODUCT(--ISNUMBER(FIND(","&LOWER(TRIM(F4))&",",$D$4:$D$203)),$C$4:$C$203)`

5. Fill **G4:H4** down beside each tag.

The `--` turns each TRUE or FALSE into 1 or 0, and SUMPRODUCT adds them up. An order that lists the same tag twice, such as `gift, gift`, counts once.

Capitals and spaces around commas do not matter. Spaces inside a tag do, so `gift wrap` is one tag and `Gift Wrap` matches it.

The template wraps these formulas so a blank cell in **D** stays blank and a tag typed with a comma shows Check instead of a count. The core formulas above are what do the work.

Help pages: Microsoft's [FIND function](https://support.microsoft.com/en-us/office/find-find-b-function-c7912941-af2a-4bdf-a553-d0d89b0a0628), [SUMPRODUCT function](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) and [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34), and Google's [COUNTIF](https://support.google.com/docs/answer/3093480). They describe the functions, not our test.

## Example

Sample data is fictional: 12 orders numbered 1001 to 1012, with tags typed the way people type them.

| Order | Tags as typed | Amount |
|---|---|---|
| 1001 | gift, rush | $24.00 |
| 1002 | Gift Wrap | $18.00 |
| 1003 | gifted, repeat | $32.50 |
| 1004 | wholesale | $210.00 |
| 1005 | gift , personalized | $45.00 |
| 1006 | (blank) | $15.00 |
| 1007 | rush, gift wrap, repeat | $62.00 |
| 1008 | personalized | $28.00 |
| 1009 | `  GIFT  ` | $19.50 |
| 1010 | repeat, wholesale | $180.00 |
| 1011 | gift, gift | $21.00 |
| 1012 | rush | $36.00 |

Counting each tag gives:

| Tag | Orders (exact) | Amount | COUNTIF with * | Difference |
|---|---|---|---|---|
| gift | 4 | $109.50 | 7 | 3 |
| gift wrap | 2 | $80.00 | 2 | 0 |
| rush | 3 | $122.00 | 3 | 0 |
| repeat | 3 | $274.50 | 3 | 0 |
| wholesale | 2 | $390.00 | 2 | 0 |
| personalized | 2 | $73.00 | 2 | 0 |

Gift matches orders 1001, 1005, 1009 and 1011, which add up to $24.00 + $45.00 + $19.50 + $21.00 = $109.50. Order 1011 lists gift twice and counts once.

The COUNTIF column says 7 because `*gift*` also matches 1002 (Gift Wrap), 1003 (gifted) and 1007 (gift wrap). That is the 3 in the Difference column.

The summary block shows 12 orders listed, 1 order with no tags (1006), 0 amounts that are not numbers, 0 tags with a comma, and CHECK OK.

## Troubleshooting

### A tag shows Check instead of a count
The tag you typed in column **F** contains a comma. The formula splits tags on commas, so one tag cannot hold one. Type each tag in its own cell.

### The count is lower than expected
The tag in the order and the tag you typed must be the whole text between commas. `gift` does not match `gift wrap`, `gifts` or `a gift`. Spaces inside a tag count, so `gift  wrap` with two spaces does not match `gift wrap`.

### The dollar amount is lower than the order count suggests
A text amount, such as `$24.00` typed with a leading apostrophe, is skipped in the Amount column and counted in the summary block as an amount that is not a number. Convert those cells to numbers first.

### My tags use * or ? and the COUNTIF column acts strangely
In COUNTIF, the characters `*`, `?` and `~` are wildcards. In the FIND formulas in **G** and **H**, they are searched as plain characters, which the test confirmed for orders and amounts. The COUNTIF column was not compared for tags containing them, so do not rely on it for those.

### A pasted tag list still does not match
Pasted text can carry a tab character or a non-breaking space, which look like a space but are different characters. TRIM removes only ordinary spaces. Neither a tab nor a non-breaking space was tested here, in LibreOffice or anywhere else. If a paste does not match, retype the spaces by hand.

## Template

[Download the .xlsx](templates/tidy-tabs-count-exact-tag.xlsx)

The file has two tabs. **Orders** has 200 slots from row 4 to row 203 and 12 fictional sample orders. Columns **A** to **D** hold the order, the tags as typed, the amount and the cleaned tags. The tag table in **F3:J13** has 10 slots (6 filled): the tag to count, orders with this tag, the amount of those orders, the COUNTIF with * result and the difference. A summary block in **L3:M7** counts orders listed, orders with no tags, amounts that are not numbers, tags with a comma, and CHECK. **How to use** has notes on each column.

Yellow cells are inputs and gray cells are formulas. The page prints on US Letter, landscape.

To open it in Google Sheets, go to **File > Import > Upload** and select the file. The file was checked in LibreOffice Calc only.
