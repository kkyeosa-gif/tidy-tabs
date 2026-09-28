---
title: How to Highlight Duplicate Values in Excel and Google Sheets
labels: conditional-formatting, data-cleanup
tested_in: LibreOffice Calc 24.2.7 (Linux). Google Sheets/Excel conditional-formatting steps from Google Docs Editors Help / Microsoft Support (not tested here).
image_prompts: A small office counter with a stack of printed order slips fanned out, two identical slips pulled to the side next to each other, slip text blurred and unreadable, no visible logos.
image_alt: Two identical order slips pulled aside from a fanned out stack
search_description: Highlight duplicate values in Excel or Google Sheets with a COUNTIF conditional formatting rule, without deleting or moving any rows.
---
Select the column, open conditional formatting, and use the custom formula `=COUNTIF($B$2:$B$16,B2)>1` to highlight every value that repeats, without deleting or moving any rows.

> Works in: LibreOffice Calc 24.2.7 (Linux). Google Sheets/Excel conditional-formatting steps from Google Docs Editors Help / Microsoft Support (not tested here).

![Order list with repeated customer email addresses highlighted in pink across several rows](images/2026-09-29-team-05-highlight-duplicate-values-excel-google-sheets/highlight-duplicate-emails-template.png) *LibreOffice Calc 24.2 PDF export of the Orders tab with conditional formatting applied.*

This is different from deleting duplicate rows outright. If you want to remove repeat rows from a list entirely, that's a separate job for a Remove Duplicates tool. This post is for when a repeat might be legitimate, like a customer placing a second order, and you just want to see which rows share a value before deciding what to do with them. Every row stays put. Nothing gets deleted.

Sample data in this post is fictional, made up for illustration.

## Steps

### In Excel

1. Select the range you want to check, starting from the first data row (skip the header). For a list with emails in column B from row 2 to row 16, select `B2:B16`.
2. Go to **Home > Conditional Formatting > New Rule**.
3. Choose **Use a formula to determine which cells to format**.
4. Enter the formula `=COUNTIF($B$2:$B$16,B2)>1`, using your actual range and the first selected cell.
5. Click **Format**, pick a fill color on the **Fill** tab, and click **OK** twice.
6. Any cell whose value appears more than once in the range gets highlighted immediately.

### In Google Sheets

1. Select the range you want to check, such as `B2:B16`.
2. Go to **Format > Conditional formatting**.
3. Under "Format rules," set the rule type to **Custom formula is**.
4. Enter the formula `=COUNTIF($B$2:$B$16,B2)>1`.
5. Under "Formatting style," pick a fill color.
6. Click **Done**. Sheets applies the highlight to every matching cell right away.

The dollar signs in `$B$2:$B$16` matter. They lock the range so every row in your selection checks against the full list, not just the rows above it. The `B2` part (no dollar signs) is relative, so Excel and Sheets adjust it automatically as the rule applies to each row in your selection.

## Example

| Order # | Customer email | Duplicate? |
|---|---|---|
| 1001 | priya.nair@example.com | Duplicate |
| 1002 | sam.patel@example.com | Duplicate |
| 1003 | dana.ruiz@example.com | Duplicate |
| 1005 | emma.lee@example.com | (blank) |
| 1006 | jamie.chen@example.com | (blank) |
| 1008 | new.customer@example.com | (blank) |

The full sample has 15 rows. Four email addresses repeat: `priya.nair@example.com` (3 times), `sam.patel@example.com` (3 times), `dana.ruiz@example.com` (2 times), and `ana.gomez@example.com` (2 times). Those rows get highlighted, or return "Duplicate" in a helper column like the template's column C. The other 7 emails are unique and stay blank.

## Troubleshooting

### The highlight breaks after you copy or insert a row
This usually means a dollar sign is missing from the range part of the formula. If you typed `=COUNTIF(B2:B16,B2)>1` without locking the range, pasting or inserting rows shifts the range along with the cell, so the count goes wrong. Rewrite the formula with the range locked, like `=COUNTIF($B$2:$B$16,B2)>1`, and reapply the rule.

### Emails with different capitalization aren't flagged as duplicates, or the reverse
`COUNTIF` in both Excel and Google Sheets treats text comparisons as case-insensitive, so `Priya.Nair@example.com` and `priya.nair@example.com` should already count as the same value. If you're seeing something else, check for a trailing space or a typo in the domain instead of assuming a case problem. Adding `=TRIM(B2)` in a helper column can rule out stray spaces.

### The wrong column gets the color
This happens when the conditional formatting range doesn't match the column your formula references. If you applied the rule to `A2:A16` but the formula checks `B2`, the highlight lands on order numbers instead of emails. Open the rule, check the applies-to range, and make sure it matches the column in your formula.

### Every cell gets highlighted, even unique ones
If `COUNTIF` is pointed at the wrong range, such as the whole sheet or a range that includes the header row, it can inflate every count past 1. Double check the range in the formula matches your actual data rows, and that the header row itself isn't included in the highlighted selection.

### The rule disappears when you sort the data
Some spreadsheet versions anchor a conditional formatting range to specific rows rather than the data itself, so sorting can leave the rule covering the wrong rows or losing rows that moved below it. After a sort, open the conditional formatting rule and confirm the range still covers all your data rows.

## Template

[Download the .xlsx](templates/tidy-tabs-highlight-duplicates-sample.xlsx). It has one tab, **Orders**, with columns for Order #, Customer email, and a `Duplicate?` formula column, plus the matching conditional formatting rule already applied to the Customer email column.

To use it in Google Sheets, go to **File > Import > Upload** and select the downloaded file.

Official references:
- [Add, edit, or delete conditional formatting rules - Google Docs Editors Help](https://support.google.com/docs/answer/78413)
- [Add conditional formatting - Microsoft Support](https://support.microsoft.com/en-us/office/add-conditional-formatting-1948134f-28d3-4bb3-ac93-a2a1c33cb50e)
