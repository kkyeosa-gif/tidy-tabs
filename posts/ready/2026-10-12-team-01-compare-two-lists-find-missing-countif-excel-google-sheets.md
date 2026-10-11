---
title: How to Compare Two Lists and Find Missing Items in Excel and Google Sheets
labels: excel-formulas, data-cleanup, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there on 10/11/2026 and every result was checked against hand math, including a trailing-space case and a wildcard case. Excel and Google Sheets were not opened; the one help link is from Microsoft Support.
search_description: Compare two lists and flag what is missing with =IF(COUNTIF($D$2:$D$200,A2)=0,"Not found","Found"). Case is ignored. Free .xlsx template.
image_prompts: Two small stacks of paper index cards side by side on a wooden desk, one card pulled out and lying apart, a brass paperclip nearby, soft window light, no screens, no readable text, no logos
image_alt: Two stacks of index cards on a desk with one card set apart
threads: Two email lists and no idea who is on only one of them?\n=IF(COUNTIF($D$2:$D$200,A2)=0,"Not found","Found") flags every item in list A that list B is missing.\nCapitals do not matter, but one trailing space makes a false mismatch.
---
Flag every item in list A that is missing from list B with `=IF(COUNTIF($D$2:$D$200,A2)=0,"Not found","Found")`, then fill the formula down. COUNTIF ignores capital letters, so only a different spelling or a stray space at the end of a cell counts as a mismatch.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against hand math. The Excel and Google Sheets steps use the same function. Neither Excel nor Google Sheets was opened for this post.

![Customer emails in list A and list B emails with four Not found flags](images/2026-10-12-team-01-compare-two-lists-find-missing-countif-excel-google-sheets/compare-two-lists-find-missing-template.png) *Render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

`COUNTIF` counts how many cells in a range equal a value. If an item shows up zero times in the other list, it is missing from that list.

1. Type list A in **A2:A200** and put a header in **A1**.
2. Type list B in **D2:D200** and put a header in **D1**. Columns B, C and E are left for the results.
3. In **B2**, type `=IF(COUNTIF($D$2:$D$200,A2)=0,"Not found","Found")` and fill it down next to list A.
4. In **E2**, type `=IF(COUNTIF($A$2:$A$200,D2)=0,"Not found","Found")` and fill it down next to list B.
5. Add conditional formatting on the result columns so Not found cells turn red.

The `$` signs lock the other list's range, so it stays `$D$2:$D$200` as you fill down. The reference to `A2` has no `$`, so it moves to `A3`, `A4` and so on.

The template wraps the formula in a blank guard so empty rows stay empty: `=IF(A2="","",IF(COUNTIF($D$2:$D$200,A2)=0,"Not found","Found"))`. It is filled to row 15. If your lists are longer, fill down further. The ranges stop at row 200.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) page. This was not opened in Excel for this post.

### In Google Sheets

The same formulas should work as typed. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. List A is 10 customer emails and list B is 8 subscriber emails.

| Cell | List A (customers) | Result in B | List B (subscribers) | Result in E |
|---|---|---|---|---|
| Row 2 | ana.gomez | Not found | | |
| Row 4 | | | TINA.BROOKS | Found |
| Row 7 | maya.cole | Not found | | |
| Row 8 | omar.said | Not found | chris.doe | Not found |
| Row 9 | jo.hart | Not found | pat.lane | Not found |
| Row 10 | Tina.Brooks | Found | | |

Only the interesting rows are shown. So four customers are not in the subscriber list (**B2**, **B7**, **B8**, **B9**), and two subscribers are not in the customer list (**E8**, **E9**).

Tina.Brooks in **A10** matches TINA.BROOKS in **D4** and returns Found, because capitals are ignored.

The Summary tab counts the results:

| Cell | What it shows | Value |
|---|---|---|
| B2 | Items in list A | 10 |
| B3 | Found | 6 |
| B4 | Not found | 4 |
| B5 | Items in list B | 8 |
| B6 | Found | 6 |
| B7 | Not found | 2 |
| B8 | Check for list A | OK |
| B9 | Check for list B | OK |

**B8** and **B9** use `=IF(B3+B4=B2,"OK","CHECK")` and `=IF(B6+B7=B5,"OK","CHECK")`. They show CHECK if the found and not found counts do not add up to the list size. In the recalculated file, every value matched a hand check.

## Troubleshooting

### Two items look identical but show Not found
There is probably a space at the end of one cell. In the recalculated file, adding a trailing space to D5 ("sam.patel@example.com ") made **B4** Not found and **E5** Not found, a false mismatch on both sides. Wrap the value in `TRIM`, or clean the cells first. Microsoft documents it on its [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) page.

### An item returns Found when it should not
Check for a `*`, `?` or `~` in the item. These are wildcards inside COUNTIF. In the recalculated file, a `*` typed as an item in **A9** returned Found because it matches everything in list B. With that edit and the trailing space together, the Summary showed B3 6, B4 4, B6 5 and B7 3, so the totals no longer tell the whole story.

### Near spellings show as Not found
COUNTIF matches whole cells exactly, apart from capitals. It does not do fuzzy matching, so "Jon Hart" and "Jo Hart" are different items. Fix the spelling or add a clean-up column first.

### Items past row 15 have no result
The template formulas are filled to row 15. Fill **B2** and **E2** down as far as your lists go. Lists longer than 199 items also need the `$200` in the ranges changed.

### Not found does not say why
The sheet shows what is missing, not the reason. An unsubscribed customer or a typo in an email address looks the same. Check your records for each one.

## Template

[Download the .xlsx](templates/tidy-tabs-compare-two-lists.xlsx)

The file has three tabs. **Lists** holds the 10 sample customer emails in A2:A11 and 8 subscriber emails in D2:D9, with the results in columns B and E and red fill on Not found. **Summary** has the counts and checks in B2:B9. **How to use** has short fill-in notes.

The emails are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
