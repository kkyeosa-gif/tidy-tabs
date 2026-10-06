---
threads_url: https://www.threads.com/@tin_ylab/post/DeKD7sEG6qP
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-clean-up-customer-list-with-trim.html
title: How to Clean Up a Customer List With TRIM and PROPER in Excel and Google Sheets
labels: data-cleanup, excel-formulas, spreadsheet-basics
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets steps come from Microsoft Support and Google Docs Editors Help (not tested here).
search_description: Clean up a customer list with =PROPER(TRIM(CLEAN(A2))): remove extra spaces, fix ALL CAPS, and handle non-breaking spaces and line breaks. Free .xlsx template.
image_prompts: A neat stack of paper index cards with handwritten names in uneven capitals and spacing, a few cards fanned out next to a rubber band and a mug of pens, small shop counter, all writing blurred and unreadable, no logos, no laptop
image_alt: Fanned handwritten index cards with uneven writing beside a rubber band and pen mug
threads: Pasted names from a web page and TRIM did nothing?\nThe culprit is usually a non-breaking space. SUBSTITUTE(A2,UNICHAR(160)," ") swaps it for a normal one first.\nUNICHAR, not CHAR(160), which missed it in LibreOffice here.
---
Type `=PROPER(TRIM(CLEAN(A2)))` next to the messy name and copy it down. It removes extra spaces, turns ALL CAPS into Proper Case, and strips most hidden characters.

This formula does not fix non-breaking spaces, and it joins words that were split by a line break. Both fixes are in the Steps below.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every result below was recalculated and checked against the expected text. The Excel and Google Sheets steps use the same functions and link to Microsoft Support and Google Docs Editors Help. Neither app was opened for this post.

![Clean names sheet showing messy customer names next to the cleaned versions](images/2026-10-02-team-05-clean-customer-list-trim-proper-excel-google-sheets/clean-customer-list-trim-proper-template.png) *LibreOffice Calc 24.2 PDF export of the Clean names sheet with sample data (fake). The length-check columns are hidden.*

## Steps

Put the messy names in column **A**, starting in **A2**. Leave your original column alone and build the clean version beside it.

1. In **B1**, type a header such as Basic clean. In **B2**, type `=PROPER(TRIM(CLEAN(A2)))`.
2. Copy **B2** down to the last row. Double-click the fill handle (the small square at the bottom right of the cell) to do it in one move.
3. Read the results. Extra spaces, ALL CAPS and all lowercase are fixed. Check any name that looks run together or still has a wide gap.
4. For names pasted from web pages, emails or a CRM export, use the full formula in **C2**: `=PROPER(TRIM(CLEAN(SUBSTITUTE(SUBSTITUTE(A2,UNICHAR(160)," "),CHAR(10)," "))))`.
5. Copy **C2** down the same way.
6. To keep the clean names, select column C, copy it, then paste as values only. Delete columns B and C if you no longer need the formulas.

What each function does:

- `TRIM` removes leading, trailing and doubled spaces.
- `CLEAN` removes most non-printing characters, including line breaks.
- `PROPER` capitalizes the first letter of every word and lowercases the rest.
- `SUBSTITUTE` swaps one character for another. Here it turns non-breaking spaces and line breaks into normal spaces before TRIM runs.

Use `UNICHAR(160)`, not `CHAR(160)`. In LibreOffice, `CHAR(160)` depends on the computer's locale and missed the non-breaking space on a UTF-8 setup, while `UNICHAR(160)` worked on both. How `CHAR(160)` behaves in Excel and Google Sheets was not checked.

### In Excel

The formulas work as typed. `UNICHAR` is available in Excel 2013 and later. For paste as values, use **Home > Paste > Paste Values**.

Microsoft's [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9) page says TRIM removes only the standard space (character 32) and not the non-breaking space (character 160). Its [CLEAN function](https://support.microsoft.com/en-us/office/clean-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41) page describes the non-printing characters CLEAN removes.

### In Google Sheets

The same formulas work as typed. For paste as values, use **Edit > Paste special > Values only**. Google documents the arguments on its [TRIM function](https://support.google.com/docs/answer/3094140) page.

## Example

Sample data below is fictional. A "line break" is Alt+Enter in Excel or Ctrl+Enter in Google Sheets inside one cell, and a "non-breaking space" is the invisible character 160 that often comes along with text copied from a web page.

| Original (column A) | Basic formula (column B) | Full formula (column C) |
|---|---|---|
| `  dana   ruiz ` | Dana Ruiz | Dana Ruiz |
| `EMMA LEE` | Emma Lee | Emma Lee |
| `sam` + non-breaking space + `patel` | Sam Patel (the space is still a non-breaking space) | Sam Patel |
| `jamie` + line break + `chen` | Jamiechen | Jamie Chen |
| `  PRIYA   NAIR  ` | Priya Nair | Priya Nair |
| `ana gomez` | Ana Gomez | Ana Gomez |
| `ronald McDONALD` | Ronald Mcdonald | Ronald Mcdonald |
| `sean o'neil` | Sean O'Neil | Sean O'Neil |
| `TOM  REYES` + trailing non-breaking space | Tom Reyes (the trailing space is still there) | Tom Reyes |
| `   lee wu` | Lee Wu | Lee Wu |

The template also counts characters before and after with `=LEN(A2)` and `=LEN(C2)`. For the Tom Reyes row, the original is 11 characters and the full formula leaves 9.

Row 3 shows why the basic formula can look fine and still cause trouble. "Sam Patel" in column B looks right, but a sort, a lookup, or a duplicate check can treat the non-breaking space as a different character from a normal one.

## Troubleshooting

### Two names run together, like Jamiechen
CLEAN deletes a line break instead of replacing it with a space, so "jamie" + line break + "chen" becomes "Jamiechen". Use the full formula in step 4. It swaps `CHAR(10)` for a space before CLEAN runs.

### A gap or trailing space is still there after TRIM
The space is probably a non-breaking space (character 160), which TRIM does not remove. Wrap the cell in `SUBSTITUTE(A2,UNICHAR(160)," ")` first, as in the full formula. `=LEN(A2)` against `=LEN(C2)` shows whether anything was removed.

### McDONALD became Mcdonald
PROPER lowercases everything after the first letter of each word, so it cannot know that "McDonald" has a second capital. Fix names like that by hand after you paste the values. The same goes for names such as "DeShawn" and "van der Berg".

### O'Neil or hyphenated names look wrong
In the LibreOffice check, `o'neil` came out as O'Neil as expected. Always scan your own list, because PROPER also capitalizes the letter after any non-letter character. Excel and Google Sheets were not checked on this.

### The formula returns #NAME?
Your app may not recognize `UNICHAR`, which Excel added in 2013. In an older version, use the basic formula and clear non-breaking spaces with **Find & Replace** instead.

## Template

[Download the .xlsx](templates/tidy-tabs-clean-customer-list.xlsx)

The file has two tabs. **Clean names** holds ten fictional names in column A, the basic formula in column B, the full formula in column C, and the Length before and Length after columns. The formula columns are shaded gray. **How to use** has short notes on each function and the McDonald limit.

The names are fictional. To open it in Google Sheets, go to **File > Import > Upload** and select the file, as described in Google's [import help](https://support.google.com/docs/answer/40608).

Once the names are clean, remove repeats with [How to Remove Duplicate Rows From a Customer Email List](https://tidytabs.blogspot.com/2026/09/how-to-remove-duplicate-rows-from.html), or separate first and last names with [How to Split a Full Name Into First and Last Name Columns](https://tidytabs.blogspot.com/2026/09/how-to-split-full-name-into-first-and.html).
