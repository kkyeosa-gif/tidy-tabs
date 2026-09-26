---
title: How to Remove Duplicate Rows From a Customer Email List in Excel and Google Sheets
labels: data-cleanup, spreadsheet-basics
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A small business desk with a laptop showing a blurred spreadsheet of customer contacts, a coffee cup, and a notepad with scribbled email addresses, all text unreadable
image_alt: Desk scene with laptop open to a blurred customer contact spreadsheet
search_description: Remove duplicate customer emails in Excel or Google Sheets using the built-in Remove Duplicates tool, plus a formula way to flag duplicates before you delete anything.
threads: Sending the same newsletter twice to one customer? Check your list for duplicate emails first.\nExcel and Sheets both have a built-in Remove Duplicates tool, no formulas needed. Just watch out for trailing spaces and mismatched capitalization, those slip past it.
---
Select your customer list, then use **Data > Remove Duplicates** in Excel or **Data > Data cleanup > Remove duplicates** in Google Sheets. Both tools compare entire rows or just the columns you pick, and remove every repeat after the first match.

> Works in: Excel (Microsoft 365 and Excel 2016+) and Google Sheets, current version. Steps confirmed against Microsoft and Google help pages, not hands-on tested.

## Steps

### In Excel

1. Select the range with your customer data, including the header row.
2. Go to **Data > Remove Duplicates**.
3. In the dialog box, check **My data has headers** if your top row has column labels.
4. Check the boxes for the columns to compare. For an email list, check only the **Email** column so any row with a repeated address counts as a duplicate.
5. Click **OK**. Excel tells you how many duplicate values it removed and how many unique values remain.

### In Google Sheets

1. Select the range with your customer data, including the header row.
2. Go to **Data > Data cleanup > Remove duplicates**.
3. In the panel, confirm **Data has header row** is checked.
4. Under "Analyze which columns," uncheck any column you don't want compared, then keep only **Email** checked if that's your duplicate test.
5. Click **Remove duplicates**. Sheets shows how many duplicate rows it found and deleted.

## Example

Sample data below is fictional, made up for this post.

| Name | Email | Phone | Signup Date |
|---|---|---|---|
| Maria Chen | maria@brightbakery.com | 617-555-0134 | 03/02/2026 |
| James Ortiz | james.ortiz@gmail.com | 508-555-0198 | 03/04/2026 |
| Maria Chen | maria@brightbakery.com | 617-555-0134 | 03/09/2026 |
| Priya Nair | priya.n@outlook.com | 401-555-0177 | 03/11/2026 |

After running Remove Duplicates on the Email column, the second Maria Chen row is deleted and three rows remain.

## Troubleshooting

### It didn't catch obvious duplicates
Check for trailing spaces or a stray tab character in the email column. Both Excel and Sheets treat `maria@brightbakery.com ` (with a space) as different from `maria@brightbakery.com`. Run **Find & Replace** first to strip extra spaces, or wrap the email column in a helper column with `=TRIM(B2)` and paste the results as values before removing duplicates.

### It caught too many rows
If you selected the whole row instead of just the Email column, any tiny difference in Name or Phone formatting keeps a row from matching, but if you only checked Email, two different customers who happen to share a shared inbox (like a shared family or business address) will get merged into one. Recheck which columns you selected for comparison.

### Uppercase and lowercase emails aren't matching as duplicates
Excel's Remove Duplicates treats text comparisons as case-insensitive for this feature, so `MARIA@brightbakery.com` and `maria@brightbakery.com` should already match there. Case handling in Google Sheets may not behave the same way, so if two entries that differ only in capitalization aren't being flagged, try normalizing the column first with `=LOWER(B2)` in a helper column before deduping, or check for a hidden space or a domain typo instead.

### The header row got deleted or treated as data
If **My data has headers** (Excel) or **Data has header row** (Sheets) wasn't checked, your column labels get treated as a regular row and can be flagged as a duplicate or left in the wrong place. Undo with **Ctrl+Z** (Windows) or **Cmd+Z** (Mac), then rerun with the header box checked.

### You need to see duplicates before deleting anything
Use a helper column instead of deleting right away. In an empty column next to your data, type `=COUNTIF($B$2:B2,B2)>1` (adjust B for your Email column) and fill it down. Any row marked TRUE is a repeat of an earlier row, so you can review the list before running Remove Duplicates.

## Copy-paste setup

Use these headers in row 1 of a new sheet:

| A | B | C | D | E |
|---|---|---|---|---|
| Name | Email | Phone | Signup Date | Is Duplicate |

In cell E2, type:

`=COUNTIF($B$2:B2,B2)>1`

Fill this down through every row of data. Any row showing TRUE has an email address that already appeared above it. Sort or filter by column E to review those rows, then delete them manually, or select the whole range and run **Data > Remove Duplicates** (Excel) or **Data > Data cleanup > Remove duplicates** (Google Sheets) on the Email column once you're confident.

Official references:
- [Find and remove duplicates - Microsoft Support](https://support.microsoft.com/en-us/office/find-and-remove-duplicates-00e35bea-b46a-4d5d-b28e-66a552dc138d)
- [Remove duplicate values - Microsoft Support](https://support.microsoft.com/en-us/office/remove-duplicate-values-ccf664b0-81d6-449b-bbe1-8b16050e4550)
- [Find and remove duplicate data - Google Docs Editors Help](https://support.google.com/docs/answer/3403979)
