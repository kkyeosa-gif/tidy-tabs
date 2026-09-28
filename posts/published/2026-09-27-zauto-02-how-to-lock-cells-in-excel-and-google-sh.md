---
threads_url: https://www.threads.com/@tin_ylab/post/Dd1JxcxloLg
blogger_url: https://tidytabs.blogspot.com/2026/09/how-to-lock-cells-in-excel-and-google.html
title: How to Lock Cells in Excel and Google Sheets So Formulas Don't Get Overwritten
labels: sheet-protection, spreadsheet-basics
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A small desk with a printed spreadsheet page covered in red pen circles around formula cells, next to a padlock key sitting on top of a laptop keyboard, no visible screen text.
image_alt: Printed spreadsheet with circled formula cells and a padlock key on a keyboard
search_description: Lock formula cells and protect a worksheet in Excel or Google Sheets so coworkers and clients can only edit the input cells you allow.
threads: Ever open a shared invoice template and find someone typed over your SUMIFS formula\nLock the formula cells and protect the sheet so only the blank input cells stay editable\nExcel calls it Protect Sheet, Google Sheets calls it Protect sheets and ranges, same idea either way
---

Lock the cells that hold formulas or headers, then turn on sheet protection so no one can type over them by accident. In Excel, select the cells to protect, open **Format Cells > Protection**, check **Locked**, then go to **Review > Protect Sheet**. In Google Sheets, select the range, go to **Data > Protect sheets and ranges**, and set who can edit it.

> Works in: Excel (Windows and Mac, current Microsoft 365 versions) and Google Sheets. Steps come from Microsoft and Google's own help pages; not tested hands-on for this post.

![Spreadsheet example with columns Item, Quantity (editable), Unit Price, Total (locked formula)](images/2026-09-27-zauto-02-how-to-lock-cells-in-excel-and-google-sh/how-to-lock-cells-in-excel-and-google-sheets-so-fo-example.png) *The example from this post laid out in a spreadsheet. Rendered with LibreOffice Calc.*

## Steps

By default, every cell in a new Excel worksheet is already marked "locked" in its cell format, but that setting has no effect until you protect the sheet — so you need to unlock the cells you want people to type in before you protect the rest. Google Sheets works a bit differently: cells aren't marked locked by default, so you explicitly choose which sheet or range to protect using **Protect sheets and ranges**.

### In Excel

1. Select the cells you want to stay editable, such as blank input cells on an invoice or order form.
2. Press **Ctrl+1** on Windows (**Cmd+1** on Mac), or right-click and choose **Format Cells**.
3. Click the **Protection** tab and uncheck **Locked**. Click **OK**.
4. Go to **Review > Protect Sheet**.
5. Check **Protect worksheet and contents of locked cells**.
6. Type a password if you want one, or leave it blank to just prevent casual edits.
7. Click **OK**. Everything you didn't unlock in step 3 is now locked.

### In Google Sheets

1. Select the range or entire sheet you want to protect.
2. Go to **Data > Protect sheets and ranges**.
3. Click **Set permissions**.
4. Choose **Only you**, or click **Custom** to pick specific people who can still edit.
5. Click **Done**.
6. To leave some cells open for everyone, protect the whole sheet first, then use the **Except certain cells** option in the same panel and select the input range to exclude.

## Example

A freelancer sends this order form to a client to fill in. The quantity and notes columns stay open. The formula column is locked so the total can't get overwritten by a stray keystroke.

| Item | Quantity (editable) | Unit Price | Total (locked formula) |
|---|---|---|---|
| Custom logo mug | 24 | $12.50 | $300.00 |
| Tote bag | 50 | $8.00 | $400.00 |
| Sticker sheet | 100 | $2.25 | $225.00 |

Sample data above is fictional.

## Troubleshooting

### Excel still lets people edit locked cells
Locking a cell only takes effect after you turn on **Review > Protect Sheet**. If you skipped that step, every cell behaves as if it's unlocked.

### Google Sheets shows a warning but still allows edits
"Warning" permission lets anyone edit after clicking through a pop-up. Choose **Restrict who can edit this range** instead if you want a hard block.

### Formulas disappear after a client pastes data
Pasting can overwrite a locked cell's format along with its content if the paste happens before protection is active, or if the cell was never actually locked. Re-check the cell's Protection setting after any bulk paste and re-protect the sheet.

### Forgot the Excel password
Microsoft does not offer password recovery for worksheet protection. Keep a copy of the password in a password manager, or skip the password and rely on locking alone if the goal is just to prevent accidents, not enforce security.

### Protection blocks your own edits later
If you protected the sheet as "Only you" in Google Sheets from a different Google account, or Excel forgot your saved password, you'll need the original account or password to unprotect it. Keep an unprotected backup copy of any template you plan to keep editing yourself.

## Copy-paste setup

Use this as a checklist when protecting a shared template:

- Unlock the input cells first (Excel: **Format Cells > Protection > uncheck Locked**; Sheets: exclude them under **Except certain cells**).
- Protect the rest of the sheet (Excel: **Review > Protect Sheet**; Sheets: **Data > Protect sheets and ranges**).
- Leave the password blank in Excel if the goal is just to stop accidental typing, not to secure sensitive data.
- In Sheets, set permissions to **Only you** for templates you send out, or **Custom** if a coworker also needs edit rights.
- Test the result by opening the file as a different user (or in incognito, for Sheets) and confirming the input cells work but the formula cells don't.

Official references:
- [Protect a worksheet – Microsoft Support](https://support.microsoft.com/en-us/office/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4d815d5b18f)
- [Lock cells to protect them – Microsoft Support](https://support.microsoft.com/en-us/office/lock-cells-to-protect-them-fa383f3c-f907-4f96-8bef-a72c0fb2ea41)
- [Protect sheets and ranges – Google Docs Editors Help](https://support.google.com/docs/answer/1218656)
