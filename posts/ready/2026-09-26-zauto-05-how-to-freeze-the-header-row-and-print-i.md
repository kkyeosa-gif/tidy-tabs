---
title: How to Freeze the Header Row and Print It on Every Page in Excel and Google Sheets
labels: page-setup, excel-printing
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A home office desk with a printed multi-page spreadsheet report fanned out, header row visible on each page but text unreadable, coffee cup nearby
image_alt: Printed spreadsheet pages fanned across a desk showing repeated header row
search_description: Freeze the header row on screen and force it to repeat on every printed page in Excel and Google Sheets, with exact menu paths for both apps.
threads: Your spreadsheet header row disappears after page 1 when you print it\nFreezing a row only fixes the on-screen view, not the printout\nExcel needs Page Layout > Print Titles, Sheets needs a checkbox in the print settings sidebar\nHere's exactly where to click
---
Freeze the header row so it stays visible while you scroll, and separately set it to repeat on every printed page. In Excel, use **Page Layout > Print Titles > Rows to repeat at top**. In Google Sheets, freeze the row from **View > Freeze**, then check **Repeat frozen rows** in the print settings.

> Works in: Excel for Windows and Excel for Mac (Microsoft 365 and other recent desktop versions), and in Google Sheets. Print Titles is a desktop feature; it isn't available in Excel for the web, so use the desktop app for that step. Steps come from Microsoft and Google support pages, not hands-on testing.

## Steps

### In Excel

1. Click any cell in row 1 (your header row).
2. Go to the **View** tab.
3. Click **Freeze Panes > Freeze Top Row**. This keeps the header visible while you scroll on screen.
4. Go to the **Page Layout** tab.
5. Click **Print Titles** in the Page Setup group.
6. In the **Rows to repeat at top** field, type `$1:$1`.
7. Click **OK**.
8. Go to **File > Print** and check the preview pane to confirm the header shows on every page.

### In Google Sheets

1. Click any cell in row 1.
2. Go to **View > Freeze > 1 row**. This freezes the header on screen.
3. Go to **File > Print**.
4. In the print settings sidebar, open the **Formatting** section.
5. Check the box for **Repeat frozen rows** (or **Repeat row headers**, depending on your version).
6. Check the preview on the right to confirm the header appears on every page before you print or export to PDF.

## Example

A two-page invoice log with headers Date, Customer, Amount should show the same header row at the top of page 1 and page 2 of the printout.

| Date | Customer | Amount |
|---|---|---|
| 09/02/2026 | Blue Harbor Cafe | $145.00 |
| 09/09/2026 | Maple & Vine Florist | $220.00 |
| ... (page break) | | |
| 09/16/2026 | Riverside Hardware | $310.00 |

Sample data is fictional, used only to show the layout.

## Troubleshooting

### Header shows on screen but not in the printout (Excel)
Freeze Panes only controls the on-screen view. You still need to set **Rows to repeat at top** under Print Titles, or the header won't repeat on paper.

### Print Titles option is grayed out
This happens when multiple sheet tabs are grouped. Right-click any sheet tab and click **Ungroup Sheets**, then try again.

### Google Sheets prints the header only on page 1
Open **File > Print** and check that **Repeat frozen rows** is checked under Formatting. If no rows are frozen yet, freeze row 1 first from **View > Freeze**.

### Wrong row repeats after inserting a new row above the header
Inserting a row can shift what's inside the Print Titles range. Reopen **Page Layout > Print Titles** and reset the range to the correct row number.

### Frozen row settings look off after applying a filter view in Sheets
Filter views can sometimes interact with how the frozen row displays. If this happens, turn off the filter view, confirm the freeze is still set under **View > Freeze**, then reapply the filter.

## Copy-paste setup

**Excel**
- View tab > Freeze Panes > Freeze Top Row
- Page Layout tab > Print Titles > Rows to repeat at top: `$1:$1`

**Google Sheets**
- View > Freeze > 1 row
- File > Print > Formatting section > check "Repeat frozen rows"

Helpful references:
- [Freeze or lock rows and columns (Microsoft)](https://support.microsoft.com/en-us/office/freeze-panes-to-lock-rows-and-columns-dab2ffc9-020d-4026-8121-67dd25f2508f)
- [Print rows with column headings on every printed page (Microsoft)](https://support.microsoft.com/en-us/office/print-rows-with-column-headings-on-every-page-fb17c051-582e-4831-8e1e-f45c0eddf341)
- [Freeze or unfreeze rows and columns (Google)](https://support.google.com/docs/answer/9060444)
