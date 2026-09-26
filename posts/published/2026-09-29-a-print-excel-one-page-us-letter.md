---
threads_url: https://www.threads.com/@tin_ylab/post/DduBYXWCco7
blogger_url: https://tidytabs.blogspot.com/2026/09/how-to-print-excel-sheet-on-one-page-us.html
title: How to Print an Excel Sheet on One Page (US Letter)
labels: excel-printing, page-setup
tested_in: LibreOffice Calc 24.2.7 (Linux), 09/25/2026; Excel and Google Sheets steps from Microsoft/Google help pages
image_prompts: A single landscape-oriented printed spreadsheet page on US Letter paper coming out of a small home-office laser printer, the table too small to read
image_alt: One printed spreadsheet page coming out of a home-office printer
search_description: To print an Excel sheet on one page, set Paper Size to Letter and pick Fit Sheet on One Page under File > Print. Google Sheets steps too.
threads: A 48-row price list printed on 6 pages with default settings. Same file, 1 page after three changes.\nLetter paper, Landscape, fit to 1 page wide by 1 tall.\nFor long lists, pick Fit All Columns on One Page instead so the text stays readable.
---
To print an Excel sheet on one page, go to **File > Print**, check that the paper size is Letter, and change the scaling setting to **Fit Sheet on One Page**. The ribbon route is **Page Layout > Scale to Fit**: set **Width** to 1 page and **Height** to 1 page.

> Works in: Excel for Microsoft 365 on Windows (Excel 2016 and later use the same Page Setup dialog), Google Sheets in a desktop browser. Tested in: LibreOffice Calc 24.2.7 only. The Excel and Google Sheets steps come from Microsoft and Google help pages.

![Price list printed on six pages with default settings next to the same list fitted on one page](images/2026-09-29-a-print-excel-one-page-us-letter/print-excel-one-page-before-after.png) *Left: default settings, 6 pages. Right: the same file on one US Letter page. PDF exports from LibreOffice Calc 24.2.*

## Steps

### In Excel

1. Open **Page Layout** and click the small launcher arrow in the **Page Setup** group.
2. On the **Page** tab, set **Paper size** to **Letter**. Downloaded templates sometimes arrive set to A4, so check this first.
3. Under **Orientation**, pick **Landscape** if the sheet is wider than it is tall.
4. Under **Scaling**, select **Fit to** and enter 1 page(s) wide by 1 tall. Click **OK**.
5. Go to **Page Layout > Margins** and click **Narrow**, or click **Custom Margins** and type your own sizes.
6. To print only part of the sheet, select those cells and click **Page Layout > Print Area > Set Print Area**.
7. Click **File > Print** and check the preview before you print.

Microsoft's [Fit to one page in Excel](https://support.microsoft.com/en-us/excel/fit-to-one-page-in-excel) page covers steps 1 to 4. The **File > Print** scaling menu also offers **Fit All Columns on One Page**, described on the [Scale a worksheet](https://support.microsoft.com/en-us/excel/scale-a-worksheet) page. These steps are for the Excel desktop app.

**Fit Sheet on One Page or Fit All Columns on One Page?** A short sheet fits nicely on one page. A long list squeezed onto one page prints in text too small to read. For a long list, choose **Fit All Columns on One Page**: every column fits across and the rows continue onto page 2 and beyond.

When rows run onto more pages, repeat the header row:

1. Click **Page Layout > Print Titles**.
2. On the **Sheet** tab, click in **Rows to repeat at top** and type `$1:$1`.
3. Click **OK**.

To see where each page ends, click **View > Page Break Preview**. Dashed lines are breaks Excel added. Drag one to move it.

### In Google Sheets

1. Click **File > Print**.
2. Under **Paper size**, choose **Letter**.
3. Under **Orientation**, choose **Landscape**.
4. In the scale setting, choose the option that fits the sheet to the page width, or the one that fits the whole page. Google's help page does not list these option names, so match the wording you see.
5. Under **Margins**, click **Normal** and pick **Narrow**, or pick **Custom numbers**.
6. To repeat the header row, freeze row 1 first. Then in the print settings, open **Headers & footers** and choose **Repeat frozen rows**.
7. Click **Next**.

Google's [Print from Google Sheets](https://support.google.com/docs/answer/7663148?hl=en&co=GENIE.Platform%3DDesktop) page covers paper size, orientation, margins, frozen rows, and custom page breaks.

## Example

We checked this in LibreOffice Calc 24.2, using a sample price list with 48 rows and 14 columns: SKU, Item, Category, Size, Color, Unit cost, Wholesale, Retail, Margin %, On hand, Reorder at, Supplier, Lead time (days), and Last ordered.

| Setting | Before | After |
| --- | --- | --- |
| Orientation | Portrait | Landscape (11 x 8.5 in) |
| Scaling | None | 1 page wide by 1 tall |
| Margins | Default | 0.25 in left/right, 0.5 in top/bottom |
| Header row repeats | No | Yes, row 1 |
| Pages on US Letter | 6 | 1 |

With the default settings, the PDF came out to 6 US Letter pages, and the columns were split across pages. The prices landed on pages 3 and 4, and Supplier and Last ordered on pages 5 and 6, away from the item names.

![Six printed pages with the price list columns split across them](images/2026-09-29-a-print-excel-one-page-us-letter/before-six-pages-libreoffice.png) *Before: default settings, 6 pages. PDF export from LibreOffice Calc 24.2.*

After the changes in the table, the same file printed on one US Letter page turned sideways.

![The same price list fitted on one sideways US Letter page](images/2026-09-29-a-print-excel-one-page-us-letter/after-one-page-libreoffice.png) *After: Landscape, fit to 1 page, narrow margins. PDF export from LibreOffice Calc 24.2.*

Both images are LibreOffice exports of the same .xlsx file. The page setup is stored inside the file as standard Excel settings (Letter paper, Landscape, fit to 1 page wide by 1 tall, row 1 as the print title), so Excel should open it with the same setup. We have not printed it from Excel ourselves.

## Troubleshooting

### The text is too small to read
Fitting a long sheet onto one page shrinks everything. Switch to **Fit All Columns on One Page**, use Landscape, or hide columns you don't need on paper (right-click the column letter, then **Hide**).

### An extra blank page prints at the end
Stray formatting or a stray space in a far-off cell makes Excel think the sheet is bigger. Set a print area around your real data, or open **Page Break Preview** to see which cells Excel plans to print.

### The header row is missing on page 2
Print titles are not turned on. Add `$1:$1` in **Rows to repeat at top** under **Page Layout > Print Titles**. In Google Sheets, freeze row 1 and choose **Repeat frozen rows**.

### It still prints on A4
The paper size is saved per sheet, so a template made outside the US can still be set to A4. Change **Paper size** to **Letter** in the Page Setup dialog, and check that your printer has Letter paper loaded.

### Google Sheets cuts off the right-hand columns
The scale setting is still on normal size. Change it to the fit-to-width option, switch to Landscape, or use **Set custom page breaks** to control where each page ends.

## Template

The download is the price list from the example, already set up to print: US Letter, Landscape, fit to one page, header row repeating, and narrow margins (0.25 in left/right, 0.5 in top/bottom). It has 48 products; all names, suppliers, and prices are fictional sample data. Replace them with your own and the page setup stays in place.

[Download the one-page price list (.xlsx)](templates/tidy-tabs-print-one-page-us-letter.xlsx)

To use it in Google Sheets, click **File > Import > Upload** and select the file. Google Sheets uses its own print settings, so set Letter, Landscape, and the scale option in **File > Print** before you print.
