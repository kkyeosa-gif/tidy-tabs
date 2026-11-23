---
title: How to Use DSUM With Multiple Criteria (AND and OR) in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 36 workbooks and 18,900 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Criteria text matching and blank criteria rows were checked in LibreOffice Calc only; Excel's rule is quoted from Microsoft's help page. Not tested in Excel or Google Sheets.
search_description: Total and count sales with DSUM and DCOUNT and a criteria box: conditions on one row must all match (AND), each extra row is another way to match (OR).
image_prompts: A wooden craft fair table with rows of jar candles and paper-wrapped soap bars, a metal cash box with its lid open and a small spiral sales notebook lying beside it, warm indoor light, realistic photo, no screens, no readable text, no logos
image_alt: Craft fair table with jar candles, soap bars and a metal cash box
threads: In a DSUM criteria box, conditions on one row are AND, and each extra row is OR.\nUnder Channel, ="=Etsy" totals $104.50 across 4 sales in the sample.\nAdd ="=Market" on the row below and it becomes $256.50 across 7.
---
Type the sales headers over a small criteria box, put each condition under its header, and total with `=DSUM(Sales!$A$4:$E$504,"Amount",$A$3:$G$4)` (count with DCOUNT and the same arguments): conditions on one row must all match (AND), and each extra row is another way to match (OR). In the made-up shop report, ="=Etsy" under Channel totals $104.50 across 4 sales; add ="=Market" on the next row and the total becomes $256.50 across 7 sales.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (36 workbooks, 18,900 cell comparisons, 0 mismatches), plus the edge cases below. Text matching and blank criteria rows are LibreOffice Calc results; Excel's text rule is quoted from Microsoft's help page. Excel and Google Sheets were not opened for this post.

![Criteria box with =Etsy and =Market rows and DSUM total above sales list](images/2026-11-23-team-04-total-sales-with-criteria-box-and-or-conditions-dsum-dcount-excel-google-sheets/dsum-criteria-box-report-template.png) *Report tab with =Etsy on criteria row 4 and =Market typed on row 5 (the download opens with =Etsy alone) (Etsy OR Market), giving a total of $256.50 across 7 sales with a $36.64 average, above the 10 fictional sales it reads; the $240.00 Etsy Wholesale sale is left out. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. A criteria box is one row of labels with one or more rows of conditions under it.

1. On the **Sales** tab, type one sale per row from row 5 under the row 4 headers Date, Channel, Product, State, Amount.
2. On the **Report** tab, type the box labels in **A3:G3**: Date, Date, Channel, Product, State, Amount, Amount. Each must match a Sales header exactly. Date and Amount appear twice so one row can hold a from and a to.
3. Type conditions in **A4:G6**, each under the label it is about. Conditions on the same row are AND. A condition on the next row is OR.
4. For exact text, type `="=Etsy"`. The cell then shows =Etsy. Microsoft's [advanced criteria help](https://support.microsoft.com/en-us/office/filter-by-using-advanced-criteria-4c9222fe-8529-4cd7-a898-3f16abdff32b) says text typed without an equal sign matches values that begin with it, and it uses ="=Davolio" for an exact match. In LibreOffice Calc, plain Etsy matched whole cells only (Etsy and etsy) and left out Etsy Wholesale.
5. For numbers, type `>50` under Amount. For a range, type `>=20` under the first Amount and `<=29` under the second, on the same row.
6. For dates, type `=">="&DATE(2026,11,1)` under the first Date and `="<="&DATE(2026,11,30)` under the second. Dates typed as text in the box depend on your locale settings, so the DATE form is the safe one.
7. In **J3**, find the last row of the box you filled with `=IF(COUNTA(A6:G6)>0,3,IF(COUNTA(A5:G5)>0,2,1))`.
8. In **J4**, total with `=DSUM(Sales!$A$4:$E$504,"Amount",$A$3:INDEX($A$4:$G$6,$J$3,7))`. The INDEX part makes the criteria range end at the last filled row, so the box grows when you add an OR row. See the DSUM help from [Microsoft](https://support.microsoft.com/en-us/office/dsum-function-53181285-0c4b-4f5a-aaa3-529a322be41b) and [Google](https://support.google.com/docs/answer/3094281).
9. In **J5:J8**, use the same three arguments with DCOUNT (number of sales), DAVERAGE, DMAX and DMIN. DAVERAGE with no match is #DIV/0!, so the template wraps it in IFERROR.
10. Type at least one condition. **J9** counts them with `=COUNTA(A4:G6)` and must be 1 or more.

## Example

The ten sales are made up and total $603.50:

| Date | Channel | Product | State | Amount |
|---|---|---|---|---|
| 10/03/2026 | Etsy | Candle | MA | $28.00 |
| 10/04/2026 | Market | Soap | MA | $12.00 |
| 10/11/2026 | Etsy Wholesale | Candle | NY | $240.00 |
| 10/18/2026 | Market | Candle | MA | $56.00 |
| 11/01/2026 | Etsy | Soap | CA | $18.50 |
| 11/02/2026 | Website | Candle | NJ | $42.00 |
| 11/08/2026 | Market | Candle | MA | $84.00 |
| 11/09/2026 | Etsy | Candle | TX | $28.00 |
| 11/15/2026 | Website | Gift box | NY | $65.00 |
| 11/20/2026 | etsy | Candle | MA | $30.00 |

What different boxes returned in LibreOffice Calc:

| Criteria box | Total (DSUM) | Sales (DCOUNT) |
|---|---|---|
| Channel ="=Etsy" | $104.50 | 4 |
| Channel Etsy* | $344.50 | 5 |
| ="=Etsy" on row 4, ="=Market" on row 5 | $256.50 | 7 |
| Product Candle, State MA | $198.00 | 4 |
| Amount >50 | $445.00 | 4 |
| Candle and MA on row 4, Website on row 5 | $305.00 | 6 |
| ="=Etsy", Amount >=20, Amount <=29 | $56.00 | 2 |
| Channel <>Market | $451.50 | 7 |
| Dates 11/01/2026 to 11/30/2026 | $267.50 | 6 |

As downloaded, with only ="=Etsy" in the box, the Report shows $104.50, 4 sales, a $26.13 average, $30.00 largest, $18.50 smallest, a $0.00 Difference against SUMIFS, and CHECK OK.

## Troubleshooting

### The results show #VALUE!
A box label does not match a Sales header. A label typed "Amt" gave #VALUE! in the total, count, largest and smallest, and row 7 said Check: not a Sales label. Copy the header text exactly, or use a formula such as `=Sales!E4`.

### The total is higher than expected
A wildcard or begins-with match picked up more sales. Etsy* also counts Etsy Wholesale ($344.50 instead of $104.50). Type `="=Etsy"` for one exact channel. For a one-row box, **J11** shows DSUM minus a SUMIFS cross-check; in a test copy with one sale missing its product, <>Website gave $484.50 from DSUM and $400.50 from SUMIFS, because the cross-check leaves out sales with a blank field.

### CHECK says Check and Conditions typed is 0
The box is empty. Type at least one condition in **A4:G6**.

### CHECK says Check after clearing a row
A blank row sits between filled rows. With Etsy on row 4, row 5 blank and Market plus Candle on row 6, LibreOffice ignored the blank row ($244.50), and **J13** showed 1. Move the conditions up so the filled rows touch.

### A sale is missing from the total
Its amount or date is text. DSUM skipped a text amount, and a date criterion of 11/01/2026 or later left out a text date ($249.00). Sales **F** shows Check on those rows; retype them.

## Template

[Download the .xlsx](templates/tidy-tabs-dsum-criteria-box-report.xlsx)

The file has three tabs. **Report** holds the criteria box in **A3:G6** with label checks in row 7, the results and checks in **J3:J15**, and examples of each kind of condition in **A9:A18**. It opens with ="=Etsy" in **C4**; type ="=Market" in **C5** to see the OR total.

**Sales** holds up to 500 sales in rows 5 to 504, with a Status in **F** that flags text dates and amounts and a summary in **H4:I8**. **How to use** explains each formula. Both data tabs print Landscape on US Letter.

Yellow cells are inputs, gray cells are formulas, and every Check turns red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
