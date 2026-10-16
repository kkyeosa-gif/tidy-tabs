---
threads_url: https://www.threads.com/@tin_ylab/post/Dejz4h0FhLT
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-find-median-order-value-in-excel.html
title: How to Find the Median Order Value in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python (mean, statistics.median and a filtered mean), including an even count, a blank row, a total typed as text, a single order and an empty list. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Median order value: =MEDIAN(C2:C100). One $640.00 order lifts the average to $83.35 while the median stays $35.25. Free .xlsx template.
image_prompts: Row of small padded mailers on a wooden shelf with one oversized cardboard carton standing beside them and a ball of kraft twine, soft window light, no screens, no readable text, no logos
image_alt: One oversized carton standing beside a row of small padded mailers on a shelf
threads: One $640.00 wholesale order pushed a sample shop's average order to $83.35.\n=MEDIAN(C2:C100) still says $35.25, which is what a typical order looks like.\nThe post shows both formulas side by side, plus a $500.00 cutoff version.\nA total typed as text gets skipped without any warning.
---
Use `=MEDIAN(C2:C100)` for the middle order and `=AVERAGE(C2:C100)` for the mean. In the sample, one $640.00 wholesale order lifts the average to $83.35 while the median stays at $35.25.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Thirteen orders beside a summary of average $83.35, median $35.25 and largest order](images/2026-10-15-team-05-median-order-value-vs-average-excel-google-sheets/median-order-value-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The average adds every order and divides by the count, so one huge order drags it up. The median is the middle order after sorting, so one huge order barely moves it.

1. On a tab named **Orders**, type these headers in row 1: Order, Date, Order total.
2. Type the order number in column **A**, the date as MM/DD/YYYY in column **B**, and the total as a plain number in column **C**.
3. In **F4**, type `=COUNT(C2:C100)`. This counts the orders.
4. In **F5**, type `=IFERROR(AVERAGE(C2:C100),"")`. This is the average order value.
5. In **F6**, type `=IFERROR(MEDIAN(C2:C100),"")`. This is the median order value.
6. In **F7**, type `=IF(F5="","",F5-F6)`. This is how far the big orders pull the average above the median.
7. Put a cutoff such as $500.00 in **F2**. In **F9**, type `=IFERROR(AVERAGEIF(C2:C100,"<="&F2),"")` for the average without the big orders. In **F10**, type `=COUNTIF(C2:C100,">"&F2)` to count how many orders were left out.

With an even number of orders, MEDIAN averages the two middle ones. The ranges stop at row 100, so make them longer if you will log more orders.

### In Excel

The formulas work as typed. Microsoft describes them on its [MEDIAN function](https://support.microsoft.com/en-us/office/median-function-d0916313-4753-414c-8537-ce85bdd967d2) and [AVERAGE function](https://support.microsoft.com/en-us/office/average-function-047bac88-d466-426c-a32b-8f33eb960cf6) pages. Format the result cells as dollars from **Home > Number Format**.

### In Google Sheets

The same formulas should work as typed. Format the results with **Format > Number > Currency**. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small online shop's September orders, every other day.

| Order | Date | Order total |
|---|---|---|
| ORD-4001 | 09/01/2026 | $24.50 |
| ORD-4002 | 09/03/2026 | $38.00 |
| ORD-4003 | 09/05/2026 | $42.75 |
| ORD-4004 | 09/07/2026 | $18.00 |
| ORD-4005 | 09/09/2026 | $55.25 |
| ORD-4006 | 09/11/2026 | $31.50 |
| ORD-4007 | 09/13/2026 | $640.00 |
| ORD-4008 | 09/15/2026 | $27.00 |
| ORD-4009 | 09/17/2026 | $46.50 |
| ORD-4010 | 09/19/2026 | $35.25 |
| ORD-4011 | 09/21/2026 | $29.00 |
| ORD-4012 | 09/23/2026 | $62.00 |
| ORD-4013 | 09/25/2026 | $33.75 |

| Result | Cell | Value |
|---|---|---|
| Orders counted | F4 | 13 |
| Average order value | F5 | $83.35 |
| Median order value | F6 | $35.25 |
| Average minus median | F7 | $48.10 |
| Largest order | F8 | $640.00 |
| Average without orders above $500.00 | F9 | $36.96 |
| Orders above $500.00 | F10 | 1 |

Twelve of the 13 orders are $62.00 or less. The median of $35.25 describes them well. The $83.35 average describes almost no real order.

## Troubleshooting

### The count is lower than my number of orders
A total stored as text is skipped without any warning. In the recalculated copy, the $640.00 typed as text ("640.00") left F4 at 12, F5 at $36.96, F6 at $34.50 and F8 at $62.00. Compare F4 with the number of rows you filled in, and retype any left-aligned total as a plain number.

### The median is a number nobody paid
With an even count, MEDIAN averages the two middle orders. In the recalculated copy, a 14th order of $50.00 gave a median of $36.63, the average of the middle orders $35.25 and $38.00. That is correct, even if no single order matches.

### A blank row changes the result
Blank cells are ignored, not counted as $0.00. Clearing one order dropped F4 to 12 and moved the median to $36.63, the same as the even-count case, because the list now has 12 orders.

### The average and median are identical
With one order of $40.00, F5 and F6 both show $40.00 and F7 shows $0.00. That is expected. The two numbers only separate when the totals are uneven.

### The results are blank
With every total cleared, F4 shows 0, F5 to F9 are blank and F10 shows 0. The IFERROR wrappers keep an empty list from showing an error.

## Limits

The formulas read totals only. They do not remove refunds, sales tax or shipping, so subtract those first if you want product revenue alone. The $500.00 cutoff is your choice, not a rule. Pick a number that fits your own shop.

## Template

[Download the .xlsx](templates/tidy-tabs-median-order-value.xlsx)

The file has two tabs. **Orders** holds the 13 sample orders in yellow input cells, room down to row 100, the cutoff in **F2**, and the seven results in **F4** to **F10**. **How to use** has short fill-in notes. Both tabs print portrait on US Letter.

The orders are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
