---
threads_url: https://www.threads.com/@tin_ylab/post/DeriCkIldar
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-price-custom-job-by-square-foot.html
title: How to Price a Custom Job by Square Foot With a Minimum Charge in Excel and Google Sheets
labels: excel-formulas, small-business, pricing
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the 8 sample jobs were checked against a Python Decimal calculation (round half up), plus rates of $3.35 and $2.75, a quantity of 0 or blank, a width typed as text, a fractional quantity, a rate of 0 or text, and sizes typed in feet. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Price a job by square foot with a minimum charge using =MAX($B$2,ROUND(E6*D6*$B$1,2)) in Excel and Google Sheets. Free quote template.
image_prompts: Roll of blank white vinyl banner material beside a yellow tape measure and a stack of plain cardboard yard sign blanks on a workshop table, soft overhead shop light, no screens, no readable text, no logos
image_alt: Vinyl banner roll next to a yellow tape measure and stacked yard sign blanks
threads: Two 12 x 12 in decals at $2.50 per sq ft come to $5.00. Your minimum is $25.00.\n=MAX($B$2,ROUND(E6*D6*$B$1,2)) charges whichever is larger, and the sheet marks which lines hit the minimum.
---
Multiply width by height in inches, divide by 144 for square feet, then price it and keep the larger of that and your minimum with `=MAX($B$2,ROUND(E6*D6*$B$1,2))`. **B1** is the rate per square foot, **B2** is the minimum charge, **E6** is the area of one piece and **D6** is the quantity.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Quote table of eight sign shop jobs with area, line price and minimum flag](images/2026-10-18-team-05-price-per-square-foot-with-minimum-charge-excel-google-sheets/price-per-square-foot-quote-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The settings sit at the top, with labels in **A** and values in **B**. Below them is one job per row, starting in row 6: job in **A**, width in **B**, height in **C**, quantity in **D**, area in **E**, line price in **F** and a minimum flag in **G**.

1. On a tab named **Quote**, type the labels Rate per sq ft, Minimum charge per line, Setup fee and Settings OK? in **A1:A4**. Type your rate in **B1** (for example $2.50), the minimum in **B2** ($25.00) and the setup fee in **B3** ($15.00).
2. In **B4**, type `=IF(AND(ISNUMBER(B1),ISNUMBER(B2),ISNUMBER(B3)),IF(AND(B1>0,B2>=0,B3>=0),"OK","Check settings"),"Check settings")`. It shows OK only when all three settings are numbers and the rate is above 0.
3. In row 5, type the headers: Job, Width (in), Height (in), Qty, Area per piece (sq ft), Line price, Minimum applied, Size check.
4. From row 6 down, type each job with its width and height in inches and the quantity.
5. In **E6**, type `=IF(OR(H6="",H6="Check"),"",ROUND(B6*C6/144,2))`. The part that does the work is `ROUND(B6*C6/144,2)`. There are 144 square inches in a square foot, so 24 x 36 in is 6.00 sq ft.
6. In **F6**, type `=IF(OR(E6="",$B$4<>"OK"),"",MAX($B$2,ROUND(E6*D6*$B$1,2)))`. `MAX` returns the minimum whenever area x quantity x rate comes to less.
7. In **G6**, type `=IF(F6="","",IF(ROUND(E6*D6*$B$1,2)<$B$2,"Yes","No"))`. It says Yes when the minimum raised the price.
8. In **H6**, type `=IF(A6&B6&C6&D6="","",IF(OR(NOT(ISNUMBER(B6)),NOT(ISNUMBER(C6)),NOT(ISNUMBER(D6))),"Check",IF(OR(B6<=0,C6<=0,D6<=0,D6<>INT(D6)),"Check",IF(B6*C6<36,"Under 36 sq in: feet?","OK"))))`. A Check row gets no price instead of a wrong one.
9. Select **E6:H6** and drag the fill handle down to row 25.
10. Beside the table, add the lines in **K2** with `=SUM(F6:F25)`, count the priced lines in **K6** with `=COUNT(F6:F25)`, and charge the setup fee once in **K3** with `=IF(K6=0,0,B3)`. The total in **K4** is `=K2+K3`.

### Why the dollar signs matter

`E6` and `D6` are relative, so each row prices its own area and quantity. `$B$1` and `$B$2` are absolute, so every row uses the same rate and the same minimum. Without the `$` signs, row 7 would read B2 as the rate and B3 as the minimum.

Press **F4** while typing a reference in Excel to add the `$` signs. Microsoft explains this on its page about [relative, absolute and mixed references](https://support.microsoft.com/en-us/office/switch-between-relative-absolute-and-mixed-references-dfec08cd-ae65-4f56-839e-5f0d8d0baca9).

### In Excel

The formulas work as typed. Microsoft describes the function that sets the floor on its [MAX function](https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098) page.

### In Google Sheets

The same formulas should work as typed, and you can fill down by dragging the small square at the corner of the selected cell. Google describes [MAX](https://support.google.com/docs/answer/3094013) in its help. None of this was opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is a small sign shop with a rate of $2.50 per square foot, a $25.00 minimum per line and a $15.00 setup fee.

| Job | Width (in) | Height (in) | Qty | Area (sq ft) | Line price | Minimum applied |
|---|---|---|---|---|---|---|
| Vinyl banner | 36 | 72 | 1 | 18.00 | $45.00 | No |
| Yard signs | 18 | 24 | 10 | 3.00 | $75.00 | No |
| Window decals | 12 | 12 | 2 | 1.00 | $25.00 | Yes |
| Poster | 24 | 36 | 1 | 6.00 | $25.00 | Yes |
| Canvas prints | 30 | 40 | 2 | 8.33 | $41.65 | No |
| Table banner | 24 | 60 | 1 | 10.00 | $25.00 | No |
| Door signs | 11 | 17 | 3 | 1.30 | $25.00 | Yes |
| Feather flags | 30 | 96 | 2 | 20.00 | $100.00 | No |

The lines add up to $361.65, three lines were raised to the minimum, and the $15.00 setup fee brings the quote to $376.65.

The poster is 6.00 sq ft x $2.50 = $15.00, so it is raised to $25.00. The table banner comes to exactly $25.00, so it shows No, because the test is "less than" the minimum. The canvas area is 8.33 because the area is rounded to 2 decimals before pricing. At other rates the totals still matched Python: $459.11 at $3.35, and $348.05 at $2.75 with no minimum.

## Troubleshooting

### A row has no price and Size check says Check
The quantity is blank, 0, negative or a fraction such as 1.5, or the width or height is blank, 0, negative or typed as text such as "18". Fix the cell and the row prices itself. In the recalculated copy, those rows were left out of the subtotal instead of being priced as $0.00.

### Every line price is blank and B4 says Check settings
The rate in **B1** is 0 or text, or **B2** or **B3** is not a number. Type a number in each. Until then **K2** to **K4** show $0.00 and the check cell in **K9** says CHECK.

### Size check says Under 36 sq in: feet?
A size under 36 square inches is probably a banner typed in feet. A 3 x 6 entry is 0.13 sq ft and gets the $25.00 minimum, which is a warning only: the row is still priced. It misses large sizes typed in feet, so 6 x 8 (48 sq in) is not flagged.

### K9 says CHECK but all the rows look right
**K9** recomputes the subtotal from area x quantity x rate and allows half a cent per priced line. It also shows CHECK whenever **B4** is not OK. Fix the settings first, then look for a Check row.

## Limits

This is arithmetic only: no sales tax, discounts or advice on choosing a rate. Sizes are in inches, the minimum applies per line and not per quote, and the formulas run from row 6 to row 25.

## Template

[Download the .xlsx](templates/tidy-tabs-price-per-square-foot.xlsx)

The file has two tabs. **Quote** holds the rate, minimum and setup fee in yellow cells **B1:B3**, the settings check in **B4**, and 8 sample jobs in rows 6 to 13. Yellow input columns **A:D** and formulas in **E:H** run to row 25, with the quote totals in **J:K**. **How to use** has short notes. The Quote tab prints landscape on US Letter.

The jobs are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
