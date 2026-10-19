---
title: How to Calculate Return on Ad Spend (ROAS) and Break Even ROAS in Excel and Google Sheets
labels: excel-formulas, small-business, advertising
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the 8 sample campaigns were checked against a Python Decimal model, plus 22 edited variants (exact break-even, spend with no sales, zero spend, text amounts, bad margins, a 41.25% margin, 12 full rows, an empty sheet). The conditional formatting colors were seen only in the LibreOffice PDF export. Excel and Google Sheets were not opened.
search_description: Calculate ROAS with =D5/B5 and break even ROAS with =1/B1 in Excel and Google Sheets. Free ad spend template with profit after ads.
image_prompts: A bundle of folded flyers tied with twine, a tray of wrapped soap bars and a small folding table sign holder on a farmers market table before opening, soft natural morning light, no screens, no readable text, no logos
image_alt: Twine-tied flyers and wrapped soap bars on a folding market table
threads: 2.50 is the ROAS a 40% margin has to beat, and =1/B1 finds it.\nOne sample campaign sits at 1.63 and loses $52.00 after ads, while the overall ROAS of 3.95 looks fine.\nCheck each row, not the total.
---
ROAS is sales divided by ad spend, so `=D5/B5` gives it for each campaign. Break even ROAS is 1 divided by your gross margin: with a 40% margin, `=1/B1` returns 2.50.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every number below was recalculated and checked against Python. The formulas are plain division, SUM and IF, so the same steps should work in Excel and Google Sheets, but neither app was opened for this post. The light red rows were seen only in the LibreOffice PDF export.

![Campaign table showing ROAS, profit after ads and status, with two rows shaded red](images/2026-10-20-team-04-return-on-ad-spend-roas-break-even-excel-google-sheets/roas-calculator-template.png) *PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

Break even ROAS is the sales each $1.00 of ads must bring back to avoid a loss. It depends on your gross margin, so type that first.

1. In **A1**, type Gross margin before ads. In **B1**, type your margin as a percent, such as 40%. Work it out as (price - product cost - fees) / price, using your own numbers.
2. In **A2**, type Break even ROAS. In **B2**, type `=IF(AND(ISNUMBER(B1),B1>0,B1<=1),1/B1,"Check margin")`. The core is `=1/B1`. The `IF` shows a message instead of an error when the margin is blank, zero, text or above 100%.
3. In row 4, type the headers Campaign, Ad spend, Orders, Sales, ROAS, Cost per order, Profit after ads and Status in **A4:H4**.
4. From row 5 down, type one campaign per row: name in **A**, ad spend in **B**, orders in **C**, sales in **D**. Copy these from the ad report.
5. In **E5**, type `=IF(AND(ISNUMBER(B5),ISNUMBER(D5),B5>0),D5/B5,"")`. This is ROAS. The `B5>0` test leaves the cell blank for a campaign with no spend, so you never see #DIV/0!.
6. In **F5**, type `=IF(AND(ISNUMBER(B5),ISNUMBER(C5),C5>0),B5/C5,"")`. This is the cost per order.
7. In **G5**, type `=IF(AND(ISNUMBER(B5),ISNUMBER(D5),ISNUMBER($B$2)),D5*$B$1-B5,"")`. This is profit after ads: sales times margin, minus spend.
8. In **H5**, type `=IF(AND(B5="",C5="",D5=""),"",IF(NOT(ISNUMBER(G5)),"Check",IF(B5=0,"No spend",IF(ROUND(G5,2)>=0,"Above break-even","Below break-even"))))`.
9. Select **E5:H5** and drag the fill handle down to row 16.
10. In **B17**, type `=SUM(B5:B16)` and fill right to **D17**. In **E17**, type `=IF(SUM(B5:B16)>0,SUM(D5:D16)/SUM(B5:B16),"")` for the overall ROAS. In **G17**, type `=SUM(G5:G16)`.

Status reads the profit, not the ROAS, on purpose. A campaign sitting exactly on break even then shows Above break-even instead of Below because of a tiny rounding difference.

### In Excel

The formulas work as typed. Microsoft describes the functions on its [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) and [SUM function](https://support.microsoft.com/en-us/office/sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89) pages. To color the losing rows, use **Home > Conditional Formatting > New Rule** with the formula `=$H5="Below break-even"` applied to **A5:H16**.

### In Google Sheets

The same formulas should work as typed. For the color, use **Format > Conditional formatting**, choose **Custom formula is**, and enter `=$H5="Below break-even"` for the range **A5:H16**. These menu paths were not opened in Google Sheets for this post, and no Google help page was reachable when the links were checked, so none is linked.

## Example

Sample data below is fictional: eight campaigns for a small candle and soap shop, with a 40% gross margin, so break even ROAS is 2.50.

| Campaign | Ad spend | Orders | Sales | ROAS | Cost per order | Profit after ads | Status |
|---|---|---|---|---|---|---|---|
| Etsy Ads, candles | $180.00 | 31 | $1,240.00 | 6.89 | $5.81 | $316.00 | Above break-even |
| Facebook holiday gift sets | $420.00 | 38 | $1,260.00 | 3.00 | $11.05 | $84.00 | Above break-even |
| Google Shopping, soap | $250.00 | 29 | $725.00 | 2.90 | $8.62 | $40.00 | Above break-even |
| Instagram reel boost | $150.00 | 7 | $245.00 | 1.63 | $21.43 | -$52.00 | Below break-even |
| Pinterest pins, wraps | $90.00 | 12 | $450.00 | 5.00 | $7.50 | $90.00 | Above break-even |
| Etsy Ads, mugs | $120.00 | 9 | $300.00 | 2.50 | $13.33 | $0.00 | Above break-even |
| Farmers market flyers | $60.00 | 0 | $0.00 | 0.00 | (blank) | -$60.00 | Below break-even |
| Email sale announcement | $40.00 | 22 | $960.00 | 24.00 | $1.82 | $344.00 | Above break-even |
| **Total** | $1,310.00 | 148 | $5,180.00 | 3.95 | $8.85 | $762.00 | |

The Instagram row has a ROAS of 1.63, under the 2.50 line, so it loses $52.00 after ads. The mugs row sits exactly on the line: $300.00 of sales at 40% is $120.00, the same as the spend, so profit is $0.00.

The overall ROAS of 3.95 hides the losing row, so check each campaign.

**B19** shows OK when total profit equals total sales times margin minus total spend ($5,180.00 x 40% - $1,310.00 = $762.00). It says CHECK when any row needs checking.

## Troubleshooting

### B2 and every Status say Check margin or Check
The margin in **B1** is blank, zero, above 100%, or typed as text. A margin typed as 40 instead of 40% is read as 4,000%, so B2 shows Check margin. Retype it as 40% or 0.40.

### A campaign row says Check
Ad spend or sales is blank or typed as text, such as "$180" or "n/a". ROAS and profit stay blank for that row, and **B19** says CHECK until you retype the amount as a number.

### ROAS is blank and Status says No spend
The ad spend is 0, so there is nothing to divide by. The guard blanks ROAS instead of showing #DIV/0!. Cost per order goes blank the same way when Orders is 0, and Orders never changes the Status.

### My ROAS does not match the ad platform
The sheet divides the numbers you typed. Ad platforms credit sales to ads in their own ways, and the sheet ignores returns, shipping and taxes. It says nothing about what a good ROAS is.

## Template

[Download the .xlsx](templates/tidy-tabs-roas-calculator.xlsx)

The file has two tabs. **Campaigns** holds the margin in yellow cell **B1**, the break even ROAS in **B2**, eight sample campaigns in yellow input cells **A5:D12** with room to row 16, the formula columns **E:H**, the totals in row 17 and the check in **B19**. Rows below break even turn light red. **How to use** has short notes. Campaigns prints landscape on US Letter.

The campaigns, spend and sales are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
