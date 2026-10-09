---
title: How to Make a Bar Chart Inside Cells With REPT in Excel and Google Sheets
labels: excel-formulas, charts, small-business
tested_in: LibreOffice Calc 24.2.7 (Linux) only. Excel and Google Sheets were not opened; the REPT and LEN help links are from Microsoft Support.
search_description: Draw a bar chart inside cells with =REPT("█",ROUND(B2/MAX($B$2:$B$7)*20,0)). Includes a | fallback for fonts without the block. Free .xlsx template.
image_prompts: Candle workshop table with a tall stack, a medium stack and a short stack of boxed candles in descending heights next to a wooden ruler, soft window light, no screens, no readable text, no logos
image_alt: Three stacks of boxed candles in descending heights beside a wooden ruler
threads: Want a bar chart without inserting a chart?\n=REPT("█",ROUND(B2/MAX($B$2:$B$7)*20,0)) draws it in the cell. The biggest value gets 20 blocks.\nIf your font shows an empty box, swap the block for "|".
---
Type `=REPT("█",ROUND(B2/MAX($B$2:$B$7)*20,0))` next to each number to draw a bar made of block characters. The biggest value gets 20 blocks and every other row is scaled to match.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built, recalculated and exported to PDF. The Excel and Google Sheets steps use the same REPT, ROUND, MAX and LEN functions. Neither app was opened for this post, so how the block character looks there is unchecked.

![Product list with units sold and a row of solid block bars beside each number](images/2026-10-10-team-05-in-cell-bar-chart-rept-excel-google-sheets/in-cell-bars-template.png) *PDF export from LibreOffice Calc 24.2*

## Steps

You need a list of labels in column A and a number for each in column B. Rows 2 to 7 hold six products here.

1. In **A1:B1**, type the headers Product and Units sold. Type the six products in A2:A7 and their units in B2:B7.
2. In **C2**, type `=REPT("█",ROUND(B2/MAX($B$2:$B$7)*20,0))`. Copy the block character from this line if you cannot type it.
3. Select **C2** and drag the fill handle down to row 7.
4. In **D2**, type `=LEN(C2)`. This counts the blocks in the bar, as a check. Fill it down to row 7.
5. Widen column C until the longest bar fits on one line.

REPT repeats a piece of text a set number of times. The part inside ROUND is the count. `B2/MAX($B$2:$B$7)` turns the value into a share of the biggest value, and `*20` turns that share into a bar length. The `$` signs keep the MAX range fixed as you fill down.

To make the longest bar shorter or longer, change the 20. A bar is rounded to whole characters, so a very small value can show 0 blocks.

### In Excel

The formulas work as typed. Microsoft documents the arguments on its [REPT function](https://support.microsoft.com/en-us/office/rept-function-04c4d778-e712-43b4-9c15-d656582bb061) page and its [LEN function](https://support.microsoft.com/en-us/office/len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb) page.

### In Google Sheets

The same formulas should work as typed, since Sheets has REPT, ROUND, MAX and LEN. This was not opened in Google Sheets for this post.

### If the block does not show

The block character depends on the cell's font. In the LibreOffice check, the default font substitute had no block glyph, so the app drew it from another font. If you see an empty box or an odd look, use the `|` character instead:

`=REPT("|",ROUND(B2/MAX($B$2:$B$7)*20,0))`

The `|` bars have the same lengths as the block bars. They look thinner, but they appear in any font.

## Example

Sample data below is fictional. It is six products from a small candle shop.

| Product | Units sold | Blocks (LEN) |
|---|---|---|
| Soy candle, 8 oz | 3,200 | 20 |
| Beeswax melts | 1,600 | 10 |
| Gift tags (retired) | 0 | 0 |
| Wick trimmer | 450 | 3 |
| Lip balm | 2,400 | 15 |
| Mini jar candle | 800 | 5 |

Check two rows by hand. Soy candle is the biggest value, so it gets all 20 blocks. Wick trimmer is 450 divided by 3,200, which is 0.140625, times 20, which is 2.8125. ROUND makes that 3.

In the recalculated file, the `=LEN` check matched hand math on all six rows. The sample avoids values that land exactly on a half block, where rounding could differ between apps.

The retired product has 0 units, so its bar is empty. That is correct. REPT with a count of 0 returns nothing.

## Troubleshooting

### The bars are all the same length
The MAX range is probably not locked. Make sure it reads `$B$2:$B$7` with `$` signs. Without them, the range slides down as you fill and each row compares against a different maximum.

### The cell shows #VALUE!
REPT returns #VALUE! when the repeat count is negative. A negative number in column B causes this. Fix the data, or wrap the count in MAX, as in `ROUND(MAX(0,B2)/MAX($B$2:$B$7)*20,0)`. This case was not reproduced in a test here.

### The bar is cut off or shows `###`
The column is too narrow, or the bar is wider than the cell. Widen column C. Text bars stay on one line, so a 20 block bar needs a column wide enough for 20 characters.

### The block looks like an empty box or has gaps
That is a font issue. In the LibreOffice PDF export, the blocks showed faint vertical seams between characters. Switch the column to another font, or use the `|` version in column E of the template.

### A new, larger value shrinks every other bar
That is by design. Every bar is scaled to the current maximum. If you add a row below row 7, extend the range in every formula to include it.

## Template

[Download the .xlsx](templates/tidy-tabs-in-cell-bars.xlsx)

The file has two tabs. **Units sold** holds six sample products with units, the block bar, the `LEN` check and the same bar built with `|`. **How to use** has short notes on changing the bar length and swapping the character.

The file was recalculated only in LibreOffice Calc 24.2.7, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
