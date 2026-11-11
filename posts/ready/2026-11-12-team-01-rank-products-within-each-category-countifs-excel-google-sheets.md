---
title: How to Rank Products Within Each Category in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 21 workbooks and 16,905 comparisons with 0 mismatches. Checked: Top N settings 3, 1, 5, 14 and 200 and the invalid settings 0, -1, 2.5, "abc", "3" typed as text and blank, three-way and two-way ties, Candles / candles / CANDLES, "Candles " and " Candles", a blank category, text sales, blank sales with a name, 0 and negative sales, So*p, Soap? and Soap~, very small and close decimals, a row on row 203, and 8 random lists of 200 rows. Excel and Google Sheets were not used. Not run: decimals with more than 15 significant digits, more rows than the template holds, a non-US language setting.
search_description: Rank each product against only its own category with COUNTIFS, with no sorting. Ties share a rank. Free Excel template with a Top N mark.
image_prompts: A craft-fair shelf with candle jars grouped by scent in three rows and soap bars stacked beside them, one small blank ribbon rosette pinned next to the front jar of each group, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Candle jars in three rows on a craft fair shelf with blank ribbon rosettes
threads: Want the best seller in each category, not just the best seller overall?\nCount the products in the same category that sold more, then add 1:\n=COUNTIFS($B$4:$B$203,B4,$C$4:$C$203,">"&C4)+1\nNo sorting needed. Two soaps tied at $720.00 both get 2, and Soap has no 3.
---
Rank each product against only the products in its own category with `=COUNTIFS($B$4:$B$203,B4,$C$4:$C$203,">"&C4)+1`, where B holds the category and C the sales. It counts the products in the same category that sold more and adds 1, so the list does not need sorting. In the made-up sample, Lavender soap bar ($802.00) is 1 in Soap, and the two soaps tied at $720.00 are both 2.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was built and every number below was recalculated and compared with a Python model (21 workbooks, 16,905 comparisons). The formulas use only IF, AND, OR, NOT, TRIM, ISNUMBER, FIND, SUMPRODUCT, INT, COUNT, COUNTIF and COUNTIFS, which also exist in Excel and Google Sheets, but neither was opened for this post.

![Sample product sales table with a rank per category and Top 3 marks](images/2026-11-12-team-01-rank-products-within-each-category-countifs-excel-google-sheets/rank-within-category-template.png) *Fourteen products in four categories with the rank in each category, the overall rank and the Top 3 mark. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets. The template has a **Sales** tab and a **How to use** tab.

1. Type one product per row: the product name in **A4**, the category in **B4** and the sales in **C4**, and so on down to row 203. Any order works.
2. In **D4**, type `=COUNTIFS($B$4:$B$203,B4,$C$4:$C$203,">"&C4)+1`.
3. Fill **D4** down to the last product.
4. In **B1**, type how many products per category to mark, for example `3`.
5. In **G4**, type `=IF(D4<=$B$1,"Top "&$B$1,"")` and fill it down. The template uses a longer version that also checks B1.

How it works: the first pair, `$B$4:$B$203,B4`, keeps only rows in the same category. The second pair, `$C$4:$C$203,">"&C4`, keeps only rows that sold more than this one. If no product in the category sold more, the count is 0 and the rank is 1.

The dollar signs lock both ranges, so every row looks at the whole list. That is why sorting is not needed.

Ties share a rank and the next number is skipped, the same way RANK handles ties: two products tied for 2 are both 2 and the next one is 4. There is no tie-breaker in the file.

The template wraps the core formula so blank rows stay blank and a row with a blank category, a sales cell that is not a number, or a `*`, `?` or `~` in the category shows Check. It also adds two columns:

- **Products in category** (E4): `=IF(ISNUMBER(D4),COUNTIFS($B$4:$B$203,B4,$D$4:$D$203,">0"),"")`, the number of ranked rows in the same category.
- **Overall rank** (F4): `=IF(ISNUMBER(D4),COUNTIFS($C$4:$C$203,">"&C4,$D$4:$D$203,">0")+1,"")`, the rank against every ranked row, with the same tie rule.

Help pages: Microsoft's [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [RANK function](https://support.microsoft.com/en-us/office/rank-function-6a2fc49d-1831-4a03-9d8c-c279cf99f723), and Google's [COUNTIFS](https://support.google.com/docs/answer/3256550). They describe the functions, not our test.

## Example

Sample data is fictional: 14 made-up products in four categories, typed in no particular order, with **B1** set to 3.

| Product | Category | Sales | Rank in category | Products in category | Overall rank | Top N mark |
|---|---|---|---|---|---|---|
| Lavender soy candle 8 oz | Candles | $1,240.00 | 2 | 4 | 3 | Top 3 |
| Oatmeal goat milk soap | Soap | $720.00 | 2 | 4 | 8 | Top 3 |
| Brass hoop earrings | Jewelry | $1,890.00 | 1 | 3 | 1 | Top 3 |
| Letterpress birthday card | Cards | $310.00 | 3 | 3 | 14 | Top 3 |
| Cedar soy candle 8 oz | Candles | $980.00 | 3 | 4 | 6 | Top 3 |
| Charcoal soap bar | Soap | $720.00 | 2 | 4 | 8 | Top 3 |
| Pearl drop earrings | Jewelry | $1,125.00 | 2 | 3 | 4 | Top 3 |
| Thank you card set | Cards | $492.00 | 2 | 3 | 12 | Top 3 |
| Vanilla bean candle 12 oz | Candles | $1,410.50 | 1 | 4 | 2 | Top 3 |
| Lemon poppy seed soap | Soap | $455.25 | 4 | 4 | 13 | |
| Beaded stack bracelet | Jewelry | $560.00 | 3 | 3 | 11 | Top 3 |
| Holiday card 6 pack | Cards | $1,005.00 | 1 | 3 | 5 | Top 3 |
| Pumpkin spice candle tin | Candles | $615.00 | 4 | 4 | 10 | |
| Lavender soap bar | Soap | $802.00 | 1 | 4 | 7 | Top 3 |

Look at Soap. Lavender soap bar is 1, the two bars at $720.00 are both 2, and Lemon poppy seed soap is 4. There is no 3 in Soap, and the Top 3 mark lands on three soaps: the 1 and both 2s. The same tie shows in Overall rank: both $720.00 soaps are 8 and the next one is 10.

Holiday card 6 pack is 1 in Cards with $1,005.00, while Cedar soy candle 8 oz sells less ($980.00) and is only 3 in Candles. That is the point of ranking within a category.

The summary block in **I3:J7** shows 14 rows ranked, 0 rows showing Check, 4 rows ranked 1, 12 rows marked Top 3, and a CHECK cell that says OK.

## Troubleshooting

### Every product in a category shows 1, or the ranks look too low
Some rows have a trailing space in the category. Matching ignores upper and lower case, so Candles, candles and CANDLES are one group, but "Candles " with a space at the end is its own group. In a test, "Candles " with $2,000.00 got rank 1 with Products in category 1, apart from the other Candles rows. Retype the category without the space.

### A row shows Check
The category is blank, the category contains `*`, `?` or `~` ("Cand*" showed Check), or the sales cell holds text. Sales pasted as the text "$900.00" show Check; the [convert text to numbers post](posts/published/2026-10-07-team-05-convert-text-to-numbers-dollar-signs-excel-google-sheets.md) shows how to turn them into real numbers.

### The short formula counts the wrong rows
COUNTIFS reads `*`, `?` and `~` in a criterion as wildcards. In a test, `=COUNTIFS($B$4:$B$203,B23,$C$4:$C$203,">"&C23)+1` with B23 = "Cand*" and C23 = 10 returned 7, because it counted 6 rows starting with "Cand", including "candles" and "Candles ". Keep those characters out of category names. The template flags them instead of ranking.

### Two products get the same rank and a number is skipped
They have the same sales. In the sample, Soap reads 1, 2, 2, 4. This is the tie rule, not an error. If you need a unique order, decide by hand which tied product goes first.

### Top N mark shows Check on every row
**B1** is not a whole number of 1 or more. In testing, 0, -1, 2.5, "abc", "3" typed as text and a blank B1 all gave Check. Type a plain number such as 3.

## Template

[Download the .xlsx](templates/tidy-tabs-rank-within-category.xlsx)

The file has two tabs. **Sales** has the Top N setting in **B1**, products, categories and sales in **A4:C203**, the formula columns **D:G** (Rank in category, Products in category, Overall rank, Top N mark), and a summary with a CHECK cell in **I3:J7**. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. A green fill marks rank 1 and a red fill marks Check. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
