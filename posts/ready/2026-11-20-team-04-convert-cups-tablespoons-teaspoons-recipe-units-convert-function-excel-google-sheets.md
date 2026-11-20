---
title: How to Convert Cups, Tablespoons and Teaspoons in Excel and Google Sheets
labels: excel-formulas, small-business, food-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 30 workbooks and 8,550 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Convert cups, tablespoons and teaspoons with =CONVERT(amount,"cup","tbs"). The exact codes (tbs, not tbsp), why oz means fluid ounces, and a free template.
image_prompts: A home kitchen counter set up for small-batch production with a glass measuring cup half full of golden honey, a ring of stainless steel measuring spoons fanned out beside it, and a small amber dropper bottle of fragrance oil next to a wooden spoon, bright daylight, realistic photo, no screens, no readable text, no logos
image_alt: Glass measuring cup of honey beside measuring spoons and a small amber bottle
threads: Type "tbs" in CONVERT, not "tbsp".\n=CONVERT(1,"cup","tbs") gives 16. "tbsp", or "Cup" with a capital C, returned #VALUE! in LibreOffice.\nAnd "oz" means fluid ounces, so it will not turn cups of flour into weight.
---
Use `=CONVERT(amount,"cup","tbs")` with the exact lowercase codes "tsp", "tbs", "oz", "cup", "pt", "qt" and "gal", so 1 cup returns 16 tablespoons and 24 teaspoons returns 0.5 cup. Type "tbs", not "tbsp", and remember that "oz" here means fluid ounces (1 cup is 8 oz). The download lets you pick friendly names like Tablespoon from a list and shows "Check units" when a code does not match.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (30 workbooks, 8,550 cell comparisons, 0 mismatches), plus the edge cases below. CONVERT also exists in Excel and Google Sheets, and the unit codes below come from their help pages, but neither app was opened for this post.

![Six ingredients converted between cups, teaspoons, tablespoons, quarts, gallons and fluid ounces](images/2026-11-20-team-04-convert-cups-tablespoons-teaspoons-recipe-units-convert-function-excel-google-sheets/recipe-cups-tablespoons-teaspoons-converter-template.png) *Six sample conversions: 2.5 cups of honey is 20 fluid ounces, 3 teaspoons of vanilla is 1 tablespoon, 24 teaspoons of fragrance oil is 0.5 cup and 1 gallon of water is 16 cups (the two unit-code helper columns, the Status column and the check summary are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Converter** tab, one amount per row starting in row 5.

1. Type the ingredient in **A5**, such as Honey.
2. Type the amount as a plain number in **B5**: 0.5, not 1/2. A typed fraction can turn into text or a date.
3. Pick the unit you have in **C5** (From unit) and the unit you want in **D5** (To unit) from the dropdowns: Teaspoon, Tablespoon, Fluid ounce, Cup, Pint, Quart or Gallon.
4. Read the answer in **G5**. Behind it, **E5** and **F5** look up the CONVERT code for each name on the **Units** tab, and **G5** runs `=CONVERT(B5,E5,F5)`.
5. Check **H5**. It shows OK, or a Check message that says which input to fix. **G5** stays blank on a Check row.
6. To skip the template, type the codes straight into the formula: `=CONVERT(1,"cup","tbs")` gives 16 and `=CONVERT(24,"tsp","cup")` gives 0.5.

The codes must be exact. Microsoft's [CONVERT function](https://support.microsoft.com/en-us/office/convert-function-d785bef1-808e-4aac-bdcd-666c810f9af2) page lists Teaspoon "tsp", Tablespoon "tbs", Fluid ounce "oz", Cup "cup", U.S. pint "pt", Quart "qt" and Gallon "gal", and says unit names are case-sensitive. Google's [CONVERT](https://support.google.com/docs/answer/6055540) help page lists the same tsp, tbs, oz, cup, pt, qt and gal under Volume. Both pages also list look-alike codes such as "tspm" (modern teaspoon) and "uk_pt" (U.K. pint). Those are different sizes, so stick to the seven above for US recipes.

CONVERT handles volume only. Cups of flour or sugar to ounces of weight needs the ingredient's density, which CONVERT does not know, and this sheet does not try. The "oz" code is a fluid ounce. The weight ounce is "ozm", and CONVERT will not turn cups into it.

If you are converting because you scaled a batch, the Tidy Tabs post on [scaling a recipe up or down](https://tidytabs.blogspot.com/2026/10/how-to-scale-recipe-up-or-down-in-excel.html) covers the multiplying part.

## Example

The batch below is fictional.

| Ingredient | Amount | From unit | To unit | Codes | Result |
|---|---|---|---|---|---|
| Honey | 2.5 | Cup | Fluid ounce | cup to oz | 20 |
| Vanilla extract | 3 | Teaspoon | Tablespoon | tsp to tbs | 1 |
| Water | 1 | Gallon | Cup | gal to cup | 16 |
| Fragrance oil | 24 | Teaspoon | Cup | tsp to cup | 0.5 |
| Milk | 1 | Quart | Cup | qt to cup | 4 |
| Lemon juice | 0.0625 | Cup | Teaspoon | cup to tsp | 3 |

Every row shows OK. The summary in **J4:K9** shows 6 rows listed, 0 rows showing Check, 0 results showing Check units, 0 Units tab codes showing Check, 0 typed below row 54, and CHECK shows OK.

The **Units** tab shows how many teaspoons are in one of each unit: 1, 3, 6, 48, 96, 192 and 768, from Teaspoon down to Gallon. Its grid in **E4:L11** gives every pair. One cup is 48 teaspoons, 16 tablespoons, 8 fluid ounces, 0.5 pint, 0.25 quart and 0.0625 gallon. One gallon is 768 teaspoons, 256 tablespoons, 128 fluid ounces and 16 cups.

## Troubleshooting

### CONVERT returns #VALUE!
A code is off. In LibreOffice, "tbsp", "Cup" with a capital C, and " cup" with a leading space each returned #VALUE!. Use the exact lowercase code and wrap a typed code in `TRIM`: once trimmed, " cup" gave 16 tablespoons again. In the template, a bad code on the Units tab shows Check code in column **C**, and any Converter row using it shows Check units.

### Status says Check: from unit or Check: to unit
The unit name is not on the Units tab, or the cell is blank. "Tbsp" is not on the list, so it showed Check: from unit. Add your own name in **Units** rows 12 to 20 with its code: with Tbsp and tbs added, 2 Tbsp converted to 6 teaspoons and CHECK stayed OK. Names on the list ignore case and extra spaces: "cup", "Cup " and "teaspoon" all worked.

### Status says Check: amount
The amount is not a plain number. Text such as "2 cups" or "1/2", a blank amount and a negative number all showed Check: amount. Retype it as a number, such as 0.5.

### Cups of flour will not convert to ounces
CONVERT does volume or weight, not one into the other. In LibreOffice, `CONVERT(1,"cup","ozm")` returned #VALUE!, and in the template it shows Check units. Weight needs a density for each ingredient. Even weight-to-weight results are not always round: 1 lbm came back as 16.0000023429409 ozm in LibreOffice.

## Template

[Download the .xlsx](templates/tidy-tabs-recipe-unit-converter.xlsx)

The file has three tabs. **Converter** has rows 5 to 54 with Ingredient, Amount, From unit and To unit (inputs in **A:D**, the unit columns as dropdowns), the From code and To code helpers in **E:F**, Result in **G** and Status in **H**: OK, Check: amount, Check: from unit or Check: to unit. The summary in **J4:K9** counts rows listed, rows showing Check, results showing Check units, Units tab codes showing Check, entries typed below row 54, and the overall CHECK.

**Units** holds the friendly names and CONVERT codes in **A5:B20** (7 filled, rows 12 to 20 free for your own), Teaspoons in one in **C**, and the grid of every pair in **E4:L11**. The dropdowns warn rather than block, so a name you type is still checked. **How to use** walks through each formula. Both data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Check units, Check code and every Check status turn red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
