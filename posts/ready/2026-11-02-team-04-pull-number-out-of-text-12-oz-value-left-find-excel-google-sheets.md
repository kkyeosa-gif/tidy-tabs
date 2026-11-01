---
title: How to Pull the Number Out of Text Like 12 oz or 3 lb in Excel and Google Sheets
labels: excel-formulas, text-functions, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python mirror of the formulas over 8 workbooks and 512 comparisons, with 0 mismatches. Checked: spacing, capitalization, trailing periods, decimals, unit-first text, fl oz, fractions, hyphens, two units in one cell and blank weights. The PDF export pages are US Letter. Excel and Google Sheets were not used.
search_description: Pull the number out of text like 12 oz or 3 lb with VALUE, LEFT and FIND in Excel and Google Sheets. Free template with a Check flag and ounces total.
image_prompts: A small kitchen scale with a lump of soy wax on it and a metal candle tin beside it on a wooden workbench, soft window light, no screens, no readable text, no logos
image_alt: Kitchen scale holding a lump of soy wax beside an empty metal candle tin
threads: Cells that say 12 oz won't add up because they're text.\nFor 12 oz, take everything before the first space with LEFT and FIND, then wrap it in VALUE. For 12oz, strip the unit with SUBSTITUTE instead.\nThe template flags 12 fl oz and 1/2 lb as Check instead of guessing.
---
For text like `12 oz` with a space, take everything before the space with `=VALUE(LEFT(TRIM(B5),FIND(" ",TRIM(B5)&" ")-1))`. For text like `12oz` with no space, strip the unit instead: `=VALUE(SUBSTITUTE(LOWER(TRIM(B5)),"oz",""))`. Either way you get a real number you can add up.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every result below was recalculated and compared with a Python mirror of the formulas. VALUE, LEFT, FIND, TRIM, SUBSTITUTE and IFERROR exist in current Excel and Google Sheets, but neither app was opened for this post.

![Typed weights like 12 oz and 3 lb beside number, unit and ounces columns](images/2026-11-02-team-04-pull-number-out-of-text-12-oz-value-left-find-excel-google-sheets/weights-text-to-number-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template: **Item** in column A and the **typed weight** in column B, starting in row 5. Format column B as Plain text first so a value like `1.5 lb` stays exactly as typed.

1. In **A5:A24**, type the item names. In **B5:B24**, type each weight with its unit, such as `12 oz` or `3 lb`.
2. In **C5**, use the space form to read the number: `=VALUE(LEFT(TRIM(B5),FIND(" ",TRIM(B5)&" ")-1))`.
3. For weights with no space, such as `12oz`, use the unit-stripping form: `=VALUE(SUBSTITUTE(LOWER(TRIM(B5)),"oz",""))`.
4. To handle both shapes in one column, nest them as the template does (see the full formula below).
5. In **D5**, find the unit: `=IF(OR(A5="",NOT(ISNUMBER(C5))),"",IF(ISNUMBER(SEARCH("lb",B5)),"lb","oz"))`.
6. In **E5**, convert everything to ounces: `=IF(ISNUMBER(C5),IF(D5="lb",ROUND(C5*16,2),C5),"")`.
7. Fill **C5:E5** down to row 24, then add them up with `=SUM(E5:E24)`.

How the space form works: `TRIM` removes extra spaces, `FIND` locates the first space, `LEFT` keeps what comes before it, and `VALUE` turns that text into a number. The added `&" "` guarantees `FIND` always has a space to find.

Each form fails on the other shape. Run alone, the space form gives 12 for `12 oz` and #VALUE! for `12oz`. The unit-stripping form gives 12 for `12oz` and #VALUE! for ` 3 lb `.

That is why the template wraps both in `IFERROR`. Here is the full formula in **C5**:

`=IF(A5="","",IF(AND(ISNUMBER(SEARCH("oz",B5))<>ISNUMBER(SEARCH("lb",B5)),ISERROR(SEARCH("fl",B5)),ISERROR(SEARCH("/",B5)),ISERROR(SEARCH("-",B5)),ISERROR(SEARCH(":",B5))),IFERROR(VALUE(LEFT(TRIM(B5),FIND(" ",TRIM(B5)&" ")-1)),IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(LOWER(TRIM(B5)),"oz",""),"lb","")),"Check")),"Check"))`

The first test requires exactly one of oz or lb, and none of fl, a slash, a hyphen or a colon. Without that guard, `1/2 lb` came back as the date number 46024 and `12-oz` as -12 in LibreOffice. Anything the guard rejects shows Check instead of a wrong number.

Help pages: Microsoft's [VALUE function](https://support.microsoft.com/en-us/office/value-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2) and [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. These links were not link-checked.

## Example

Sample data is fictional: made up supplies for a small candle maker.

| Item | Typed weight | Number | Unit | Ounces |
|---|---|---|---|---|
| Soy wax | 10 lb | 10 | lb | 160 |
| Fragrance oil | 12 oz | 12 | oz | 12 |
| Cotton wicks | 3oz | 3 | oz | 3 |
| Tins, 8 oz | 16 OZ | 16 | oz | 16 |
| Shipping box | ` 3 lb ` | 3 | lb | 48 |
| Dye chips | 1.5 lb | 1.5 | lb | 24 |
| Mica powder | .5 lb | 0.5 | lb | 8 |
| Lip balm tubes | 12 oz. | 12 | oz | 12 |
| Jar lids | 4  oz | 4 | oz | 4 |
| Gift tags | 2lb | 2 | lb | 32 |
| Ribbon | 6 oz | 6 | oz | 6 |
| Packing paper | 1 lb | 1 | lb | 16 |

The ounces add up to 341. The template shows that as `21 lb 5 oz` with `=INT(ROUND(H4,2)/16)&" lb "&ROUND(MOD(ROUND(H4,2),16),2)&" oz"`, where **H4** holds the total. The rounding comes first so 15.999 oz cannot show as 16 oz.

## Troubleshooting

### A row says Check
The text is something the formula refuses to guess at. In the tests these all showed Check: `n/a`, a blank weight, `oz` alone, `12 fl oz`, `1 lb 4 oz`, `12 ounces`, `5 kg`, `twelve oz`, `-3 lb`, `1/2 lb` and `12-oz`. Retype the weight as a number and one unit, such as `12 oz`.

### A plain number shows Check
A cell with just `12` or `12.5` has no unit, so the formula cannot tell ounces from pounds. Add the unit.

### The total is off after a unit like lbs
A unit word that contains oz or lb is read as that unit. `12 lbs` came out as 12 pounds. Units other than oz and lb are not handled: `kg` shows Check.

### Commas in numbers give odd results
These depend on the computer's region setting. In LibreOffice with US settings, `1,5 lb` showed Check, `1,200 oz` was read as 1200 and `1e3 oz` as 1000. Excel and Google Sheets were not tested, so do not assume they match. Type decimals with a point.

### The check cell says Check but no row does
An item with no weight, or a weight with no item name, makes the check cell **H7** show Check. The row is ignored by **C:E**, so fill in the missing half.

## Template

[Download the .xlsx](templates/tidy-tabs-number-from-text-units.xlsx)

The file has two tabs. **Weights** has 12 fictional items in yellow cells **A5:B24** (20 slots), gray formula columns for number, unit and ounces, and a summary block: total ounces in **H4**, pounds and ounces in **H5**, the count of Check rows in **H6**, and the **CHECK** cell, **H7**. **How to use** has notes on each formula.

**H7** says OK when no row shows Check and every item has a weight. Fix every Check before you trust the total. The sample shows OK. Limits: 20 rows, and only oz and lb. **Weights** prints on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
