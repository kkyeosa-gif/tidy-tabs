---
threads_url: https://www.threads.com/@tin_ylab/post/DfDPxTbgUYj
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-convert-decimal-inches-to.html
title: How to Convert Decimal Inches to Fractions (1/16 or 1/8) in Excel and Google Sheets
labels: excel-formulas, small-business, measurements
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Fraction model on the sample and 14 edited copies. Checked: precision 8, 16 and 32, values exactly halfway between steps, 0, 0.03, 12.999, values that round up to the next whole inch, blank rows, negative numbers, numbers typed as text, and a bad precision cell. Excel and Google Sheets output of TEXT with a fraction pattern was not tested.
search_description: Convert decimal inches to fractions in Excel and Google Sheets with MROUND and TEXT. Free template with a 1/8, 1/16 or 1/32 dropdown, feet and inches, and a rounding check.
image_prompts: A folding wooden ruler opened flat on a pine board next to a carpenter pencil and a pair of metal calipers, close view with warm shop lighting, no screens, no readable text, no logos
image_alt: Folding wooden ruler on a pine board with a carpenter pencil and metal calipers
threads: Your cut list says 2.7 inches and the ruler says nothing of the sort.\nMROUND(2.7,1/16) snaps it to 2.6875, and TRIM(TEXT(...,"# ?/16")) writes it as 2 11/16.\nTEXT leaves stray spaces around the fraction, so the TRIM is not optional.
---
Round a decimal inch to the nearest sixteenth with `MROUND` and show it as a fraction with `=TRIM(TEXT(MROUND(A5,1/16),"# ?/16"))`, where A5 is the decimal inch. The `TRIM` removes the extra spaces `TEXT` leaves around the fraction. Use `1/8` and `"# ?/8"` for eighths.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed and link to Microsoft and Google help. Neither Excel nor Google Sheets was opened for this post, so how their `TEXT` writes fractions is not tested.

![Inches sheet listing decimal inches with rounded value, fraction text, reduced fraction and feet and inches columns](images/2026-10-28-team-04-convert-decimal-inches-to-fractions-sixteenths-mround-text-excel-google-sheets/decimal-inches-to-fractions-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The idea: `MROUND(A5,1/16)` snaps the number to the nearest multiple of 1/16, which is 0.0625. `TEXT` with the pattern `# ?/16` then writes the whole inches, a space and the sixteenths. This is for display and cut lists, not tolerances or fit.

### In Excel

1. In **C1**, type the step you want as a number: 8, 16 or 32.
2. In row 4, type the headings: Typed, Rounded, Fraction, Whole inches, Parts, Reduced, Feet and inches, Difference.
3. Type your decimal inches in **A5** and down.
4. In **B5**, type `=MROUND(A5,1/$C$1)`. Format the column as a number with 4 decimals.
5. In **C5**, type `=TRIM(TEXT(B5,"# ?/"&$C$1))`. The `&` builds the pattern from C1, so changing C1 to 8 changes the whole column.
6. Select **B5:C5** and drag the fill handle down the list.

`TEXT` keeps the denominator you give it, so 0.5 comes out as `8/16`, not `1/2`. If you want ruler style, add two helper columns:

- **D5**: `=INT(B5)` for the whole inches.
- **E5**: `=ROUND((B5-D5)*$C$1,0)` for the parts of an inch.
- **F5**: `=IF(E5=0,D5&"",IF(D5=0,"",D5&" ")&E5/GCD(E5,$C$1)&"/"&$C$1/GCD(E5,$C$1))`. `GCD` divides the top and bottom by the same number, so 8/16 becomes 1/2.

Microsoft's pages: [MROUND function](https://support.microsoft.com/en-us/office/mround-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427) and [TEXT function](https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c). These pages describe the functions, not our test.

### In Google Sheets

1. Enter the same cells and formulas as in the Excel steps, typed exactly as written.
2. Drag the fill handle down the list.

Google's help page: [TEXT function](https://support.google.com/docs/answer/3094139). `MROUND` works the same way there, but the fraction text from `TEXT` was not tested in Sheets.

### The version in the template

The template's B5 refuses negative numbers and text, because `MROUND` errors on mixed signs. Its C5 writes `0` when the rounded value is 0, because `TEXT` returns only spaces for 0:

`=IF(ISNUMBER(B5),IF(B5=0,"0",TRIM(TEXT(B5,"# ?/"&$C$1))),"")`

The summary sits in **J5:K9**: numbers rounded, rounded up, rounded down, the largest difference, and a check that says OK or CHECK. **C2** says OK when **C1** is 8, 16 or 32.

## Example

Sample data is fictional: 20 made up cuts for signs, frames and fabric, rounded to 1/16.

| Typed | Rounded | Fraction (TEXT) | Reduced | Feet and inches | Difference |
|---|---|---|---|---|---|
| 0.3125 | 0.3125 | 5/16 | 5/16 | 5/16 in | 0 |
| 2.7 | 2.6875 | 2 11/16 | 2 11/16 | 2 11/16 in | -0.0125 |
| 5 | 5 | 5 | 5 | 5 in | 0 |
| 11.875 | 11.875 | 11 14/16 | 11 7/8 | 11 7/8 in | 0 |
| 0.03 | 0 | 0 | 0 | 0 in | -0.03 |
| 12.999 | 13 | 13 | 13 | 1 ft 1 in | 0.001 |
| 0.09375 | 0.125 | 2/16 | 1/8 | 1/8 in | 0.03125 |
| 72 | 72 | 72 | 72 | 6 ft 0 in | 0 |
| 100.4375 | 100.4375 | 100 7/16 | 100 7/16 | 8 ft 4 7/16 in | 0 |

For all 20 rows, **K5** is 20 rounded, **K6** is 4 rounded up, **K7** is 5 rounded down, **K8** is 0.03125, and **K9** is OK.

Three rows show the rules. 0.03 is under half a step (0.03125), so it rounds down to 0. 0.09375 is exactly halfway, and `MROUND` rounds halfway values up, so it lands on 0.125. 12.999 rounds up to a whole 13, which the feet column shows as 1 ft 1 in.

## Troubleshooting

### The fraction has extra spaces or looks misaligned
`TEXT` pads the result. In the test, `5/16` came back as a leading space plus `5/16`, and a whole number like 5 came back with trailing spaces. Wrap the formula in `TRIM`.

### A small number shows 0
The value is below half a step. At 1/16, 0.03 and anything up to 0.03125 round to 0. Pick 32 in C1 if you need finer cuts.

### The fraction shows 8/16 instead of 1/2
`TEXT` keeps the denominator from the pattern. Use the reduced column (**F**), which divides both parts by `GCD`.

### Rounded value is higher than the typed value
Values exactly halfway between two steps round up, so 0.03125 becomes 0.0625. Values like 11.99 round up to the next whole inch (12), and the feet column then shows 1 ft 0 in.

### The row says Check number or Check precision
Check number means the cell is negative or text, such as "2.5" typed as text. Retype it as a positive number. Check precision means C1 is not 8, 16 or 32, and K9 turns to CHECK.

## Template

[Download the .xlsx](templates/tidy-tabs-inches-to-fractions.xlsx)

The **Inches** tab holds the yellow step dropdown in **C1** and yellow decimal inches in **A5:A104**, with 20 fictional rows filled in. Gray cells are formulas: columns **B** to **H** and the summary in **J5:K9**. **How to use** has short notes.

The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, landscape for the Inches tab. To open it in Google Sheets, go to **File > Import > Upload** and select the file, then compare the Fraction column with the table above.
