---
threads_url: https://www.threads.com/@tin_ylab/post/DeeqSySFUVi
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-split-product-code-into-parts.html
title: How to Split a Product Code Into Parts With LEFT, MID and FIND in Excel and Google Sheets
labels: excel-formulas, handmade-business, small-business
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the results for all 10 sample codes were checked against Python str.split, including one-dash, no-dash, extra-dash, leading-space and lowercase cases. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Split CND-LAV-8OZ with =LEFT(A2,FIND("-",A2)-1) for the first part, MID plus a second FIND for the middle. Free .xlsx with sample SKUs.
image_prompts: Craft table with paper product tags, a roll of small price stickers and a few wrapped bars of soap, soft window light, no screens, no readable text, no logos
image_alt: Paper product tags and a sticker roll beside wrapped soap bars on a craft table
threads: One column of codes like CND-LAV-8OZ, and you need the category on its own?\n=LEFT(A2,FIND("-",A2)-1) gives CND. A second FIND inside MID pulls LAV.\nThe sample sheet adds up units by category: 51 candles, 60 soaps, 27 earrings, 68 cards.
---
Take everything before the first dash with `=LEFT(A2,FIND("-",A2)-1)`, so CND-LAV-8OZ gives CND. MID with a second FIND returns the middle part, and the rest of the code comes after the second dash.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Product codes split into category, style and size columns with units sold totaled by category code](images/2026-10-13-team-05-split-sku-code-into-parts-left-mid-find-excel-google-sheets/split-sku-code-into-parts-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The idea: FIND tells you where a dash sits. LEFT and MID then cut the text at those positions.

This post assumes every code has exactly two dashes and three parts, such as category, style and size (CND-LAV-8OZ).

1. In row 1, type these headers in A to E: SKU, Units sold, Category, Style, Size or color.
2. In **A2:A11**, type or paste one code per row. Put the units sold for each code in **B2:B11**.
3. In **C2**, type `=IF(A2="","",IFERROR(LEFT(A2,FIND("-",A2)-1),"Check SKU"))`. FIND("-",A2) is the position of the first dash, and LEFT keeps the characters before it.
4. In **D2**, type `=IF(A2="","",IFERROR(MID(A2,FIND("-",A2)+1,FIND("-",A2,FIND("-",A2)+1)-FIND("-",A2)-1),"Check SKU"))`. This starts one character after the first dash and stops before the second.
5. In **E2**, type `=IF(A2="","",IFERROR(MID(A2,FIND("-",A2,FIND("-",A2)+1)+1,LEN(A2)),"Check SKU"))`. This starts after the second dash and takes the rest.
6. Select **C2:E2** and drag the fill handle down to row 30 (or as far as your list goes).

The third argument of FIND, `FIND("-",A2)+1`, tells it to start looking just past the first dash. That is how the second FIND lands on the second dash instead of the first.

IFERROR replaces the #VALUE! you would get from a bad code with the words Check SKU, so one odd row does not look like a broken sheet.

### Add up units by category

7. In **G2:G5**, type your category codes: CND, SOP, EAR, CRD.
8. In **H2**, type `=SUMIF($C$2:$C$30,G2,$B$2:$B$30)` and fill it down to **H5**.
9. In **H6**, type `=SUM(H2:H5)`. In **H7**, type `=SUM(B2:B30)`. In **H8**, type `=IF(H6=H7,"OK","CHECK")`.

H8 says CHECK when some units sold under a code that is not in your list.

### In Excel

The formulas work as typed. Microsoft documents both functions: [LEFT](https://support.microsoft.com/en-us/office/left-leftb-functions-9203d2d2-7960-479b-84c6-1ea52b99640c) and [FIND](https://support.microsoft.com/en-us/office/find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628).

### In Google Sheets

The same formulas should work as typed, since LEFT, MID, FIND, IFERROR and SUMIF are all Sheets functions. This was not opened in Google Sheets for this post.

## Example

Sample data below is fictional. It is the SKU list of a small shop that sells candles, soap, earrings and cards.

| SKU | Units sold | Category | Style | Size or color |
|---|---|---|---|---|
| CND-LAV-8OZ | 24 | CND | LAV | 8OZ |
| CND-LAV-16OZ | 9 | CND | LAV | 16OZ |
| CND-CED-8OZ | 18 | CND | CED | 8OZ |
| SOP-OAT-4OZ | 31 | SOP | OAT | 4OZ |
| SOP-CHR-4OZ | 22 | SOP | CHR | 4OZ |
| SOP-OAT-8OZ | 7 | SOP | OAT | 8OZ |
| EAR-HOOP-GLD | 12 | EAR | HOOP | GLD |
| EAR-STUD-SLV | 15 | EAR | STUD | SLV |
| CRD-BDAY-A6 | 40 | CRD | BDAY | A6 |
| CRD-THNX-A6 | 28 | CRD | THNX | A6 |

The category totals from column H:

| Category | Cell | Units sold |
|---|---|---|
| CND | H2 | 51 |
| SOP | H3 | 60 |
| EAR | H4 | 27 |
| CRD | H5 | 68 |
| **Total** | H6 | **206** |

H7 also shows 206, so H8 reads OK. All 10 split codes matched a Python `str.split("-")` check.

## Troubleshooting

### Check SKU appears in the Style and Size columns
The code has only one dash, such as CND-LAV. There is no second dash for FIND to land on. In the recalculated file, C still returned CND while D and E showed Check SKU. Add the missing part to the code.

### Check SKU appears in all three columns
The code has no dash at all, such as NOHYPHEN. FIND cannot locate the dash, so every formula falls back to the message. Check for a different dash character, which can come in when you paste from another app.

### The last column holds more than one part
The code has a third dash. In the recalculated file, CND-LAV-8OZ-GIFT put 8OZ-GIFT in column E, because E takes everything after the second dash. Split that column again with a copy of the same pattern if you need the extra part.

### H8 says CHECK
Some units are not landing in a listed category. A space before the code is the usual cause: in the recalculated file, " SOP-OAT-4OZ" gave " SOP" in column C, which missed the SOP total. Clean pasted codes with `=TRIM(A2)` first. A new category also needs its own row in **G2:G5**.

### Lowercase codes are still counted
This is expected. SUMIF ignores capitals, so in the recalculated file "cnd-ced-8oz" returned cnd in column C and still counted under CND. Wrap the code in UPPER if you want the split parts in capitals too.

## Template

[Download the .xlsx](templates/tidy-tabs-split-sku-code.xlsx)

The file has two tabs. **SKUs** holds the 10 sample codes and units in A2:B11, the split formulas in C2:E30, the yellow category codes in G2:G5, and the totals and check in H2:H8. Each tab is set to print on one US Letter page wide. **How to use** has short fill-in notes.

The codes and quantities are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
