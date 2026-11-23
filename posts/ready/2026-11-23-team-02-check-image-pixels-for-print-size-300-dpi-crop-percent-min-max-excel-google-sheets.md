---
title: How to Check If an Image Has Enough Pixels for a Print Size (300 DPI) in Excel and Google Sheets
labels: excel-formulas, small-business, etsy-sellers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 19 workbooks and 132,772 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Sort image pixels and print sizes long side first with MAX and MIN, then check each size's DPI and crop percent against your 300 DPI minimum.
image_prompts: A stack of matted art prints with blurred floral images on a green cutting mat in a small home print studio, next to a metal ruler and an empty 16x20 wooden frame leaning against the table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Matted art prints on a cutting mat beside a ruler and empty frame
threads: 3600 x 5400 px prints 12 x 18 in at 300 DPI, but 16x20 drops to 225 DPI.\nPut both sides long side first with MAX and MIN, then =MIN(long_px/long_in,short_px/short_in) is the DPI.\nThe same file loses 17% of the image to fill 8x10.
---
Turn both the image and the print size long side first with `=MAX(width,height)` and `=MIN(width,height)`, then the DPI you get when the print is filled edge to edge is `=MIN(long_px/long_in,short_px/short_in)`, and a size works if that is at least your minimum, such as 300. In the made-up 3600 x 5400 px file, 8x10 shows OK (crops 17%), 16x20 shows Too small: 225 DPI, and the largest print at 300 DPI is 12 x 18 in.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (19 workbooks, 132,772 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened for this post.

![Six image files with pixel sizes and OK or Too small for five prints](images/2026-11-23-team-02-check-image-pixels-for-print-size-300-dpi-crop-percent-min-max-excel-google-sheets/print-size-dpi-checker-template.png) *Images tab at a 300 DPI minimum: wildflowers-print.jpg (3,600 x 5,400 px) prints up to 12 x 18 in, shows OK (crops 17%) at 8x10 and Too small: 225 DPI at 16x20, while small-scan.png (1,200 x 1,500 px) fits no listed size. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

DPI here means pixels per printed inch. The 300 in the template is a sample minimum, not a rule: type the number your print lab's guide gives. The DPI number saved inside a file does not change its pixel count.

The steps are the same in Excel and Google Sheets.

1. On the **Print sizes** tab, type your minimum DPI in **B1**.
2. Type one print size per row from row 5: the name in **A** (such as 8x10) and the two sides in inches in **B** and **C**. Either side can come first, because **D5** `=IF(G5<>"OK","",MIN(B5,C5))` and **E5** `=IF(G5<>"OK","",MAX(B5,C5))` sort them, and **F5** gives the area.
3. On the **Images** tab, type one file per row from row 5: the file name in **A** and the width and height in pixels in **B** and **C**.
4. In **D5** and **E5**, sort the pixels the same way: `=IF($X5<>"OK","",MAX(B5,C5))` and `=IF($X5<>"OK","",MIN(B5,C5))`. In LibreOffice Calc, 3600 x 5400 and 5400 x 3600 gave identical rows. See the [MAX help](https://support.microsoft.com/en-us/office/max-function-e0012414-9ac8-4b34-9a47-73e662c08098) and MIN help from [Microsoft](https://support.microsoft.com/en-us/office/min-function-61635d12-920f-4ce2-a70f-96f202dcc152) and [Google](https://support.google.com/docs/answer/3094017).
5. **F5** gives the largest print at your minimum: `=IF(OR($X5<>"OK",'Print sizes'!$J$4<>"OK"),"",ROUNDDOWN(E5/'Print sizes'!$B$1,1)&" x "&ROUNDDOWN(D5/'Print sizes'!$B$1,1)&" in")`. For 3600 x 5400 px at 300 DPI that is 12 x 18 in.
6. The effective DPI per size sits in the grouped, hidden helper columns **Z:AO**; click the **+** above them to expand. **Z5** holds `=IF(OR($X5<>"OK",'Print sizes'!$J$4<>"OK",'Print sizes'!$G$5<>"OK"),"",MIN($D5/'Print sizes'!$E$5,$E5/'Print sizes'!$D$5))` for the size on row 5; each column to the right uses the next size row. There you see the 450 DPI behind the wildflowers 8x10 cell.
7. The crop share sits in **AP:BE**: **AP5** is `=IF(Z5="","",1-('Print sizes'!$D$5*'Print sizes'!$E$5*Z5^2)/($D5*$E5))`. It is the share of the image cut off to fill the print's shape: how much, not where.
8. The visible size columns **H:W** turn those into a label. **H5** is `=IF(Z5="","",IF(ROUND(Z5,6)>='Print sizes'!$B$1,"OK"&IF(AP5>=0.005," (crops "&TEXT(AP5,"0%")&")",""),"Too small: "&INT(Z5)&" DPI"))`. ROUND(,6) lets exactly 300 DPI count as OK. INT rounds down on purpose: 272.7 shows as 272 DPI.
9. **G5** names the largest listed size that prints: `=IF(F5="","",IF(MAX(BF5:BU5)=0,"None",INDEX('Print sizes'!$A$5:$A$20,MATCH(MAX(BF5:BU5),BF5:BU5,0))))`, where **BF:BU** hold each OK size's area (0 if not).
10. Read the Status in column **X** for each file, then check **Print sizes!J11**. CHECK says OK when the minimum is a number above 0 and no size or image row shows Check.

## Example

The six files are made up. Results at a 300 DPI minimum:

| File | Pixels | Largest print | Largest listed | 5x7 | 8x10 | 11x14 | 16x20 | 24x36 |
|---|---|---|---|---|---|---|---|---|
| wildflowers-print.jpg | 3600 x 5400 | 12 x 18 in | 12x18 | OK (crops 7%) | OK (crops 17%) | OK (crops 15%) | Too small: 225 DPI | Too small: 150 DPI |
| cabin-landscape.jpg | 6000 x 4800 | 16 x 20 in | 16x20 | OK (crops 11%) | OK | OK (crops 2%) | OK | Too small: 166 DPI |
| square-logo.png | 3000 x 3000 | 10 x 10 in | 8x10 | OK (crops 29%) | OK (crops 20%) | Too small: 214 DPI | Too small: 150 DPI | Too small: 83 DPI |
| phone-photo.jpg | 3024 x 4032 | 10 x 13.4 in | 8.5x11 | OK (crops 5%) | OK (crops 6%) | Too small: 274 DPI | Too small: 189 DPI | Too small: 112 DPI |
| small-scan.png | 1200 x 1500 | 4 x 5 in | None | Too small: 214 DPI | Too small: 150 DPI | Too small: 107 DPI | Too small: 75 DPI | Too small: 41 DPI |
| botanical-set-04.jpg | 4961 x 7016 | 16.5 x 23.3 in | 16x20 | OK (crops 1%) | OK (crops 12%) | OK (crops 10%) | OK (crops 12%) | Too small: 194 DPI |

The 2:3 wildflowers file fills 12x18 with no crop but loses 17% to fill 8x10. The square logo shows Too small: 272 DPI for 8.5x11 (272.7 rounded down). **Print sizes J4:J11** reads OK, 11, 0, 0, 6, 0, 0 and CHECK OK.

## Troubleshooting

### A size shows Too small at 299 DPI
The file is a few pixels short. 2399 x 3000 px on 8x10 is 299.875 DPI, shown as 299 and under 300. At 2400 x 3000 it says OK.

### A sideways photo fails a portrait size
Compare long side to long side. With the MAX and MIN columns, the 6000 x 4800 px cabin file prints 16x20 at exactly 300 DPI with no crop.

### Every size column is blank
The Minimum DPI in **B1** is text (such as "300" typed as text) or 0, so **Print sizes J4** says Check and the labels stay blank. Retype it as a number. A single blank row means that file's Status in **X** shows Check.

### Largest listed size says None
No listed size has enough pixels at your minimum. The 1200 x 1500 px small-scan.png gets only 250 DPI even at 4x6, so **G** shows None while **F** still gives 4 x 5 in.

## Template

[Download the .xlsx](templates/tidy-tabs-print-size-dpi-checker.xlsx)

**Images** holds up to 100 files (rows 5 to 104; you type **A:C**, helpers grouped in **Z:BU**). **Print sizes** holds the minimum DPI, up to 16 sizes (rows 5 to 20) and the summary in **I4:J11**. **How to use** explains each formula. Both data tabs print Landscape on US Letter.

Yellow cells are inputs, gray cells are formulas, OK sizes are light green, Too small is orange, and every Check turns red. To price sizes between your listed ones, see [how to price a custom size between two sizes](https://tidytabs.blogspot.com/2026/11/how-to-price-custom-size-between-two.html). To open the template in Google Sheets, go to **File > Import > Upload** and select the file.
