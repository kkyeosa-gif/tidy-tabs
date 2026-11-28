---
title: How to Count Barcode Scans in Excel and Google Sheets
labels: excel-formulas, small-business, inventory
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 216 workbooks and 2,272,536 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets, and no barcode scanner was used. The COUNTIF number-versus-text results were seen in LibreOffice Calc only.
search_description: Turn a column of barcode scans into a count per item, add typed case quantities, fix scans that lost a leading zero and flag codes not on your list.
image_prompts: A gift shop stockroom at year-end with a corded handheld barcode scanner resting on a sealed cardboard case of boxed greeting cards, a wire shelf behind it lined with amber candle jars and paper-wrapped soap bars, and a clipboard with a pencil-ticked count sheet propped against the case, cool evening light from a back window, realistic photo, no screens, no readable text, no logos
image_alt: Handheld barcode scanner on a sealed card case beside candle jars and a count clipboard
threads: Your scanner typed 012345678905 and the cell kept 12345678905?\nA key column looks the number up in your item list and gets the code back, leading zeros included.\nSUMPRODUCT then adds the units, with a blank Qty counted as 1. In the sample, 3 scans had lost a leading zero.
---
Give every scan a clean key in Scans **C5** with `=IF(A5="","",IF(ISNUMBER(A5),IFERROR(INDEX(Items!$A$5:$A$504,MATCH(A5,Items!$H$5:$H$504,0)),TEXT(A5,"0")),TRIM(CLEAN(A5))))`, so a scan that lost its leading zero matches its item again, then add each item's units on the Items tab with `=IF(A5="","",SUMPRODUCT(--(Scans!$C$5:$C$2004=A5),Scans!$D$5:$D$2004))`, where a blank Qty counts as 1. In the made-up gift shop's count, the 8 oz cedar candle comes to 9 units from 4 scans, and 3 scans had lost a leading zero.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (216 workbooks, 2,272,536 cell comparisons, 0 mismatches), plus the edge cases below. No barcode scanner was used. The COUNTIF results below were seen in LibreOffice Calc only. Excel and Google Sheets were not opened for this post.

![Five gift shop items with expected and counted units, four short and one never scanned](images/2026-11-29-team-02-tally-barcode-scans-into-inventory-counts-lost-leading-zeros-sumproduct-countif-excel-google-sheets/barcode-scan-tally-template.png) *Items tab of a made-up gift shop's year-end count from 17 barcode scans: per item the expected and counted units, the number of scans, the difference and the status (four Short and the gift wrap Not scanned in orange), with the summary of 56 units counted against 80 expected, 1 scan Not on list and CHECK OK. PDF export of the template from LibreOffice Calc 24.2.7.2.*

## Steps

A USB barcode scanner usually types the code like a keyboard into the active cell and presses Enter, so each beep fills the next row. In a General cell, 012345678905 can turn into the number 12345678905.

The steps are the same in Excel and Google Sheets.

1. On the **Items** tab, select column **A** and format it as Plain text (see [How to Keep Leading Zeros in Excel ZIP Codes](https://tidytabs.blogspot.com/2026/09/how-to-keep-leading-zeros-in-excel-zip.html)).
2. From row 5, type or paste each barcode exactly as printed in **A**, the item name in **B**, and the expected count in **C** if you have one.
3. On the **Scans** tab, click **A5** and scan. Blank rows between scans are fine.
4. Type a number in **Qty** (column **B**) only for a full case or items you counted by hand.
5. Column **C** builds the key. A text scan is cleaned with TRIM (spaces) and CLEAN (Tab, line feed and other control characters). A scan that became a number is looked up by its number value in Items column **H** (Code as number) and gets the item's own code back, leading zeros included, whatever the code's length.
6. Column **D** (Units) is the Qty, or 1 when Qty is blank. Column **E** (Status) shows OK, Not on list, or a Check message (text or bad Qty, a Qty with no scan, a decimal scan, more than 15 digits); Check rows add no units.
7. Column **F** (Note) marks repaired scans: "Lost leading zero (fixed)" or "Spaces or Tab removed". Both still count.
8. Back on **Items**, column **D** (Counted) adds the units with the SUMPRODUCT above, **E** counts the scan rows, **F** is Counted minus Expected and **G** shows Not scanned, Short, Over, Matches expected or Counted. See the SUMPRODUCT help from [Microsoft](https://support.microsoft.com/en-us/office/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e) or [Google](https://support.google.com/docs/answer/3094294), and the [CLEAN help from Microsoft](https://support.microsoft.com/en-us/office/clean-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41).
9. Read the totals in Scans **H4:I14**. **CHECK** in **I14** (repeated in Items **K13**) is OK only when no row shows Check, nothing is typed below row 2004, and units scanned equal the units matched to items plus the units on Not on list scans.

To put a value on the gap or adjust your book count, take the Difference column to [How to Compare a Physical Inventory Count With Expected Stock](https://tidytabs.blogspot.com/2026/10/how-to-compare-physical-inventory-count.html). This sheet counts units only. For stock between counts, see [stock on hand from an in and out log](https://tidytabs.blogspot.com/2026/10/how-to-track-stock-on-hand-from-log-of.html).

## Example

The count is made up: a gift shop's 17 scans of five items at the year-end count, with fictional codes. The five item codes pass the UPC check digit rule (checked in Python; see [How to Check a UPC Barcode Number's Check Digit](https://tidytabs.blogspot.com/2026/11/how-to-check-upc-barcode-numbers-check.html)), and 999999999999 is a deliberate stray scan.

Some of the scans:

| Row | Scan | Qty | Key | Units | Status | Note |
|---|---|---|---|---|---|---|
| 5 | 012345678905 | | 012345678905 | 1 | OK | |
| 6 | 12345678905 (number) | | 012345678905 | 1 | OK | Lost leading zero (fixed) |
| 8 | 012345678905 | 6 | 012345678905 | 6 | OK | |
| 13 | 008800123455 | 24 | 008800123455 | 24 | OK | |
| 19 | "400012000010 " (trailing space) | | 400012000010 | 1 | OK | Spaces or Tab removed |
| 20 | 999999999999 | | 999999999999 | 1 | Not on list | |
| 21 | 8800123455 (number) | 2 | 008800123455 | 2 | OK | Lost leading zero (fixed) |

The Items tab result:

| Code | Item | Expected | Counted | Scans | Difference | Status |
|---|---|---|---|---|---|---|
| 012345678905 | Cedar candle 8 oz | 14 | 9 | 4 | -5 | Short |
| 071234500017 | Lavender soap bar | 22 | 15 | 4 | -7 | Short |
| 008800123455 | Holiday card 6 pack | 30 | 29 | 5 | -1 | Short |
| 400012000010 | Wool dryer balls | 9 | 3 | 3 | -6 | Short |
| 400012000027 | Gift wrap roll | 5 | 0 | 0 | -5 | Not scanned |

The 17 scans come to 57 units: 56 matched to items and 1 on the Not on list scan. Three scans had lost a leading zero (rows 6, 12 and 21; the card code lost two). Counted totals 56 against 80 expected, 24 units short across four items, and the gift wrap never scanned. CHECK shows OK.

## Troubleshooting

### COUNTIF gives a different count than Counted
COUNTIF counts rows, not units. In LibreOffice Calc, `=COUNTIF(Scans!$A$5:$A$2004,A5)` on the raw sample scans gave 4 for the cedar candle, 4 for the lavender soap, 5 for the holiday cards and 2 for the dryer balls, while Counted shows 9, 15, 29 and 3. It ignored the typed Qty values and missed the dryer ball scan with a trailing space. Check the [COUNTIF help from Microsoft](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) for how your app matches, or use the template's key column.

### An exact match on the raw scans comes up even lower
In LibreOffice Calc, COUNTIF also counted the number 12345678905 as a match for the text "012345678905". An exact `SUMPRODUCT(--(Scans!$A$5:$A$2004=A5))` did not, and gave 3, 3, 4 and 2 rows. That is why the template compares the cleaned key, not the raw scan. These are LibreOffice results only; this post makes no claim about how Excel or Google Sheets treat number-like text in COUNTIF.

### A scan shows Not on list but the item is on your list
The scan does not match any code on Items. Check for a typo or an extra character in the Items code (its Status shows "Check: spaces in the code"), a new item that is not on the list yet, or a misread scan. Add the item or scan again.

### Status says "Check: over 15 digits (format A as text)"
The scan became a number with more than 15 digits. A number cell holds only 15 significant digits, so the last digits are already gone and the template does not guess them. Format column A as Plain text and scan that item again.

## Template

[Download the .xlsx](templates/tidy-tabs-barcode-scan-tally.xlsx)

**Scans** takes 2,000 scans (rows 5 to 2004) with Qty, key, units, status and note, and the summary and CHECK in **H4:I14**. **Items** holds up to 500 items (rows 5 to 504) with Counted, Scans, Difference, Status and the helper Code as number in **H**, plus totals in **J4:K13** (items listed, units counted and expected, items not scanned, short and over). **How to use** explains each formula. Scans and Items print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Not on list scans, repaired scans and items that are short, over or not scanned turn light orange, Matches expected and the unit totals are light green, and every Check turns red.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
