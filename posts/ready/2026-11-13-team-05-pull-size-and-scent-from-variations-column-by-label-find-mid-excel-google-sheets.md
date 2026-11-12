---
title: How to Pull Size and Scent Out of an Order Variations Column in Excel and Google Sheets
labels: excel-formulas, text-functions, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 14 workbooks and 11,256 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Pull the value after a label such as Size: or Scent: from a variations column with FIND, MID and LOWER, in any order. Free Excel template.
image_prompts: Three soy candles in different jar sizes on a packing table, each with a small blank kraft hang tag on twine, next to a spool of gift ribbon and folded tissue paper.
image_alt: Three soy candles in different jar sizes beside gift ribbon and tissue paper
threads: Wick size: Double keeps coming back as the Size.\nSearch for the label with a comma in front, ",size:", on LOWER text, then MID out what sits before the next comma. The options can be in any order.\nOne catch: a value with its own comma gets cut, so "Happy birthday, Mom" comes back as Happy birthday.
---
Pull the value after a label such as Size: with `=IFERROR(TRIM(MID($C4,FIND(","&LOWER(D$3)&":",LOWER($C4))+LEN(D$3)+2,FIND(",",$C4,FIND(","&LOWER(D$3)&":",LOWER($C4))+1)-FIND(","&LOWER(D$3)&":",LOWER($C4))-LEN(D$3)-2)),"")`, where D3 holds the label and C4 is the variations text with a comma added at each end. Because it looks for the comma before the label, Size does not match Wick size, and the options can be in any order.

In the made-up sample, "Size: 12 oz, Scent: Cedar & Sage, Gift wrap: Yes" gives 12 oz, Cedar & Sage and Yes.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was recalculated and compared with an independent Python model (14 workbooks, 11,256 cell comparisons, 0 mismatches), plus the edge cases below. Excel and Google Sheets have the same functions but were not opened for this post; the open questions there are TRIM on a range inside SUMPRODUCT and how pasted line breaks come through.

![Order list with scent, size and gift wrap pulled from mixed variation text](images/2026-11-13-team-05-pull-size-and-scent-from-variations-column-by-label-find-mid-excel-google-sheets/variation-values-by-label-template.png) *Scent, size and gift wrap pulled from twelve variation texts written in different ways. PDF export from LibreOffice Calc 24.2.*

## Steps

The steps are the same in Excel and Google Sheets.

1. Type the option names across the header row: **Scent** in **D3**, **Size** in **E3**, **Gift wrap** in **F3**, spelled as in your order text.
2. Paste order numbers into **A4** and the variations text into **B4** and down. The template holds rows 4 to 203.
3. In **C4**, clean the text so every option sits between two commas:
   `=IF(TRIM(B4)="","",","&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B4,CHAR(10),","),CHAR(13),","),";",",")),", ",",")," ,",",")," :",":")&",")`
   Line breaks and semicolons become commas, extra spaces and spaces next to a comma or before a colon go, and a comma is added at each end. For #2101 it gives `,Scent: Lavender,Size: 8 oz,`.
4. In **D4**, type the formula from the top of this page.
5. Copy **D4** right to **F4**, then fill **C4:F4** down to the last order.
6. In **G4**, count the empty options: `=IF(C4="",IF(TRIM(A4)="","","No variations"),SUMPRODUCT((TRIM($D$3:$F$3)<>"")*(D4:F4="")))`. Fill it down.

How it works:

- `LOWER` on both the label and the text makes the match ignore case. The value comes from the original text, so it keeps its capitals.
- `FIND(","&LOWER(D$3)&":", ...)` finds `,size:`. The comma in front is the trick: `,Wick size: Double,` has no comma right before `size:`, so nothing matches.
- The value starts at that position plus the label length plus 2 (the comma and the colon).
- The length is the position of the next comma minus that start. The comma added at the end of C4 closes the last option too.
- `TRIM` drops the space after the colon; `IFERROR` gives a blank when the label is missing.

The template's D4 also trims the label and stays blank when C4 or the header is empty; a blank header is not counted in G.

Help pages: Microsoft's [FIND function](https://support.microsoft.com/en-us/office/find-function-06213f91-b5be-4544-8b0b-2fd5a775436f) and [MID function](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028), and Google's [FIND](https://support.google.com/docs/answer/3094126). They describe the functions, not our test.

Newer functions such as TEXTAFTER or REGEXEXTRACT can do similar work in some versions. We did not test them, and the template does not use them.

## Example

Sample data is fictional: 12 made-up candle orders. The formats are examples, not copied from any marketplace export.

| Order no. | Variations as exported | Scent | Size | Gift wrap | Not found |
|---|---|---|---|---|---|
| #2101 | Scent: Lavender, Size: 8 oz | Lavender | 8 oz | | 1 |
| #2102 | Size: 12 oz, Scent: Cedar & Sage, Gift wrap: Yes | Cedar & Sage | 12 oz | Yes | 0 |
| #2103 | scent:Vanilla bean,size:4 oz,gift wrap:No | Vanilla bean | 4 oz | No | 0 |
| #2104 | Scent: Pumpkin Spice / Size: 8 oz / Gift wrap: Yes (three lines in one cell) | Pumpkin Spice | 8 oz | Yes | 0 |
| #2105 | Wick size: Double, Scent: Lavender | Lavender | | | 2 |
| #2106 | Size : 8 oz , Scent : Cedar & Sage | Cedar & Sage | 8 oz | | 1 |
| #2107 | Gift wrap: No; Scent: Fresh Linen; Size: 12 oz | Fresh Linen | 12 oz | No | 0 |
| #2108 | (blank) | | | | No variations |
| #2109 | Scent: Lavender, Scent: Cedar & Sage, Size: 4 oz | Lavender | 4 oz | | 1 |
| #2110 | Size: 8 oz, Scent: | | 8 oz | | 2 |
| #2111 | Scent: Bay Rum, Size: 8 oz, Pickup time: 2:30 PM | Bay Rum | 8 oz | | 1 |
| #2112 | Color: Gold lid, Size: 8 oz | | 8 oz | | 2 |

The summary in **I3:J6** shows 12 orders listed, 4 with every option found, 7 with some options not found and 1 with no variations.

Rows #2102, #2104 and #2107 show that order and separators do not matter. Row #2109 lists Scent twice, and the first one wins.

## Troubleshooting

### Size shows Double or another wrong value
A formula without the leading comma matches the end of a longer label. In LibreOffice, `=TRIM(MID(B8,SEARCH("Size:",B8)+5,...))` returned " Double" from "Wick size: Double, Scent: Lavender". The template returns nothing for Size there.

### A value is cut short
The value contains a comma. Commas, semicolons and line breaks all count as separators, so "Note: Happy birthday, Mom, Size: 8 oz" gives Happy birthday for Note. Fix free-text options like this by hand.

### An option comes back empty although it is in the text
The header is spelled differently. Labels match whole, so Size does not find Sizes or Size (oz); copy the spelling from the text. A tab before the colon also stops the match.

### Odd spaces are left in a value, or a space vanished
The cleanup handles regular spaces only. Tabs and non-breaking spaces stay; see [How to Remove Non-Breaking Spaces in Excel and Google Sheets](https://tidytabs.blogspot.com/2026/10/how-to-remove-non-breaking-spaces-in.html). It also drops a space just before a colon inside a value, so "Ratio: 1 :2" gives 1:2.

## Template

[Download the .xlsx](templates/tidy-tabs-variation-values-by-label.xlsx)

**Orders** has order numbers in **A4:A203**, variations text in **B4:B203**, the Cleaned helper in **C**, option names in **D3:F3** with values below, Options not found in **G** and a summary in **I3:J6**. It prints Landscape on US Letter with the header row repeated. **How to use** explains each column.

Yellow cells are inputs, gray cells are formulas. Options not found is green at 0, red above 0 or for No variations. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
