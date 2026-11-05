---
title: How to Look Up a Shipping Zone From a ZIP Code in Excel and Google Sheets
labels: excel-formulas, small-business, shipping
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 21 workbooks and 17,367 comparisons with 0 mismatches. Checked: the sample, no orders, 41 odd ZIPs (2108 as a number and as text, 7030, spaces around a ZIP, ZIP+4, decimals, negatives, 100000, 99999, 0, 4 and 6 digits, letters, 1e345, 12E45, an inner space, a dash in the wrong place, an empty string), every zone boundary (99/100, 199/200, 299/300, 499/500, 699/700, 849/850, 959/960), a table with one zone and with 10 zones, a table that does not start at 0, starts that do not rise, equal starts, a gap, a text price, a blank zone name, a repeated zone name, an empty table, 200 orders, and six random sets. Excel and Google Sheets were not used. Not run: a TRUE or FALSE value in the ZIP column, more than 200 orders.
search_description: Look up a shipping zone and price from a ZIP code in Excel and Google Sheets with TEXT, INDEX and MATCH. Fixes lost leading zeros. Free template.
image_prompts: A kraft paper mailer with a blank white address label, a small brass postal scale and a folded paper zone map with no readable text on a wooden packing table, soft daylight, realistic photo, no screens, no readable text, no logos
image_alt: Kraft mailer with a blank label beside a brass postal scale and a folded map
threads: Two steps turn 2108 into 02108 and then into Zone 1 at $5.90 in the sample.\nTEXT puts the zero back, and INDEX with MATCH looks up the zone from the first three digits.\nThe zone table in the template is made up, so swap in your carrier prices.
---
Restore a ZIP that lost its zero with `=TEXT(C5,"00000")`, take the first three digits with `=VALUE(LEFT(D5,3))`, then look up the zone with `=INDEX(Zones!$B$5:$B$14,MATCH(E5,Zones!$A$5:$A$14,1))`. In the sample, 2108 becomes 02108 and lands in Zone 1 at $5.90 (a made-up zone table).

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. TEXT, VALUE, LEFT, INDEX and MATCH exist in current Excel and Google Sheets, but neither app was opened for this post.

![Orders with ZIP codes restored to five digits and the shipping zone and price](images/2026-11-06-team-04-shipping-zone-from-zip-code-index-match-excel-google-sheets/shipping-zone-from-zip-code-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Orders** has one row per order from row 5: order number in **A**, customer in **B**, the ZIP as typed in **C**. **Zones** holds your zone table in **A5:C14**.

1. On **Zones**, list the zones in **A5:C14**. Column **A** is the first three digits where the zone starts, as numbers, smallest first. The first row must be 0. Column **B** is the zone name and column **C** is the price.
2. On **Orders**, paste or type each ZIP in column **C**. A ZIP that lost its zero shows up as a number, such as 2108.
3. In **D5**, type `=TEXT(C5,"00000")`. This pads a number to five digits, so 2108 becomes 02108.
4. In **E5**, type `=VALUE(LEFT(D5,3))`. This is the first three digits as a number, such as 21.
5. In **F5**, type `=INDEX(Zones!$B$5:$B$14,MATCH(E5,Zones!$A$5:$A$14,1))` for the zone.
6. In **G5**, type `=INDEX(Zones!$C$5:$C$14,MATCH(E5,Zones!$A$5:$A$14,1))` for the price.
7. Fill **D5:G5** down.

The `1` in MATCH finds the last start that is not above your number. Prefix 164 is not in the table, so MATCH uses the 100 row and returns Zone 2. That is why the table must rise from top to bottom. The dollar signs lock the table when you fill down.

The template uses a longer D5 so it also accepts ZIPs stored as text, including ZIP+4 such as 02134-1234:

`=IF(C5="","",IF(ISNUMBER(C5),IF(AND(C5>=0,C5<=99999,C5=INT(C5)),TEXT(C5,"00000"),"Check"),IF(OR(LEN(TRIM(C5))=5,AND(LEN(TRIM(C5))=10,MID(TRIM(C5),6,1)="-")),IF(SUMPRODUCT(--ISNUMBER(FIND(MID(TRIM(C5),{1,2,3,4,5},1),"0123456789")))=5,LEFT(TRIM(C5),5),"Check"),"Check")))`

It returns Check for anything that is not a clean ZIP. The template also wraps E, F and G in tests, so they show Check too instead of an error.

Use the zone chart and prices from your own carrier. The zones in this post are invented for the example and are not a real carrier chart.

Help pages: Microsoft's [MATCH function](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a) and [INDEX function](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a4-a822-b56b061328bd), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. The pages were not read for this post.

## Example

Sample data is fictional: 16 orders and a made-up zone table for a shop shipping from Boston, MA.

| Starts at first 3 digits | Zone | Price |
|---|---|---|
| 0 | Zone 1 | $5.90 |
| 100 | Zone 2 | $6.40 |
| 200 | Zone 3 | $7.10 |
| 300 | Zone 4 | $7.80 |
| 500 | Zone 5 | $8.60 |
| 700 | Zone 6 | $9.30 |
| 850 | Zone 7 | $10.20 |
| 960 | Zone 8 | $11.40 |

| Customer | ZIP as typed | ZIP (5 digits) | First 3 | Zone | Price |
|---|---|---|---|---|---|
| Maria Alvarez | 02108 (text) | 02108 | 21 | Zone 1 | $5.90 |
| Ben Carter | 2108 (number) | 02108 | 21 | Zone 1 | $5.90 |
| Tom Nguyen | 7030 (number) | 07030 | 70 | Zone 1 | $5.90 |
| Lena Fischer | 10001 | 10001 | 100 | Zone 2 | $6.40 |
| Omar Haddad | 33101 | 33101 | 331 | Zone 4 | $7.80 |
| Chloe Park | 60601 (text) | 60601 | 606 | Zone 5 | $8.60 |
| Jamal Reed | 94105 | 94105 | 941 | Zone 7 | $10.20 |
| Sofia Rossi | 98101 | 98101 | 981 | Zone 8 | $11.40 |
| Dan Whitaker | 02134-1234 (text) | 02134 | 21 | Zone 1 | $5.90 |
| Grace Kim | 85001 (text) | 85001 | 850 | Zone 7 | $10.20 |

Grace Kim's prefix is 850, exactly on a boundary, so she lands in Zone 7. The Zones tab then counts all 16 orders: Zone 1 has 6 orders totaling $35.40, Zone 2 has 1 ($6.40), Zone 3 has 0, Zone 4 has 2 ($15.60), Zone 5 has 1 ($8.60), Zone 6 has 2 ($18.60), Zone 7 has 3 ($30.60) and Zone 8 has 1 ($11.40). The total is $126.60.

## Troubleshooting

### The ZIP loses its leading zero
A ZIP stored as a number loses its zero, so 02108 shows as 2108. Step 3 puts the zero back, but it cannot know the original. In the test, a number such as 123 became 00123. If a ZIP could have been shorter than it looks, fix the source column.

### The row says Check
A number with decimals, a negative number, or a number above 99999 gives Check. Text must be 5 characters or ZIP+4 (10 characters with a dash in position 6) after extra spaces are trimmed, and the first 5 must be digits. In the test, 1e345 and 12E45 gave Check, and so did letters. Only the first 5 digits of a ZIP+4 are checked.

### Every row says Check
The zone table is the usual cause. In the test, a table that did not start at 0, did not rise, had a gap, had a text price or had a blank zone name made **H4** on Zones say Check and every order row say Check. An empty table did the same.

### The CHECK cell on Zones says Check
**H8** says OK when the zone table is valid, no order has a ZIP to check and every order is counted in a zone. A repeated zone name also made it say Check in the test, so keep zone names unique.

## Template

[Download the .xlsx](templates/tidy-tabs-shipping-zone-from-zip.xlsx)

The file has three tabs. **Orders** has 200 rows from row 5 with 16 fictional sample orders and yellow input cells for the ZIP as typed. **Zones** has 10 table slots in **A5:C14** (8 used in the sample), a count and a shipping total per zone, an **All zones** row, and a summary in **H4:H8**: zone table check, orders listed, orders with a ZIP to check, orders counted in a zone, and a **CHECK** cell. **How to use** has notes on each formula.

Replace the sample zones and prices with your carrier's. A list longer than 200 orders was not tested.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
