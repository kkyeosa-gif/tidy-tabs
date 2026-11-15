---
title: How to Calculate Delivery Distance in Miles From Latitude and Longitude in Excel and Google Sheets
labels: excel-formulas, small-business, shipping, pricing
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python haversine model, 19 workbooks and 7,771 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Get straight-line miles between two coordinates with the haversine formula (ASIN, RADIANS), then mark each delivery in area and add a fee. Free template.
image_prompts: A navy insulated delivery bag unzipped on the passenger seat of a parked car, a white bakery cake box tied with twine inside it, and a folded paper road map with a small pushpin tucked beside the bag, late afternoon light through the car window, realistic photo, no screens, no readable text, no logos
image_alt: Cake box in an open insulated bag on a car seat beside a folded map
threads: Drop one minus sign and a Cambridge customer lands 6,127.60 miles from the shop.\nStraight-line miles from two coordinates: haversine with ASIN, SQRT, RADIANS and 3958.8 for Earth's radius.\nThen base + per mile x ROUNDUP(miles,0): 2.76 miles at $5.00 + $1.50 is $9.50.
---
Get the straight-line miles from your shop to a customer with `=2*3958.8*ASIN(SQRT(SIN(RADIANS(lat2-lat1)/2)^2+COS(RADIANS(lat1))*COS(RADIANS(lat2))*SIN(RADIANS(lon2-lon1)/2)^2))`, where 3958.8 is Earth's mean radius in miles, and charge a base fee plus a fee per started mile with `=base+per_mile*ROUNDUP(miles,0)`. In the made-up sample, a shop at 42.3601, -71.0589 is 2.76 miles from a Cambridge customer at 42.3736, -71.1097, which gives a $9.50 fee at $5.00 plus $1.50 per started mile, and the download flags a longitude typed without its minus sign (6,127.60 miles) in the Status column.

This is the haversine formula, and the result is a close estimate of the straight-line distance on a round Earth. It is not road miles: the drive is longer, so set your radius and fee with that in mind. The fee rule is your own choice.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python haversine model (19 workbooks, 7,771 cell comparisons, 0 mismatches), plus the edge cases below. ASIN, SQRT, SIN, COS, RADIANS and ROUNDUP also exist in Excel and Google Sheets, but neither app was opened for this post.

![Six delivery rows with coordinates, straight-line miles, in-area answer and delivery fee](images/2026-11-16-team-01-local-delivery-distance-miles-from-latitude-longitude-fee-radians-asin-excel-google-sheets/delivery-distance-and-fee-template.png) *Six fictional deliveries measured from one shop: Cambridge at 2.76 miles costs $9.50, and Worcester at 38.58 miles shows Out of area (the summary columns are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. They use the **Deliveries** tab of the template.

1. Type the shop's latitude in **B1** (42.3601) and longitude in **C1** (-71.0589).
2. Type the delivery radius in miles in **B2** (15), the base fee in **B3** ($5.00) and the fee per started mile in **B4** ($1.50).
3. Starting in row 7, type the customer in **A7**, the ZIP in **B7** (the column is formatted as text, so 02139 keeps its zero), and the customer's latitude and longitude in **C7** and **D7** as plain decimal numbers. To get them, right-click the spot in Google Maps on a computer and click the coordinates in the pop-up to copy them, as [Google Maps Help](https://support.google.com/maps/answer/18539) describes. Latitude comes first.
4. In **E7**, type `=2*3958.8*ASIN(SQRT(SIN(RADIANS(C7-$B$1)/2)^2+COS(RADIANS($B$1))*COS(RADIANS(C7))*SIN(RADIANS(D7-$C$1)/2)^2))` for the miles. RADIANS turns degrees into the unit SIN and COS expect, and the dollar signs keep the shop's cells fixed when you fill down.
5. In **F7**, type `=IF(ROUND(E7,2)<=$B$2,"Yes","Out of area")`. The miles are compared as shown with two decimals, so a point at 15.00 miles is in a 15-mile area.
6. In **G7**, type `=IF(F7="Yes",$B$3+$B$4*ROUNDUP(ROUND(E7,2),0),"")`. ROUNDUP charges every started mile, so 2.76 miles counts as 3 and the fee is $5.00 + $1.50 x 3 = $9.50.
7. Fill **E7:G7** down to the last row.

The template wraps E7 in checks. It stays blank when both coordinates are blank, shows Check when one is blank, typed as text or out of range, and only calculates when the settings check in **K6** says OK. It also puts `MIN(1,...)` inside ASIN as a safety step: rounding can push the value under the square root a hair above 1, and ASIN of anything above 1 is an error.

Help pages: Microsoft's [ASIN function](https://support.microsoft.com/en-us/office/asin-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347) and [RADIANS function](https://support.microsoft.com/en-us/office/radians-function-ac409508-3d48-45f5-ac02-1497c92de5bf), and Google's [RADIANS](https://support.google.com/docs/answer/3093437).

## Example

The customers are fictional and the coordinates are sample points. The shop is at 42.3601, -71.0589 with a 15-mile radius, a $5.00 base fee and $1.50 per started mile.

| Customer | ZIP | Latitude | Longitude | Miles (straight line) | In area? | Delivery fee | Status |
|---|---|---|---|---|---|---|---|
| Cambridge order | 02139 | 42.3736 | -71.1097 | 2.76 | Yes | $9.50 | OK |
| Somerville order | 02143 | 42.3876 | -71.0995 | 2.81 | Yes | $9.50 | OK |
| Newton order | 02458 | 42.3370 | -71.2092 | 7.84 | Yes | $17.00 | OK |
| Quincy order | 02169 | 42.2529 | -71.0023 | 7.95 | Yes | $17.00 | OK |
| Lynn order | 01902 | 42.4668 | -70.9495 | 9.25 | Yes | $20.00 | OK |
| Worcester order | 01608 | 42.2626 | -71.8023 | 38.58 | Out of area | | OK |

The summary in **J6:K14** shows the settings check OK, 6 deliveries listed, 5 in area, 1 out of area, 9.25 as the farthest in-area miles, $73.00 in fees for in-area deliveries, 0 rows showing Check, nothing typed below row 106 and CHECK OK.

Newton at 7.84 miles and Quincy at 7.95 miles both round up to 8 started miles, so both pay $17.00.

## Troubleshooting

### A customer shows thousands of miles away
The minus sign on the longitude was dropped. The Cambridge point typed as 42.3736, 71.1097 lands in Asia, 6,127.60 miles from the shop, and shows Out of area with Status Check: sign. US longitudes are negative, so retype it as -71.1097.

### A row shows 9,473.32 miles and Check: sign
The latitude and longitude were typed in the wrong columns. Map apps list latitude first, so the 42 value goes in **C** and the -71 value goes in **D**. The template flags any row with a negative latitude or a positive longitude, which also catches this swap.

### Status shows Check: text
A coordinate is text, not a number, for example "42.3736" pasted with quotes, 42.3736° with a degree sign, or 42.3736 N with a letter. Retype it as a plain decimal number. Check: blank means one of the two coordinates is missing, and Check: range means a latitude is outside -90 to 90 or a longitude outside -180 to 180.

### Every row shows Check: settings
The shop's coordinates, the radius or a fee in **B1:B4** or **C1** is blank, text or out of range, so **K6** shows Check. A fee typed as the text "$1.50" counts as text. Type 1.50 and let the cell's dollar format show it.

### A customer at your own address
The shop's own point gives 0.00 miles, Yes, and the base fee of $5.00, because ROUNDUP of 0 is 0 started miles. If you want a minimum charge above the base fee, raise **B3**.

## Template

[Download the .xlsx](templates/tidy-tabs-delivery-distance-from-coordinates.xlsx)

The file has two tabs. **Deliveries** holds the shop settings in **B1:C1** and **B2:B4**, then rows 7 to 106: Customer, ZIP, Latitude and Longitude (inputs in **A:D**), then Miles (straight line), In area?, Delivery fee and Status (formulas in **E:H**), with the summary in **J6:K14**. **How to use** explains each column and the haversine formula. The Deliveries tab prints Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Status turns red for any Check, and Out of area rows are shaded orange. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
