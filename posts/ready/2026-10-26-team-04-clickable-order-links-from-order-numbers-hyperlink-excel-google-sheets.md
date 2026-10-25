---
title: How to Turn Order Numbers Into Clickable Links With HYPERLINK in Excel and Google Sheets
labels: excel-formulas, small-business, orders
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and the displayed text and address of each row were compared with Python string building on the sample and on 15 edited copies (spaces inside and around order numbers, numbers, blank rows, a base address without the trailing slash, an http:// base, a 255 and a 256 character address). No link was clicked, no web page was opened, and the formula makes no web request. Excel and Google Sheets were not opened.
search_description: Turn order numbers into clickable links with HYPERLINK in Excel and Google Sheets. Free template with a base address cell and a column showing each full address.
image_prompts: A shipping desk with a stack of paper packing slips held by a steel binder clip, a roll of kraft tape with a dispenser, and a small brown parcel tied with twine, soft window light, no screens, no readable text, no logos
image_alt: Brown parcel, kraft tape roll and clipped packing slips on a small shipping desk
threads: Want every order number in your sheet to open its own page?\nKeep the shared part of the address in B1 and join each number to it inside HYPERLINK.\nThe template also swaps spaces for %20, so "A 12" still works, and shows the plain address next to each link so you can check it first.
---
Type the shared part of your address in **B1**, then build each link with `=IF(TRIM(A4)="","",HYPERLINK($B$1&SUBSTITUTE(TRIM(A4)," ","%20"),"Order "&TRIM(A4)))`, where A4 is the order number. The cell shows "Order" and the number, and the address it opens is the base address with the number added at the end. The template points at the placeholder `https://example.com/orders/`, so replace it with your own store or carrier address.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and the displayed text and addresses below were recalculated and compared with Python string building. The steps for Excel and Google Sheets use `HYPERLINK`, an old function in both. Neither Excel nor Google Sheets was opened, and no link was clicked.

![Orders sheet with a base address cell, order numbers, blue Order links and a column of full addresses](images/2026-10-26-team-04-clickable-order-links-from-order-numbers-hyperlink-excel-google-sheets/clickable-order-links-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The simplest version is `=HYPERLINK($B$1&A4,"Order "&A4)`. It joins the base address and the order number with `&`, and `HYPERLINK` turns the result into a clickable cell. The template formula wraps the number in two extra functions, so it also handles order numbers with spaces or stray spaces and leaves the cell blank on empty rows.

1. Type your base address in **B1**. It must start with `https://` and end with `/` or `=`, for example `https://example.com/orders/` or `https://track.example.com/?id=`.
2. Type your order numbers in **A4**, one per row. If your numbers have leading zeros, format column **A** as Text first, because a number typed as a number loses them.
3. In **B4**, type `=IF(TRIM(A4)="","",HYPERLINK($B$1&SUBSTITUTE(TRIM(A4)," ","%20"),"Order "&TRIM(A4)))`.
4. In **C4**, type `=IF(TRIM(A4)="","",$B$1&SUBSTITUTE(TRIM(A4)," ","%20"))`. This shows the exact address the link points to.
5. Copy **B4:C4** down. The template goes to row 203.
6. Read column **C** before you send a link to anyone. If the address looks wrong, fix **B1**.

How the formula reads:

- `TRIM(A4)` removes extra spaces around the number and between its parts.
- `SUBSTITUTE(...," ","%20")` swaps each remaining space for `%20`, the code web addresses use for a space.
- `$B$1&...` puts the base address in front. The `$` keeps it fixed when you copy down.
- `HYPERLINK(address, text)` makes the clickable cell. The second part is what the cell shows.
- `IF(TRIM(A4)="","",...)` leaves the row empty when there is no order number.

Only the space is changed. The characters `&`, `#`, `/` and `é` in an order number were left as typed, so a number that contains them may need different handling on your site.

### In Excel

Type the formulas as shown. Microsoft's page for [HYPERLINK](https://support.microsoft.com/en-us/office/hyperlink-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f) describes the function and says Excel limits the address to 255 characters. We did not test that limit in Excel. That page is the source, not our test.

### In Google Sheets

Type the same formulas. Google's page for [HYPERLINK](https://support.google.com/docs/answer/3093313) describes the function. To start from our file, use **File > Import > Upload**.

## Example

Sample data is fictional: 15 order numbers, and `example.com` is a placeholder, not a real shop. With the base address in B1, the first rows looked like this after recalculation. The Link column shows the text, and the Address column shows where it points.

| Order number | Link shows | Address | Length |
|---|---|---|---|
| 10482 | Order 10482 | https://example.com/orders/10482 | 32 |
| 10483 | Order 10483 | https://example.com/orders/10483 | 32 |
| A 12 | Order A 12 | https://example.com/orders/A%2012 | 33 |
| 10490 | Order 10490 | https://example.com/orders/10490 | 32 |
| SH-2031 | Order SH-2031 | https://example.com/orders/SH-2031 | 34 |
| C 104 | Order C 104 | https://example.com/orders/C%20104 | 34 |

The summary beside the table read:

| Item | Result |
|---|---|
| Order numbers listed | 15 |
| Links built | 15 |
| Longest address | 34 |
| Check | OK |

Only the displayed text and the address were read back. The links were not clicked, so we do not know what any page would do. The formula itself makes no web request.

## Troubleshooting

### A link opens the wrong page
Check the base address in **B1**. Without the trailing `/`, the address joins as `https://example.com/orders10482`, and the check cell shows CHECK. Column **C** shows the built address, so you can see the problem before clicking. Add the slash, or the `=` if your site uses `?id=`.

### A row has no link
A cell that holds only spaces counts as empty to `TRIM`, so the row stays blank. In a copy with one such row, "Order numbers listed" was higher than "Links built" and the check cell showed CHECK. Clear the cell, or type the number.

### The address is too long
Microsoft's help page gives a limit of 255 characters for Excel. The template shows each length in column D and the longest in the summary. In copies with addresses of 255 and 256 characters, the check read OK for 255 and CHECK for 256. The limit itself was not tested in Excel.

### The check says CHECK for the base address
The check wants `https://` at the start, a `/` or `=` at the end, and no spaces. In copies with an `http://` base, a base with a space and an empty base, it showed CHECK each time. A base ending in `=` passed.

### Leading zeros disappeared
A number typed into a General cell drops its leading zeros, so 00123 becomes 123. Format column **A** as Text before you type or paste the numbers.

## Template

[Download the .xlsx](templates/tidy-tabs-order-links.xlsx)

The file has two tabs. **Orders** holds the base address in **B1**, the yellow order numbers in **A4:A203**, the gray formulas in **B:D** (link, address, length) and a summary in **F4:G7** with the check in **G7**. **How to use** has short notes.

The shop and the orders are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints landscape on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
