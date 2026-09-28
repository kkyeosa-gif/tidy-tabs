---
title: How to Create a Dependent Drop-Down List in Excel and Google Sheets
labels: data-validation, excel-formulas, google-sheets
tested_in: LibreOffice Calc 24.2.7 (Linux), INDIRECT/named-range mechanism only. Excel steps from Microsoft Support, Google Sheets steps from Google Docs Editors Help (not tested here).
image_prompts: 
image_alt: 
search_description: Make a Subcategory dropdown follow Category by naming each list after its category and setting the second cell's validation to =INDIRECT(A2).
---
Make a second dropdown follow the first by naming each list after its category, then pointing the second cell's data validation at `=INDIRECT(A2)`. Pick "Candles" in column A and column B's dropdown limits itself to candle subcategories only.

> Works in: Excel and Google Sheets (standard named ranges, data validation, and the INDIRECT function). Checked in: LibreOffice Calc 24.2.7 (Linux) only, and only the INDIRECT/named-range mechanism, not the dropdown click itself (see Example).

![Spreadsheet example with columns Category, Subcategory options, COUNTA(INDIRECT())](images/2026-09-29-team-01-dependent-dropdown-list-excel-google-sheets/how-to-create-a-dependent-drop-down-list-in-excel-example.png) *The example from this post laid out in a spreadsheet. Rendered with LibreOffice Calc.*

## Steps

Build the source lists once, then set up two linked dropdowns. The named-range step is the same idea in both apps; the dropdown menus look different.

### In Excel

1. On a tab you'll use for source lists (call it **Lists**), type each category name across the top row: Candles in A1, Soap in B1, Jewelry in C1, Cards in D1.
2. Under each header, list that category's items in the cells below it.
3. Select one category's items (not the header), then click the Name Box to the left of the formula bar, type the header text exactly, for example `Candles`, and press Enter. Repeat for every column. Microsoft's [guide to defining names](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64) covers this.
4. On your order form, select the cell for Category (for example A2) and go to **Data > Data Validation**.
5. Under **Allow**, choose **List**, and set **Source** to your category headers, for example `=Lists!$A$1:$D$1`. Click **OK**.
6. Select the cell for Subcategory (B2) and open **Data > Data Validation** again.
7. Under **Allow**, choose **List**, and set **Source** to `=INDIRECT(A2)`. Click **OK**. Microsoft documents the function on its [INDIRECT function](https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261) page.
8. Copy B2's validation down the rest of the order form.

### In Google Sheets

1. Build the same **Lists** tab: category names across row 1, that category's items in the column below each one.
2. Select one category's items, then go to **Data > Named ranges**, click **Add a range**, and name it exactly the same as the header above it, for example `Candles`. Repeat for every column. See Google's [named ranges](https://support.google.com/docs/answer/63175) help page.
3. On your order form, select the Category cell and go to **Data > Data validation**.
4. Under **Criteria**, choose **Dropdown (from a range)** and point it at your category headers, for example `Lists!A1:D1`. Click **Done**.
5. Select the Subcategory cell and open **Data > Data validation** again.
6. Under **Criteria**, choose **Dropdown (from a range)**, then type `=INDIRECT(A2)` into the range field instead of selecting cells. Click **Done**. Google's [INDIRECT function](https://support.google.com/docs/answer/3093377) page explains how it turns text into a reference.
7. Copy the Subcategory cell's validation down the rest of the order form.

## Example

The template's **Lists** tab has four categories with a different number of items each. A check formula, `=COUNTA(INDIRECT(A2))`, sits next to the order form and counts how many subcategories INDIRECT found for that row's category:

| Category | Subcategory options | COUNTA(INDIRECT()) |
| --- | --- | --- |
| Candles | Lavender soy 8 oz, Cedar soy 8 oz, Vanilla soy 4 oz | 3 |
| Soap | Oatmeal goat milk bar, Charcoal bar | 2 |
| Jewelry | Brass hoop earrings, Beaded necklace, Cuff bracelet, Stacking ring | 4 |
| Cards | Letterpress birthday card, Letterpress thank you card | 2 |

Product names are fictional. Clicking the actual dropdown arrow to confirm it opens and filters correctly isn't something that can be done headlessly, so that part wasn't screen-tested. What was checked: after a forced recalculation in LibreOffice Calc, the COUNTA/INDIRECT formula above returned the exact item count for all four categories, which confirms INDIRECT resolves to the right named range every time. That's the mechanism the dropdown itself relies on.

## Troubleshooting

### Subcategory dropdown is empty
The named range's name doesn't match the Category text exactly, including case and spaces. INDIRECT turns the cell's text into a name; if no range has that exact name, it finds nothing. Rename the range to match the header word for word.

### Subcategory shows items from every category mixed together
The Subcategory cell's validation Source is still set to a fixed range (like `Lists!A2:A10`) instead of `=INDIRECT(A2)`. Reopen its data validation and re-enter the INDIRECT formula.

### Excel warns "The Source currently evaluates to an error"
This shows up when the Category cell is blank or its text doesn't match any named range yet. It's a warning, not a block; click **Yes** to keep the rule, then fix the Category cell so INDIRECT has something valid to resolve.

### Adding a new category doesn't add its dropdown
Adding a column to Lists only creates the data. You also need to name that column as a range (step 3 in Excel or Google Sheets above) and add its header to the Category dropdown's source list.

### Google Sheets dropdown shows a red warning triangle instead of filtering
This usually means the range field for Subcategory still holds actual cell references instead of the `=INDIRECT(A2)` formula, so Sheets is trying to validate against a literal range that doesn't match the selected category. Retype the formula into the range box.

## Template

[Download the .xlsx](templates/tidy-tabs-dependent-dropdown-list.xlsx)

The workbook has three tabs. **Order form** holds the Category dropdown, the dependent Subcategory dropdown, and the COUNTA/INDIRECT check column from the Example above. **Lists** holds the four sample categories, each column already set up as a named range matching its header. **How to use** walks through the same setup in plain language.

To open it in Google Sheets, click **File > Import**, upload the .xlsx, and choose to convert it. Named ranges and the INDIRECT-based validation come across with the file.
