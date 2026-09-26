---
title: How to Add a Dropdown List for Status Columns in Excel and Google Sheets
labels: data-validation, spreadsheet-basics
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A small business desk with a laptop showing a spreadsheet open next to a coffee cup, screen glare hides all text and labels
image_alt: Laptop on a desk displaying a spreadsheet with an unreadable status column
search_description: Add a dropdown list to a status column in Excel or Google Sheets using data validation, with steps, sample data, and common fixes.
threads: Typing "Paid," "paid," and "PD" into the same column? Add a dropdown so every entry matches exactly.\nExcel: Data > Data Validation > List.\nSheets: Data > Data validation > Dropdown.\nTakes two minutes and stops the typos for good.
---
Add a dropdown list by selecting the column, then setting data validation to a list of allowed values such as Paid, Unpaid, Overdue. This stops typos like "paid" vs "Paid" and keeps every entry consistent.

> Works in: Excel (Microsoft 365 and Excel 2021), Google Sheets. Steps below follow the official Microsoft and Google help pages; not hands-on tested.

## Steps

### In Excel

1. Select the column or range you want to restrict, for example `C2:C200`.
2. Go to **Data > Data Validation**.
3. Under **Allow**, choose **List**.
4. In the **Source** box, type your options separated by commas, for example `Paid,Unpaid,Overdue`.
5. Check **In-cell dropdown** if it isn't already checked.
6. Click **OK**.

See Microsoft's guide: [Create a drop-down list](https://support.microsoft.com/en-us/office/create-a-drop-down-list-7693307a-59ef-400a-b769-c5402dc57b0c).

### In Google Sheets

1. Select the column or range, for example `C2:C200`.
2. Go to **Data > Data validation**.
3. If a rule editor doesn't open automatically, click **Add rule** (this may appear as **+ Add a rule**, depending on your version).
4. Under **Criteria**, choose **Dropdown**.
5. Type your options, one per line: Paid, Unpaid, Overdue.
6. Click **Done**.

See Google's guide: [Add data validation to cells](https://support.google.com/docs/answer/186103).

## Example

A freelancer tracking invoice status might set up a column like this. Sample data is fictional.

| Client | Invoice # | Status |
|---|---|---|
| Maple Street Cafe | 1042 | Paid |
| Riverside Design Co. | 1043 | Unpaid |
| North Shore Pottery | 1044 | Overdue |

Instead of typing "Paid" one time and "paid" the next, the dropdown forces every cell to match one of the three options exactly. That matters if you later run a `COUNTIF` or filter by status.

## Troubleshooting

### The dropdown arrow doesn't show up
In Excel, check that **In-cell dropdown** is still checked in the Data Validation dialog. In Google Sheets, confirm you chose **Dropdown** and not **Dropdown (from a range)** by mistake.

### Old entries typed before the dropdown don't match
Data validation only checks new entries by default; it doesn't fix text already in the cells. Retype or reselect those cells from the dropdown to bring them in line.

### Pasting data breaks the dropdown
Pasting values over a validated cell can overwrite the rule in both apps. Reapply data validation to the range after a bulk paste, or paste values only (**Ctrl+Shift+V** in Sheets, **Paste Special > Values** in Excel) into a separate area first.

### You want the list to pull from another sheet instead of typing it
In Excel, set **Source** to a range reference like `=Lists!$A$1:$A$3`. In Google Sheets, choose **Dropdown (from a range)** under Criteria and select the range.

### Excel shows an error when someone types something not on the list
By default Excel blocks the entry and shows a warning. To allow it but flag it, open Data Validation, go to the **Error Alert** tab, and change the style to **Warning** or **Information**.

## Copy-paste setup

Use these exact settings to recreate the status dropdown from scratch.

**Column header:** Status

**Excel data validation:**
- Allow: List
- Source: `Paid,Unpaid,Overdue`
- In-cell dropdown: checked

**Google Sheets data validation:**
- Criteria: Dropdown
- Options: Paid / Unpaid / Overdue (one per line)
- Show dropdown list in cell: checked

Adjust the option text to match your own workflow, for example Not Started, In Progress, Done for a project tracker, or In Stock, Low, Out of Stock for inventory.
