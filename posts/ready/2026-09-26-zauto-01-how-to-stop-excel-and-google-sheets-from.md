---
title: How to Stop Excel and Google Sheets From Turning Phone Numbers Into Scientific Notation
labels: data-entry, data-cleanup
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A small business desk with a printed client contact sheet and a laptop showing a spreadsheet, all screen text and labels unreadable, warm natural light
image_alt: Desk with laptop and printed client list showing rows of contact phone numbers
search_description: Fix phone numbers that turn into 1.23E+11 or lose the leading +1 in Excel and Google Sheets, with the exact cell format to use before you type or paste.
threads: Typed a client's phone number and got 1.23457E+11 instead? Excel and Sheets treat long digit strings as numbers unless you tell them not to.\nFormat the column as Plain text (or Text) before you type or paste, then re-enter the numbers. Fixing it after the fact won't bring back a lost +1 or leading 0.
---
Format the column as **Text** (Excel) or **Plain text** (Google Sheets) before you type or paste phone numbers. Once a phone number has already turned into scientific notation like 1.23457E+11, changing the format back won't restore the missing digits, so you have to re-enter or re-import the data after fixing the format.

> Works in: Excel (Microsoft 365, Excel 2019+) and Google Sheets, current version. Steps come from Microsoft and Google help pages, not hands-on testing. Menu paths and the Ctrl+1 shortcut below match Excel for Windows; Excel for Mac uses Cmd+1 and some menu wording may differ slightly.

## Steps

### In Excel
1. Select the column or range where you'll enter phone numbers.
2. Go to the **Home** tab, open the number format dropdown in the **Number** group (it shows **General** by default), and choose **Text**, or press **Ctrl+1** (**Cmd+1** on Mac), open **Format Cells**, pick **Text**, and click **OK**.
3. Type or paste the phone numbers. Include the plus sign for country code if you use one, like +1 617-555-0148.
4. If numbers are already broken, clear the cells, reformat them as **Text**, then retype or re-paste the original data.

### In Google Sheets
1. Select the column or range.
2. Go to **Format > Number > Plain text**.
3. Type or paste the phone numbers.
4. If a paste came in as numbers already, undo the paste, set the format to **Plain text** first, then paste again.

## Example

| Client | Entered as | Shows as (default number format) | Shows as (Text/Plain text format) |
|---|---|---|---|
| Riverside Cafe | 6175550148 | 6175550148 | 6175550148 |
| Union Sq. Tailors | 16175550148 | Scientific notation, such as 1.61756E+10 (exact digits shown depend on column width) | 16175550148 |
| North End Bakery | +16175550148 | May keep the + or may drop it, depending on the app and cell format | +16175550148 |

Sample names and numbers above are fictional, made up for this example.

Once a number turns into scientific notation, the original digits are gone. You can't just switch the format back and get 6175550148 from a value like 6.17555E+9, so always set the format before entering data, not after.

## Troubleshooting

### Long numbers show as 1.23E+11 instead of the real digits
This happens when the cell format is General or Number and the value has too many digits to display normally, which often starts around 11-12 digits, or when the column is too narrow to show the full number. Reformat the column as Text/Plain text, then retype the numbers.

### A leading + or 0 disappears after pasting
Both apps can strip a leading + or 0 if the cell isn't set to text first. Undo the paste, format the destination cells as Text or Plain text, and paste again.

### Numbers pasted from a CSV came in as scientific notation
CSV files store phone numbers as plain digits with no format info, so Excel and Sheets guess a number format on import. In Excel, use the **Get Data** / **From Text/CSV** import option and set the phone number column's data type to Text during import, instead of just double-clicking to open the file (exact wording and availability can vary between Excel for Windows and Excel for Mac). In Google Sheets, import with **File > Import > Upload**, choose an option that replaces or adds to the existing sheet, and reformat the column to Plain text right after.

### Numbers look fine in the cell but formulas or exports break
A phone number stored as text will not add, subtract, or sort as a number, which is normal and expected. If you need to sort by area code, do it as a text sort, not a numeric one.

### Some phone numbers show fine, others don't, in the same column
This usually means the column format was changed partway through data entry. Select the whole column, set it to Text/Plain text, and manually check any rows entered before the format change.

## Copy-paste setup

Use this header row and format for a simple client contact list:

| Client Name | Phone | Email | Notes |
|---|---|---|---|

- Set the **Phone** column format to Text (Excel) or Plain text (Google Sheets) before entering any data.
- Enter numbers in a consistent style, for example `617-555-0148` or `+1 617-555-0148`.
- If you need a formula to check for exactly 10 digits (ignoring dashes and spaces), use `=LEN(SUBSTITUTE(SUBSTITUTE(A2,"-","")," ",""))=10` in a helper column, in both Excel and Google Sheets.

Related help pages:
- [Format numbers as text](https://support.microsoft.com/en-us/office/format-numbers-as-text-419eb164-a71b-4c56-a5cb-40d61c93639b)
- [Create a custom number format](https://support.microsoft.com/en-us/office/create-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4)
- [Format numbers in a spreadsheet](https://support.google.com/docs/answer/78591)
