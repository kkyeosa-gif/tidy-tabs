---
title: How to Split a Full Name Into First and Last Name Columns in Excel and Google Sheets
labels: data-cleanup, excel-formulas
tested_in: Steps from Microsoft/Google help pages; not hands-on tested.
image_prompts: A printed customer sign-up sheet with handwritten full names next to a stack of blank name tag stickers and a pen on a desk, no readable text
image_alt: Printed sign-up sheet with full names beside blank sticker name tags and a pen
search_description: Split a Full Name column into separate First Name and Last Name columns in Excel and Google Sheets using formulas, Flash Fill, or Split text to columns.
threads: Customer list has one "Full Name" column and you need First and Last separate for a mail merge or CRM import.\nSkip retyping everything. Excel's Flash Fill or a simple LEFT/FIND formula splits hundreds of names in seconds.\nWatch out for suffixes like Jr. and two-word last names though, they need a manual fix after.
---

Split "First Last" into two columns with `=LEFT(A2,FIND(" ",A2)-1)` for the first name and `=TRIM(MID(A2,FIND(" ",A2)+1,LEN(A2)))` for the last name. This works for simple two-word names; names with a middle name, suffix, or two-word last name need a quick manual check afterward.

> Works in: Excel 365, Excel 2021, and Google Sheets. Steps use Text to Columns, Flash Fill, and the SPLIT function, all documented on Microsoft and Google's help pages. Flash Fill is a desktop feature; it may not behave the same way (or be available at all) in Excel for the web.

![Spreadsheet example with columns Full Name, First Name (formula result), Last Name (formula result), Needs manual](images/2026-09-27-zauto-03-how-to-split-a-full-name-into-first-and-/how-to-split-a-full-name-into-first-and-last-name-example.png) *The example from this post laid out in a spreadsheet. Rendered with LibreOffice Calc.*

## Steps

1. Add two new column headers to the right of your full name column: **First Name** and **Last Name**.
2. Decide whether your list has simple two-word names ("Maria Gonzalez") or messier ones (middle names, suffixes, hyphenated last names). Messier lists need the formula method below plus manual cleanup.

### In Excel

1. Click the cell under **First Name** and type `=LEFT(A2,FIND(" ",A2)-1)`, then press **Enter**.
2. Click the cell under **Last Name** and type `=TRIM(MID(A2,FIND(" ",A2)+1,LEN(A2)))`, then press **Enter**.
3. Select both formula cells and drag the fill handle down to cover the rest of your list.
4. For a no-formula option, type the first name in the first row by hand, then select the cell below it and press **Ctrl+E** (Windows) or use **Data > Flash Fill** from the ribbon (Mac and other versions) to trigger Flash Fill for the whole column. Repeat for last names.

### In Google Sheets

1. Click the cell under **First Name** and type `=SPLIT(A2," ")`, which splits the name by spaces into separate cells (two cells for a typical first-and-last-name entry).
2. If you'd rather keep everything in your **First Name** / **Last Name** columns, use `=REGEXEXTRACT(A2,"^(\S+)")` for the first name and `=REGEXEXTRACT(A2,"\S+$")` for the last name.
3. Alternatively, select the full name column, then click **Data > Split text to columns**. Sheets splits on the space delimiter by default and adds new columns to the right.
4. If Split text to columns overwrote data in an adjacent column, press **Ctrl+Z** to undo, insert two blank columns first, then repeat the split.

## Example

Sample data below is fictional.

| Full Name | First Name (formula result) | Last Name (formula result) | Needs manual fix? |
|---|---|---|---|
| Maria Gonzalez | Maria | Gonzalez | No |
| Tom O'Brien | Tom | O'Brien | No |
| Susan Lee Park | Susan | Lee Park | Maybe (two-word last name) |
| Robert Chen Jr. | Robert | Chen Jr. | Yes (move "Jr." out) |
| Dr. Alan Ruiz | Dr. | Alan Ruiz | Yes (title, not first name) |

## Troubleshooting

### First name column is blank
A leading space before the name pushes `FIND(" ",A2)` to position 1, so `LEFT` returns nothing. Use `=TRIM(A2)` in a helper column first, then run the split formulas on the trimmed version.

### Suffixes like Jr., Sr., or III land in the last name column
The formulas split on the first space only, so anything after the last name (including a suffix) stays attached. Sort the affected rows and fix them by hand, since suffixes are too inconsistent to catch reliably with one formula.

### Two-word last names get cut in half
Names like "Van Buren" or "De La Cruz" split after the first word, leaving "Van" as the last name. Filter your list for these cases and correct them manually rather than rewriting the formula.

### Split text to columns overwrote data in the next column
Google Sheets and Excel both write split results into whatever columns are immediately to the right, replacing existing content. Insert two blank columns before splitting, or copy the full name column to a new sheet first.

### Flash Fill guesses wrong on inconsistent names
Flash Fill in Excel learns from the pattern in your first few rows. If your list mixes formats (some with middle initials, some without), fill in 2-3 more example rows by hand so Excel has enough pattern to match.

## Copy-paste setup

Headers to type across row 1:

```
Full Name | First Name | Last Name
```

Excel formulas (row 2, then fill down):

```
First Name: =LEFT(A2,FIND(" ",A2)-1)
Last Name: =TRIM(MID(A2,FIND(" ",A2)+1,LEN(A2)))
```

Google Sheets formulas (row 2, then fill down):

```
First Name: =REGEXEXTRACT(A2,"^(\S+)")
Last Name: =REGEXEXTRACT(A2,"\S+$")
```

Helper formula to strip extra spaces before splitting:

```
=TRIM(A2)
```

Reference pages:
- [Split text into different columns with the Convert
