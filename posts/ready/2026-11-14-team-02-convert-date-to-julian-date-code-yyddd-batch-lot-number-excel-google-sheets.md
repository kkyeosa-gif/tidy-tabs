---
title: How to Convert a Date to a Julian Date Code for Lot Numbers in Excel and Google Sheets
labels: excel-formulas, small-business, dates
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 16 workbooks and 28,896 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Turn 11/14/2026 into the Julian date code 26318 for batch lot numbers, and read a YYDDD code back into a date. Free Excel template.
image_prompts: Freshly filled jam jars cooling upside down on a wire rack in a home kitchen, a small handheld roller coder for marking lot codes resting beside them and an open carton of spare jar lids, any marks too small to read, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Filled jam jars cooling upside down on a wire rack beside a handheld coder
threads: Turn a batch date into a Julian lot code with =RIGHT(YEAR(A4),2)&TEXT(A4-DATE(YEAR(A4),1,0),"000").\n11/14/2026 becomes 26318, day 318 of 2026.\nSkip the TEXT padding and 02/01/2005 comes out as 0532 instead of 05032.
---
Turn a date in A4 into a five-digit Julian code (YYDDD) with `=RIGHT(YEAR(A4),2)&TEXT(A4-DATE(YEAR(A4),1,0),"000")`, so 11/14/2026 becomes 26318, and read a code in A4 back into a date with `=DATE(2000+LEFT(A4,2),1,RIGHT(A4,3))`. The download also catches codes that lost their leading zero (05032 typed as 5032) and days that do not exist (26366).

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only, where the template was recalculated and compared with a Python model (16 workbooks, 28,896 comparisons, 0 mismatches). Excel and Google Sheets have the same functions, but neither app was opened for this post.

![Batch dates turned into Julian lot codes, and printed codes read back into dates](images/2026-11-14-team-02-convert-date-to-julian-date-code-yyddd-batch-lot-number-excel-google-sheets/julian-date-lot-codes-template.png) *Left, ten batches get a YYDDD code and lot code (a date typed as text shows Check); right, nine printed codes read back into dates, three of them Check. PDF export from LibreOffice Calc 24.2.*

## Steps

A Julian date code is the two-digit year followed by the day of the year from 001 to 366. 26318 means the 318th day of 2026. The formulas are the same in Excel and Google Sheets.

This is not the astronomical Julian Day number (a seven-digit count such as 2461358). It is also only the date part of a lot code: the post and the file do not cover any lot code rule, FDA or other.

### Make a code from a date

1. Type the batch dates in column **A**, starting in **A4**, as real dates such as 11/14/2026.
2. In **D4**, find the day of the year: `=A4-DATE(YEAR(A4),1,0)`. DATE with day 0 is the last day of the year before, so 01/01/2026 gives 1. Format **D4** as a plain number, not a date.
3. In **E4**, build the code: `=RIGHT(YEAR(A4),2)&TEXT(D4,"000")`. TEXT pads the day to three digits, so day 32 becomes 032.
4. To make a full lot code, put a product prefix in **B4** and a batch number in **C4**, then type `=B4&"-"&E4&"-"&C4` in **F4**. CND, 26318 and 1 give CND-26318-1.
5. Fill **D4:F4** down to the last batch.

`TEXT(A4,"yy")` also gave 26 in LibreOffice, but format codes can change with the language setting, so the year comes from RIGHT(YEAR()).

### Read a code back into a date

1. Paste the codes as printed into column **A** of a second sheet, starting in **A4**.
2. In **E4**, type `=DATE(2000+LEFT(A4,2),1,RIGHT(A4,3))`. DATE counts day 318 of January forward into November, which gives 11/14/2026.
3. Format **E4** as mm/dd/yyyy and fill it down.

This short formula is for clean five-digit codes. Troubleshooting covers the rest.

Help pages: Microsoft's [Insert Julian dates](https://support.microsoft.com/en-us/office/insert-julian-dates-functions-0c7fa6aa-daff-402e-9990-93a5b76ba018) and [TEXT function](https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c), and Google's [DATE](https://support.google.com/docs/answer/3092969). They describe the functions, not our test.

### Leap years

Every day after February 28 shifts by one in a leap year. 03/01/2027 is day 60 (27060), while 02/29/2028 is day 60 (28060) and 03/01/2028 is day 61 (28061). 12/31/2028 is day 366. The formulas handle this; a printed day chart from another year does not.

## Example

Sample data is fictional: ten candle, soap and jam batches.

| Made on | Prefix | Batch | Day | Code | Lot code | Status |
|---|---|---|---|---|---|---|
| 11/14/2026 | CND | 1 | 318 | 26318 | CND-26318-1 | OK |
| 11/14/2026 | CND | 2 | 318 | 26318 | CND-26318-2 | OK |
| 01/01/2026 | SOP | 1 | 1 | 26001 | SOP-26001-1 | OK |
| 12/31/2026 | JAM | 1 | 365 | 26365 | JAM-26365-1 | OK |
| 02/29/2028 | CND | 1 | 60 | 28060 | CND-28060-1 | OK |
| 03/01/2027 | SOP | 1 | 60 | 27060 | SOP-27060-1 | OK |
| 12/31/2028 | JAM | 1 | 366 | 28366 | JAM-28366-1 | OK |
| 02/01/2005 | SOP | 1 | 32 | 05032 | SOP-05032-1 | OK |
| 11/13/2026 | JAM | 2 | 317 | 26317 | JAM-26317-2 | OK |
| "11/15/2026" (text) | CND | 1 | | Check | | Check |

The Make codes summary shows 10 rows listed, 9 codes made and 1 Check, the date typed as text.

Reading codes back:

| Code as printed | Date | Status |
|---|---|---|
| "26318" | 11/14/2026 | OK |
| " 26318 " | 11/14/2026 | OK |
| "2026032" | 02/01/2026 | OK |
| 28366 (number) | 12/31/2028 | OK |
| 5032 (number, zero lost) | 02/01/2005 | OK |
| "05032" | 02/01/2005 | OK |
| "26366" | Check | Check |
| "26000" | Check | Check |
| "26A18" | Check | Check |

The Read codes summary shows 9 codes listed, 6 read and 3 Check. 2026 has 365 days, so 26366 does not exist, and day 000 never does.

## Troubleshooting

These cases were reproduced in LibreOffice 24.2.7.2.

### A code came out four digits, such as 0532
The day was not padded. Without `TEXT(...,"000")`, 02/01/2005 gave 0532. Wrap the day in TEXT as in step 3.

### An old code reads as the year 2050
The code was typed or imported as a number, so 05032 became 5032 and the short formula gave 02/01/2050. The template puts the zero back for numbers. A four-digit code stored as text shows Check, so type it again as a plain number or as "05032" text.

### A seven-digit code reads as the wrong year
Some labels print YYYYDDD. The short formula read "2026032" as 02/01/2020. The template reads seven-digit codes as a four-digit year plus the day, which gives 02/01/2026.

### A code with spaces gives a date in 2002
" 26318 " gave 01/18/2002 in the short formula. The template trims spaces before it reads the code.

### An impossible code still shows a date
DATE shows no error: 26366 became 01/01/2027 and 26000 became 12/31/2025. The template checks that the day is between 1 and the year's last day (365 or 366) and shows Check otherwise.

## Template

[Download the .xlsx](templates/tidy-tabs-julian-date-lot-codes.xlsx)

The file has three tabs. **Make codes** has inputs in **A4:C203** (Made on, Product prefix, Batch of the day) and formula columns **D:G** (Day of year, Julian code, Lot code, Status), with a summary in **I3:J5**. A blank batch number gives a lot code without the suffix, such as CND-26318. Codes are made for 2000 to 2099 only, because YYDDD cannot tell 1926 from 2026.

**Read codes** takes codes as printed in **A4:A203** and fills Cleaned, Year, Day, Date and Status in **B:F**, with a summary in **H3:I5**. It reads five-digit codes and seven-digit codes from 1900 on. **How to use** explains each column.

Yellow cells are inputs and gray cells are formulas. Status is red for Check. Both tabs print on US Letter. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
