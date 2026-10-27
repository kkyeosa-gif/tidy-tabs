---
threads_url: https://www.threads.com/@tin_ylab/post/DfAWXrwEb-O
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-flag-email-addresses-that-look.html
title: How to Flag Email Addresses That Look Wrong in a List in Excel and Google Sheets
labels: excel-text-functions, small-business, data-cleanup
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was built with a script, recalculated there and compared with a Python model of the same rules, including edited copies (a space before or after an address, two @ signs, a double dot, a final dot, uppercase, a plus sign, accented letters, a blank row in the middle, an empty list). The page was exported to PDF. Excel and Google Sheets were not opened, and the formula does not check that any address exists or receives mail.
search_description: Flag email addresses that look wrong in Excel and Google Sheets with FIND, SUBSTITUTE and LEN. Free template with a status and reason column and a count of rows to fix.
image_prompts: A wooden desk with a stack of paper index cards held by a brass binder clip, a red pencil with an eraser, and a small ink stamp pad, soft window light, no screens, no readable text, no logos
image_alt: Index cards bound with a brass clip beside a red pencil and an ink pad
threads: Pasted an email list and some rows look off?\nOne test counts @ signs: =IF(LEN(B2)-LEN(SUBSTITUTE(B2,"@",""))<>1,"Check","OK"). The template adds a space, dot and double dot test, plus a Reason column.\nIn the sample, 8 of 30 rows get flagged. It checks shape only, not whether anyone reads the address.
---
Flag an address that cannot be right with a Status formula built from `FIND`, `SUBSTITUTE` and `LEN`. One of its tests, `=IF(LEN(B2)-LEN(SUBSTITUTE(B2,"@",""))<>1,"Check","OK")`, marks any address with no @ or more than one. The template adds tests for spaces, text before the @, a dot after it, a double dot and a final dot.

This checks the shape of an address only. A row can pass every test and still not exist or not be read by anyone.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built, recalculated and exported to PDF. The Excel and Google Sheets steps use menu paths from the Microsoft and Google help pages. Neither app was opened for this post.

![Email list with a Status and Reason column and eight red Check rows beside 22 OK rows](images/2026-10-26-team-03-flag-invalid-email-addresses-find-substitute-len-excel-google-sheets/flag-bad-email-addresses-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

The idea: `SUBSTITUTE(B2,"@","")` removes every @, so the drop in `LEN` is the number of @ signs. A good address has exactly one.

1. In **A1:D1**, type the headings Name, Email, Status and Reason.
2. Paste your names in column **A** and your addresses in column **B**, starting in row 2.
3. In **C2**, type the Status formula below and fill it down.
4. In **D2**, type the Reason formula below and fill it down.
5. Fix the rows marked Check, then copy the Email column to wherever the list is going.

The Status formula in **C2** is one nested `IF`:

```
=IF(AND(A2="",B2=""),"",IF(B2="","Check",IF(ISNUMBER(FIND(" ",B2)),"Check",IF(LEN(B2)-LEN(SUBSTITUTE(B2,"@",""))<>1,"Check",IF(FIND("@",B2)<2,"Check",IF(ISERROR(FIND(".",B2,FIND("@",B2)+2)),"Check",IF(OR(RIGHT(B2,1)=".",ISNUMBER(FIND("..",B2))),"Check","OK")))))))
```

It stops at the first test that fails, in this order:

1. The row is completely empty: the cell stays blank.
2. The email is empty but the row has a name: Check.
3. The address contains a space: Check.
4. The number of @ signs is not 1: Check. This is the test from the first paragraph.
5. `FIND("@",B2)<2`, so nothing comes before the @: Check.
6. No dot after the @. `FIND(".",B2,FIND("@",B2)+2)` starts looking two characters past the @, so a dot straight after it, as in a@.com, does not count: Check.
7. The address ends with a dot or contains two dots in a row: Check.
8. Anything that passes all of these: OK.

The Reason formula in **D2** runs the same tests in the same order and returns words instead, such as Has a space or Needs exactly one @. It stays blank for an OK row.

### In Excel

1. Type the formulas in **C2** and **D2**, then double-click the fill handle to fill down.
2. To color the Check rows, select **C2:D200** and go to **Home > Conditional Formatting > New Rule**. Use the formula `=$C2="Check"`.

Microsoft's pages: [FIND and FINDB functions](https://support.microsoft.com/en-us/office/find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628) and [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332). Those are help pages, not our test.

### In Google Sheets

1. Type the formulas in **C2** and **D2**, then drag the fill handle down.
2. To color the Check rows, select **C2:D200** and go to **Format > Conditional formatting**. Choose "Custom formula is" and enter `=$C2="Check"`.

Google's page covers the rule editor: [Use conditional formatting rules in Google Sheets](https://support.google.com/docs/answer/78413).

### What the template adds

The built template has these formulas in rows 2 to 200, plus a summary in **F1:G4** with labels in F and values in G:

- **G1** counts rows checked.
- **G2** counts OK rows and **G3** counts Check rows.
- **G4** says OK when rows checked equals OK plus Check, every Check has a Reason, and nothing is typed below row 200. Otherwise it says CHECK.

Check rows turn red in **C:D** and OK turns green in **C**.

## Example

Sample data is fictional: 30 made-up names and addresses, with 8 rows broken on purpose.

| Name | Email | Status | Reason |
|---|---|---|---|
| Frank Moore | frank.moore.shop.com | Check | Needs exactly one @ |
| Hank Diaz | hank@@pinehill.com | Check | Needs exactly one @ |
| Jon Park | jon park@oaklane.com | Check | Has a space |
| Kara Webb | kara+orders@mail.com | OK | |
| Mia Scott | @birchbakery.com | Check | Nothing before the @ |
| Paul Young | paul@shop | Check | No dot after the @ |
| Sam Cole | sam@florist..com | Check | Ends with a dot or has two dots in a row |
| Uma Shah | uma@bakery. | Check | Ends with a dot or has two dots in a row |
| Xena Roy | | Check | Email is empty |

The template reports 30 rows checked, 22 OK and 8 Check, and G4 says OK. A Python model of the same rules gave the same counts.

## Troubleshooting

### An address that looks fine is flagged
Look for a space typed before or after it. The space test looks at the whole cell, so " ana@shop.com" with a leading space and "ana@shop.com " with a trailing space both come back as Check with the reason Has a space. Delete the space.

### A bad address shows OK
The tests only cover the shapes listed above. An address like a.@b.com passes, and so does one with accented letters (josé@café.com) or a tab character inside it. Only the space character is tested. The template does not flag uppercase, a plus sign or a long address either.

### The Status cell is blank
A row with no name and no email stays blank on purpose, and so does a row in the middle of your list that you cleared. A row with a name and no email shows Check with the reason Email is empty.

### The CHECK cell says CHECK
G4 says CHECK when something is typed below row 200, because the formulas stop there. Copy the formulas further down. It also says CHECK if a Check row has no Reason, which can happen if you overwrite a formula in column D.

## Template

[Download the .xlsx](templates/tidy-tabs-flag-bad-emails.xlsx)

The file has two tabs. **Emails** holds the yellow inputs, the name in **A** and the email in **B**, for rows 2 to 200. Gray cells are formulas: Status in **C**, Reason in **D**, and the summary with its check in **F1:G4**. **How to use** has short notes, including the limits of the tests.

The names and addresses are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2. It does not verify that an address exists or receives mail, and it gives no advice on sending to a list. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
