---
title: How to Flip Last, First Names to First Last in Excel and Google Sheets
labels: excel-formulas, text-functions, small-business
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 8 workbooks and 1,472 comparisons with 0 mismatches. Checked: names with extra spaces, no space after the comma, an apostrophe, multi-word last names, a middle initial, no comma, two commas, nothing before or after the comma, blank and space-only cells, and 60 rows. Excel and Google Sheets were not used.
search_description: Flip Last, First names to First Last with FIND, MID and LEFT in Excel and Google Sheets. Free template that flags names with no comma or two commas.
image_prompts: A stack of paper name tags with blank white labels and a black marker next to a small cardboard box of index cards on a folding table, soft daylight, no screens, no readable text, no logos
image_alt: Stack of blank name tags and a marker beside a box of index cards
threads: Customer list came back as "Alvarez, Maria"?\nFLIP IT: =TRIM(MID(A5,FIND(",",A5)+1,200))&" "&TRIM(LEFT(A5,FIND(",",A5)-1)). FIND locates the comma, MID takes what follows, LEFT takes what comes before.\n"Garcia Lopez, Ana" stays Ana Garcia Lopez. A name with two commas shows Check.
---
Turn `Alvarez, Maria` into `Maria Alvarez` with `=TRIM(MID(A5,FIND(",",A5)+1,200))&" "&TRIM(LEFT(A5,FIND(",",A5)-1))`. FIND locates the comma, MID takes the first name after it, LEFT takes the last name before it, and TRIM removes stray spaces.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every result below was recalculated and compared with a Python model. FIND, MID, LEFT, TRIM and LEN exist in current Excel and Google Sheets, but neither app was opened for this post.

![Typed names in Last, First order beside First, Last and First Last columns](images/2026-11-03-team-04-flip-last-first-names-to-first-last-find-mid-excel-google-sheets/flip-last-first-names-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template: the typed names in **column A**, starting in row 5. Format column A as Plain text first if you are typing names by hand.

1. In **A5**, type or paste a name as Last, First, such as `Alvarez, Maria`.
2. In **C5**, get the last name: `=TRIM(LEFT(A5,FIND(",",A5)-1))`. FIND returns the position of the comma, and LEFT keeps everything before it.
3. In **B5**, get the first name: `=TRIM(MID(A5,FIND(",",A5)+1,200))`. MID starts one character after the comma and keeps up to 200 characters.
4. In **D5**, join them: `=B5&" "&C5`.
5. Fill **B5:D5** down.

The 200 in MID is just a length longer than any name. Everything after the comma is the first name, so `Ruth Ann` stays together and so does a middle initial: `Brown, Pat A.` becomes `Pat A. Brown`.

Everything before the comma is the last name, so `Garcia Lopez, Ana` becomes `Ana Garcia Lopez` and `de la Cruz, Luis` becomes `Luis de la Cruz`.

To skip the helper columns, use the single formula at the top of this post.

To keep the result without the formulas, copy column D, then use Paste special and choose Values only.

Help pages: Microsoft's [MID function](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028) and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/02/2026; the pages were not read for this post.

## Example

Sample data is fictional: names from a small shop's customer export.

| Typed name | First | Last | First Last |
|---|---|---|---|
| Alvarez, Maria | Maria | Alvarez | Maria Alvarez |
| O'Brien, Kevin | Kevin | O'Brien | Kevin O'Brien |
| Garcia Lopez, Ana | Ana | Garcia Lopez | Ana Garcia Lopez |
| de la Cruz, Luis | Luis | de la Cruz | Luis de la Cruz |
| Nguyen,  Linh | Linh | Nguyen | Linh Nguyen |
| Brown, Pat A. | Pat A. | Brown | Pat A. Brown |
| Patel,Raj | Raj | Patel | Raj Patel |

The template turns 12 sample names into 12 flipped names, and the check cell says OK. This is text handling only. It does not know whether a name is a person or a company.

## Troubleshooting

### The cell shows #VALUE!
FIND did not find a comma. A name already written First Last, such as `Maria Alvarez`, has none. The plain formula returns #VALUE! for it. The template wraps the formulas so a name with no comma shows Check instead of an error.

### A name with Jr. or two commas looks wrong
`Smith, Jr., John` has two commas, so the first comma splits it into `Jr., John` as the first name. The template counts the commas and shows Check for any name without exactly one. Fix those by hand.

### The name has nothing on one side
`Kim,` has a last name and no first name. The template shows Check. In the test, a name with nothing before the comma (`,Lee`) did too.

### Extra spaces are still there
TRIM removes spaces at the ends and extra spaces in between. In the test, `  Patel,Raj  ` became `Raj Patel` and `Nguyen,  Linh` became `Linh Nguyen`. Spaces other than the regular space, such as a non-breaking space copied from a web page, were not tested.

### The capitalization is not what you want
The formulas keep the capitalization you typed. In the test, `de la Cruz, Luis` stayed `Luis de la Cruz`. Fix the case in the typed name, or wrap the result in `PROPER`, which was not tested here and capitalizes every word.

## Template

[Download the .xlsx](templates/tidy-tabs-flip-last-first-names.xlsx)

The file has two tabs. **Names** has room for 60 names in yellow cells (12 fictional ones filled in), the First, Last and First Last columns, and a summary of names typed, names flipped, rows showing Check and a **CHECK** cell, **G7**. **How to use** has notes on each formula.

**CHECK** says OK when every typed name was flipped. Fix every Check row before you paste the result into a mailing list.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
