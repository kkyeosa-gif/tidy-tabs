---
title: How to Compare Vendors With a Weighted Score in Excel and Google Sheets
labels: excel-formulas, small-business, vendors
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including 25 edited variants (weights that add to 100 or 25, all weights 0, a negative, blank or text weight, one weight 0, ties of identical and of different rows, a blank or text score, scores of 0, 11 and 7.5, an option name with no scores, cleared rows). No mismatches in the final run. Excel and Google Sheets were not opened.
search_description: Compare vendors with a weighted score in Excel and Google Sheets using SUMPRODUCT and RANK. Free decision matrix template with sample scores.
image_prompts: Three folded canvas tote bags in different colors, a fan of fabric swatches and a wooden ruler on a screen printing workshop table, soft window light, no screens, no readable text, no logos
image_alt: Three folded canvas tote bags, fabric swatches and a wooden ruler on a workshop table
threads: Five print shops, four things that matter, and no clear winner?\nMultiply each score by its weight with SUMPRODUCT, then RANK the totals.\nIn the sample, Harbor Print Co. wins at 7.60 even though it is not the cheapest. Change the weights and the order changes.
---
Turn several scores into one number with `=SUMPRODUCT($B$3:$E$3,B5:E5)/SUM($B$3:$E$3)`, where B3:E3 are your weights for each criterion and B5:E5 are one option's scores from 1 to 10. Copy it down, then rank the results with `=RANK(F5,$F$5:$F$9)`.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Weighted score sheet comparing five print shops with weights, scores, ranks and top row highlighted](images/2026-10-24-team-05-compare-options-weighted-score-sumproduct-rank-excel-google-sheets/weighted-score-vendor-comparison-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

The idea: multiply each score by the weight of its criterion, add the products, then divide by the total of the weights. Dividing keeps the result on the 1 to 10 scale, whatever your weights add to.

1. In **B4:E4**, type your criteria, such as Price, Turnaround, Quality and Support.
2. In **B3:E3**, type a weight for each criterion. A bigger number means that criterion counts more. For example, 0.40, 0.20, 0.30 and 0.10.
3. In **A5:A9**, type the names of the options you are comparing.
4. In **B5:E9**, type a score from 1 to 10 for each option on each criterion. These are your own opinions, so use the same scale for every row.
5. In **F5**, type `=SUMPRODUCT($B$3:$E$3,B5:E5)/SUM($B$3:$E$3)` and fill it down to **F9**. The dollar signs lock the weights row, so every option is multiplied by the same weights.
6. In **G5**, type `=RANK(F5,$F$5:$F$9)` and fill it down to **G9**.

How the formulas read:

- `SUMPRODUCT($B$3:$E$3,B5:E5)` multiplies the first weight by the first score, the second by the second, and so on, then adds them up.
- `/SUM($B$3:$E$3)` divides by the total of the weights.
- `RANK(F5,$F$5:$F$9)` gives 1 to the highest weighted score. If two options tie, both get the same rank and the next rank is skipped.

### What the template adds

The built formulas are longer than the bare ones. **F5** wraps the division in `ROUND(...,10)` and two checks. It shows "Check weights" when a weight is blank, text or negative, or when the weights add to 0. It shows "Check scores" when an option has a blank or text score.

The `ROUND` matters for ties. In our LibreOffice runs, two options whose different scores add up to the same weighted total got ranks 1 and 2 because of tiny floating point differences. With `ROUND(...,10)` they tie, as they do on paper. **G5** is the RANK formula inside an `ISNUMBER` test, so an option with a message in F is left out of the ranking.

**J1** names the option with the highest weighted score, **J2** shows that score and **J4** counts how many options share rank 1. **J5** is a check cell.

### In Excel

The formulas work as typed. SUMPRODUCT, SUM and RANK are old functions, and no newer function is needed. We did not run them in Excel.

### In Google Sheets

The same formulas should work as typed. We did not run them in Google Sheets.

## Example

The sample data is fictional: five print shops compared for tote bag orders. The weights are Price 0.40, Turnaround 0.20, Quality 0.30 and Support 0.10, and they add to 1.00.

| Option | Price | Turnaround | Quality | Support | Weighted score | Rank |
|---|---|---|---|---|---|---|
| Harbor Print Co. | 8 | 6 | 9 | 5 | 7.60 | 1 |
| Bluebird Printing | 7 | 9 | 7 | 8 | 7.50 | 2 |
| Pine Street Press | 9 | 5 | 6 | 6 | 7.00 | 4 |
| Quick Tote Shop | 6 | 10 | 5 | 7 | 6.60 | 5 |
| Maple Screen Prints | 7 | 7 | 8 | 8 | 7.40 | 3 |

Check the first row by hand: 0.4 x 8 + 0.2 x 6 + 0.3 x 9 + 0.1 x 5 = 3.2 + 1.2 + 2.7 + 0.5 = 7.6.

Change the weights and the order can change. With weights of 10, 5, 5 and 5, which add to 25, Bluebird Printing moves to rank 1 at 7.60 and Harbor Print Co. falls to 7.20. With weights of 40, 20, 30 and 10, which add to 100, the scores and ranks match the 0.40 set.

The highest weighted score only summarizes the numbers you typed. It is not the right choice, and the sheet recommends no one.

## Troubleshooting

### Two options show the same rank
Their weighted scores are equal, so both get the same rank and the next rank is skipped. Two options at rank 1 mean there is no rank 2. J4 shows how many share rank 1, and J1 names the first one listed.

### F shows Check weights
A cell in B3:E3 is blank, text or negative, or all four weights are 0. Retype the weights as plain numbers. A single weight of 0 is fine: that criterion simply does not count.

### F shows Check scores
One of the four scores for that option is blank or text, such as "n/a". The option is left out of the ranking and the other options are ranked without it. Type a number in every score cell.

### J5 says CHECK but every option has a score
A score outside 1 to 10 still gets a weighted score, but J5 flags it. A score of 11 or 0 was computed in our runs, not blocked. Scores typed in a row with no option name also trigger it.

## Template

[Download the .xlsx](templates/tidy-tabs-weighted-score-matrix.xlsx)

The file has two tabs. **Score** holds the yellow inputs: the weights in **B3:E3**, the criteria names in **B4:E4**, the option names in **A5:A9** and the scores in **B5:E9**. Gray cells are formulas: the weight total in **F3**, the weighted scores in **F5:F9**, the ranks in **G5:G9**, and a summary with the best option, its score, the count of options ranked and the check in **I1:J5**. The rank 1 row turns green and J5 turns red when it is not OK. **How to use** has short notes.

The shops and scores are fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2, and the green and red formatting appeared in its PDF export but was not seen in a spreadsheet window. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
