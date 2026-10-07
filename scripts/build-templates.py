#!/usr/bin/env python3
"""Builds the downloadable .xlsx/.csv templates in templates/.

Run from the repo root: python3 scripts/build-templates.py
Needs openpyxl (pip install openpyxl). Every template is plain formulas
(IF, SUMIFS, TODAY, TEXT) so it opens the same in Excel and after
File > Import in Google Sheets. Re-run after editing; the files are
committed so readers download exactly what was checked.
"""
import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path("templates")
USD = '"$"#,##0.00'
US_DATE = "mm/dd/yyyy"
HEADER_FILL = PatternFill("solid", fgColor="2E5E4E")
HEADER_FONT = Font(bold=True, color="FFFFFF")
ALERT_FILL = PatternFill("solid", fgColor="F8D7DA")
DONE_FILL = PatternFill("solid", fgColor="E2F0D9")
FORMULA_FILL = PatternFill("solid", fgColor="EDEDED")


def header(ws, names, widths):
    ws.append(names)
    for i, (cell, w) in enumerate(zip(ws[1], widths), start=1):
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"


def notes_sheet(wb, lines):
    ws = wb.create_sheet("How to use")
    ws.column_dimensions["A"].width = 100
    for line in lines:
        ws.append([line])
    ws["A1"].font = Font(bold=True, size=13)


def inventory():
    wb = Workbook()
    items = wb.active
    items.title = "Items"
    header(items, ["SKU", "Item", "Category", "Unit cost", "Sale price", "Starting qty",
                   "Qty in", "Qty out", "On hand", "Reorder at", "Status", "Stock value (cost)"],
           [10, 30, 14, 11, 11, 11, 9, 9, 9, 11, 13, 16])
    rows = [
        ("CND-LAV8", "Lavender soy candle, 8 oz", "Candles", 4.10, 18.00, 40, 10),
        ("CND-CED8", "Cedar soy candle, 8 oz", "Candles", 4.25, 18.00, 25, 10),
        ("SOP-OAT", "Oatmeal goat milk soap bar", "Soap", 1.85, 8.00, 60, 20),
        ("SOP-CHR", "Charcoal soap bar", "Soap", 2.05, 8.50, 30, 20),
        ("EAR-HOOP", "Brass hoop earrings", "Jewelry", 3.40, 24.00, 15, 5),
        ("CRD-BDAY", "Letterpress birthday card", "Cards", 0.95, 6.00, 80, 25),
        ("PKG-BOX6", "Kraft mailer box 6x6x4", "Supplies", 0.62, 0, 100, 40),
    ]
    for r, (sku, name, cat, cost, price, start, reorder) in enumerate(rows, start=2):
        items.append([sku, name, cat, cost, price, start,
                      f"=SUMIFS('Stock log'!$C:$C,'Stock log'!$B:$B,A{r})",
                      f"=SUMIFS('Stock log'!$D:$D,'Stock log'!$B:$B,A{r})",
                      f"=F{r}+G{r}-H{r}", reorder,
                      f'=IF(A{r}="","",IF(I{r}<=J{r},"REORDER","OK"))',
                      f"=I{r}*D{r}"])
        for col in "DEL":
            items[f"{col}{r}"].number_format = USD
        for col in "GHIKL":
            items[f"{col}{r}"].fill = FORMULA_FILL
    # Ranges run to row 500 (not just the sample rows) so a product added by
    # copying the last row down is counted, offered in the dropdown, and
    # highlighted without editing any formula. The total sits beside the
    # table, not under it, for the same reason.
    last = 500
    items.column_dimensions["N"].width = 18
    items.column_dimensions["O"].width = 14
    items["N1"], items["O1"] = "Total stock value", f"=SUM(L2:L{last})"
    items["N1"].font = Font(bold=True)
    items["O1"].number_format = USD
    items.conditional_formatting.add(f"K2:K{last}", FormulaRule(formula=[f'$K2="REORDER"'], fill=ALERT_FILL))

    log = wb.create_sheet("Stock log")
    header(log, ["Date", "SKU", "Qty in", "Qty out", "Note"], [12, 12, 9, 9, 40])
    entries = [
        (date(2026, 9, 1), "CND-LAV8", 24, 0, "Poured batch #14"),
        (date(2026, 9, 3), "CND-LAV8", 0, 12, "Etsy orders 9/1-9/3"),
        (date(2026, 9, 5), "SOP-OAT", 0, 18, "Saturday farmers market"),
        (date(2026, 9, 5), "CND-CED8", 0, 9, "Saturday farmers market"),
        (date(2026, 9, 8), "SOP-CHR", 0, 14, "Wholesale order, Main St. Gift Shop"),
        (date(2026, 9, 10), "EAR-HOOP", 0, 11, "Etsy orders"),
        (date(2026, 9, 12), "PKG-BOX6", 0, 64, "Shipped Etsy orders"),
        (date(2026, 9, 15), "CRD-BDAY", 0, 30, "Craft fair"),
    ]
    for e in entries:
        log.append(list(e))
        log.cell(row=log.max_row, column=1).number_format = US_DATE
    sku_list = DataValidation(type="list", formula1=f"=Items!$A$2:$A${last}", allow_blank=True)
    log.add_data_validation(sku_list)
    sku_list.add("B2:B1000")

    notes_sheet(wb, [
        "Tidy Tabs: Small Craft Business Inventory Tracker",
        "",
        "1. Items tab: one row per product or supply. Type your own SKU, cost, price, starting count, and reorder point.",
        "2. Stock log tab: add a row every time stock comes in (Qty in) or goes out (Qty out). Pick the SKU from the dropdown.",
        "3. On hand, Status, and Stock value update by themselves. Status turns red when On hand is at or below Reorder at.",
        "4. Don't type over the gray formula columns (Qty in, Qty out, On hand, Status, Stock value).",
        "5. To add a product, copy the last Items row down so the formulas come with it. Total stock value is in cell O1.",
        "",
        "Sample data is fictional. Delete the Stock log rows and change the Items rows to start fresh.",
        "Google Sheets: File > Import > Upload this .xlsx > Insert new sheet(s) or Replace spreadsheet.",
    ])
    wb.save(OUT / "tidy-tabs-craft-inventory-tracker.xlsx")


def projects():
    wb = Workbook()
    ws = wb.active
    ws.title = "Projects"
    header(ws, ["Client", "Project", "Start date", "Due date", "Status", "Billing", "Hours", "Rate",
                "Flat fee", "Amount", "Invoice sent", "Paid on", "Days left", "Payment"],
           [18, 28, 12, 12, 13, 10, 8, 9, 10, 12, 12, 12, 10, 12])
    rows = [
        ("Harbor Dental", "Spring newsletter design", date(2026, 9, 2), date(2026, 9, 19), "Done", "Flat", None, None, 650, date(2026, 9, 19), date(2026, 9, 24)),
        ("Pine & Co. Realty", "Listing photo edits (batch 3)", date(2026, 9, 8), date(2026, 9, 26), "In progress", "Hourly", 6.5, 55, None, None, None),
        ("Lakeside Yoga", "Website copy refresh", date(2026, 9, 15), date(2026, 10, 10), "In progress", "Hourly", 3, 60, None, None, None),
        ("Main St. Gift Shop", "Holiday flyer", date(2026, 9, 22), date(2026, 10, 3), "Not started", "Flat", None, None, 300, None, None),
        ("Harbor Dental", "Social media templates", date(2026, 8, 18), date(2026, 9, 5), "Done", "Flat", None, None, 480, date(2026, 9, 5), None),
    ]
    for r, (client, proj, start, due, status, billing, hours, rate, flat, sent, paid) in enumerate(rows, start=2):
        ws.append([client, proj, start, due, status, billing, hours, rate, flat,
                   f'=IF(F{r}="Hourly",G{r}*H{r},I{r})', sent, paid,
                   f'=IF(OR(D{r}="",E{r}="Done"),"",D{r}-TODAY())',
                   f'=IF(L{r}<>"","Paid",IF(K{r}="","Not invoiced",IF(TODAY()-K{r}>30,"Overdue","Waiting")))'])
        for col in "CDKL":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "HIJ":
            ws[f"{col}{r}"].number_format = USD
    last = 200
    status = DataValidation(type="list", formula1='"Not started,In progress,Waiting on client,Done"', allow_blank=True)
    billing = DataValidation(type="list", formula1='"Hourly,Flat"', allow_blank=True)
    ws.add_data_validation(status)
    ws.add_data_validation(billing)
    status.add(f"E2:E{last}")
    billing.add(f"F2:F{last}")
    ws.conditional_formatting.add(f"A2:N{last}", FormulaRule(formula=['$E2="Done"'], fill=DONE_FILL))
    ws.conditional_formatting.add(f"M2:M{last}", FormulaRule(formula=['AND(ISNUMBER($M2),$M2<0)'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"N2:N{last}", FormulaRule(formula=['$N2="Overdue"'], fill=ALERT_FILL))

    summary = wb.create_sheet("Summary")
    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 14
    for label, formula in [
        ("Open projects", '=COUNTIFS(Projects!A2:A200,"<>",Projects!E2:E200,"<>Done")'),
        ("Billed, not paid ($)", '=SUMIFS(Projects!J2:J200,Projects!K2:K200,"<>",Projects!L2:L200,"")'),
        ("Paid ($)", '=SUMIFS(Projects!J2:J200,Projects!L2:L200,"<>")'),
        ("Invoices overdue (30+ days)", '=COUNTIF(Projects!N2:N200,"Overdue")'),
    ]:
        summary.append([label, formula])
    for r in (2, 3):
        summary[f"B{r}"].number_format = USD

    notes_sheet(wb, [
        "Tidy Tabs: Freelance Project Tracker",
        "",
        "1. One row per project. Pick Status and Billing from the dropdowns.",
        "2. Hourly jobs: fill Hours and Rate. Flat-fee jobs: fill Flat fee. Amount picks the right one.",
        "3. Type the date you sent the invoice in Invoice sent, and the date the money arrived in Paid on.",
        "4. Days left counts down to the due date (negative = late) and hides itself once Status is Done.",
        "5. Payment shows Not invoiced / Waiting / Overdue (more than 30 days after you invoiced) / Paid.",
        "6. Days left and Payment use TODAY(), so they change every day you open the file.",
        "",
        "Sample clients are fictional. Delete rows 2-6 on the Projects tab to start fresh.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-freelance-project-tracker.xlsx")


ZIPS = [
    ("Boston", "MA", "02108", "02108-1522"),
    ("Hoboken", "NJ", "07030", "07030-3707"),
    ("Holtsville", "NY", "00501", "00501-0001"),
    ("Burlington", "VT", "05401", "05401-4417"),
    ("Chicago", "IL", "60601", "60601-3014"),
    ("Las Vegas", "NV", "89101", "89101-5019"),
]


def zip_codes():
    wb = Workbook()
    ws = wb.active
    ws.title = "Fixed (text)"
    header(ws, ["City", "State", "ZIP (text)", "ZIP+4 (text)"], [14, 7, 12, 14])
    for city, st, z, z4 in ZIPS:
        ws.append([city, st, z, z4])
    for row in ws.iter_rows(min_row=2, max_row=1000, min_col=3, max_col=4):
        for c in row:
            c.number_format = "@"

    broken = wb.create_sheet("Broken (numbers)")
    header(broken, ["City", "State", "ZIP as typed", "Repair with TEXT", "Repair ZIP+4 (9 digits)"], [14, 7, 14, 18, 22])
    for r, (city, st, z, z4) in enumerate(ZIPS, start=2):
        broken.append([city, st, int(z), f'=TEXT(C{r},"00000")', f'=TEXT({int(z4.replace("-", ""))},"00000-0000")'])

    notes_sheet(wb, [
        "Tidy Tabs: US ZIP Code Leading Zeros sample",
        "",
        "Fixed (text): the ZIP columns are formatted as Text (@) BEFORE the codes were entered, so 02108 stays 02108.",
        "Broken (numbers): the same ZIPs stored as numbers, which is what happens when you type or paste them into a General cell.",
        'Column D rebuilds a 5-digit ZIP with =TEXT(C2,"00000"). Copy it, then Paste Special > Values to keep the result.',
        "Column E shows the ZIP+4 version with the format 00000-0000.",
        "",
        "The 5-digit ZIPs are real; the +4 extensions are made up for practice.",
        "The same six ZIP codes are in tidy-tabs-zip-codes-sample.csv for practicing a CSV import.",
    ])
    wb.save(OUT / "tidy-tabs-zip-codes-leading-zeros.xlsx")

    with open(OUT / "tidy-tabs-zip-codes-sample.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["City", "State", "ZIP", "ZIP+4"])
        for row in ZIPS:
            w.writerow(row)


def print_one_page():
    wb = Workbook()
    ws = wb.active
    ws.title = "Price list"
    cols = ["SKU", "Item", "Category", "Size", "Color", "Unit cost", "Wholesale", "Retail",
            "Margin %", "On hand", "Reorder at", "Supplier", "Lead time (days)", "Last ordered"]
    header(ws, cols, [10, 30, 12, 10, 10, 10, 11, 10, 10, 9, 12, 14, 10, 13])
    colors = ["Natural", "Sage", "Navy", "Rust", "Cream"]
    sizes = ["4 oz", "8 oz", "12 oz"]
    for i in range(1, 49):
        r = i + 1
        cost = round(1.5 + (i % 7) * 0.65, 2)
        ws.append([f"ITM-{i:03d}", f"Sample product {i}", ["Candles", "Soap", "Cards", "Jewelry"][i % 4],
                   sizes[i % 3], colors[i % 5], cost, f"=ROUND(F{r}*2,2)", f"=ROUND(F{r}*{3 + (i % 3) * 0.5},2)",
                   f"=(H{r}-F{r})/H{r}", 10 + (i * 7) % 50, 15, f"Supplier {chr(65 + i % 6)}",
                   7 + (i % 4) * 7, date(2026, 8, 1 + i % 28)])
        for col in "FGH":
            ws[f"{col}{r}"].number_format = USD
        ws[f"I{r}"].number_format = "0%"
        ws[f"N{r}"].number_format = US_DATE
    # The settings the post walks through, already applied: US Letter,
    # landscape, fit all columns on one page wide, header row repeats.
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_title_rows = "1:1"
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_options.horizontalCentered = True
    wb.save(OUT / "tidy-tabs-print-one-page-us-letter.xlsx")


ADDRESSES = [
    ("Boston, MA 02108", "Harbor Dental"),
    ("Hoboken, NJ 07030", "Pine & Co. Realty"),
    ("Holtsville, NY 00501", "Main St. Gift Shop"),
    ("Burlington, VT 05401", "Lakeside Yoga"),
    ("Chicago, IL 60601", "Northside Print Co."),
    ("Las Vegas, NV 89101", "Desert Bloom Candles"),
    ("Salt Lake City, UT 84101", "Wasatch Bike Repair"),
    ("St. Louis, MO 63101", "Gateway Bakery"),
]


def split_city_state_zip():
    wb = Workbook()
    ws = wb.active
    ws.title = "Split"
    header(ws, ["Customer", "City, ST ZIP", "City", "State", "ZIP"], [22, 28, 18, 8, 9])
    for r, (addr, name) in enumerate(ADDRESSES, start=2):
        ws.append([name, addr,
                   f'=LEFT(B{r},FIND(",",B{r})-1)',
                   f'=MID(B{r},FIND(",",B{r})+2,2)',
                   f'=RIGHT(B{r},5)'])
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    notes_sheet(wb, [
        "Tidy Tabs: Split City, ST ZIP into three columns",
        "",
        'Type or paste addresses like "Boston, MA 02108" in column B.',
        "City, State, and ZIP (gray) fill in by formula. Copy a gray row down for more addresses.",
        'City = LEFT(B2,FIND(",",B2)-1): everything before the comma.',
        'State = MID(B2,FIND(",",B2)+2,2): the 2 letters after the comma and space.',
        "ZIP = RIGHT(B2,5): the last 5 characters, returned as text, so 02108 keeps its 0.",
        "Works for 5-digit ZIPs. For ZIP+4 (02108-1522), use RIGHT(B2,10).",
        "To keep only the results, copy C:E and Paste Special > Values.",
        "",
        "Customer names are fictional; the city/state/ZIP combinations are real.",
    ])
    wb.save(OUT / "tidy-tabs-split-city-state-zip.xlsx")


def dependent_dropdown():
    from openpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    ws = wb.active
    ws.title = "Order form"
    header(ws, ["Category", "Subcategory", "Item note", "Test: COUNTA(INDIRECT(cat))"],
           [14, 22, 30, 30])
    # Sample rows exercise every category so the INDIRECT lookup is checked
    # against each named range, not just one.
    rows = ["Candles", "Soap", "Jewelry", "Cards"]
    for r, cat in enumerate(rows, start=2):
        ws.append([cat, "", "pick from Subcategory dropdown", f"=COUNTA(INDIRECT(A{r}))"])
        ws[f"D{r}"].fill = FORMULA_FILL

    lists = wb.create_sheet("Lists")
    header(lists, ["Candles", "Soap", "Jewelry", "Cards"], [16, 16, 16, 16])
    cols = {
        "A": ["Lavender soy 8 oz", "Cedar soy 8 oz", "Vanilla soy 4 oz"],
        "B": ["Oatmeal goat milk bar", "Charcoal bar"],
        "C": ["Brass hoop earrings", "Beaded necklace", "Cuff bracelet", "Stacking ring"],
        "D": ["Letterpress birthday card", "Letterpress thank you card"],
    }
    for col, items in cols.items():
        for i, val in enumerate(items, start=2):
            lists[f"{col}{i}"] = val

    # Named ranges sized exactly to each list, keyed by the category text so
    # =INDIRECT(A2) resolves "Candles" to the Candles named range.
    ranges = {"Candles": "A2:A4", "Soap": "B2:B3", "Jewelry": "C2:C5", "Cards": "D2:D3"}
    for name, rng in ranges.items():
        wb.defined_names[name] = DefinedName(name, attr_text=f"Lists!${rng.split(':')[0][0]}${rng.split(':')[0][1:]}:${rng.split(':')[1][0]}${rng.split(':')[1][1:]}")

    cat_list = DataValidation(type="list", formula1="=Lists!$A$1:$D$1", allow_blank=True)
    ws.add_data_validation(cat_list)
    cat_list.add("A2:A200")
    sub_list = DataValidation(type="list", formula1="=INDIRECT($A2)", allow_blank=True)
    ws.add_data_validation(sub_list)
    sub_list.add("B2:B200")

    notes_sheet(wb, [
        "Tidy Tabs: Dependent (Category > Subcategory) Drop-Down List",
        "",
        "Order form tab: pick a Category in column A, then Subcategory in column B only shows items for that category.",
        "Lists tab: each column header is a category name; the cells below it are that category's subcategories.",
        "Behind the scenes, each Lists column is a named range with the same name as its header (Candles, Soap, Jewelry, Cards).",
        "Subcategory's data validation formula is =INDIRECT($A2), which turns the Category text into the matching named range.",
        "Column D is a check, not part of the form: =COUNTA(INDIRECT(A2)) counts how many subcategories INDIRECT found for that row's category, so you can confirm the link works before you rely on it.",
        "To add a category: add a column to Lists, name a range after it (Data > Named ranges), and add that name to the Category dropdown list.",
        "",
        "Sample categories and products are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx. Named ranges and INDIRECT dropdowns import with it.",
    ])
    wb.save(OUT / "tidy-tabs-dependent-dropdown-list.xlsx")


def client_contact_list():
    wb = Workbook()
    ws = wb.active
    ws.title = "Clients"
    header(ws, ["Client", "Company", "Last contact", "Follow up every (days)",
                "Next follow-up", "Status", "Notes"], [16, 22, 13, 15, 13, 14, 30])
    rows = [
        ("Dana Ruiz", "Harbor Dental", date(2026, 9, 1), 21, "Wants a spring newsletter quote"),
        ("Emma Lee", "Pine & Co. Realty", date(2026, 9, 20), 14, "Sent listing photo proofs"),
        ("Sam Patel", "Lakeside Yoga", date(2026, 8, 15), 30, "Website copy in review"),
        ("Jamie Chen", "Main St. Gift Shop", date(2026, 9, 25), 7, "Holiday flyer kickoff call"),
    ]
    for r, (name, company, last, every, notes) in enumerate(rows, start=2):
        ws.append([name, company, last, every, f"=C{r}+D{r}",
                   f'=IF(E{r}="","",IF(TODAY()>=E{r},"Follow up now","OK"))', notes])
        for col in "CE":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "EF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    last = 300
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['$F2="Follow up now"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Client Contact List with Follow-Up Reminders",
        "",
        "1. One row per client. Fill in Last contact and Follow up every (days).",
        "2. Next follow-up = Last contact + Follow up every (days), calculated for you.",
        "3. Status turns red and reads \"Follow up now\" once today's date reaches Next follow-up.",
        "4. Status uses TODAY(), so it updates on its own every day you open the file.",
        "5. After you reach out, update Last contact to today's date to reset the countdown.",
        "",
        "Sample clients are fictional. Delete rows 2-5 to start your own list.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-client-contact-list.xlsx")


def packing_slip():
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing slip"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws["A1"] = "Tidy Tabs Craft Co."
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "PACKING SLIP"
    ws["A2"].font = Font(bold=True, size=12)
    ws["A4"] = "Order #"
    ws["B4"] = "1042"
    ws["A5"] = "Order date"
    ws["B5"] = date(2026, 9, 26)
    ws["B5"].number_format = US_DATE
    ws["A6"] = "Ship to"
    ws["B6"] = "Priya Nair"
    ws["B7"] = "48 Willow St"
    ws["B8"] = "Burlington, VT 05401"

    headers_row = 10
    for col, val in zip("ABCD", ["Item", "SKU", "Qty ordered", "Qty packed"]):
        c = ws[f"{col}{headers_row}"]
        c.value = val
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    items = [
        ("Lavender soy candle, 8 oz", "CND-LAV8", 2, 2),
        ("Charcoal soap bar", "SOP-CHR", 3, 3),
        ("Brass hoop earrings", "EAR-HOOP", 1, 1),
        ("Letterpress birthday card", "CRD-BDAY", 4, 3),
    ]
    for i, (name, sku, qty_o, qty_p) in enumerate(items):
        r = headers_row + 1 + i
        ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = name, sku, qty_o, qty_p
    last_item_row = headers_row + len(items)
    total_row = last_item_row + 2
    ws[f"C{total_row}"] = "Total packed"
    ws[f"C{total_row}"].font = Font(bold=True)
    ws[f"D{total_row}"] = f"=SUM(D{headers_row + 1}:D{last_item_row})"
    ws[f"D{total_row}"].font = Font(bold=True)
    ws[f"A{total_row + 2}"] = "Thank you for your order! Questions: hello@example.com"

    # US Letter, portrait, one page, matching the print-one-page settings
    # used on the other printable templates in this file.
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = ws.page_margins.right = 0.75
    ws.page_margins.top = ws.page_margins.bottom = 0.75

    notes_sheet(wb, [
        "Tidy Tabs: Packing Slip Template",
        "",
        "1. Fill in Order #, Order date, and Ship to at the top.",
        "2. List one item per row: Item, SKU, Qty ordered, Qty packed.",
        "3. Total packed adds up column D on its own: =SUM(D11:D14) in this sample.",
        "4. If Qty packed is less than Qty ordered for any row, that item shipped short; note why in a blank row below the table.",
        "5. Page is already set to US Letter, portrait, and fit to one page, so File > Print needs no changes.",
        "",
        "Business name, order, and customer are fictional sample data.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-packing-slip-template.xlsx")


def running_balance():
    wb = Workbook()
    ws = wb.active
    ws.title = "Cash log"
    header(ws, ["Date", "Description", "Money in", "Money out", "Balance"], [12, 32, 11, 11, 12])
    ws.append([date(2026, 9, 1), "Opening balance", None, None, 250.00])
    entries = [
        (date(2026, 9, 6), "Farmers market cash sales", 180.00, None),
        (date(2026, 9, 9), "Bought jars and labels", None, 45.50),
        (date(2026, 9, 15), "Etsy payout deposited", 96.20, None),
        (date(2026, 9, 20), "Paid booth fee, Oct market", None, 35.00),
    ]
    for desc_row, (d, desc, in_amt, out_amt) in enumerate(entries, start=3):
        prev = desc_row - 1
        ws.append([d, desc, in_amt, out_amt, f"=E{prev}+C{desc_row}-D{desc_row}"])
    for r in range(2, 2 + 1 + len(entries)):
        ws[f"A{r}"].number_format = US_DATE
        for col in "CDE":
            ws[f"{col}{r}"].number_format = USD
        ws[f"E{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Running Balance Cash Log",
        "",
        "1. Row 2 is your opening balance, typed in directly (not a formula).",
        "2. Every row after that: type Money in or Money out (leave the other blank).",
        "3. Balance adds the row above's balance, plus Money in, minus Money out: =E2+C3-D3, copied down.",
        "4. Never type over the Balance column (gray); copy the formula down when you add a row.",
        "",
        "Sample cash entries are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-running-balance-cash-log.xlsx")


def highlight_duplicates():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order #", "Customer email", "Duplicate?"], [12, 30, 14])
    rows = [
        ("1001", "priya.nair@example.com"),
        ("1002", "sam.patel@example.com"),
        ("1003", "dana.ruiz@example.com"),
        ("1004", "priya.nair@example.com"),  # repeat customer, different order
        ("1005", "emma.lee@example.com"),
        ("1006", "jamie.chen@example.com"),
        ("1007", "sam.patel@example.com"),   # repeat customer
        ("1008", "new.customer@example.com"),
        ("1009", "dana.ruiz@example.com"),   # repeat customer
        ("1010", "ana.gomez@example.com"),
        ("1011", "ana.gomez@example.com"),   # same order pasted in twice by mistake
        ("1012", "priya.nair@example.com"),  # third order, same email
        ("1013", "tom.reyes@example.com"),
        ("1014", "lee.wu@example.com"),
        ("1015", "sam.patel@example.com"),   # repeat customer
    ]
    last = 1 + len(rows)
    for r, (order, email) in enumerate(rows, start=2):
        ws.append([order, email, f'=IF(COUNTIF($B$2:$B${last},B{r})>1,"Duplicate","")'])
        ws[f"C{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"B2:B{last}", FormulaRule(formula=[f'COUNTIF($B$2:$B${last},B2)>1'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Highlight Duplicate Values Sample",
        "",
        "Column C shows =IF(COUNTIF($B$2:$B$16,B2)>1,\"Duplicate\",\"\") for every row.",
        "The same COUNTIF formula drives the red conditional formatting fill on column B (Format > Conditional formatting > Custom formula is).",
        "This flags repeated values; it never deletes or moves rows, so repeat customers with more than one real order still show up highlighted on purpose.",
        "To find only accidental duplicates (the exact same order entered twice), also compare the Order # column, or sort by email and scan by eye.",
        "",
        "Order numbers and emails are fictional sample data.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-highlight-duplicates-sample.xlsx")


def business_days_ship_by():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order #", "Customer", "Order date", "Business days to ship", "Ship by",
                "Calendar days", "Business days check"], [10, 22, 12, 14, 12, 12, 16])
    rows = [
        ("2001", "Harbor Dental", date(2026, 9, 28), 5),
        ("2002", "Lakeside Yoga", date(2026, 10, 6), 10),
        ("2003", "Main St. Gift Shop", date(2026, 11, 9), 3),
        ("2004", "Gateway Bakery", date(2026, 11, 20), 5),
    ]
    for r, (num, cust, d, n) in enumerate(rows, start=2):
        ws.append([num, cust, d, n,
                   f"=WORKDAY(C{r},D{r},Holidays!$A$2:$A$20)",
                   f"=E{r}-C{r}",
                   f"=NETWORKDAYS(C{r},E{r},Holidays!$A$2:$A$20)-1"])
        for col in "CE":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "EFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    hol = wb.create_sheet("Holidays")
    header(hol, ["Date", "Day off"], [12, 28])
    for d, name in [(date(2026, 9, 7), "Labor Day"), (date(2026, 10, 12), "Columbus / Indigenous Peoples' Day"),
                    (date(2026, 11, 11), "Veterans Day"), (date(2026, 11, 26), "Thanksgiving Day"),
                    (date(2026, 12, 25), "Christmas Day")]:
        hol.append([d, name])
        hol.cell(row=hol.max_row, column=1).number_format = US_DATE
    hol.column_dimensions["B"].width = 34

    notes_sheet(wb, [
        "Tidy Tabs: Business Days and Ship-By Dates",
        "",
        "1. Orders tab: type the Order date and how many business days you need to ship.",
        "2. Ship by = WORKDAY(C2,D2,Holidays!$A$2:$A$20). It skips Saturdays, Sundays, and every date on the Holidays tab.",
        "3. Calendar days shows how many real days that is. Business days check = NETWORKDAYS(C2,E2,Holidays!$A$2:$A$20)-1 counts back and should equal column D.",
        "4. Holidays tab: add or remove your own days off (market days, vacation). Ranges run to row 20.",
        "5. WORKDAY and NETWORKDAYS treat Saturday and Sunday as the weekend. For other weekends use WORKDAY.INTL.",
        "",
        "Customers and orders are fictional. Holiday dates are the 2026 US federal holidays that fall in the sample range; check your own calendar.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-business-days-ship-by-date.xlsx")


def time_to_decimal_hours():
    from datetime import time

    wb = Workbook()
    ws = wb.active
    ws.title = "Shifts"
    header(ws, ["Date", "Clock in", "Clock out", "Unpaid break (min)", "Time worked (h:mm)",
                "Decimal hours", "Nearest 15 min", "Pay"], [12, 11, 11, 13, 14, 12, 13, 11])
    rows = [
        (date(2026, 9, 21), time(8, 30), time(17, 0), 30),
        (date(2026, 9, 22), time(9, 0), time(17, 45), 30),
        (date(2026, 9, 23), time(7, 15), time(15, 40), 30),
        (date(2026, 9, 24), time(22, 0), time(6, 30), 30),
        (date(2026, 9, 25), time(9, 5), time(13, 20), 0),
    ]
    for r, (d, tin, tout, brk) in enumerate(rows, start=2):
        ws.append([d, tin, tout, brk,
                   f"=MOD(C{r}-B{r},1)-D{r}/1440",
                   f"=ROUND(E{r}*24,2)",
                   f"=MROUND(E{r}*24,0.25)",
                   f"=ROUND(F{r}*$K$1,2)"])
        ws[f"A{r}"].number_format = US_DATE
        for col in "BC":
            ws[f"{col}{r}"].number_format = "h:mm AM/PM"
        ws[f"E{r}"].number_format = "h:mm"
        ws[f"F{r}"].number_format = "0.00"
        ws[f"G{r}"].number_format = "0.00"
        ws[f"H{r}"].number_format = USD
        for col in "EFGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["J1"], ws["K1"] = "Hourly rate", 22.5
    ws["J2"], ws["K2"] = "Total decimal hours", "=SUM(F2:F200)"
    ws["J3"], ws["K3"] = "Total pay", "=SUM(H2:H200)"
    ws["K1"].number_format = USD
    ws["K3"].number_format = USD
    for c in ("J1", "J2", "J3"):
        ws[c].font = Font(bold=True)
    for col, w in (("J", 20), ("K", 12)):
        ws.column_dimensions[col].width = w

    notes_sheet(wb, [
        "Tidy Tabs: Convert Clock Times to Decimal Hours",
        "",
        "1. Type Clock in, Clock out, and Unpaid break in minutes. Format the time cells as Time.",
        "2. Time worked = MOD(C2-B2,1)-D2/1440. MOD makes overnight shifts (10:00 PM to 6:30 AM) come out positive; 1440 is minutes in a day.",
        "3. Decimal hours = ROUND(E2*24,2). Excel and Sheets store time as a fraction of a day, so multiply by 24 to get hours.",
        "4. Nearest 15 min = MROUND(E2*24,0.25), for payroll that rounds to quarter hours.",
        "5. Pay = ROUND(Decimal hours x the hourly rate in K1,2). Change the rate there.",
        "",
        "Sample shifts and the $22.50 rate are fictional. Rounding rules are set by your own payroll policy; this file only does the arithmetic.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-time-to-decimal-hours.xlsx")


def receivables_aging():
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    header(ws, ["Invoice #", "Client", "Invoice date", "Terms (days)", "Due date", "Amount",
                "Paid on", "Balance", "Days past due", "Bucket"],
           [10, 22, 12, 11, 12, 12, 12, 12, 11, 11])
    rows = [
        ("1001", "Harbor Dental", date(2026, 7, 1), 30, 650, date(2026, 8, 5)),
        ("1002", "Pine & Co. Realty", date(2026, 7, 15), 30, 1200, None),
        ("1003", "Lakeside Yoga", date(2026, 8, 10), 30, 380, None),
        ("1004", "Main St. Gift Shop", date(2026, 8, 25), 15, 300, None),
        ("1005", "Northside Print Co.", date(2026, 9, 5), 30, 875, None),
        ("1006", "Desert Bloom Candles", date(2026, 5, 20), 30, 450, None),
        ("1007", "Wasatch Bike Repair", date(2026, 6, 30), 30, 220, None),
        ("1008", "Gateway Bakery", date(2026, 9, 20), 30, 540, None),
        ("1009", "Harbor Dental", date(2026, 9, 15), 30, 650, None),
    ]
    for r, (num, client, d, terms, amt, paid) in enumerate(rows, start=2):
        ws.append([num, client, d, terms, f"=C{r}+D{r}", amt, paid,
                   f'=IF(OR(F{r}="",G{r}<>""),0,F{r})',
                   f'=IF(H{r}=0,"",MAX(0,Aging!$B$1-E{r}))',
                   f'=IF(H{r}=0,"Paid",IF(I{r}=0,"Current",IF(I{r}<=30,"1-30",IF(I{r}<=60,"31-60",IF(I{r}<=90,"61-90","90+")))))'])
        for col in "CEG":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "FH":
            ws[f"{col}{r}"].number_format = USD
        for col in "EHIJ":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    ag = wb.create_sheet("Aging")
    ag.column_dimensions["A"].width = 22
    ag.column_dimensions["B"].width = 14
    ag.column_dimensions["C"].width = 12
    ag["A1"], ag["B1"] = "As of date", date(2026, 9, 30)
    ag["B1"].number_format = US_DATE
    ag["A1"].font = Font(bold=True)
    ag.append([])
    ag.append(["Bucket (days past due)", "Balance", "Invoices"])
    for c in ag[3]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for r, b in enumerate(["Current", "1-30", "31-60", "61-90", "90+"], start=4):
        ag.append([b, f"=SUMIFS(Invoices!$H$2:$H$500,Invoices!$J$2:$J$500,A{r})",
                   f"=COUNTIFS(Invoices!$J$2:$J$500,A{r})"])
        ag[f"B{r}"].number_format = USD
    ag.append(["Total outstanding", "=SUM(B4:B8)", "=SUM(C4:C8)"])
    ag["B9"].number_format = USD
    ag["A9"].font = ag["B9"].font = Font(bold=True)
    ws.conditional_formatting.add("J2:J500", FormulaRule(formula=['OR($J2="61-90",$J2="90+")'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Accounts Receivable Aging Report",
        "",
        "1. Invoices tab: one row per invoice. Fill Invoice #, Client, Invoice date, Terms (days), Amount. Type Paid on when the money arrives.",
        "2. Due date = Invoice date + Terms. Balance is the Amount until you enter a Paid on date, then 0.",
        "3. Days past due = As of date minus Due date (0 if not late yet). The As of date is cell B1 on the Aging tab.",
        "4. Bucket sorts each unpaid invoice into Current, 1-30, 31-60, 61-90, or 90+ days past due.",
        "5. Aging tab adds up Balance by bucket with SUMIFS. Type =TODAY() in B1 to always age as of today; the sample uses 09/30/2026 so the numbers stay put.",
        "",
        "Clients and amounts are fictional. This only groups invoices by how late they are; it makes no decision about collection or write-offs.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-receivables-aging-report.xlsx")


def markup_vs_margin():
    wb = Workbook()
    ws = wb.active
    ws.title = "Pricing"
    header(ws, ["Item", "Unit cost", "Markup %", "Price from markup", "Profit per unit", "Actual margin %",
                "Target margin %", "Price for target margin", "Margin check"],
           [26, 11, 11, 14, 13, 13, 13, 16, 13])
    rows = [
        ("Lavender soy candle, 8 oz", 4.10, 1.00, 0.50),
        ("Oatmeal goat milk soap bar", 1.80, 1.50, 0.60),
        ("Brass hoop earrings", 3.40, 2.00, 0.70),
        ("Letterpress birthday card", 0.95, 3.00, 0.75),
    ]
    for r, (item, cost, mk, tgt) in enumerate(rows, start=2):
        ws.append([item, cost, mk, f"=ROUND(B{r}*(1+C{r}),2)", f"=D{r}-B{r}", f"=E{r}/D{r}",
                   tgt, f"=ROUND(B{r}/(1-G{r}),2)", f"=(H{r}-B{r})/H{r}"])
        for col in "BDEH":
            ws[f"{col}{r}"].number_format = USD
        for col in "CFGI":
            ws[f"{col}{r}"].number_format = "0.0%"
        for col in "DEFHI":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Markup vs Profit Margin Calculator",
        "",
        "1. Type Unit cost and the Markup % you use (100% means price is double the cost).",
        "2. Price from markup = ROUND(B2*(1+C2),2). Profit per unit = price minus cost.",
        "3. Actual margin % = Profit / Price. A 100% markup is only a 50% margin, because margin is measured against the price, not the cost.",
        "4. Going the other way: type a Target margin %. Price for target margin = ROUND(B2/(1-G2),2). Do not use B2*(1+G2); that mixes up markup and margin.",
        "5. Margin check divides the profit at the rounded price by that price; it can differ from the target by a fraction of a percent.",
        "",
        "Items and costs are fictional. Prices are arithmetic only; they do not include shipping, fees, or taxes.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-markup-vs-margin-calculator.xlsx")


def vlookup_price_list():
    wb = Workbook()
    ws = wb.active
    ws.title = "Order lines"
    header(ws, ["SKU", "Qty", "Item", "Unit price", "Line total", "Price via INDEX/MATCH"],
           [12, 7, 30, 11, 12, 20])
    rows = [("CND-LAV8", 2), ("SOP-OAT", 3), ("EAR-HOOP", 1), ("SOP-OAT ", 1),
            ("CRD-BDAY", 5), ("CND-CED8", 1)]
    for r, (sku, qty) in enumerate(rows, start=2):
        ws.append([sku, qty,
                   f'=IFERROR(VLOOKUP(A{r},\'Price list\'!$A$2:$C$500,2,FALSE),"SKU not found")',
                   f'=IFERROR(VLOOKUP(A{r},\'Price list\'!$A$2:$C$500,3,FALSE),"")',
                   f'=IF(D{r}="","",B{r}*D{r})',
                   f'=IFERROR(INDEX(\'Price list\'!$C$2:$C$500,MATCH(A{r},\'Price list\'!$A$2:$A$500,0)),"")'])
        for col in "DEF":
            ws[f"{col}{r}"].number_format = USD
        for col in "CDEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["H1"], ws["I1"] = "Order total", "=SUM(E2:E500)"
    ws["H1"].font = Font(bold=True)
    ws["I1"].number_format = USD
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 12

    pl = wb.create_sheet("Price list")
    header(pl, ["SKU", "Item", "Price"], [12, 30, 11])
    for sku, item, price in [
        ("CND-LAV8", "Lavender soy candle, 8 oz", 18.00), ("CND-CED8", "Cedar soy candle, 8 oz", 18.00),
        ("SOP-OAT", "Oatmeal goat milk soap bar", 8.00), ("SOP-CHR", "Charcoal soap bar", 8.50),
        ("EAR-HOOP", "Brass hoop earrings", 24.00), ("CRD-BDAY", "Letterpress birthday card", 6.00),
    ]:
        pl.append([sku, item, price])
        pl.cell(row=pl.max_row, column=3).number_format = USD

    notes_sheet(wb, [
        "Tidy Tabs: Look Up a Price with VLOOKUP",
        "",
        "1. Price list tab: SKU must be the first column, then Item, then Price.",
        "2. Order lines tab: type a SKU and Qty. Item = VLOOKUP(A2,'Price list'!$A$2:$C$500,2,FALSE). FALSE means exact match only.",
        "3. Unit price uses column number 3. Line total = Qty x Unit price. Order total is in I1.",
        "4. IFERROR turns a missing SKU into \"SKU not found\" instead of #N/A. Row 5 has a trailing space after SOP-OAT on purpose, so it is not found.",
        "5. Column F does the same lookup with INDEX/MATCH, which also works when the lookup column is not first.",
        "",
        "Products and prices are fictional. XLOOKUP is not used here because it needs newer versions of Excel and LibreOffice.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-vlookup-price-list.xlsx")


def break_even_calculator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Break-even"
    header(ws, ["Product", "Price", "Variable cost per unit", "Fixed costs", "Profit per unit",
                "Break-even units", "Break-even revenue"], [28, 10, 14, 12, 12, 13, 15])
    rows = [
        ("Lavender soy candle, 8 oz", 18.00, 6.50, 120.00),
        ("Oatmeal goat milk soap bar", 8.00, 3.25, 75.00),
        ("Brass hoop earrings", 24.00, 9.60, 150.00),
        ("Letterpress birthday card", 6.00, 1.50, 60.00),
    ]
    for r, (item, price, var, fixed) in enumerate(rows, start=2):
        ws.append([item, price, var, fixed, f"=B{r}-C{r}",
                   f'=IF(E{r}<=0,"No break-even",ROUNDUP(D{r}/E{r},0))',
                   f'=IF(ISNUMBER(F{r}),F{r}*B{r},"")'])
        for col in "BCDEG":
            ws[f"{col}{r}"].number_format = USD
        for col in "EFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Break-Even Calculator for a Booth or Online Shop",
        "",
        "1. One row per product. Type Price, Variable cost per unit (materials, packaging, fees for one sale), and Fixed costs (booth fee, table rental, a one-time setup cost).",
        "2. Profit per unit = Price minus Variable cost per unit.",
        "3. Break-even units = ROUNDUP(Fixed costs / Profit per unit, 0). ROUNDUP, not ROUND, because you cannot sell part of a unit and be fully covered.",
        "4. Break-even revenue = Break-even units x Price.",
        "5. If Price is not higher than Variable cost per unit, Break-even units shows No break-even instead of a negative number.",
        "6. Copy the last row down to add products.",
        "",
        "Products, prices, and costs are fictional. This is arithmetic only; it does not include taxes or your own pay.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-break-even-calculator.xlsx")


def subscription_renewal_tracker():
    wb = Workbook()
    ws = wb.active
    ws.title = "Subscriptions"
    header(ws, ["Subscription", "Last renewed", "Months per cycle", "Cost per cycle", "Next renewal",
                "Days left", "Status", "Monthly cost", "Annual cost"],
           [28, 13, 11, 12, 13, 10, 14, 12, 12])
    rows = [
        ("Domain name", date(2026, 3, 15), 12, 18.00),
        ("Email marketing plan", date(2026, 9, 15), 1, 20.00),
        ("Online shop plan", date(2026, 9, 28), 1, 10.00),
        ("Bookkeeping software", date(2025, 11, 5), 12, 180.00),
        ("Photo storage (6 months)", date(2026, 8, 31), 6, 30.00),
    ]
    for r, (name, last, months, cost) in enumerate(rows, start=2):
        ws.append([name, last, months, cost,
                   f'=IF(B{r}="","",EDATE(B{r},C{r}))',
                   f'=IF(E{r}="","",E{r}-$K$1)',
                   f'=IF(F{r}="","",IF(F{r}<0,"Past due",IF(F{r}<=30,"Renews soon","OK")))',
                   f'=IF(D{r}="","",D{r}/C{r})',
                   f'=IF(D{r}="","",D{r}*12/C{r})'])
        for col in "BE":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "DHI":
            ws[f"{col}{r}"].number_format = USD
        for col in "EFGHI":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    last_row = 200
    ws["J1"], ws["K1"] = "As of date", date(2026, 10, 2)
    ws["J2"], ws["K2"] = "Total monthly", f"=SUM(H2:H{last_row})"
    ws["J3"], ws["K3"] = "Total annual", f"=SUM(I2:I{last_row})"
    ws["K1"].number_format = US_DATE
    ws["K2"].number_format = USD
    ws["K3"].number_format = USD
    for c in ("J1", "J2", "J3"):
        ws[c].font = Font(bold=True)
    ws.column_dimensions["J"].width = 16
    ws.column_dimensions["K"].width = 13
    ws.conditional_formatting.add(f"G2:G{last_row}", FormulaRule(formula=['OR($G2="Renews soon",$G2="Past due")'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Subscription Renewal Tracker",
        "",
        "1. One row per subscription. Type Last renewed (the date you were last charged), Months per cycle (1 = monthly, 12 = yearly), and Cost per cycle.",
        "2. Next renewal = EDATE(B2,C2), the same day of the month after Months per cycle. If that day does not exist (08/31 plus 6 months), EDATE uses the last day of the month: 02/28/2027.",
        "3. Days left = Next renewal minus the As of date in K1. Status shows Renews soon at 30 days or fewer, and Past due below 0.",
        "4. Monthly cost = Cost per cycle / Months per cycle. Annual cost = Cost per cycle x 12 / Months per cycle. Totals are in K2 and K3.",
        "5. The sample As of date is 10/02/2026 so the numbers stay put. Type =TODAY() in K1 to count from today.",
        "6. After a renewal, type the new charge date in Last renewed.",
        "",
        "Subscriptions and prices are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-subscription-renewal-tracker.xlsx")


def budget_vs_actual():
    wb = Workbook()
    ws = wb.active
    ws.title = "Budget vs actual"
    header(ws, ["Category", "Budget", "Actual", "Variance ($)", "Variance (%)", "Status"],
           [28, 12, 12, 13, 13, 14])
    rows = [
        ("Booth and market fees", 450.00, 450.00),
        ("Materials", 600.00, 683.40),
        ("Shipping", 220.00, 241.75),
        ("Packaging", 120.00, 131.20),
        ("Software subscriptions", 85.00, 85.00),
        ("Advertising", 150.00, 92.50),
        ("Payment processing fees", 90.00, 87.35),
        ("Training (not budgeted)", 0.00, 45.00),
    ]
    for r, (cat, bud, act) in enumerate(rows, start=2):
        ws.append([cat, bud, act, f"=C{r}-B{r}", f'=IF(B{r}=0,"",(C{r}-B{r})/B{r})',
                   f'=IF(C{r}>B{r},"Over budget","On or under")'])
        for col in "BCD":
            ws[f"{col}{r}"].number_format = USD
        ws[f"E{r}"].number_format = "0.0%"
        for col in "DEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    last = 200
    # Totals sit beside the table so rows can be added below without moving them.
    ws["H1"], ws["I1"] = "Total budget", f"=SUM(B2:B{last})"
    ws["H2"], ws["I2"] = "Total actual", f"=SUM(C2:C{last})"
    ws["H3"], ws["I3"] = "Total variance ($)", "=I2-I1"
    ws["H4"], ws["I4"] = "Total variance (%)", '=IF(I1=0,"",(I2-I1)/I1)'
    for c in ("I1", "I2", "I3"):
        ws[c].number_format = USD
    ws["I4"].number_format = "0.0%"
    for c in ("H1", "H2", "H3", "H4"):
        ws[c].font = Font(bold=True)
    ws.column_dimensions["H"].width = 20
    ws.column_dimensions["I"].width = 13
    ws.conditional_formatting.add(f"A2:F{last}", FormulaRule(formula=['AND($C2<>"",$C2>$B2)'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Budget vs Actual Variance",
        "",
        "1. One row per spending category. Type the Budget for the month and the Actual you spent.",
        "2. Variance ($) = Actual minus Budget. Positive means you spent more than planned.",
        "3. Variance (%) = (Actual - Budget) / Budget. It stays blank when Budget is 0, because dividing by zero is an error.",
        "4. Status says Over budget when Actual is more than Budget. Over-budget rows turn red by conditional formatting.",
        "5. Totals are in H1:I4, beside the table. Copy the last row down to add categories.",
        "",
        "Categories and amounts are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-budget-vs-actual.xlsx")


def farmers_market_sales_log():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Market date", "Item", "Qty sold", "Price", "Line total"], [13, 28, 10, 10, 12])
    d1, d2, d3 = date(2026, 9, 12), date(2026, 9, 19), date(2026, 9, 26)
    rows = [
        (d1, "Lavender soy candle, 8 oz", 9, 18.00),
        (d1, "Oatmeal goat milk soap bar", 14, 8.00),
        (d1, "Brass hoop earrings", 3, 24.00),
        (d1, "Letterpress birthday card", 12, 6.00),
        (d2, "Lavender soy candle, 8 oz", 11, 18.00),
        (d2, "Oatmeal goat milk soap bar", 10, 8.00),
        (d2, "Cedar soy candle, 8 oz", 6, 18.00),
        (d2, "Letterpress birthday card", 8, 6.00),
        (d2, "Brass hoop earrings", 2, 24.00),
        (d3, "Lavender soy candle, 8 oz", 7, 18.00),
        (d3, "Oatmeal goat milk soap bar", 16, 8.00),
        (d3, "Brass hoop earrings", 4, 24.00),
        (d3, "Letterpress birthday card", 10, 6.00),
    ]
    for r, (d, item, qty, price) in enumerate(rows, start=2):
        ws.append([d, item, qty, price, f"=C{r}*D{r}"])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = USD
        ws[f"E{r}"].number_format = USD
        ws[f"E{r}"].fill = FORMULA_FILL

    dt = wb.create_sheet("Daily totals")
    header(dt, ["Market date", "Revenue (SUMPRODUCT)", "Revenue (SUMIFS check)", "Items sold"],
           [13, 22, 22, 11])
    for r, d in enumerate((d1, d2, d3), start=2):
        dt.append([d,
                   f"=SUMPRODUCT((Sales!$A$2:$A$500=A{r})*Sales!$C$2:$C$500*Sales!$D$2:$D$500)",
                   f"=SUMIFS(Sales!$E$2:$E$500,Sales!$A$2:$A$500,A{r})",
                   f"=SUMIFS(Sales!$C$2:$C$500,Sales!$A$2:$A$500,A{r})"])
        dt[f"A{r}"].number_format = US_DATE
        dt[f"B{r}"].number_format = USD
        dt[f"C{r}"].number_format = USD
        for col in "BCD":
            dt[f"{col}{r}"].fill = FORMULA_FILL
    dt["F1"], dt["G1"] = "All days (SUMPRODUCT)", "=SUMPRODUCT(Sales!C2:C500,Sales!D2:D500)"
    dt["F2"], dt["G2"] = "All days (sum of Line total)", "=SUM(Sales!E2:E500)"
    dt["G1"].number_format = USD
    dt["G2"].number_format = USD
    dt.column_dimensions["F"].width = 28
    dt.column_dimensions["G"].width = 13
    for c in ("F1", "F2"):
        dt[c].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Farmers Market Sales Log",
        "",
        "1. Sales tab: one row per item sold at a market day. Type the Market date, Item, Qty sold, and Price. Line total = Qty x Price.",
        "2. Daily totals tab: type each market date in column A. Revenue (SUMPRODUCT) = SUMPRODUCT((Sales!$A$2:$A$500=A2)*Sales!$C$2:$C$500*Sales!$D$2:$D$500). It multiplies Qty by Price for that day's rows and adds them, without needing the Line total column.",
        "3. Revenue (SUMIFS check) adds the Line total column for that date. The two columns should match.",
        "4. G1 = SUMPRODUCT(Sales!C2:C500,Sales!D2:D500) is the grand total of Qty x Price across every day. G2 adds the Line total column; they should match too.",
        "5. Dates must be real dates, and Qty and Price must be numbers. Text in those cells makes the SUMPRODUCT formula in step 2 return #VALUE!.",
        "",
        "The vendor, items, and sales are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-farmers-market-sales-log.xlsx")


def clean_customer_list():
    wb = Workbook()
    ws = wb.active
    ws.title = "Clean names"
    header(ws, ["Messy name (paste here)", "Basic: PROPER(TRIM(CLEAN()))",
                "Full: also fixes non-breaking spaces and line breaks", "Length before", "Length after"],
           [30, 30, 40, 10, 10])
    nb = " "
    raw = [
        "  dana   ruiz ",
        "EMMA LEE",
        f"sam{nb}patel",
        "jamie\nchen",
        "  PRIYA   NAIR  ",
        "ana gomez",
        "ronald McDONALD",
        "sean o'neil",
        f"TOM  REYES{nb}",
        "   lee wu",
    ]
    for r, name in enumerate(raw, start=2):
        ws.append([name,
                   f"=PROPER(TRIM(CLEAN(A{r})))",
                   # UNICHAR(160), not CHAR(160): LibreOffice maps CHAR(160) through the
                   # system locale, so it misses non-breaking spaces under UTF-8. The
                   # _xlfn. prefix is how .xlsx files store UNICHAR (Excel 2013+, Sheets).
                   f'=PROPER(TRIM(CLEAN(SUBSTITUTE(SUBSTITUTE(A{r},_xlfn.UNICHAR(160)," "),CHAR(10)," "))))',
                   f"=LEN(A{r})", f"=LEN(C{r})"])
        ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        for col in "BCDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Clean Up a Customer List with TRIM, PROPER, and CLEAN",
        "",
        "1. Paste names in column A. Copy the gray formulas in B and C down next to them.",
        "2. TRIM removes extra spaces (leading, trailing, and doubled). CLEAN removes most hidden non-printing characters. PROPER capitalizes the first letter of each word.",
        "3. Column B is the basic formula =PROPER(TRIM(CLEAN(A2))). It does not remove non-breaking spaces (character 160, common in text copied from web pages), and CLEAN deletes a line break instead of replacing it with a space, so jamie + line break + chen becomes Jamiechen.",
        "4. Column C swaps both for normal spaces first: =PROPER(TRIM(CLEAN(SUBSTITUTE(SUBSTITUTE(A2,UNICHAR(160),\" \"),CHAR(10),\" \")))). UNICHAR(160) is used instead of CHAR(160) because LibreOffice can return a different character for CHAR(160) depending on the computer's locale.",
        "5. PROPER lowercases the rest of every word, so McDonald becomes Mcdonald. Fix names like that by hand.",
        "6. To keep the results, copy column C and Paste Special > Values only, then delete the formula columns.",
        "",
        "Names are fictional. The non-breaking space and line break behavior was checked in LibreOffice Calc only.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-clean-customer-list.xlsx")


def loan_payment_calculator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Loans"
    header(ws, ["Loan", "Amount", "Yearly rate", "Years", "Monthly payment", "Total paid", "Total interest"],
           [28, 12, 11, 8, 15, 13, 14])
    rows = [
        ("Equipment loan (fictional)", 15000.00, 0.065, 5),
        ("Kiln purchase", 8000.00, 0.0725, 3),
        ("Booth trailer", 22000.00, 0.059, 7),
        ("Inventory line", 5000.00, 0.09, 2),
    ]
    for r, (name, amt, rate, yrs) in enumerate(rows, start=2):
        ws.append([name, amt, rate, yrs, f"=PMT(C{r}/12,D{r}*12,-B{r})", f"=E{r}*D{r}*12", f"=F{r}-B{r}"])
        ws[f"B{r}"].number_format = USD
        ws[f"C{r}"].number_format = "0.00%"
        for col in "EFG":
            ws[f"{col}{r}"].number_format = USD
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Loan Payment Calculator with PMT",
        "",
        "1. One row per loan. Type the Amount, the Yearly rate (6.5% or 0.065), and the Years.",
        "2. Monthly payment = PMT(rate/12, years*12, -amount). The rate is divided by 12 and the years are multiplied by 12 because the payments are monthly.",
        "3. The minus sign before the amount makes the payment come out positive. Without it PMT returns a negative number.",
        "4. Total paid = Monthly payment x number of payments. Total interest = Total paid minus Amount.",
        "5. Results can be a cent off a lender's statement because lenders round each month's payment.",
        "",
        "Loans and rates are fictional. This is arithmetic only, not financial advice.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-loan-payment-calculator.xlsx")


def percent_change():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Sales channel", "Last month", "This month", "Change ($)", "Change (%)"], [26, 13, 13, 13, 12])
    rows = [
        ("Whole shop", 2400.00, 2760.00),
        ("Etsy orders", 1800.00, 1692.00),
        ("Farmers market", 950.00, 1140.00),
        ("Workshops (new this month)", 0.00, 300.00),
        ("Wholesale", 640.00, 640.00),
    ]
    for r, (name, old, new) in enumerate(rows, start=2):
        ws.append([name, old, new, f"=C{r}-B{r}", f'=IF(B{r}=0,"",(C{r}-B{r})/B{r})'])
        for col in "BCD":
            ws[f"{col}{r}"].number_format = USD
        ws[f"E{r}"].number_format = "0.0%"
        for col in "DE":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Percent Change Between Two Months",
        "",
        "1. Type last month in column B and this month in column C.",
        "2. Change ($) = This month minus Last month.",
        "3. Change (%) = (This month - Last month) / Last month, formatted as Percent. Divide by the earlier month, not the later one.",
        "4. If Last month is 0 the percent is undefined, so the formula leaves the cell blank (see the Workshops row).",
        "5. A negative percent means the number went down.",
        "",
        "All figures are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-percent-change.xlsx")


def rank_top_customers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Customers"
    header(ws, ["Customer", "Total sales", "Rank (ties share)", "Unique rank (ties broken)"], [26, 13, 14, 16])
    rows = [
        ("Harbor Coffee Co.", 4850.00), ("Maple Street Bakery", 3920.00), ("Oak & Ember Candles", 3920.00),
        ("Riverside Yoga", 2780.00), ("Juniper Florist", 2150.00), ("Bluebird Books", 2150.00),
        ("Cedar Hardware", 1640.00), ("Lakeview Dental Office", 3100.00), ("Pine Grove Garden", 980.00),
        ("North End Tailors", 1275.00),
    ]
    for r, (name, tot) in enumerate(rows, start=2):
        ws.append([name, tot, f"=RANK(B{r},$B$2:$B$11,0)", f"=C{r}+COUNTIF($B$2:B{r},B{r})-1"])
        ws[f"B{r}"].number_format = USD
        for col in "CD":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["F1"], ws["G1"], ws["H1"] = "Top 3", "Customer", "Total sales"
    for c in ("F1", "G1", "H1"):
        ws[c].fill, ws[c].font = HEADER_FILL, HEADER_FONT
    for i in range(1, 4):
        r = i + 1
        ws[f"F{r}"] = i
        ws[f"G{r}"] = f"=INDEX($A$2:$A$11,MATCH(F{r},$D$2:$D$11,0))"
        ws[f"H{r}"] = f"=INDEX($B$2:$B$11,MATCH(F{r},$D$2:$D$11,0))"
        ws[f"H{r}"].number_format = USD
        ws[f"G{r}"].fill = ws[f"H{r}"].fill = FORMULA_FILL
    ws.column_dimensions["F"].width = 8
    ws.column_dimensions["G"].width = 26
    ws.column_dimensions["H"].width = 13

    notes_sheet(wb, [
        "Tidy Tabs: Rank Customers by Sales",
        "",
        "1. Type customer names in A and total sales in B.",
        "2. Rank = RANK(B2,$B$2:$B$11,0). The 0 ranks the biggest total as 1. Use 1 instead of 0 to rank the smallest as 1.",
        "3. Tied totals get the same rank, and the next rank is skipped (two customers at rank 2 means no rank 3).",
        "4. Unique rank = Rank + COUNTIF($B$2:B2,B2) - 1. The first customer with a total keeps the rank, the next one with the same total gets rank + 1.",
        "5. The Top 3 block uses INDEX/MATCH on the unique rank, so it never hits a missing rank.",
        "6. Lock the range with $ signs ($B$2:$B$11) so it does not slide when you fill down. Extend it when you add customers.",
        "",
        "Customers and totals are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-rank-top-customers.xlsx")


def invoice_number_generator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    header(ws, ["Client", "Invoice date", "Invoice number (TEXT)", "Invoice number (YEAR)", "Amount"],
           [26, 13, 22, 22, 12])
    rows = [
        ("Harbor Coffee Co.", date(2026, 9, 28), 450.00),
        ("Maple Street Bakery", date(2026, 9, 30), 1200.00),
        ("Riverside Yoga", date(2026, 10, 2), 315.00),
        ("Juniper Florist", date(2026, 10, 5), 780.00),
        ("Bluebird Books", date(2026, 10, 9), 95.00),
        ("Cedar Hardware", date(2026, 10, 12), 640.00),
    ]
    for r, (client, d, amt) in enumerate(rows, start=2):
        ws.append([client, d,
                   f'="INV-"&TEXT(B{r},"yyyy")&"-"&TEXT(ROWS($A$2:A{r}),"0000")',
                   f'="INV-"&YEAR(B{r})&"-"&TEXT(ROWS($A$2:A{r}),"0000")', amt])
        ws[f"B{r}"].number_format = US_DATE
        ws[f"E{r}"].number_format = USD
        for col in "CD":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Automatic Invoice Numbers",
        "",
        "1. Type the client and invoice date. Copy the gray formulas down for each new row.",
        "2. Column C: =\"INV-\"&TEXT(B2,\"yyyy\")&\"-\"&TEXT(ROWS($A$2:A2),\"0000\"). TEXT(...,\"0000\") pads the count to four digits; ROWS counts how many rows are in the range, so each new row gets the next number.",
        "3. Column D does the year with YEAR(B2) instead of a TEXT format code. Use it if a date format code such as yyyy does not work in your language or locale.",
        "4. The number follows the row position, not a saved counter. If you sort or delete rows, the numbers shift. Once an invoice is sent, copy the numbers and Paste Special > Values only so they stay fixed.",
        "5. The year comes from the invoice date, so the count does not restart on its own each January. Start a new sheet or tab each year.",
        "",
        "Clients and amounts are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-invoice-number-generator.xlsx")


def combine_address_columns():
    wb = Workbook()
    ws = wb.active
    ws.title = "Addresses"
    header(ws, ["Street", "Unit", "City", "State", "ZIP", "TEXTJOIN (skips blanks)", "With ZIP", "Plain & join (stray commas)"],
           [22, 9, 14, 7, 8, 38, 44, 38])
    rows = [
        ("118 Elm Street", "Apt 4B", "Boston", "MA", "02108"),
        ("2450 Harbor Road", "", "Portland", "ME", "04101"),
        ("77 Mill Lane", "Suite 210", "Hoboken", "NJ", "07030"),
        ("910 Cedar Avenue", "", "Providence", "RI", "02903"),
        ("36 Orchard Way", "Unit 2", "Albany", "NY", "12207"),
    ]
    for r, row in enumerate(rows, start=2):
        ws.append(list(row) + [f'=_xlfn.TEXTJOIN(", ",TRUE,A{r}:D{r})',
                               f'=_xlfn.TEXTJOIN(", ",TRUE,A{r}:D{r})&" "&E{r}',
                               f'=A{r}&", "&B{r}&", "&C{r}&", "&D{r}'])
        ws[f"E{r}"].number_format = "@"
        for col in "FGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Combine Address Columns with TEXTJOIN",
        "",
        "1. Put street, unit, city, and state in A to D and the ZIP in E (formatted as Text so 02108 keeps its zero).",
        "2. Column F: =TEXTJOIN(\", \",TRUE,A2:D2). The first argument is the separator, TRUE skips empty cells, and the last is the range to join.",
        "3. Column G adds the ZIP after a space, so the state and ZIP are not split by a comma.",
        "4. Column H shows the older way, joining with &. When Unit is empty it leaves a double comma, which is what TEXTJOIN avoids.",
        "5. TEXTJOIN is in Excel 2019 and Microsoft 365, and in Google Sheets. It is not in Excel 2016 or earlier.",
        "",
        "Addresses are fictional. Checked in LibreOffice Calc only.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-combine-address-columns.xlsx")


def sale_price_discount():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sale prices"
    header(ws, ["Item", "Original price", "Percent off", "Sale price", "You save"],
           [30, 14, 12, 12, 12])
    rows = [
        ("Lavender soy candle, 8 oz", 18.00, 0.10),
        ("Cedar soy candle, 8 oz", 18.00, 0.15),
        ("Oatmeal goat milk soap bar", 8.00, 0.25),
        ("Charcoal soap bar", 8.50, 0.10),
        ("Brass hoop earrings", 24.00, 0.15),
        ("Letterpress birthday card", 6.00, 0.25),
        ("Beaded necklace", 48.00, 0.25),
        ("Candle and soap gift set", 32.00, 0.10),
    ]
    for r, (item, price, pct) in enumerate(rows, start=2):
        ws.append([item, price, pct, f"=B{r}*(1-C{r})", f"=B{r}-D{r}"])
        for col in "BDE":
            ws[f"{col}{r}"].number_format = USD
        ws[f"C{r}"].number_format = "0%"
        for col in "DE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["G1"], ws["H1"] = "Total at original prices", "=SUM(B2:B500)"
    ws["G2"], ws["H2"] = "Total at sale prices", "=SUM(D2:D500)"
    ws["G3"], ws["H3"] = "Total saved", "=SUM(E2:E500)"
    for r in (1, 2, 3):
        ws[f"G{r}"].font = Font(bold=True)
        ws[f"H{r}"].number_format = USD
    ws.column_dimensions["G"].width = 24
    ws.column_dimensions["H"].width = 13

    rv = wb.create_sheet("Reverse and stacked")
    header(rv, ["Item", "Where you saw it", "Percent off", "Sale price", "Original price"],
           [28, 18, 12, 12, 14])
    for r, (item, where, pct, sale) in enumerate([
        ("Beaded necklace", "Market tag", 0.25, 36.00),
        ("Brass hoop earrings", "Etsy listing", 0.15, 20.40),
        ("Charcoal soap bar", "Receipt", 0.10, 7.65),
    ], start=2):
        rv.append([item, where, pct, sale, f"=D{r}/(1-C{r})"])
        rv[f"C{r}"].number_format = "0%"
        for col in "DE":
            rv[f"{col}{r}"].number_format = USD
        rv[f"E{r}"].fill = FORMULA_FILL
    rv["A7"] = "Two discounts in a row"
    rv["A7"].font = Font(bold=True, size=12)
    labels = [
        ("Original price", 100.00),
        ("First discount", 0.20),
        ("Second discount", 0.10),
        ("Price after both, one at a time", "=B8*(1-B9)*(1-B10)"),
        ("Wrong: add the percents (30% off)", "=B8*(1-(B9+B10))"),
        ("Difference", "=B11-B12"),
        ("Real total discount", "=1-B11/B8"),
    ]
    for r, (label, val) in enumerate(labels, start=8):
        rv[f"A{r}"], rv[f"B{r}"] = label, val
    for r in (8, 11, 12, 13):
        rv[f"B{r}"].number_format = USD
    for r in (9, 10, 14):
        rv[f"B{r}"].number_format = "0%"
    for r in (11, 12, 13, 14):
        rv[f"B{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Sale Price and Discount Calculator",
        "",
        "1. Sale prices tab: type the Original price and Percent off. Sale price = B2*(1-C2). You save = B2-D2.",
        "2. Format Percent off as a percentage and type 25% (or 0.25), not 25.",
        "3. Reverse and stacked tab, top: you know the sale price and the percent off and want the original price. Original price = D2/(1-C2).",
        "4. Reverse and stacked tab, bottom: 20% off and then 10% off is not 30% off. $100.00 becomes $80.00, then $72.00, a real discount of 28%.",
        "5. Totals for the item list are in H1:H3.",
        "",
        "Items and prices are fictional. This is arithmetic only; it does not include sales tax or shipping.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-sale-price-discount-calculator.xlsx")


def shipping_weight_tier_lookup():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order #", "Weight (oz)", "Rate (VLOOKUP)", "Rate (INDEX/MATCH)", "VLOOKUP without IFERROR",
                "", "From (oz)", "Rate"],
           [10, 12, 15, 18, 22, 4, 11, 10])
    # Lower edge of each band, smallest to largest. An exact hit on an edge
    # belongs to that band, so 4 oz is in the 4 oz band.
    table = [(1, 4.50), (4, 5.75), (8, 7.25), (16, 9.80), (32, 13.40), (64, 18.90)]
    weights = [2.5, 3.99, 4, 4.01, 7.9, 8, 12.5, 16, 20, 40, 64, 0.5]
    for r, w in enumerate(weights, start=2):
        ws.append([str(2100 + r - 1), w,
                   f'=IFERROR(VLOOKUP(B{r},$G$2:$H$7,2,TRUE),"Below minimum")',
                   f'=IFERROR(INDEX($H$2:$H$7,MATCH(B{r},$G$2:$G$7,1)),"Below minimum")',
                   f"=VLOOKUP(B{r},$G$2:$H$7,2,TRUE)"])
        for col in "CDE":
            ws[f"{col}{r}"].number_format = USD
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"B{r}"].number_format = "0.00"
    for r, (lo, rate) in enumerate(table, start=2):
        ws[f"G{r}"], ws[f"H{r}"] = lo, rate
        ws[f"H{r}"].number_format = USD
    for c in ("G1", "H1"):
        ws[c].fill, ws[c].font = HEADER_FILL, HEADER_FONT

    un = wb.create_sheet("Unsorted table demo")
    header(un, ["Weight (oz)", "VLOOKUP result", "Correct rate (sorted table)", "", "From (oz)", "Rate"],
           [12, 16, 24, 4, 11, 10])
    unsorted = [(1, 4.50), (8, 7.25), (4, 5.75), (16, 9.80), (32, 13.40), (64, 18.90)]
    for r, w in enumerate([5, 10], start=2):
        un.append([w, f"=VLOOKUP(A{r},$E$2:$F$7,2,TRUE)",
                   f"=VLOOKUP(A{r},Orders!$G$2:$H$7,2,TRUE)"])
        for col in "BC":
            un[f"{col}{r}"].number_format = USD
        un[f"B{r}"].fill = ALERT_FILL
    for r, (lo, rate) in enumerate(unsorted, start=2):
        un[f"E{r}"], un[f"F{r}"] = lo, rate
        un[f"F{r}"].number_format = USD

    notes_sheet(wb, [
        "Tidy Tabs: Shipping Cost by Weight Tier (approximate match lookup)",
        "",
        "1. Rate table in G:H, Orders tab. From (oz) is the smallest weight in each band. A package at or above that weight and below the next row uses that rate.",
        "2. Rate (VLOOKUP) = IFERROR(VLOOKUP(B2,$G$2:$H$7,2,TRUE),\"Below minimum\"). TRUE (or leaving it out) means approximate match: the largest From value that is less than or equal to the weight.",
        "3. Rate (INDEX/MATCH) does the same with MATCH(...,1). Both columns agree.",
        "4. The From column must be sorted smallest to largest. Approximate match skips through the list instead of checking every row, so an unsorted table can return a wrong rate with no error. See the Unsorted table demo tab: in LibreOffice 5 oz happens to come out right ($5.75) but 10 oz returns $5.75 instead of $7.25. Other apps can pick differently on an unsorted table.",
        "5. A weight under the first band (0.50 oz here) has no band, so VLOOKUP returns #N/A. Column E shows it; IFERROR in column C turns it into text. IFERROR also hides other errors, so test the table first.",
        "6. Exactly 4 oz is in the 4 oz band ($5.75). 3.99 oz is in the 1 oz band ($4.50). 4.01 oz is in the 4 oz band.",
        "",
        "Rates and weight bands are made up for this sample, not a carrier's rates. Use your own.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-shipping-weight-tier-lookup.xlsx")


def reorder_point_calculator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Reorder"
    header(ws, ["Supply", "Units sold (last 30 days)", "Daily sales", "Lead time (days)", "Safety stock",
                "Reorder point", "On hand", "Flag"],
           [28, 14, 11, 12, 11, 12, 10, 12])
    rows = [
        ("Glass jars, 8 oz", 90, 7, 6, 40),
        ("Soy wax, 1 lb bags", 60, 10, 8, 28),
        ("Candle wicks, pack of 50", 33, 14, 4, 21),
        ("Cotton ribbon spools", 75, 5, 10, 18),
        ("Kraft mailer boxes", 120, 12, 15, 80),
        ("Shipping label rolls", 45, 21, 5, 30),
    ]
    for r, (item, sold, lead, safety, onhand) in enumerate(rows, start=2):
        ws.append([item, sold, f"=B{r}/$K$1", lead, safety,
                   f"=ROUNDUP(C{r}*D{r}+E{r},0)", onhand,
                   f'=IF(G{r}<=F{r},"REORDER","OK")'])
        ws[f"C{r}"].number_format = "0.00"
        for col in "CFH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    last = 200
    ws["J1"], ws["K1"] = "Days in sales period", 30
    ws["J1"].font = Font(bold=True)
    ws.column_dimensions["J"].width = 22
    ws.conditional_formatting.add(f"H2:H{last}", FormulaRule(formula=['$H2="REORDER"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Reorder Point Calculator",
        "",
        "1. One row per supply. Type Units sold for the last 30 days, Lead time (days) from order to delivery, Safety stock (spare units you want to keep), and On hand.",
        "2. Daily sales = Units sold / days in the period. The period length is in K1 (30).",
        "3. Reorder point = ROUNDUP(Daily sales x Lead time + Safety stock, 0). ROUNDUP, because a fraction of a unit still needs a whole unit.",
        "4. Flag = IF(On hand <= Reorder point, \"REORDER\", \"OK\"). At exactly the reorder point the flag says REORDER.",
        "5. Example: 90 units in 30 days is 3 a day. 3 x 7 lead days + 6 safety stock = 27.",
        "6. Copy the last row down to add supplies.",
        "",
        "Supplies and numbers are fictional. Safety stock is your own judgment; this file only does the arithmetic.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-reorder-point-calculator.xlsx")


def freelance_hourly_rate():
    wb = Workbook()
    ws = wb.active
    ws.title = "Hourly rate"
    header(ws, ["", "Your numbers", "What if 1: higher goal", "What if 2: more time off", "What if 3: more billable"],
           [36, 15, 18, 18, 18])
    inputs = [
        ("Yearly income goal", 60000, 75000, 60000, 60000, USD),
        ("Yearly overhead (software, supplies)", 6000, 6000, 6000, 6000, USD),
        ("Weeks off per year", 4, 4, 6, 4, "0"),
        ("Hours per week you work", 30, 30, 30, 30, "0"),
        ("Share of those hours you bill", 0.65, 0.65, 0.65, 0.75, "0%"),
    ]
    for r, (label, *vals, fmt) in enumerate(inputs, start=2):
        ws.append([label] + vals)
        for col in "BCDE":
            ws[f"{col}{r}"].number_format = fmt
    calc = [
        ("Weeks worked", "=52-{c}4", "0"),
        ("Hours worked per year", "={c}7*{c}5", "#,##0"),
        ("Billable hours per year", "={c}8*{c}6", "#,##0"),
        ("Money you need to bring in", "={c}2+{c}3", USD),
        ("Hourly rate", "={c}10/{c}9", USD),
        ("Rate rounded up to a whole dollar", "=ROUNDUP({c}11,0)", USD),
    ]
    for r, (label, f, fmt) in enumerate(calc, start=7):
        ws[f"A{r}"] = label
        for col in "BCDE":
            ws[f"{col}{r}"] = f.format(c=col)
            ws[f"{col}{r}"].number_format = fmt
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["A11"].font = Font(bold=True)
    for col in "BCDE":
        ws[f"{col}11"].font = Font(bold=True)
    ws.freeze_panes = "B2"

    notes_sheet(wb, [
        "Tidy Tabs: Freelance Hourly Rate Calculator",
        "",
        "1. Type your numbers in column B, rows 2 to 6. Columns C to E are what-if copies: change any input there and the rate below it updates.",
        "2. Weeks worked = 52 - weeks off. Hours worked = weeks worked x hours per week. Billable hours = hours worked x billable share.",
        "3. Money you need = income goal + overhead. Hourly rate = money you need / billable hours.",
        "4. Example: (52-4) x 30 x 65% = 936 billable hours. ($60,000.00 + $6,000.00) / 936 = $70.51 an hour.",
        "5. Billable share is the part of your work week you can actually charge for. Time spent finding clients, invoicing and emailing is not billable.",
        "",
        "Illustration only. This does not cover taxes, self-employment costs, or benefits; ask an accountant about those. All numbers are made up.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-freelance-hourly-rate-calculator.xlsx")


def convert_text_to_numbers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Pasted amounts"
    header(ws, ["Pasted text", "Is it a number?", "Basic clean", "Clean with (negatives)", "Is it a number now?"],
           [16, 14, 14, 22, 18])
    texts = ["$1,250.00", " 18.00", "(45.50)", "$89.95", "$2,400.00 ", "12.50",
             "$7.25", "$1,075.40", "$310.00", "(22.00)"]
    for r, t in enumerate(texts, start=2):
        ws.append([t, f"=ISNUMBER(A{r})",
                   f'=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A{r}),"$",""),",",""))',
                   f'=IF(LEFT(TRIM(A{r}),1)="(",-1,1)*VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(A{r}),"$",""),",",""),"(",""),")",""))',
                   f"=ISNUMBER(D{r})"])
        ws[f"A{r}"].number_format = "@"
        ws[f"A{r}"].alignment = Alignment(horizontal="left")
        for col in "CD":
            ws[f"{col}{r}"].number_format = USD
        for col in "BCDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    labels = [("SUM of the pasted text", "=SUM(A2:A500)"),
              ("SUM of Basic clean", "=SUM(C2:C500)"),
              ("SUM of Clean with (negatives)", "=SUM(D2:D500)")]
    for r, (label, f) in enumerate(labels, start=1):
        ws[f"G{r}"], ws[f"H{r}"] = label, f
        ws[f"G{r}"].font = Font(bold=True)
        ws[f"H{r}"].number_format = USD
    ws.column_dimensions["G"].width = 30
    ws.column_dimensions["H"].width = 14

    notes_sheet(wb, [
        "Tidy Tabs: Convert Text to Numbers (pasted dollar amounts)",
        "",
        "1. Column A holds amounts as text, the way they arrive when you paste from a bank or Etsy export. SUM ignores text, so H1 is $0.00.",
        "2. Column B: =ISNUMBER(A2) is FALSE for text and TRUE for a real number.",
        "3. Column C: =VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),\"$\",\"\"),\",\",\"\")). TRIM removes stray spaces, SUBSTITUTE removes $ and commas, VALUE turns the text into a number.",
        "4. Column D also handles amounts in parentheses, such as (45.50), as negatives: it removes the ( and ) and multiplies by -1. LibreOffice reads (45.50) as -45.50 even in column C, but do not count on that in other apps; column D does not depend on it.",
        "5. Column E checks that the cleaned values are real numbers. Copy C or D, then Paste Special > Values to keep the numbers.",
        "",
        "Amounts are fictional. Checked in LibreOffice Calc with a US locale only; other regional settings can read $, commas and periods differently.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-convert-text-to-numbers.xlsx")


def sales_tax_calculator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales tax"
    header(ws, ["Item", "Price before tax", "Sales tax", "Total"], [28, 16, 12, 12])
    rows = [
        ("Ceramic mug", 18.00),
        ("Sticker pack", 4.50),
        ("Canvas tote bag", 24.00),
        ("Soy candle, 8 oz", 16.00),
        ("Art print, 8 x 10", 12.99),
        ("Beaded earrings", 32.50),
    ]
    for r, (item, price) in enumerate(rows, start=2):
        ws.append([item, price, f"=ROUND(B{r}*$F$1,2)", f"=B{r}+C{r}"])
        for col in "BCD":
            ws[f"{col}{r}"].number_format = USD
        for col in "CD":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["A8"] = "Total"
    ws["A8"].font = Font(bold=True)
    for col in "BCD":
        ws[f"{col}8"] = f"=SUM({col}2:{col}7)"
        ws[f"{col}8"].number_format = USD
        ws[f"{col}8"].font = Font(bold=True)
    ws["E1"], ws["F1"] = "Sales tax rate", 0.0625
    ws["E1"].font = Font(bold=True)
    ws["F1"].number_format = "0.00%"
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 10

    inc = wb.create_sheet("Tax-inclusive totals")
    header(inc, ["Item", "Total paid (tax included)", "Tax inside the total", "Price before tax"], [28, 16, 16, 16])
    paid = [
        ("Ceramic mug", 19.13),
        ("Sticker pack", 4.78),
        ("Canvas tote bag", 25.50),
        ("Soy candle, 8 oz", 17.00),
        ("Art print, 8 x 10", 13.80),
        ("Beaded earrings", 34.53),
        ("Gift set, one price at a booth", 50.00),
    ]
    for r, (item, total) in enumerate(paid, start=2):
        inc.append([item, total, f"=ROUND(B{r}-B{r}/(1+$F$1),2)", f"=B{r}-C{r}"])
        for col in "BCD":
            inc[f"{col}{r}"].number_format = USD
        for col in "CD":
            inc[f"{col}{r}"].fill = FORMULA_FILL
    inc["E1"], inc["F1"] = "Sales tax rate", "='Sales tax'!F1"
    inc["E1"].font = Font(bold=True)
    inc["F1"].number_format = "0.00%"
    inc.column_dimensions["E"].width = 16
    inc.column_dimensions["F"].width = 10

    notes_sheet(wb, [
        "Tidy Tabs: Sales Tax Calculator",
        "",
        "1. Type your rate in F1 on the Sales tax tab (6.25% is a made-up example). Rates vary by state, county and city, so check your state's revenue department site for the real rate.",
        "2. Sales tax = ROUND(Price x rate, 2). ROUND to cents because you cannot collect a fraction of a cent.",
        "3. Total = Price + Sales tax. Row 8 adds up each column.",
        "4. Tax-inclusive totals tab: type what the customer paid in column B. Tax inside = ROUND(Total - Total / (1 + rate), 2). Price before tax = Total - Tax.",
        "5. The rate on the second tab follows F1 on the first tab.",
        "6. Example: $18.00 x 6.25% = $1.125, rounded to $1.13, so the total is $19.13.",
        "",
        "Items, prices and the rate are fictional. This sheet only does arithmetic and is not tax advice.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-sales-tax-calculator.xlsx")




def sum_expenses_by_category():
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    header(ws, ["Date", "Vendor", "Category", "Amount", "", "Category", "Total spent", "% of total"],
           [12, 24, 13, 12, 3, 16, 14, 11])
    log = [
        (date(2026, 9, 1), "Greenline Office Supply", "Supplies", 84.50),
        (date(2026, 9, 2), "Lakeview Web Hosting", "Software", 29.00),
        (date(2026, 9, 3), "Pioneer Postal Center", "Shipping", 46.25),
        (date(2026, 9, 5), "Metro Transit Pass", "Travel", 90.00),
        (date(2026, 9, 6), "Harbor Coffee", "Meals", 18.40),
        (date(2026, 9, 8), "Northeast Electric", "Utilities", 112.30),
        (date(2026, 9, 9), "Social Ads Co", "Advertising", 150.00),
        (date(2026, 9, 10), "Greenline Office Supply", "Supplies", 37.80),
        (date(2026, 9, 12), "Pioneer Postal Center", "Shipping", 62.10),
        (date(2026, 9, 13), "Cloud Backup Plus", "Software", 12.99),
        (date(2026, 9, 15), "Oak Street Deli", "Meals", 24.75),
        (date(2026, 9, 16), "Harbor Parking Garage", "Travel", 22.00),
        (date(2026, 9, 17), "Social Ads Co", "Advertising", 75.00),
        (date(2026, 9, 19), "Craft Paper Wholesale", "Supplies", 128.60),
        (date(2026, 9, 20), "Northeast Gas", "Utilities", 58.45),
        (date(2026, 9, 22), "Pioneer Postal Center", "Shipping", 33.90),
        (date(2026, 9, 23), "Design Template Shop", "Software", 49.00),
        (date(2026, 9, 24), "Amtrak Regional", "Travel", 134.00),
        (date(2026, 9, 25), "Maple Cafe", "Meals", 31.20),
        (date(2026, 9, 26), "Craft Paper Wholesale", "Supplies", 66.40),
        (date(2026, 9, 27), "Local Print Shop", "Advertising", 95.50),
        (date(2026, 9, 29), "Lakeview Internet", "Utilities", 69.99),
    ]
    for r, (d, vendor, cat, amt) in enumerate(log, start=2):
        ws.append([d, vendor, cat, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = USD
    cats = ["Supplies", "Software", "Shipping", "Travel", "Meals", "Utilities", "Advertising"]
    for r, cat in enumerate(cats, start=2):
        ws[f"F{r}"] = cat
        ws[f"G{r}"] = f"=SUMIF($C$2:$C$40,F{r},$D$2:$D$40)"
        ws[f"H{r}"] = f"=G{r}/$G$9"
        ws[f"G{r}"].number_format = USD
        ws[f"H{r}"].number_format = "0.0%"
        ws[f"G{r}"].fill = ws[f"H{r}"].fill = FORMULA_FILL
    ws["F9"], ws["G9"], ws["H9"] = "Total", "=SUM(G2:G8)", "=SUM(H2:H8)"
    ws["G9"].number_format, ws["H9"].number_format = USD, "0.0%"
    ws["F11"], ws["G11"] = "Log total", "=SUM($D$2:$D$40)"
    ws["F12"], ws["G12"] = "Difference", "=ROUND(G11-G9,2)"
    ws["F13"], ws["G13"] = "Check", '=IF(G12=0,"OK","CHECK CATEGORIES")'
    for c in ("F9", "G9", "H9", "F11", "F12", "F13"):
        ws[c].font = Font(bold=True)
    ws["G11"].number_format = ws["G12"].number_format = USD
    ws["G11"].fill = ws["G12"].fill = ws["G13"].fill = FORMULA_FILL
    ws.conditional_formatting.add("G13", FormulaRule(formula=['$G$13<>"OK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Sum Expenses by Category",
        "",
        "1. Log one expense per row in A to D (Date, Vendor, Category, Amount). The formulas read rows 2 to 40, so you can add up to 39 rows without editing anything.",
        "2. Type each category exactly the same way every time. The category list is in F2:F8.",
        "3. Total spent = SUMIF($C$2:$C$40,F2,$D$2:$D$40). It adds column D wherever column C matches the category in column F.",
        "4. % of total = G2/$G$9, where G9 is the sum of the category totals.",
        "5. Check: G11 = SUM($D$2:$D$40), every amount in the log. G12 is G11 minus G9. If it is not $0.00, G13 says CHECK CATEGORIES and some expense has a category that is not in the list.",
        "6. Common cause: a typo or a trailing space, such as \"Supplies \" with a space after it. Retype the category, or clean the column with TRIM.",
        "7. To add a category, insert a row between rows 2 and 8 (so the Total row still covers it) and copy the formulas from the row above.",
        "",
        "Vendors and amounts are fictional. This file only adds up what you type; it is not tax or accounting advice.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-sum-expenses-by-category.xlsx")


def round_prices_nearest_nickel_99():
    wb = Workbook()
    ws = wb.active
    ws.title = "Prices"
    header(ws, ["Item", "Price after markup math", "Nearest cent", "Nearest $0.05", "Up to next $0.05",
                "Ends in .99"],
           [30, 16, 13, 14, 15, 13])
    rows = [
        ("Beeswax candle, 4 oz", 13.4167),
        ("Cotton tote bag", 8.2333),
        ("Hand-poured soap set", 24.7083),
        ("Ceramic mug", 17.9625),
        ("Greeting card, single", 6.13),
        ("Linen napkins, set of 4", 31.4467),
        ("Sticker sheet", 12.9833),
        ("Embroidered patch", 22.0),
    ]
    for r, (item, price) in enumerate(rows, start=2):
        ws.append([item, price, f"=ROUND(B{r},2)", f"=MROUND(B{r},0.05)", f"=CEILING(B{r},0.05)",
                   f"=ROUNDUP(B{r},0)-0.01"])
        ws[f"B{r}"].number_format = '"$"#,##0.0000'
        for col in "CDEF":
            ws[f"{col}{r}"].number_format = USD
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Round Prices to the Nearest $0.05 or .99",
        "",
        "1. Type your unrounded prices in column B (the result of your markup or discount math).",
        "2. Nearest cent: =ROUND(B2,2). Nearest nickel: =MROUND(B2,0.05). Always up to the next nickel: =CEILING(B2,0.05).",
        "3. Ending in .99: =ROUNDUP(B2,0)-0.01. It rounds up to the next whole dollar, then takes off a cent.",
        "4. Watch out: a price that is already a whole dollar, such as $22.00, becomes $21.99 with the .99 formula.",
        "5. Example: $13.4167 becomes $13.42 (cent), $13.40 (nickel), $13.45 (up) and $13.99 (.99).",
        "6. Copy the last row down to add items.",
        "",
        "Items and prices are fictional. Checked in LibreOffice Calc only.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-round-prices-nearest-nickel-99.xlsx")


def remove_non_breaking_spaces():
    wb = Workbook()
    ws = wb.active
    ws.title = "Pasted names"
    header(ws, ["Pasted name", "LEN before", "Code at position 5", "TRIM only", "Cleaned",
                "LEN after", "COUNTIF before", "COUNTIF after", "Plan price before", "Plan price after",
                "", "Customer list", "Plan price"],
           [18, 10, 12, 14, 14, 10, 12, 12, 14, 14, 3, 16, 12])
    nb = chr(160)  # built in with Python so the cells hold a real non-breaking space
    raw = [
        f"Dana{nb}Ruiz",
        f"Emma{nb}Lee",
        f"Sara{nb}Kim{nb}",
        "Mark Diaz\n",
        f"Lena{nb}Ortiz\t",
        f"Josh{nb}{nb}Park",
        f"Anna{nb}Cho",
        "Mike Stone",
    ]
    book = [("Dana Ruiz", 120), ("Emma Lee", 45), ("Sara Kim", 89.5), ("Mark Diaz", 210),
            ("Lena Ortiz", 65), ("Josh Park", 150), ("Anna Cho", 38), ("Mike Stone", 72.25)]
    for r, name in enumerate(raw, start=2):
        ws.append([name, f"=LEN(A{r})", f"=_xlfn.UNICODE(MID(A{r},5,1))", f"=TRIM(A{r})",
                   # UNICHAR/UNICODE, not CHAR/CODE: LibreOffice maps CHAR(160) and CODE through
                   # the system locale (seen: 160 in one profile, 194 in a fresh one).
                   f'=TRIM(CLEAN(SUBSTITUTE(A{r},_xlfn.UNICHAR(160)," ")))', f"=LEN(E{r})",
                   f"=COUNTIF($L$2:$L$9,A{r})", f"=COUNTIF($L$2:$L$9,E{r})",
                   f'=IFERROR(VLOOKUP(A{r},$L$2:$M$9,2,FALSE),"not found")',
                   f'=IFERROR(VLOOKUP(E{r},$L$2:$M$9,2,FALSE),"not found")'])
        ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        for col in "BCDEFGHIJ":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        for col in "IJ":
            ws[f"{col}{r}"].number_format = USD
            ws[f"{col}{r}"].alignment = Alignment(horizontal="right")
        ws[f"L{r}"], ws[f"M{r}"] = book[r - 2]
        ws[f"M{r}"].number_format = USD
    ws["L11"], ws["M11"] = "Found before", "=SUM(G2:G9)"
    ws["L12"], ws["M12"] = "Found after", "=SUM(H2:H9)"
    for c in ("L11", "L12"):
        ws[c].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Remove Non-Breaking Spaces (CHAR 160) in Excel and Google Sheets",
        "",
        "1. Column A holds fictional names pasted the way web pages and PDFs deliver them: most have a non-breaking space (character 160) where a normal space should be, and some have a trailing line break or tab.",
        "2. Column B: =LEN(A2) counts every character, including the invisible ones. Dana Ruiz shows 9 instead of 8.",
        "3. Column C: =UNICODE(MID(A2,5,1)) shows the character code at position 5 (=CODE(MID(A2,5,1)) is the older version, but in LibreOffice it returned 160 in one setup and 194 in another, depending on the locale). A normal space is 32, a non-breaking space is 160. This works here because every first name has 4 letters; change the 5 to point at the character you suspect.",
        "4. Column D: =TRIM(A2) only removes the regular space character (code 32), so the non-breaking spaces stay. Compare its length to column B.",
        "5. Column E: =TRIM(CLEAN(SUBSTITUTE(A2,UNICHAR(160),\" \"))). UNICHAR(160) is used instead of CHAR(160) because LibreOffice returned a different character for CHAR(160) in one setup. SUBSTITUTE swaps character 160 for a normal space, CLEAN removes line breaks and tabs, TRIM removes extra spaces.",
        "6. CLEAN deletes a line break or tab instead of replacing it with a space. If one sits between two words, swap it first: SUBSTITUTE(A2,CHAR(10),\" \").",
        "7. Columns G to J look each name up in the customer list in L:M. Before cleaning, only Mike Stone is found. After cleaning, all 8 are found.",
        "8. To keep the results, copy column E and Paste Special > Values only, then delete the gray formula columns.",
        "",
        "Names and prices are fictional. Checked in LibreOffice Calc 24.2.7 (Linux) only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-remove-non-breaking-spaces.xlsx")


def year_to_date_sales_total():
    rows = [
        (date(2026, 1, 12), "Harbor Coffee", 1250.00),
        (date(2026, 1, 28), "Maple Street Bakery", 480.00),
        (date(2026, 2, 10), "Oak & Ember Candles", 920.50),
        (date(2026, 2, 24), "Harbor Coffee", 1100.00),
        (date(2026, 4, 7), "Maple Street Bakery", 640.00),
        (date(2026, 5, 19), "Harbor Coffee", 1380.75),
        (date(2026, 6, 23), "Oak & Ember Candles", 705.00),
        (date(2026, 7, 8), "Lakeview Florist", 450.00),
        (date(2026, 7, 30), "Harbor Coffee", 1015.00),
        (date(2026, 8, 4), "Maple Street Bakery", 520.50),
        (date(2026, 8, 18), "Oak & Ember Candles", 860.00),
        (date(2026, 8, 31), "Lakeview Florist", 275.00),
        (date(2026, 9, 12), "Harbor Coffee", 1190.00),
        (date(2026, 10, 3), "Lakeview Florist", 330.00),
    ]

    def fill_sheet(ws, text_row=None):
        header(ws, ["Sale date", "Customer", "Amount"], [13, 26, 12])
        for r, (d, who, amt) in enumerate(rows, start=2):
            if r == text_row:
                ws.append([d.strftime("%m/%d/%Y"), who, amt])
                ws[f"A{r}"].number_format = "@"
            else:
                ws.append([d, who, amt])
                ws[f"A{r}"].number_format = US_DATE
            ws[f"C{r}"].number_format = USD
        ws.column_dimensions["E"].width = 34
        ws.column_dimensions["F"].width = 14
        ws["E1"], ws["F1"] = "Report date (type a date)", date(2026, 8, 31)
        ws["F1"].number_format = US_DATE
        ws["F1"].fill = PatternFill("solid", fgColor="FFF2CC")
        labels = [
            ("Year-to-date sales", '=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),1,1),A2:A500,"<="&F1)'),
            ("Month-to-date sales", '=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),MONTH(F1),1),A2:A500,"<="&F1)'),
            ("Whole report month (EOMONTH)", '=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),MONTH(F1),1),A2:A500,"<="&EOMONTH(F1,0))'),
            ("Start of year", "=DATE(YEAR(F1),1,1)"),
            ("Start of month", "=DATE(YEAR(F1),MONTH(F1),1)"),
            ("Month end", "=EOMONTH(F1,0)"),
            ("Rows with a real date (check)", "=COUNT(A2:A500)"),
            ("Rows with an amount (check)", "=COUNT(C2:C500)"),
        ]
        for i, (lab, f) in enumerate(labels, start=2):
            ws[f"E{i}"], ws[f"F{i}"] = lab, f
            ws[f"E{i}"].font = Font(bold=True)
            ws[f"F{i}"].fill = FORMULA_FILL
            ws[f"F{i}"].number_format = USD if i <= 4 else (US_DATE if i <= 7 else "0")
        ws["E1"].font = Font(bold=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    fill_sheet(ws)
    demo = wb.create_sheet("Text date demo")
    fill_sheet(demo, text_row=12)  # 08/18/2026 typed as text
    demo["E11"] = "Row 12 (08/18/2026) is stored as text on purpose."
    demo["E11"].font = Font(italic=True, color="C00000")

    notes_sheet(wb, [
        "Tidy Tabs: Year-to-Date Sales Total",
        "",
        "1. Sales tab: one row per sale. Type the Sale date (a real date), Customer, and Amount.",
        "2. F1 is the Report date. Type the date you want totals through. The sample uses 08/31/2026 so the numbers stay put; you can type =TODAY() instead for a live total.",
        "3. Year-to-date (F2) = SUMIFS(C2:C500, A2:A500, \">=\"&DATE(YEAR(F1),1,1), A2:A500, \"<=\"&F1). It adds amounts dated January 1 of the report year through the report date.",
        "4. Month-to-date (F3) is the same with DATE(YEAR(F1),MONTH(F1),1) as the start. F4 uses EOMONTH(F1,0) as the end to total the whole month.",
        "5. Change F1 and watch F2 to F4 move. Sales after the report date, like the 09/12/2026 and 10/03/2026 rows, are left out.",
        "6. Text date demo tab: the same data, but the 08/18/2026 date is stored as text. SUMIFS skips it; compare the totals with the Sales tab. F8 and F9 count real dates and amounts, so a gap means a text date.",
        "7. Add rows under the last sale. The formulas read down to row 500.",
        "",
        "Customers and amounts are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-year-to-date-sales-total.xlsx")


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    inventory()
    projects()
    zip_codes()
    print_one_page()
    split_city_state_zip()
    dependent_dropdown()
    client_contact_list()
    packing_slip()
    running_balance()
    highlight_duplicates()
    business_days_ship_by()
    time_to_decimal_hours()
    receivables_aging()
    markup_vs_margin()
    vlookup_price_list()
    break_even_calculator()
    subscription_renewal_tracker()
    budget_vs_actual()
    farmers_market_sales_log()
    clean_customer_list()
    sale_price_discount()
    shipping_weight_tier_lookup()
    reorder_point_calculator()
    freelance_hourly_rate()
    convert_text_to_numbers()
    sales_tax_calculator()
    sum_expenses_by_category()
    round_prices_nearest_nickel_99()
    remove_non_breaking_spaces()
    year_to_date_sales_total()
    print("\n".join(sorted(p.name for p in OUT.iterdir())))
