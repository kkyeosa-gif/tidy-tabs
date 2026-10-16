#!/usr/bin/env python3
"""Builds the downloadable .xlsx/.csv templates in templates/.

Run from the repo root: python3 scripts/build-templates.py
Needs openpyxl (pip install openpyxl). Every template is plain formulas
(IF, SUMIFS, TODAY, TEXT) so it opens the same in Excel and after
File > Import in Google Sheets. Re-run after editing; the files are
committed so readers download exactly what was checked.
"""
import csv
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path("templates")
USD = '"$"#,##0.00'
US_DATE = "mm/dd/yyyy"
US_DATETIME = "mm/dd/yyyy h:mm AM/PM"
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


def quantity_price_tiers():
    wb = Workbook()
    tiers = wb.active
    tiers.title = "Tiers"
    header(tiers, ["Min quantity", "Unit price", "Tier"], [14, 12, 16])
    for r, (q, p, label) in enumerate([(1, 12.00, "1 to 11"), (12, 10.50, "12 to 35"),
                                        (36, 9.25, "36 to 71"), (72, 8.00, "72 and up")], start=2):
        tiers.append([q, p, label])
        tiers[f"B{r}"].number_format = USD
    tiers["E1"], tiers["E2"], tiers["E3"], tiers["E4"] = "Try a quantity", "Quantity", "Unit price", "Without IFERROR"
    tiers["F2"] = 0
    tiers["F3"] = '=IFERROR(VLOOKUP(F2,$A$2:$B$5,2,TRUE),"Below first tier")'
    tiers["F4"] = "=VLOOKUP(F2,$A$2:$B$5,2,TRUE)"
    tiers["F2"].fill = PatternFill("solid", fgColor="FFF2CC")
    for c in ("F3", "F4"):
        tiers[c].fill = FORMULA_FILL
        tiers[c].number_format = USD
    tiers["E1"].font = Font(bold=True)
    tiers.column_dimensions["E"].width = 18
    tiers.column_dimensions["F"].width = 18

    ws = wb.create_sheet("Orders")
    header(ws, ["Order date", "Customer", "Quantity", "VLOOKUP", "LOOKUP", "INDEX + MATCH",
                "Order total", "All 3 agree?"], [12, 24, 10, 12, 12, 15, 13, 13])
    orders = [
        (date(2026, 9, 2), "Harbor Coffee", 1),
        (date(2026, 9, 4), "Maple Street Bakery", 11),
        (date(2026, 9, 9), "Oak & Ember Candles", 12),
        (date(2026, 9, 14), "Lakeview Florist", 35),
        (date(2026, 9, 18), "Harbor Coffee", 36),
        (date(2026, 9, 23), "Maple Street Bakery", 71),
        (date(2026, 9, 26), "Oak & Ember Candles", 72),
        (date(2026, 9, 30), "Lakeview Florist", 500),
    ]
    for r, (d, who, qty) in enumerate(orders, start=2):
        ws.append([d, who, qty,
                   f"=VLOOKUP(C{r},Tiers!$A$2:$B$5,2,TRUE)",
                   f"=LOOKUP(C{r},Tiers!$A$2:$A$5,Tiers!$B$2:$B$5)",
                   f"=INDEX(Tiers!$B$2:$B$5,MATCH(C{r},Tiers!$A$2:$A$5,1))",
                   f"=C{r}*D{r}",
                   f'=IF(AND(D{r}=E{r},E{r}=F{r}),"Yes","CHECK")'])
        ws[f"A{r}"].number_format = US_DATE
        for col in "DEFG":
            ws[f"{col}{r}"].number_format = USD
        for col in "DEFGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    demo = wb.create_sheet("Unsorted demo")
    header(demo, ["Min quantity", "Unit price", "", "Quantity", "VLOOKUP result", "Correct price"],
           [14, 12, 3, 10, 16, 14])
    for r, (q, p) in enumerate([(36, 9.25), (1, 12.00), (72, 8.00), (12, 10.50)], start=2):
        demo.append([q, p])
        demo[f"B{r}"].number_format = USD
    for r, (qty, right) in enumerate([(5, 12.00), (20, 10.50), (40, 9.25), (80, 8.00)], start=2):
        demo[f"D{r}"] = qty
        demo[f"E{r}"] = f"=VLOOKUP(D{r},$A$2:$B$5,2,TRUE)"
        demo[f"F{r}"] = right
        demo[f"E{r}"].number_format = demo[f"F{r}"].number_format = USD
        demo[f"E{r}"].fill = FORMULA_FILL
    demo["A7"] = "The tier table above is NOT sorted from smallest to largest, so approximate match can return wrong prices."
    demo["A7"].font = Font(italic=True, color="C00000")

    notes_sheet(wb, [
        "Tidy Tabs: Wholesale Price Tiers by Quantity",
        "",
        "1. Tiers tab: list the smallest quantity of each tier in A and its unit price in B. Keep column A sorted from smallest to largest. Approximate match depends on that order.",
        "2. Orders tab: type the quantity in column C. The three unit price columns find the last tier whose minimum is less than or equal to the quantity.",
        "3. D: =VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE). The TRUE (or leaving it out) turns on approximate match.",
        "4. E: =LOOKUP(C2,Tiers!$A$2:$A$5,Tiers!$B$2:$B$5). F: =INDEX(Tiers!$B$2:$B$5,MATCH(C2,Tiers!$A$2:$A$5,1)). The 1 in MATCH means the same approximate match.",
        "5. G: =C2*D2 is the order total. H checks that all three methods give the same price.",
        "6. Edge quantities in the sample: 11 pays the first tier, 12 pays the second; 35 and 36, 71 and 72 work the same way. Quantity 500 uses the last tier.",
        "7. A quantity below the first minimum (such as 0) returns #N/A. Tiers!F2 to F4 let you try one; F3 wraps the lookup in IFERROR.",
        "8. Unsorted demo tab: the same tiers in the wrong order. Compare the VLOOKUP result with the correct price. What an unsorted table returns is not guaranteed, so sort it.",
        "9. To add a tier, insert a row inside the table (not below it) so the ranges grow, and keep the minimums sorted.",
        "",
        "Customers, quantities, and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-quantity-price-tiers.xlsx")


def weighted_average_cost():
    wb = Workbook()
    ws = wb.active
    ws.title = "Purchases"
    header(ws, ["Purchase date", "Item", "Quantity", "Unit cost", "Line total"], [14, 20, 11, 11, 13])
    buys = [
        (date(2026, 7, 6), "Soy wax (lb)", 10, 4.00),
        (date(2026, 7, 20), "Cotton wicks", 200, 0.12),
        (date(2026, 8, 3), "Soy wax (lb)", 40, 3.50),
        (date(2026, 8, 12), "Glass jars", 48, 1.85),
        (date(2026, 8, 25), "Cotton wicks", 500, 0.09),
        (date(2026, 9, 7), "Soy wax (lb)", 25, 3.80),
        (date(2026, 9, 15), "Glass jars", 96, 1.60),
        (date(2026, 9, 22), "Cotton wicks", 300, 0.10),
        (date(2026, 9, 29), "Glass jars", 24, 2.10),
    ]
    for r, (d, item, qty, cost) in enumerate(buys, start=2):
        ws.append([d, item, qty, cost, f"=C{r}*D{r}"])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = '"$"#,##0.00'
        ws[f"E{r}"].number_format = USD
        ws[f"E{r}"].fill = FORMULA_FILL

    s = wb.create_sheet("Summary")
    header(s, ["Item", "Total quantity", "Total spent", "Weighted avg (SUMIFS)", "Weighted avg (SUMPRODUCT)",
               "Simple average", "Difference"], [20, 14, 13, 20, 24, 15, 12])
    for r, item in enumerate(["Soy wax (lb)", "Cotton wicks", "Glass jars"], start=2):
        s.append([item,
                  f"=SUMIFS(Purchases!$C$2:$C$40,Purchases!$B$2:$B$40,A{r})",
                  f"=SUMIFS(Purchases!$E$2:$E$40,Purchases!$B$2:$B$40,A{r})",
                  f"=C{r}/B{r}",
                  f"=SUMPRODUCT((Purchases!$B$2:$B$40=A{r})*Purchases!$C$2:$C$40*Purchases!$D$2:$D$40)/B{r}",
                  f"=AVERAGEIFS(Purchases!$D$2:$D$40,Purchases!$B$2:$B$40,A{r})",
                  f"=ROUND(D{r}-F{r},4)"])
        s[f"C{r}"].number_format = USD
        for col in "DEFG":
            s[f"{col}{r}"].number_format = '"$"#,##0.0000'
        for col in "BCDEFG":
            s[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Weighted Average Cost With SUMPRODUCT",
        "",
        "1. Purchases tab: one row per purchase with the date, item, quantity, and unit cost you paid. Column E is quantity times unit cost.",
        "2. Summary tab, column E: =SUMPRODUCT((Purchases!$B$2:$B$40=A2)*Purchases!$C$2:$C$40*Purchases!$D$2:$D$40)/B2. It multiplies quantity by unit cost for the rows that match the item, adds them up, and divides by the total quantity.",
        "3. Column D gives the same answer with SUMIFS: total spent divided by total quantity (=C2/B2).",
        "4. Column F, =AVERAGEIFS(...), is the simple average of the unit costs. It ignores how many units each purchase covered, so it is off whenever the quantities differ.",
        "5. Example: 10 lb at $4.00 and 40 lb at $3.50 cost $180.00 for 50 lb, so the weighted average is $3.60 per lb. The simple average of $4.00 and $3.50 is $3.75.",
        "6. The formulas read rows 2 to 40, so you can add purchases without editing them. Type each item name exactly the same way every time.",
        "7. This is an average unit cost of what you bought. It is not an inventory valuation method or tax advice.",
        "",
        "Items, quantities, and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-weighted-average-cost.xlsx")


def last_order_date_maxifs():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order date", "Customer", "Amount"], [13, 26, 12])
    orders = [
        (date(2026, 6, 8), "Harbor Coffee", 640.00),
        (date(2026, 6, 19), "Maple Street Bakery", 310.50),
        (date(2026, 7, 14), "Oak & Ember Candles", 455.00),
        (date(2026, 7, 28), "Harbor Coffee", 720.00),
        (date(2026, 8, 11), "Lakeview Florist", 180.25),
        (date(2026, 8, 26), "Maple Street Bakery", 395.00),
        (date(2026, 9, 9), "Oak & Ember Candles", 510.75),
        (date(2026, 9, 17), "Harbor Coffee", 865.00),
        (date(2026, 9, 24), "Lakeview Florist", 242.00),
        (date(2026, 9, 30), "Maple Street Bakery", 428.40),
    ]
    for r, (d, who, amt) in enumerate(orders, start=2):
        ws.append([d, who, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD

    s = wb.create_sheet("Summary")
    header(s, ["Customer", "Last order date", "Last order amount", "Days since last order", "Orders"],
           [26, 16, 18, 20, 9])
    for r, who in enumerate(["Harbor Coffee", "Maple Street Bakery", "Oak & Ember Candles", "Lakeview Florist"],
                            start=2):
        s.append([who,
                  f"=_xlfn.MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A{r})",
                  f"=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A{r},Orders!$A$2:$A$200,B{r})",
                  f"=$H$2-B{r}",
                  f"=COUNTIFS(Orders!$B$2:$B$200,A{r})"])
        s[f"B{r}"].number_format = US_DATE
        s[f"C{r}"].number_format = USD
        s[f"D{r}"].number_format = "0"
        for col in "BCDE":
            s[f"{col}{r}"].fill = FORMULA_FILL
    s.column_dimensions["G"].width = 14
    s.column_dimensions["H"].width = 14
    s["G2"], s["H2"] = "Report date", date(2026, 10, 9)
    s["H2"].number_format = US_DATE
    s["H2"].fill = PatternFill("solid", fgColor="FFF2CC")
    s["G2"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Last Order Date per Customer With MAXIFS",
        "",
        "1. Orders tab: one row per order with a real date, the customer name, and the amount. The Summary formulas read rows 2 to 200.",
        "2. Summary B: =MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A2). It returns the latest (largest) date among the rows whose customer matches A2.",
        "3. Summary C: =SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2,Orders!$A$2:$A$200,B2). It adds the amounts for that customer on that last date. If a customer placed two orders on the same day, you get both added together.",
        "4. Summary D: report date minus the last order date. Change H2 (or type =TODAY()) to see it move.",
        "5. If a customer has no orders, MAXIFS returns 0, which shows as a date in 1899 or 1900. Wrap it in IF(COUNTIFS(...)=0,\"\",...) if your list can include new customers.",
        "6. MAXIFS is in Excel 2019 and Microsoft 365, and in Google Sheets. It is not in Excel 2016 or earlier; there, use an array formula with MAX and IF.",
        "7. Dates typed as text (left-aligned, no date format) are ignored by MAXIFS, so convert them to real dates first.",
        "",
        "Customers and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-last-order-date-maxifs.xlsx")


def count_orders_by_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order date", "Customer", "Amount"], [13, 26, 12])
    orders = [
        (date(2026, 7, 8), "Harbor Coffee", 120.00),
        (date(2026, 7, 22), "Maple Street Bakery", 245.50),
        (date(2026, 8, 5), "Oak & Ember Candles", 89.00),
        (date(2026, 8, 19), "Lakeview Florist", 310.25),
        (date(2026, 8, 31), "Harbor Coffee", 150.00),
        (date(2026, 9, 1), "Maple Street Bakery", 75.00),
        (date(2026, 9, 14), "Oak & Ember Candles", 420.00),
        (date(2026, 9, 28), "Lakeview Florist", 198.75),
        (date(2026, 10, 2), "Harbor Coffee", 260.00),
        (date(2026, 10, 7), "Maple Street Bakery", 135.50),
    ]
    for r, (d, who, amt) in enumerate(orders, start=2):
        ws.append([d, who, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD

    s = wb.create_sheet("Monthly")
    header(s, ["Month start", "Orders", "Revenue"], [14, 10, 14])
    for r, m in enumerate([date(2026, 7, 1), date(2026, 8, 1), date(2026, 9, 1), date(2026, 10, 1)], start=2):
        s.append([m,
                  f'=COUNTIFS(Orders!$A$2:$A$200,">="&A{r},Orders!$A$2:$A$200,"<"&EDATE(A{r},1))',
                  f'=SUMIFS(Orders!$C$2:$C$200,Orders!$A$2:$A$200,">="&A{r},Orders!$A$2:$A$200,"<"&EDATE(A{r},1))'])
        s[f"A{r}"].number_format = US_DATE
        s[f"C{r}"].number_format = USD
        s[f"B{r}"].fill = s[f"C{r}"].fill = FORMULA_FILL
    s["A6"], s["B6"], s["C6"] = "Total", "=SUM(B2:B5)", "=SUM(C2:C5)"
    s["A7"], s["B7"], s["C7"] = "Orders tab total", "=COUNT(Orders!A2:A200)", "=SUM(Orders!C2:C200)"
    s["C7"].number_format = s["C6"].number_format = USD
    for c in ("A6", "A7"):
        s[c].font = Font(bold=True)
    s.column_dimensions["A"].width = 18

    notes_sheet(wb, [
        "Tidy Tabs: Count Orders by Month With COUNTIFS",
        "",
        "1. Orders tab: one row per order with a real date, the customer, and the amount. The formulas read rows 2 to 200.",
        "2. Monthly tab, column A: type the first day of each month (07/01/2026). Format the cells as dates.",
        "3. Orders: =COUNTIFS(Orders!$A$2:$A$200,\">=\"&A2,Orders!$A$2:$A$200,\"<\"&EDATE(A2,1)). It counts dates on or after the first of the month and before the first of the next month.",
        "4. Revenue: the same two date tests inside SUMIFS, adding Orders!$C$2:$C$200.",
        "5. Why \"before the first of the next month\" and not \"on or before the last day\": it also works if a date carries a time of day, such as 08/31/2026 3:30 PM.",
        "6. Rows 6 and 7 compare the monthly totals with the whole Orders tab. If they differ, an order falls outside the months listed or its date is stored as text.",
        "7. Sample edge: 08/31/2026 counts in August and 09/01/2026 counts in September.",
        "",
        "Customers and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-count-orders-by-month.xlsx")


def average_order_value_by_channel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order date", "Channel", "Amount"], [13, 18, 12])
    orders = [
        (date(2026, 9, 3), "Etsy", 34.00),
        (date(2026, 9, 5), "Farmers market", 22.00),
        (date(2026, 9, 8), "Shopify", 78.00),
        (date(2026, 9, 12), "Etsy", 52.50),
        (date(2026, 9, 13), "Farmers market", 35.50),
        (date(2026, 9, 17), "Shopify", 45.00),
        (date(2026, 9, 20), "Farmers market", 18.50),
        (date(2026, 9, 24), "Etsy", 29.00),
        (date(2026, 9, 27), "Shopify", 96.50),
        (date(2026, 9, 28), "Farmers market", 40.00),
    ]
    for r, (d, ch, amt) in enumerate(orders, start=2):
        ws.append([d, ch, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD

    s = wb.create_sheet("By channel")
    header(s, ["Channel", "Orders", "Revenue", "Average order value", "Check (revenue / orders)"],
           [18, 10, 13, 20, 24])
    for r, ch in enumerate(["Etsy", "Shopify", "Farmers market"], start=2):
        s.append([ch,
                  f"=COUNTIFS(Orders!$B$2:$B$200,A{r})",
                  f"=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A{r})",
                  f"=AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A{r})",
                  f"=C{r}/B{r}"])
        for col in "CDE":
            s[f"{col}{r}"].number_format = USD
        for col in "BCDE":
            s[f"{col}{r}"].fill = FORMULA_FILL
    s["A5"], s["B5"], s["C5"], s["D5"] = "All channels", "=SUM(B2:B4)", "=SUM(C2:C4)", "=AVERAGE(Orders!C2:C200)"
    s["C5"].number_format = s["D5"].number_format = USD
    s["A5"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Average Order Value by Sales Channel With AVERAGEIFS",
        "",
        "1. Orders tab: one row per order with the date, the sales channel, and the order amount. Type each channel name exactly the same way every time.",
        "2. By channel, column D: =AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2). The first argument is the range to average, then the range to test and the value it must match.",
        "3. Column B: =COUNTIFS(Orders!$B$2:$B$200,A2) counts the orders, and column C adds the revenue with SUMIFS.",
        "4. Column E divides revenue by orders. It should equal column D; if it does not, a row has a text amount or a blank.",
        "5. If a channel has no orders, AVERAGEIFS returns #DIV/0!. Wrap it in IFERROR(...,\"\") if you list channels you have not sold through yet.",
        "6. To average only one month, add two more date conditions, for example Orders!$A$2:$A$200,\">=\"&DATE(2026,9,1).",
        "7. Row 5 is the average across all channels. It is the average of orders, not the average of the three channel averages.",
        "",
        "Orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-average-order-value-by-channel.xlsx")


def letter_setup(ws, landscape=False, wrap_col_a=False):
    """US Letter, fit to one page wide, header row repeated."""
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    if wrap_col_a:
        for row in ws.iter_rows():
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
    else:
        ws.print_title_rows = "1:1"


def finish_letter(wb, landscape=False):
    for ws in wb.worksheets:
        letter_setup(ws, landscape=landscape and ws.title != "How to use",
                     wrap_col_a=ws.title == "How to use")


def subtotal_filtered_sales():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Sale date", "Channel", "Amount"], [13, 18, 12])
    sales = [
        (date(2026, 9, 2), "Shopify", 12.00),
        (date(2026, 9, 3), "Etsy", 34.00),
        (date(2026, 9, 5), "Farmers market", 22.00),
        (date(2026, 9, 6), "Farmers market", 8.50),
        (date(2026, 9, 8), "Shopify", 15.00),
        (date(2026, 9, 10), "Etsy", 18.50),
        (date(2026, 9, 12), "Farmers market", 6.00),
        (date(2026, 9, 14), "Shopify", 9.50),
        (date(2026, 9, 17), "Farmers market", 14.00),
        (date(2026, 9, 19), "Shopify", 11.00),
        (date(2026, 9, 21), "Etsy", 23.00),
        (date(2026, 9, 23), "Farmers market", 7.50),
        (date(2026, 9, 25), "Shopify", 5.00),
        (date(2026, 9, 27), "Farmers market", 10.50),
        (date(2026, 9, 29), "Shopify", 4.00),
    ]
    for r, (d, ch, amt) in enumerate(sales, start=2):
        ws.append([d, ch, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD
    ws.auto_filter.ref = "A1:C16"

    t = wb.create_sheet("Totals")
    header(t, ["Measure", "Formula", "Result"], [30, 28, 14])
    rows = [
        ("Visible total (SUBTOTAL 109)", "=SUBTOTAL(109,C2:C16)", "=SUBTOTAL(109,Sales!C2:C16)", USD),
        ("Visible count (SUBTOTAL 103)", "=SUBTOTAL(103,C2:C16)", "=SUBTOTAL(103,Sales!C2:C16)", "0"),
        ("Visible average (SUBTOTAL 101)", "=SUBTOTAL(101,C2:C16)", "=SUBTOTAL(101,Sales!C2:C16)", USD),
        ("All rows total (SUM, ignores filter)", "=SUM(C2:C16)", "=SUM(Sales!C2:C16)", USD),
        ("All rows count (COUNT)", "=COUNT(C2:C16)", "=COUNT(Sales!C2:C16)", "0"),
    ]
    for r, (label, shown, f, fmt) in enumerate(rows, start=2):
        t.append([label, None, f])
        t[f"B{r}"].value = shown
        t[f"B{r}"].data_type = "s"
        t[f"C{r}"].number_format = fmt
        t[f"C{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Sum Only the Visible Rows After Filtering",
        "",
        "1. Sales tab: one row per sale. The header row already has filter arrows (Data > Filter in Excel, Data > Create a filter in Google Sheets).",
        "2. Click the arrow on Channel and keep only one channel, for example Etsy. The other rows are hidden.",
        "3. Totals tab, C2: =SUBTOTAL(109,Sales!C2:C16) adds only the rows still visible. 109 means SUM, skipping rows hidden by a filter or by hand.",
        "4. C3: =SUBTOTAL(103,...) counts visible numbers. C4: =SUBTOTAL(101,...) averages the visible rows.",
        "5. C5: =SUM(Sales!C2:C16) adds every row, so it does not change when you filter. Compare it with C2.",
        "6. If you add rows below row 16, extend the ranges (or insert the new rows inside the list).",
        "7. Do not put SUBTOTAL cells inside the filtered range: they would be hidden or counted twice. That is why the totals sit on their own tab.",
        "",
        "Sales and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-subtotal-filtered-sales.xlsx")


def rolling_average_sales():
    wb = Workbook()
    ws = wb.active
    ws.title = "Monthly sales"
    header(ws, ["Month", "Sales", "3-month average", "Check (SUM / 3)", "Matches?"], [12, 12, 17, 17, 11])
    sales = [1200, 950, 1100, 2400, 3100, 1800, 1650, 1400, 2250, 2900, 3800, 4200]
    for r, amt in enumerate(sales, start=2):
        m = date(2026 if r < 2 + 12 else 2027, r - 1, 1)
        ws.append([m, amt])
        ws[f"A{r}"].number_format = "mmm yyyy"
        ws[f"B{r}"].number_format = USD
        if r >= 4:
            ws[f"C{r}"] = f'=IF(COUNT(B{r-2}:B{r})<3,"",AVERAGE(B{r-2}:B{r}))'
            ws[f"D{r}"] = f"=SUM(B{r-2}:B{r})/3"
            ws[f"E{r}"] = f'=IF(ROUND(C{r}-D{r},6)=0,"Yes","CHECK")'
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = USD

    notes_sheet(wb, [
        "Tidy Tabs: Rolling 3-Month Average of Sales",
        "",
        "1. Monthly sales tab: type each month's sales in column B, one row per month, oldest first. Months with no sales need a 0, not a blank.",
        "2. C4: =IF(COUNT(B2:B4)<3,\"\",AVERAGE(B2:B4)). It averages this month and the two before it. Fill it down. The first two months have no formula because they do not have three months yet.",
        "3. The COUNT test keeps the cell blank if any of the three months is empty or text, so a half-filled window is not averaged by mistake.",
        "4. D4: =SUM(B2:B4)/3 is a check. E says Yes when it equals column C.",
        "5. Each average moves one row at a time: row 5 drops January and adds April. That smooths out busy and slow seasons.",
        "6. To use 6 months instead of 3, widen the range (B2:B7) and change both the 3 in COUNT's test and the divisor.",
        "",
        "Months and sales are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-rolling-average-sales.xlsx")


def prorate_partial_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Proration"
    header(ws, ["Start date", "Monthly fee", "Days charged", "Days in month", "Prorated amount", "Example"],
           [13, 13, 13, 14, 16, 34])
    rows = [
        (date(2026, 10, 12), 150.00, "Retainer starts mid month"),
        (date(2026, 2, 20), 150.00, "February (28 days)"),
        (date(2026, 11, 1), 150.00, "Starts on the 1st: full month"),
        (date(2026, 12, 31), 150.00, "Starts on the last day: 1 day"),
        (date(2026, 9, 15), 85.00, "Farmers market booth, 30 day month"),
        (date(2026, 3, 10), 1200.00, "Studio rent"),
        (date(2028, 2, 20), 150.00, "Leap year February (29 days)"),
    ]
    for r, (d, fee, note) in enumerate(rows, start=2):
        ws.append([d, fee, f"=EOMONTH(A{r},0)-A{r}+1", f"=DAY(EOMONTH(A{r},0))", f"=ROUND(B{r}*C{r}/D{r},2)", note])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"B{r}"].number_format = USD
        ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = "0"
        ws[f"E{r}"].number_format = USD
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Prorate a Partial Month of Rent or a Retainer",
        "",
        "1. Proration tab: type the start date in A and the full monthly fee in B. Columns C to E are formulas; fill them down for more rows.",
        "2. C, days charged: =EOMONTH(A2,0)-A2+1. EOMONTH(A2,0) is the last day of the start month; the +1 counts the start day itself.",
        "3. D, days in that month: =DAY(EOMONTH(A2,0)). It is 28, 29, 30 or 31 depending on the month.",
        "4. E, prorated amount: =ROUND(B2*C2/D2,2), the fee times days charged over days in the month, rounded to cents.",
        "5. Example: $150.00 starting 10/12/2026 is 20 of 31 days, $96.77. $150.00 starting 02/20/2026 is 9 of 28 days, $48.21.",
        "6. Columns C and D must be formatted as Number. If they show a date like 01/20/1900, change the format to Number.",
        "7. This is arithmetic only. Some leases and contracts count a flat 30 day month or use other rules, so follow your own agreement. It is not legal advice.",
        "",
        "Dates and fees are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-prorate-partial-month.xlsx")


def two_way_rate_card():
    wb = Workbook()
    rc = wb.active
    rc.title = "Rate card"
    header(rc, ["Service", "Standard", "Rush", "Same day"], [16, 13, 13, 13])
    prices = [("Logo", 300, 420, 600), ("Brochure", 350, 500, 700), ("Website", 1200, 1600, 2100)]
    for r, (svc, *p) in enumerate(prices, start=2):
        rc.append([svc, *p])
        for col in "BCD":
            rc[f"{col}{r}"].number_format = USD

    q = wb.create_sheet("Quote")
    q.column_dimensions["A"].width = 34
    for col in "BCD":
        q.column_dimensions[col].width = 16
    q["A1"], q["A1"].font = "Pick a service and a turnaround (yellow cells)", Font(bold=True)
    q["A2"], q["B2"] = "Service", "Brochure"
    q["A3"], q["B3"] = "Turnaround", "Rush"
    for c in ("B2", "B3"):
        q[c].fill = PatternFill("solid", fgColor="FFF2CC")
    dv1 = DataValidation(type="list", formula1="='Rate card'!$A$2:$A$4", allow_blank=False)
    dv2 = DataValidation(type="list", formula1="='Rate card'!$B$1:$D$1", allow_blank=False)
    q.add_data_validation(dv1)
    q.add_data_validation(dv2)
    dv1.add("B2")
    dv2.add("B3")
    q["A5"] = "Price (INDEX + MATCH)"
    q["B5"] = "=INDEX('Rate card'!$B$2:$D$4,MATCH(B2,'Rate card'!$A$2:$A$4,0),MATCH(B3,'Rate card'!$B$1:$D$1,0))"
    q["A6"] = "Price (with IFERROR)"
    q["B6"] = ("=IFERROR(INDEX('Rate card'!$B$2:$D$4,MATCH(B2,'Rate card'!$A$2:$A$4,0),"
               "MATCH(B3,'Rate card'!$B$1:$D$1,0)),\"Check spelling\")")
    q["A7"] = "Price (SUMPRODUCT, comparison)"
    q["B7"] = "=SUMPRODUCT(('Rate card'!$A$2:$A$4=B2)*('Rate card'!$B$1:$D$1=B3)*'Rate card'!$B$2:$D$4)"
    for c in ("B5", "B6", "B7"):
        q[c].number_format = USD
        q[c].fill = FORMULA_FILL
        q[c].alignment = Alignment(horizontal="right")

    q["A9"] = "All 9 prices through the same formula"
    q["A9"].font = Font(bold=True)
    for i, col in enumerate("BCD"):
        q[f"{col}10"] = f"='Rate card'!{col}1"
        q[f"{col}10"].font = Font(bold=True)
    for r in (11, 12, 13):
        q[f"A{r}"] = f"='Rate card'!A{r-9}"
        q[f"A{r}"].font = Font(bold=True)
        for col in "BCD":
            q[f"{col}{r}"] = (f"=INDEX('Rate card'!$B$2:$D$4,MATCH($A{r},'Rate card'!$A$2:$A$4,0),"
                              f"MATCH({col}$10,'Rate card'!$B$1:$D$1,0))")
            q[f"{col}{r}"].number_format = USD
            q[f"{col}{r}"].fill = FORMULA_FILL
    q["A14"] = "Cells that match the rate card (should be 9)"
    q["B14"] = "=SUMPRODUCT(--(B11:D13='Rate card'!B2:D4))"

    notes_sheet(wb, [
        "Tidy Tabs: Two-Way Rate Card Lookup With INDEX and MATCH",
        "",
        "1. Rate card tab: services down column A, turnaround speeds across row 1, prices in B2:D4. Keep the labels spelled the same everywhere.",
        "2. Quote tab: pick a service in B2 and a turnaround in B3 from the dropdown lists.",
        "3. B5: =INDEX('Rate card'!$B$2:$D$4,MATCH(B2,'Rate card'!$A$2:$A$4,0),MATCH(B3,'Rate card'!$B$1:$D$1,0)). The first MATCH finds the row number, the second finds the column number, and INDEX returns the price where they cross.",
        "4. The 0 at the end of each MATCH means exact match.",
        "5. B6 wraps the same formula in IFERROR so a misspelled service or speed shows Check spelling instead of #N/A. The dropdowns stop most typos, but pasted text can still be wrong.",
        "6. B7 gets the same price with SUMPRODUCT. A misspelled label returns $0.00 there, not an error, so use it only as a cross-check.",
        "7. Rows 10 to 14 run the INDEX and MATCH formula on all 9 cells and count how many equal the rate card.",
        "8. To add a service or speed, insert a row or column inside the table and extend the ranges.",
        "",
        "Services and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-two-way-rate-card.xlsx")


def in_cell_bars():
    wb = Workbook()
    ws = wb.active
    ws.title = "Units sold"
    header(ws, ["Product", "Units sold", "Bar (block characters)", "Blocks (LEN)", "Bar (| fallback)"],
           [24, 12, 30, 13, 30])
    items = [("Soy candle, 8 oz", 3200), ("Beeswax melts", 1600), ("Gift tags (retired)", 0),
             ("Wick trimmer", 450), ("Lip balm", 2400), ("Mini jar candle", 800)]
    for r, (name, units) in enumerate(items, start=2):
        ws.append([name, units,
                   f'=REPT("█",ROUND(B{r}/MAX($B$2:$B$7)*20,0))',
                   f"=LEN(C{r})",
                   f'=REPT("|",ROUND(B{r}/MAX($B$2:$B$7)*20,0))'])
        ws[f"B{r}"].number_format = "#,##0"
        ws[f"C{r}"].font = Font(color="2E5E4E")
        ws[f"E{r}"].font = Font(color="2E5E4E")
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Bar Chart Inside Cells With REPT",
        "",
        "1. Units sold tab: product names in A, units in B (rows 2 to 7).",
        "2. C2: =REPT(\"█\",ROUND(B2/MAX($B$2:$B$7)*20,0)). REPT repeats the block character. The biggest value gets 20 blocks and every other row is scaled to it.",
        "3. Copy the █ character from the formula if you need to retype it. The $ signs keep the MAX range fixed when you fill down.",
        "4. D2: =LEN(C2) counts the blocks, as a check. With 3,200 / 1,600 / 0 / 450 units the bars are 20 / 10 / 0 / 3 blocks.",
        "5. Column E uses the | character instead. Use it if the block shows as an empty box or looks different in your font.",
        "6. To change the full bar length, replace the 20 in both the formula and your width. Widen the column so the longest bar fits.",
        "7. A bar is rounded to whole characters, so very small values can show 0 blocks.",
        "",
        "Products and units are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-in-cell-bars.xlsx")


def quarterly_sales_totals():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Sale date", "Amount", "Quarter"], [13, 13, 12])
    sales = [
        (date(2025, 12, 20), 90.00), (date(2026, 1, 15), 120.00), (date(2026, 2, 3), 85.50),
        (date(2026, 3, 31), 200.00), (date(2026, 4, 1), 150.00), (date(2026, 4, 18), 95.25),
        (date(2026, 5, 9), 310.00), (date(2026, 6, 30), 75.00), (date(2026, 7, 1), 220.00),
        (date(2026, 7, 22), 60.50), (date(2026, 8, 14), 180.00), (date(2026, 9, 30), 140.00),
        (date(2026, 10, 1), 99.99), (date(2026, 10, 5), 250.00), (date(2026, 11, 11), 130.00),
        (date(2026, 12, 31), 300.00),
    ]
    for r, (d, amt) in enumerate(sales, start=2):
        ws.append([d, amt, f'=YEAR(A{r})&" Q"&ROUNDUP(MONTH(A{r})/3,0)'])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"B{r}"].number_format = USD
        ws[f"C{r}"].fill = FORMULA_FILL

    q = wb.create_sheet("By quarter")
    header(q, ["Quarter", "Quarter starts", "Total (helper column)", "Orders", "Total (date range)", "Same?"],
           [12, 15, 21, 9, 19, 9])
    starts = [date(2025, 10, 1), date(2026, 1, 1), date(2026, 4, 1), date(2026, 7, 1), date(2026, 10, 1)]
    for r, d in enumerate(starts, start=2):
        q.append([f'=YEAR(B{r})&" Q"&ROUNDUP(MONTH(B{r})/3,0)', d,
                  f"=SUMIFS(Sales!$B$2:$B$200,Sales!$C$2:$C$200,A{r})",
                  f"=COUNTIFS(Sales!$C$2:$C$200,A{r})",
                  f'=SUMIFS(Sales!$B$2:$B$200,Sales!$A$2:$A$200,">="&B{r},Sales!$A$2:$A$200,"<"&EDATE(B{r},3))',
                  f'=IF(ROUND(C{r}-E{r},2)=0,"Yes","CHECK")'])
        q[f"B{r}"].number_format = US_DATE
        q[f"C{r}"].number_format = q[f"E{r}"].number_format = USD
        for col in "ACDEF":
            q[f"{col}{r}"].fill = FORMULA_FILL
    q["A7"], q["C7"], q["D7"], q["E7"] = "All quarters", "=SUM(C2:C6)", "=SUM(D2:D6)", "=SUM(E2:E6)"
    q["A8"], q["C8"] = "All sales on Sales tab", "=SUM(Sales!B2:B200)"
    q["D8"] = "=COUNT(Sales!B2:B200)"
    for c in ("C7", "E7", "C8"):
        q[c].number_format = USD
    for c in ("A7", "A8"):
        q[c].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Total Sales by Quarter",
        "",
        "1. Sales tab: type the sale date in A (a real date, not text) and the amount in B. Column C is a formula: fill it down.",
        "2. C2: =YEAR(A2)&\" Q\"&ROUNDUP(MONTH(A2)/3,0). ROUNDUP(MONTH/3,0) turns months 1 to 3 into 1, 4 to 6 into 2, 7 to 9 into 3 and 10 to 12 into 4. The year is added so 2025 Q4 and 2026 Q4 stay separate.",
        "3. By quarter tab, C2: =SUMIFS(Sales!$B$2:$B$200,Sales!$C$2:$C$200,A2) adds every sale with that quarter label. D counts the orders with COUNTIFS.",
        "4. E2 gets the same total without a helper column, by date range: =SUMIFS(amounts,dates,\">=\"&B2,dates,\"<\"&EDATE(B2,3)). B2 is the first day of the quarter. F says Yes when both agree.",
        "5. Row 7 adds the quarters and row 8 adds the whole Sales tab. If they differ, a date is outside the quarters listed or is stored as text.",
        "6. Quarters here are calendar quarters (January to March is Q1). If your business year starts in another month, this sheet does not apply as is.",
        "7. 03/31 and 04/01 fall in different quarters, and 12/31 is in Q4. Both methods handle the boundary days.",
        "",
        "Sales and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-quarterly-sales-totals.xlsx")


def missing_invoice_numbers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    header(ws, ["Invoice number", "Client", "Amount"], [16, 24, 13])
    rows = [(1001, "Maple Street Bakery", 450.00), (1002, "Dana Ruiz Design", 300.00),
            (1003, "Harbor Yoga", 180.00), (1005, "Cedar Hill Farm", 620.00),
            (1006, "Maple Street Bakery", 450.00), (1008, "Harbor Yoga", 180.00),
            (1009, "Sam Patel LLC", 975.00), (1009, "Sam Patel LLC", 975.00),
            (1011, "Dana Ruiz Design", 300.00), (1012, "Cedar Hill Farm", 620.00)]
    for r, (n, c, a) in enumerate(rows, start=2):
        ws.append([n, c, a])
        ws[f"C{r}"].number_format = USD

    ck = wb.create_sheet("Check")
    header(ck, ["Expected number", "Times found", "Status"], [17, 13, 14])
    for r in range(2, 14):
        ck.append([f"=MIN(Invoices!$A$2:$A$200)+ROW()-2",
                   f"=COUNTIF(Invoices!$A$2:$A$200,A{r})",
                   f'=IF(B{r}=0,"Missing",IF(B{r}>1,"Duplicate","OK"))'])
        ck[f"A{r}"].number_format = "0"
        for col in "ABC":
            ck[f"{col}{r}"].fill = FORMULA_FILL
    ck.conditional_formatting.add("C2:C13", FormulaRule(formula=['$C2="Missing"'], fill=ALERT_FILL))
    ck.conditional_formatting.add("C2:C13", FormulaRule(formula=['$C2="Duplicate"'], fill=PatternFill("solid", fgColor="FFF2CC")))
    ck["E1"], ck["F1"] = "Summary", None
    ck["E1"].font = Font(bold=True)
    ck.column_dimensions["E"].width = 28
    ck.column_dimensions["F"].width = 10
    summary = [
        ("Lowest number", "=MIN(Invoices!A2:A200)"),
        ("Highest number", "=MAX(Invoices!A2:A200)"),
        ("Numbers expected in range", "=F3-F2+1"),
        ("Invoices entered", "=COUNT(Invoices!A2:A200)"),
        ("Missing numbers", '=COUNTIF(C2:C13,"Missing")'),
        ("Numbers used twice", '=COUNTIF(C2:C13,"Duplicate")'),
        ("Check: entered - duplicate extras + missing", "=F5-(SUMPRODUCT((B2:B13>1)*(B2:B13-1)))+F6"),
    ]
    for i, (label, f) in enumerate(summary, start=2):
        ck[f"E{i}"], ck[f"F{i}"] = label, f
        ck[f"F{i}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Find Missing Invoice Numbers",
        "",
        "1. Invoices tab: list every invoice number you issued in column A (any order). Use plain numbers such as 1001. If your numbers have a prefix like INV-1001, keep the number in its own column.",
        "2. Check tab, A2: =MIN(Invoices!$A$2:$A$200)+ROW()-2 lists each expected number from the lowest one upward. Fill it down for as many numbers as the range needs.",
        "3. B2: =COUNTIF(Invoices!$A$2:$A$200,A2) counts how often that number appears.",
        "4. C2: =IF(B2=0,\"Missing\",IF(B2>1,\"Duplicate\",\"OK\")). Missing numbers turn red, duplicates yellow.",
        "5. F6 counts the missing numbers. F5 minus the extra copies plus F6 should equal F4 (the numbers expected in the range).",
        "6. The check only looks between your lowest and highest number. A missing number before the first or after the last one is not detected.",
        "7. A missing number is a question to answer, not proof of a problem: it may be a voided invoice. This sheet does not decide that.",
        "8. The Check tab lists 12 numbers. If your lowest-to-highest range is longer, fill the Check formulas down and widen the ranges in the summary; until then F8 will not equal F4, which tells you rows are missing from the list.",
        "",
        "Clients and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-missing-invoice-numbers.xlsx")


def sumif_contains_text():
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    header(ws, ["Date", "Description", "Amount"], [13, 34, 13])
    rows = [
        (date(2026, 9, 1), "Uber to client meeting", 18.40), (date(2026, 9, 2), "Printer ink cartridge", 42.99),
        (date(2026, 9, 4), "UBER EATS lunch", 14.25), (date(2026, 9, 7), "Pink ribbon spool", 6.50),
        (date(2026, 9, 9), "Etsy listing fees", 3.20), (date(2026, 9, 12), "Etsy shipping labels", 27.80),
        (date(2026, 9, 15), "Domain renewal", 12.00), (date(2026, 9, 18), "Uber airport", 36.10),
        (date(2026, 9, 21), "Packing tape and ink pen", 9.75), (date(2026, 9, 24), "Stamps*", 11.60),
        (date(2026, 9, 26), "Ink refill kit", 19.00), (date(2026, 9, 30), "Software subscription", 15.00),
    ]
    for r, (d, desc, amt) in enumerate(rows, start=2):
        ws.append([d, desc, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD

    t = wb.create_sheet("Totals")
    header(t, ["Text to find", "Contains (total)", "Contains (count)", "Starts with (total)", "Ends with (total)"],
           [18, 17, 17, 19, 18])
    for r, word in enumerate(["Uber", "Etsy", "ink", "Ink pen", "Stamps~*", "Zoom"], start=2):
        t.append([word,
                  f'=SUMIF(Expenses!$B$2:$B$200,"*"&A{r}&"*",Expenses!$C$2:$C$200)',
                  f'=COUNTIF(Expenses!$B$2:$B$200,"*"&A{r}&"*")',
                  f'=SUMIF(Expenses!$B$2:$B$200,A{r}&"*",Expenses!$C$2:$C$200)',
                  f'=SUMIF(Expenses!$B$2:$B$200,"*"&A{r},Expenses!$C$2:$C$200)'])
        for col in "BDE":
            t[f"{col}{r}"].number_format = USD
        for col in "BCDE":
            t[f"{col}{r}"].fill = FORMULA_FILL
    t["A9"], t["B9"] = "All expenses", "=SUM(Expenses!C2:C200)"
    t["B9"].number_format = USD
    t["A9"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Sum Cells That Contain Certain Text",
        "",
        "1. Expenses tab: date, description and amount. Totals tab column A holds the word to look for.",
        "2. B2: =SUMIF(Expenses!$B$2:$B$200,\"*\"&A2&\"*\",Expenses!$C$2:$C$200). The * wildcard stands for any text, so the description only has to contain the word somewhere. Letter case is ignored.",
        "3. C2: =COUNTIF(Expenses!$B$2:$B$200,\"*\"&A2&\"*\") counts the matching rows.",
        "4. D2 drops the first * (text starts with the word). E2 drops the last * (text ends with the word).",
        "5. The match is on letters, not on words. \"ink\" also matches \"Pink ribbon spool\". Use a longer phrase such as \"Printer ink\" if that is a problem.",
        "6. To search for a real * or ?, put a ~ in front of it: Stamps~* finds the text Stamps*.",
        "7. A blank cell in column A makes the pattern ** , which matches every row that has text.",
        "8. SUMIF does not read cell colors and only matches text, so a number in the description column is not found by a * pattern.",
        "",
        "Expenses are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-sumif-contains-text.xlsx")


def invoice_installments():
    wb = Workbook()
    ws = wb.active
    ws.title = "Payment plan"
    ws.column_dimensions["A"].width = 24
    for col, w in zip("BCD", (16, 16, 16)):
        ws.column_dimensions[col].width = w
    ws["A1"], ws["A1"].font = "Invoice total", Font(bold=True)
    ws["B1"] = 2500.00
    ws["A2"], ws["A2"].font = "Number of payments (1 to 12)", Font(bold=True)
    ws["B2"] = 6
    ws["A3"], ws["A3"].font = "First payment date", Font(bold=True)
    ws["B3"] = date(2026, 11, 1)
    ws["B1"].number_format = USD
    ws["B3"].number_format = US_DATE
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="12", allow_blank=False,
                        showErrorMessage=True, errorTitle="Payments", error="Enter a whole number from 1 to 12.")
    ws.add_data_validation(dv)
    dv.add("B2")
    for c in ("B1", "B2", "B3"):
        ws[c].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["A4"] = "Regular payment"
    ws["B4"] = "=ROUND(B1/B2,2)"
    ws["A5"] = "Last payment"
    ws["B5"] = "=B1-B4*(B2-1)"
    for c in ("B4", "B5"):
        ws[c].number_format = USD
        ws[c].fill = FORMULA_FILL
    for col, name in zip("ABCD", ["Payment #", "Due date", "Amount", "Balance after"]):
        ws[f"{col}7"] = name
        ws[f"{col}7"].fill, ws[f"{col}7"].font = HEADER_FILL, HEADER_FONT
    for n in range(1, 13):
        r = 7 + n
        ws[f"A{r}"] = n
        ws[f"B{r}"] = f'=IF(A{r}>$B$2,"",EDATE($B$3,A{r}-1))'
        ws[f"C{r}"] = f'=IF(A{r}>$B$2,"",IF(A{r}=$B$2,$B$5,$B$4))'
        ws[f"D{r}"] = f'=IF(A{r}>$B$2,"",$B$1-SUM($C$8:C{r}))'
        ws[f"B{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = USD
        for col in "BCD":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["A21"], ws["C21"] = "Total of payments", "=SUM(C8:C19)"
    ws["A22"], ws["C22"] = "Difference from invoice", "=ROUND(C21-B1,2)"
    ws["C21"].number_format = ws["C22"].number_format = USD
    ws["A21"].font = ws["A22"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Split an Invoice Into Equal Monthly Payments",
        "",
        "1. Payment plan tab: change the three yellow cells: invoice total (B1), number of payments from 1 to 12 (B2) and the first payment date (B3).",
        "2. B4: =ROUND(B1/B2,2) is the regular payment rounded to cents.",
        "3. B5: =B1-B4*(B2-1) is the last payment. It takes whatever cents are left over so the payments add up to the invoice exactly. $1,000.00 in 3 payments is $333.33, $333.33, $333.34.",
        "4. Due dates: =IF(A8>$B$2,\"\",EDATE($B$3,A8-1)) moves the first date forward one month at a time. Rows beyond your number of payments stay blank.",
        "5. EDATE keeps the day of the month, but a month that is too short uses its last day: a first date of 01/31/2026 gives 02/28/2026 and then 03/31/2026.",
        "6. Balance after: =$B$1-SUM($C$8:C8), which should end at $0.00. C22 shows the difference between the payments and the invoice (should be $0.00).",
        "7. Only B2 values from 1 to 12 are supported. This is arithmetic for a plan you already agreed with the client; it adds no interest or fees.",
        "",
        "Amounts and dates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-invoice-installments.xlsx")


def marketplace_net_payout():
    wb = Workbook()
    ws = wb.active
    ws.title = "Net payout"
    header(ws, ["Order", "Sale price", "Fees", "Net payout", "Net % of price"], [26, 13, 12, 14, 15])
    orders = [("Soy candle, 8 oz", 18.00), ("Cedar soap bar", 8.50), ("Brass hoop earrings", 24.00),
              ("Gift tag set", 1.00), ("Candle gift box", 60.00), ("Sticker sample", 0.25)]
    for r, (name, price) in enumerate(orders, start=2):
        ws.append([name, price, f"=ROUND(B{r}*($H$2+$H$3)+$H$4,2)", f"=B{r}-C{r}", f"=D{r}/B{r}"])
        for col in "BCD":
            ws[f"{col}{r}"].number_format = USD
        ws[f"E{r}"].number_format = "0.0%"
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["A8"], ws["B8"], ws["C8"], ws["D8"] = "Total", "=SUM(B2:B7)", "=SUM(C2:C7)", "=SUM(D2:D7)"
    ws["E8"] = "=D8/B8"
    for col in "BCD":
        ws[f"{col}8"].number_format = USD
    ws["E8"].number_format = "0.0%"
    ws["A8"].font = Font(bold=True)

    ws.column_dimensions["G"].width = 34
    ws.column_dimensions["H"].width = 12
    ws["G1"], ws["G1"].font = "Fee settings (fictional rates)", Font(bold=True)
    for r, (label, val, fmt) in enumerate([("Marketplace fee, % of price", 0.065, "0.0%"),
                                           ("Payment processing, % of price", 0.03, "0.0%"),
                                           ("Fixed fee per order", 0.25, USD)], start=2):
        ws[f"G{r}"], ws[f"H{r}"] = label, val
        ws[f"H{r}"].number_format = fmt
        ws[f"H{r}"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["G6"], ws["G6"].font = "Price needed for a target payout", Font(bold=True)
    ws["G7"], ws["H7"] = "Target net payout", 15.00
    ws["G8"], ws["H8"] = "Price to charge", "=ROUNDUP((H7+H4)/(1-H2-H3),2)"
    ws["G9"], ws["H9"] = "Check: net payout at that price", "=H8-ROUND(H8*(H2+H3)+H4,2)"
    ws["H7"].number_format = ws["H8"].number_format = ws["H9"].number_format = USD
    ws["H7"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["H8"].fill = ws["H9"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Net Payout After Marketplace Fees",
        "",
        "1. Net payout tab: yellow cells H2:H4 hold the fee settings. The rates here are made up. Replace them with the rates shown in your own seller account, and update them when the marketplace changes its fees.",
        "2. C2: =ROUND(B2*($H$2+$H$3)+$H$4,2) is the fee on that order: both percentages of the price plus the fixed fee, rounded to cents.",
        "3. D2: =B2-C2 is the net payout. E2: =D2/B2 is the share of the price you keep.",
        "4. The fixed fee hurts small orders most: the $1.00 gift tag set keeps only 65.0% of its price, and the $0.25 sticker sample ends up at a net payout of -$0.02, a loss.",
        "5. H8: =ROUNDUP((H7+H4)/(1-H2-H3),2) works backward to the price that nets your target, rounded up to the next cent. H9 checks it by running the forward formula on that price. Because fees are rounded to cents, the check can land a cent above the target (here $15.01 for a $15.00 target).",
        "6. This sheet covers only the fees you type in. It does not include shipping, sales tax collected, advertising, refunds, or income tax.",
        "",
        "Orders and rates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-marketplace-net-payout.xlsx")


INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")


def invoice_due_dates():
    wb = Workbook()
    ws = wb.active
    ws.title = "Due dates"
    header(ws, ["Invoice", "Client", "Invoice date", "Terms", "Days", "Due date", "Days until due", "Status"],
           [12, 26, 13, 17, 14, 13, 14, 12])
    invoices = [
        ("INV-1001", "Maple Street Bakery", date(2026, 9, 1), "Net 30"),
        ("INV-1002", "Harbor Yoga Studio", date(2026, 9, 15), "Net 15"),
        ("INV-1003", "Birch Design Co", date(2026, 9, 28), "Net 45"),
        ("INV-1004", "Lantern Print Shop", date(2026, 10, 5), "Net 30"),
        ("INV-1005", "Cedar Row Realty", date(2026, 10, 13), "Due on receipt"),
        ("INV-1006", "Willow Pet Care", date(2026, 10, 1), "Net 60"),
        ("INV-1007", "Granite Fitness", date(2026, 8, 20), "Net 60"),
        ("INV-1008", "Juniper Books", date(2026, 10, 10), "Net 15"),
    ]
    last = 20
    for r in range(2, last + 1):
        if r - 2 < len(invoices):
            ws.append(list(invoices[r - 2]))
        else:
            ws.append([])
        ws[f"C{r}"].number_format = US_DATE
        ws[f"E{r}"] = f'=IF(D{r}="","",IFERROR(VLOOKUP(D{r},$J$2:$K$6,2,FALSE),"Check terms"))'
        ws[f"F{r}"] = f'=IF(OR(C{r}="",NOT(ISNUMBER(E{r}))),"",C{r}+E{r})'
        ws[f"G{r}"] = f'=IF(F{r}="","",F{r}-$K$8)'
        ws[f"H{r}"] = f'=IF(G{r}="","",IF(G{r}<0,"Past due",IF(G{r}=0,"Due today","Open")))'
        ws[f"F{r}"].number_format = US_DATE
        for col in "EFGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1="=$J$2:$J$6", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"D2:D{last}")
    ws.conditional_formatting.add(f"H2:H{last}", FormulaRule(formula=['H2="Past due"'], fill=ALERT_FILL))
    ws.column_dimensions["I"].width = 4
    ws.column_dimensions["J"].width = 20
    ws.column_dimensions["K"].width = 12
    ws["J1"], ws["K1"] = "Terms", "Days"
    for c in (ws["J1"], ws["K1"]):
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for r, (name, days) in enumerate([("Due on receipt", 0), ("Net 15", 15), ("Net 30", 30),
                                      ("Net 45", 45), ("Net 60", 60)], start=2):
        ws[f"J{r}"], ws[f"K{r}"] = name, days
        ws[f"J{r}"].fill = ws[f"K{r}"].fill = INPUT_FILL
    ws["J8"], ws["K8"] = "As-of date", date(2026, 10, 13)
    ws["K8"].number_format = US_DATE
    ws["J8"].font = Font(bold=True)
    ws["K8"].fill = INPUT_FILL
    ws["J10"], ws["K10"] = "Open or due today", f'=COUNTIF(H2:H{last},"Open")+COUNTIF(H2:H{last},"Due today")'
    ws["J11"], ws["K11"] = "Past due", f'=COUNTIF(H2:H{last},"Past due")'

    notes_sheet(wb, [
        "Tidy Tabs: Calculate an Invoice Due Date From Payment Terms",
        "",
        "1. Due dates tab: type the invoice number, client, invoice date (C) and pick the terms from the dropdown in column D. Rows 2 to 20 have the formulas ready.",
        "2. The terms table sits in J2:K6 (yellow). Each name has a number of days. Add or change terms by editing those cells; keep the names exactly as the dropdown shows them.",
        "3. E2: =IF(D2=\"\",\"\",IFERROR(VLOOKUP(D2,$J$2:$K$6,2,FALSE),\"Check terms\")) looks up the days. A term that is not in the table shows Check terms.",
        "4. F2: =IF(OR(C2=\"\",NOT(ISNUMBER(E2))),\"\",C2+E2) is the due date: the invoice date plus the days. Dates are plain numbers, so adding days works across month ends (08/20/2026 plus 60 days is 10/19/2026).",
        "5. The short version of the whole sheet is =C2+VLOOKUP(D2,$J$2:$K$6,2,FALSE).",
        "6. K8 is the as-of date. Type =TODAY() there to compare with the current day. G2 is the due date minus the as-of date (negative means days late). H2 says Past due, Due today or Open.",
        "7. The due date is not moved for weekends or holidays, and it does not read terms such as 2/10 net 30 or end of month. Use WORKDAY if you want business days.",
        "8. This is date arithmetic only. The terms on a real invoice come from your contract with the client.",
        "",
        "Invoices, clients and dates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-invoice-due-dates.xlsx")


def sales_by_weekday():
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily sales"
    header(ws, ["Date", "Sales", "Weekday"], [13, 12, 11])
    amounts = [42.00, 38.50, 51.25, 65.00, 148.00, 212.50, 96.00,
               40.00, 35.00, 47.50, 70.25, 155.00, 230.00, 88.50,
               44.50, 41.00, 55.00, 62.00, 162.75, 248.00, 101.00,
               39.00, 36.25, 49.00, 68.00, 171.00, 226.50, 93.00]
    last = 60
    for i in range(2, last + 1):
        if i - 2 < len(amounts):
            ws.append([date(2026, 9, 1 + i - 2), amounts[i - 2]])
        else:
            ws.append([])
        ws[f"A{i}"].number_format = US_DATE
        ws[f"B{i}"].number_format = USD
        ws[f"C{i}"] = f'=IF(A{i}="","",TEXT(A{i},"ddd"))'
        ws[f"C{i}"].fill = FORMULA_FILL

    s = wb.create_sheet("By weekday")
    header(s, ["Weekday", "Days counted", "Total sales", "Average per day", "Share of sales", "Best day"],
           [14, 14, 14, 17, 15, 12])
    rng = "'Daily sales'!$C$2:$C$200"
    for r, day in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"], start=2):
        s.append([day, f"=COUNTIF({rng},A{r})",
                  f"=SUMIF({rng},A{r},'Daily sales'!$B$2:$B$200)",
                  f'=IF(B{r}=0,"",C{r}/B{r})',
                  f'=IF($C$9=0,"",C{r}/$C$9)',
                  f'=IF(D{r}="","",IF(D{r}=MAX($D$2:$D$8),"Best day",""))'])
        s[f"C{r}"].number_format = s[f"D{r}"].number_format = USD
        s[f"E{r}"].number_format = "0.0%"
        for col in "BCDEF":
            s[f"{col}{r}"].fill = FORMULA_FILL
    s["A9"], s["B9"], s["C9"] = "Total", "=SUM(B2:B8)", "=SUM(C2:C8)"
    s["A9"].font = Font(bold=True)
    s["C9"].number_format = USD
    s["A11"] = "Check: days counted match the Daily sales tab"
    s["F11"] = "=IF(B9=COUNT('Daily sales'!A2:A200),\"OK\",\"CHECK\")"
    s["A12"] = "Check: total matches the Daily sales tab"
    s["F12"] = "=IF(ROUND(C9-SUM('Daily sales'!B2:B200),2)=0,\"OK\",\"CHECK\")"

    notes_sheet(wb, [
        "Tidy Tabs: Total Sales by Day of the Week",
        "",
        "1. Daily sales tab: one row per day, with the date in A and that day's sales in B. Column C names the weekday.",
        "2. C2: =IF(A2=\"\",\"\",TEXT(A2,\"ddd\")) turns a date into Mon, Tue and so on. Fill it down as far as your data goes (it is ready to row 60).",
        "3. By weekday tab, C2: =SUMIF('Daily sales'!$C$2:$C$200,A2,'Daily sales'!$B$2:$B$200) adds the sales for each weekday. B2 counts the days with COUNTIF.",
        "4. D2: =IF(B2=0,\"\",C2/B2) is the average per day. The short form is =AVERAGEIF('Daily sales'!$C$2:$C$200,\"Sat\",'Daily sales'!$B$2:$B$200).",
        "5. Compare averages, not totals, when some weekdays appear more often in your data. F2 marks the weekday with the highest average.",
        "6. Dates should be real dates. A date stored as text may still get a weekday name in column C, but the COUNT of real dates skips it, so check cell F11 shows CHECK. Re-enter it as a date.",
        "7. A day with $0.00 sales still counts as a day if it has a row. Days with no row are not counted.",
        "8. The weekday names are English (Mon to Sun). They work the same for a US-English file in Excel and Google Sheets.",
        "",
        "Dates and sales are fictional (28 days from 09/01/2026). Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-sales-by-weekday.xlsx")


def repeat_customers():
    wb = Workbook()
    o = wb.active
    o.title = "Orders"
    header(o, ["Order", "Customer", "Order date", "Total"], [10, 20, 13, 12])
    orders = [
        ("1001", "Ana Ruiz", date(2026, 8, 4), 42.00), ("1002", "Eli Brooks", date(2026, 8, 6), 18.00),
        ("1003", "Ben Ortiz", date(2026, 8, 12), 65.50), ("1004", "Ana Ruiz", date(2026, 8, 20), 30.00),
        ("1005", "Cara Lee", date(2026, 8, 25), 24.00), ("1006", "Eli Brooks", date(2026, 9, 2), 36.00),
        ("1007", "Dev Patel", date(2026, 9, 9), 54.25), ("1008", "Fay Nolan", date(2026, 9, 14), 12.00),
        ("1009", "Eli Brooks", date(2026, 9, 21), 28.75), ("1010", "Cara Lee", date(2026, 9, 28), 48.00),
        ("1011", "Hana Cole", date(2026, 10, 1), 90.00), ("1012", "Ana Ruiz", date(2026, 10, 5), 22.50),
        ("1013", "Eli Brooks", date(2026, 10, 9), 31.00),
    ]
    for r, row in enumerate(orders, start=2):
        o.append(list(row))
        o[f"C{r}"].number_format = US_DATE
        o[f"D{r}"].number_format = USD

    c = wb.create_sheet("Customers")
    header(c, ["Customer", "Orders", "Type", "Total spent"], [20, 10, 12, 14])
    names = ["Ana Ruiz", "Ben Ortiz", "Cara Lee", "Dev Patel", "Eli Brooks", "Fay Nolan", "Gus Hall", "Hana Cole"]
    last = 31
    for r in range(2, last + 1):
        c.append([names[r - 2]] if r - 2 < len(names) else [])
        c[f"B{r}"] = f'=IF(A{r}="","",COUNTIF(Orders!$B$2:$B$200,A{r}))'
        c[f"C{r}"] = f'=IF(A{r}="","",IF(B{r}=0,"No orders",IF(B{r}=1,"One time","Repeat")))'
        c[f"D{r}"] = f'=IF(A{r}="","",SUMIF(Orders!$B$2:$B$200,A{r},Orders!$D$2:$D$200))'
        c[f"D{r}"].number_format = USD
        for col in "BCD":
            c[f"{col}{r}"].fill = FORMULA_FILL
    c.column_dimensions["E"].width = 4
    c.column_dimensions["F"].width = 36
    c.column_dimensions["G"].width = 14
    c["F1"] = "Summary"
    c["F1"].font = Font(bold=True)
    summary = [
        ("Customers with at least one order", f'=COUNTIF(B2:B{last},">0")', "0"),
        ("Repeat customers (2 or more orders)", f'=COUNTIF(B2:B{last},">=2")', "0"),
        ("Repeat customer rate", '=IF(G2=0,"",G3/G2)', "0.0%"),
        ("Orders counted above", f"=SUM(B2:B{last})", "0"),
        ("Orders on the Orders tab", "=COUNTA(Orders!B2:B200)", "0"),
        ("Check: every order matched a customer", '=IF(G5=G6,"OK","CHECK")', "General"),
        ("Spent by repeat customers", f'=SUMIF(C2:C{last},"Repeat",D2:D{last})', USD),
        ("Share of spending from repeat customers", f'=IF(SUM(D2:D{last})=0,"",G8/SUM(D2:D{last}))', "0.0%"),
    ]
    for r, (label, f, fmt) in enumerate(summary, start=2):
        c[f"F{r}"], c[f"G{r}"] = label, f
        c[f"G{r}"].number_format = fmt
        c[f"G{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Repeat Customer Rate With COUNTIF",
        "",
        "1. Orders tab: one row per order, with the customer name in column B. Paste or type your orders under the sample rows.",
        "2. Customers tab: list each customer once in column A . Rows 2 to 31 have formulas ready.",
        "3. B2: =IF(A2=\"\",\"\",COUNTIF(Orders!$B$2:$B$200,A2)) counts that customer's orders. C2 labels them No orders, One time or Repeat. D2 adds up what they spent with SUMIF.",
        "4. G2: =COUNTIF(B2:B31,\">0\") counts customers who ordered at all. G3: =COUNTIF(B2:B31,\">=2\") counts repeat customers.",
        "5. G4: =IF(G2=0,\"\",G3/G2) is the repeat customer rate. The short form is =COUNTIF(B2:B9,\">=2\")/COUNTIF(B2:B9,\">0\").",
        "6. G5 to G7 check that every order on the Orders tab matched a name on the Customers tab. A misspelled name or an extra space at the end of a cell makes G7 show CHECK.",
        "7. COUNTIF ignores capital letters, so ana ruiz and Ana Ruiz are the same customer. It does not ignore spaces or different spellings.",
        "8. The rate covers only the orders you pasted in. Choose the same date range for every rate you compare.",
        "",
        "Names, orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-repeat-customers.xlsx")


def ounces_to_pounds_ounces():
    wb = Workbook()
    ws = wb.active
    ws.title = "Weights"
    header(ws, ["Item", "Packed ounces", "Pounds", "Leftover ounces", "Label"], [26, 14, 9, 16, 16])
    items = [("Lavender candle box", 50), ("Oat soap bundle", 16), ("Earring card", 3.5), ("Canvas tote bag", 22),
             ("Ceramic mug pair", 41), ("Wax melt case", 96), ("Greeting card set", 12.25), ("Tea towel", 9),
             ("Gift crate", 135), ("Print tube", 28)]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(items[r - 2]) if r - 2 < len(items) else [])
        ws[f"B{r}"].fill = INPUT_FILL
        ws[f"C{r}"] = f'=IF(B{r}="","",INT(B{r}/16))'
        ws[f"D{r}"] = f'=IF(B{r}="","",MOD(B{r},16))'
        ws[f"E{r}"] = f'=IF(B{r}="","",INT(B{r}/16)&" lb "&MOD(B{r},16)&" oz")'
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"D{r}"].number_format = "0.00"
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 34
    ws.column_dimensions["H"].width = 16
    ws["G1"] = "Whole shipment"
    ws["G1"].font = Font(bold=True)
    summary = [
        ("Packages with a weight", f"=COUNT(B2:B{last})", "0"),
        ("Total ounces", f"=ROUND(SUM(B2:B{last}),2)", "0.00"),
        ("Total pounds", "=INT(H3/16)", "0"),
        ("Leftover ounces", "=ROUND(MOD(H3,16),2)", "0.00"),
        ("Total label", '=H4&" lb "&H5&" oz"', "General"),
        ("Check: pounds and ounces add back up", f'=IF(ROUND(SUM(C2:C{last})*16+SUM(D2:D{last})-H3,6)=0,"OK","CHECK")', "General"),
    ]
    for r, (label, f, fmt) in enumerate(summary, start=2):
        ws[f"G{r}"], ws[f"H{r}"] = label, f
        ws[f"H{r}"].number_format = fmt
        ws[f"H{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Ounces to Pounds and Ounces With INT and MOD",
        "",
        "1. Weights tab: type each item and its packed weight in ounces. Yellow cells are yours (B2:B30). Gray cells are formulas.",
        "2. C2: =IF(B2=\"\",\"\",INT(B2/16)) gives the whole pounds. 16 ounces make 1 pound.",
        "3. D2: =IF(B2=\"\",\"\",MOD(B2,16)) gives the ounces left over after the whole pounds.",
        "4. E2: =IF(B2=\"\",\"\",INT(B2/16)&\" lb \"&MOD(B2,16)&\" oz\") joins them into text such as 3 lb 2 oz.",
        "5. H3 to H6 do the same for the total of all packages (rounded to 2 decimals so totals do not show long decimals). H7 turns the pounds and ounces back into ounces and shows CHECK if they do not match the total.",
        "6. The label in column E is text. Use columns C and D when you need numbers for a label printer or a carrier form.",
        "7. A blank weight stays blank. A weight of 0 shows 0 lb 0 oz.",
        "8. This only converts units. It does not look up a carrier rate or a weight limit. Read those on your carrier's own pages.",
        "",
        "Items and weights are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-ounces-to-pounds-ounces.xlsx")


def listing_length_check():
    wb = Workbook()
    ws = wb.active
    ws.title = "Listings"
    header(ws, ["Title", "Tags (separated by commas)", "Characters", "Tags", "Title check", "Tag check"],
           [50, 52, 11, 8, 13, 12])
    listings = [
        ("Handmade Lavender Soy Candle 8 oz Gift for Her", "lavender candle, soy candle, gift for her, 8 oz candle"),
        ("Cedar and Sage Soy Candle Small Batch 8 oz Hand Poured Gift", "cedar candle, sage candle, hand poured, soy wax, small batch, housewarming, gift, 8 oz"),
        ("Oatmeal Goat Milk Soap Bar Handmade", "oatmeal soap, goat milk soap, handmade soap, gentle soap, bath gift, soap bar, natural soap, sensitive skin, farmers market"),
        ("Brass Hoop Earrings Minimalist Gold Tone Everyday Jewelry for Women Gift Idea", "brass hoops, gold hoop earrings, minimalist, everyday earrings"),
        ("Letterpress Birthday Card Blank Inside", "birthday card,letterpress,blank card"),
        ("Charcoal Soap Bar Cold Process Vegan", ""),
        ("Eucalyptus Mint Bath Salts 12 oz Jar", "bath salts, eucalyptus, spa gift,"),
        ("Beeswax Taper Candles Pair Natural Unscented", "taper candles, beeswax, natural candle, dinner candles, unscented"),
        ("Linen Wick Soy Candle Tin Travel Size 4 oz", "travel candle, tin candle, soy candle, stocking stuffer, 4 oz, gift, small candle, linen"),
        ("Kraft Gift Box Set Soap and Candle Combo Birthday", "gift set, soap and candle, birthday gift, kraft box, self care"),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(listings[r - 2]) if r - 2 < len(listings) else [])
        ws[f"C{r}"] = f'=IF(A{r}="","",LEN(A{r}))'
        ws[f"D{r}"] = (f'=IF(AND(A{r}="",B{r}=""),"",IF(TRIM(B{r})="",0,LEN(B{r})-LEN(SUBSTITUTE(B{r},",",""))+1'
                       f'-(RIGHT(TRIM(B{r}),1)=",")))')
        ws[f"E{r}"] = f'=IF(A{r}="","",IF(C{r}>$I$2,"Too long","OK"))'
        ws[f"F{r}"] = f'=IF(AND(A{r}="",B{r}=""),"",IF(D{r}>$I$3,"Too many","OK"))'
        for col in "CDEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"E2:F{last}", FormulaRule(formula=['OR(E2="Too long",E2="Too many")'], fill=ALERT_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 36
    ws.column_dimensions["I"].width = 12
    ws["H1"] = "Limits and summary"
    ws["H1"].font = Font(bold=True)
    rows = [
        ("Title limit (characters), type yours", 60, "0"),
        ("Tag limit, type yours", 8, "0"),
        ("Listings checked", f'=COUNTA(A2:A{last})', "0"),
        ("Titles too long", f'=COUNTIF(E2:E{last},"Too long")', "0"),
        ("Listings with too many tags", f'=COUNTIF(F2:F{last},"Too many")', "0"),
        ("Check: characters add up", f'=IF(SUM(C2:C{last})=SUMPRODUCT(LEN(A2:A{last})),"OK","CHECK")', "General"),
    ]
    for r, (label, v, fmt) in enumerate(rows, start=2):
        ws[f"H{r}"], ws[f"I{r}"] = label, v
        ws[f"I{r}"].number_format = fmt
        ws[f"I{r}"].fill = INPUT_FILL if r in (2, 3) else FORMULA_FILL
    ws["H9"] = "Sample limits only. Copy the real"
    ws["H10"] = "numbers from your marketplace's help page."

    notes_sheet(wb, [
        "Tidy Tabs: Count Characters and Tags in a Product Listing",
        "",
        "1. Listings tab: type or paste each title in column A and its tags in column B, separated by commas. Rows 2 to 30 have formulas ready.",
        "2. I2 and I3 are yellow input cells. Type the title limit and the tag limit from your marketplace's own help page. The 60 and 8 in the sample are made-up numbers, not any marketplace's rules.",
        "3. C2: =IF(A2=\"\",\"\",LEN(A2)) counts every character in the title, including spaces.",
        "4. D2 counts the tags: =LEN(B2)-LEN(SUBSTITUTE(B2,\",\",\"\"))+1 counts the commas and adds 1. The template wraps it so an empty tags cell gives 0 and a comma left at the end is not counted as a tag.",
        "5. E and F show OK, Too long or Too many. A title exactly at the limit is OK.",
        "6. I7 adds up the character counts a second way. If it shows CHECK, a formula in column C was overwritten.",
        "7. LEN counts spaces at the start and end of a cell. Spaces after commas do not change the tag count, but two commas in a row do (an empty tag is still counted).",
        "",
        "Listings are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-listing-length-check.xlsx")


def supplier_price_compare():
    wb = Workbook()
    ws = wb.active
    ws.title = "Compare"
    header(ws, ["Supplier", "Package price", "Package size (oz)", "Shipping", "Cost per oz",
                "Cost for amount needed", "Rank"], [24, 14, 14, 11, 13, 17, 8])
    quotes = [
        ("Pine Ridge Wax Co", 42.50, 160, 12.00),
        ("Lakeside Supply", 27.99, 80, 9.50),
        ("Harbor Craft Wholesale", 64.00, 240, 0.00),
        ("Birchwood Candle Goods", 18.50, 48, 7.25),
        ("Maple Lane Supply", 33.00, 100, 8.00),
    ]
    for r, q in enumerate(quotes, start=2):
        ws.append(list(q))
        ws[f"E{r}"] = f'=IF(OR(B{r}="",C{r}<=0),"",(B{r}+D{r})/C{r})'
        ws[f"F{r}"] = f'=IF(E{r}="","",ROUND(E{r}*$I$2,2))'
        ws[f"G{r}"] = f'=IF(E{r}="","",RANK(E{r},$E$2:$E$6,1))'
        ws[f"B{r}"].number_format = ws[f"D{r}"].number_format = ws[f"F{r}"].number_format = USD
        ws[f"E{r}"].number_format = '"$"0.0000'
        for col in "EFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        for col in "ABCD":
            ws[f"{col}{r}"].fill = INPUT_FILL
    ws.column_dimensions["H"].width = 36
    ws.column_dimensions["I"].width = 20
    ws["H1"] = "Compare"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["I2"] = "Amount you need (oz)", 200
    ws["I2"].fill = INPUT_FILL
    ws["H4"] = "Cheapest supplier per oz"
    ws["I4"] = '=IFERROR(INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0)),"Enter prices")'
    ws["H5"], ws["I5"] = "Lowest cost per oz", '=IF(COUNT(E2:E6)=0,"",MIN(E2:E6))'
    ws["H6"], ws["I6"] = "Highest cost per oz", '=IF(COUNT(E2:E6)=0,"",MAX(E2:E6))'
    ws["H7"], ws["I7"] = "Cost of amount needed, cheapest", '=IF(COUNT(F2:F6)=0,"",MIN(F2:F6))'
    ws["H8"], ws["I8"] = "Extra paid with the priciest", '=IF(COUNT(F2:F6)=0,"",MAX(F2:F6)-MIN(F2:F6))'
    for a in ("I5", "I6"):
        ws[a].number_format = '"$"0.0000'
    for a in ("I7", "I8"):
        ws[a].number_format = USD
    for a in ("I4", "I5", "I6", "I7", "I8"):
        ws[a].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Find the Cheapest Supplier by Cost Per Ounce",
        "",
        "1. Compare tab: type each supplier's package price (B), package size in ounces (C) and shipping for that package (D). The yellow cells are inputs.",
        "2. E2: =IF(OR(B2=\"\",C2<=0),\"\",(B2+D2)/C2) is the cost per ounce including shipping. A blank row or a size of 0 stays blank.",
        "3. I4: =INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0)) returns the name of the supplier with the lowest cost per ounce. MIN finds the lowest number, MATCH finds its row, INDEX reads the name.",
        "4. If two suppliers tie, the one listed first is returned. Column G (RANK) shows the tie.",
        "5. I2 is how many ounces you need. F2 multiplies each supplier's cost per ounce by it, rounded to cents.",
        "6. Use one unit for every row (all ounces, or all pounds). Mixing units gives a wrong answer without any error.",
        "7. Shipping is counted once per package. If you buy several packages, or shipping is free above an order size, adjust the shipping cell yourself. Minimum order sizes and delivery times are not compared.",
        "",
        "Suppliers and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-supplier-price-compare.xlsx")


def split_sku_code():
    wb = Workbook()
    ws = wb.active
    ws.title = "SKUs"
    header(ws, ["SKU", "Units sold", "Category", "Style", "Size or color"], [20, 11, 12, 12, 15])
    skus = [("CND-LAV-8OZ", 24), ("CND-LAV-16OZ", 9), ("CND-CED-8OZ", 18), ("SOP-OAT-4OZ", 31),
            ("SOP-CHR-4OZ", 22), ("SOP-OAT-8OZ", 7), ("EAR-HOOP-GLD", 12), ("EAR-STUD-SLV", 15),
            ("CRD-BDAY-A6", 40), ("CRD-THNX-A6", 28)]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(skus[r - 2]) if r - 2 < len(skus) else [])
        ws[f"C{r}"] = f'=IF(A{r}="","",IFERROR(LEFT(A{r},FIND("-",A{r})-1),"Check SKU"))'
        ws[f"D{r}"] = (f'=IF(A{r}="","",IFERROR(MID(A{r},FIND("-",A{r})+1,'
                       f'FIND("-",A{r},FIND("-",A{r})+1)-FIND("-",A{r})-1),"Check SKU"))')
        ws[f"E{r}"] = f'=IF(A{r}="","",IFERROR(MID(A{r},FIND("-",A{r},FIND("-",A{r})+1)+1,LEN(A{r})),"Check SKU"))'
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["H"].width = 12
    ws["G1"], ws["H1"] = "Category code", "Units sold"
    for cell in (ws["G1"], ws["H1"]):
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    for r, code in enumerate(["CND", "SOP", "EAR", "CRD"], start=2):
        ws[f"G{r}"] = code
        ws[f"G{r}"].fill = INPUT_FILL
        ws[f"H{r}"] = f"=SUMIF($C$2:$C${last},G{r},$B$2:$B${last})"
        ws[f"H{r}"].fill = FORMULA_FILL
    ws["G6"], ws["H6"] = "Total of the categories", "=SUM(H2:H5)"
    ws["G7"], ws["H7"] = "All units on the sheet", f"=SUM(B2:B{last})"
    ws["G8"], ws["H8"] = "Check", '=IF(H6=H7,"OK","CHECK")'
    ws["G6"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Split a Product Code Into Parts With LEFT, MID and FIND",
        "",
        "1. SKUs tab: type or paste codes in column A. Each code has three parts joined by two hyphens, such as CND-LAV-8OZ (category, style, size or color).",
        "2. C2: =IFERROR(LEFT(A2,FIND(\"-\",A2)-1),\"Check SKU\") takes everything before the first hyphen. FIND returns the position of the hyphen.",
        "3. D2 takes the text between the first and second hyphen: =MID(A2,FIND(\"-\",A2)+1,FIND(\"-\",A2,FIND(\"-\",A2)+1)-FIND(\"-\",A2)-1). The second FIND starts looking just after the first hyphen.",
        "4. E2: =MID(A2,FIND(\"-\",A2,FIND(\"-\",A2)+1)+1,LEN(A2)) takes everything after the second hyphen.",
        "5. A code with fewer than two hyphens shows Check SKU. A code with a third hyphen keeps the extra part in column E (CND-LAV-8OZ-GIFT gives 8OZ-GIFT).",
        "6. H2: =SUMIF($C$2:$C$30,G2,$B$2:$B$30) adds units sold for each category code. H8 compares the category total with all units: CHECK means a code did not land in any listed category.",
        "7. A space before or after a code breaks the match. Clean pasted codes with TRIM first.",
        "8. The formulas are ready to row 30; fill them further down if you have more codes.",
        "",
        "Codes and quantities are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-split-sku-code.xlsx")


def last_number_in_column():
    wb = Workbook()
    ws = wb.active
    ws.title = "Log"
    header(ws, ["Date", "Balance", "Note"], [13, 14, 30])
    rows = {
        2: (date(2026, 9, 1), 1000.00, "Opening cash"),
        3: (date(2026, 9, 3), 1085.50, "Saturday market"),
        5: (date(2026, 9, 8), 1210.25, "Online orders"),
        6: (date(2026, 9, 12), 1148.00, "Supply order"),
        9: (date(2026, 9, 19), 1320.40, "Saturday market"),
        10: (date(2026, 9, 24), 1257.90, "Booth fee"),
        12: (date(2026, 10, 2), 1402.75, "Online orders"),
    }
    for r in range(2, 31):
        if r in rows:
            d, b, n = rows[r]
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = d, b, n
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        ws[f"A{r}"].number_format = US_DATE
        ws[f"B{r}"].number_format = USD

    s = wb.create_sheet("Summary")
    s.column_dimensions["A"].width = 40
    s.column_dimensions["B"].width = 20
    s["A1"], s["B1"] = "Latest from the Log tab", "Result"
    for c in s[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    items = [
        ("Latest balance", "=LOOKUP(9.99E+307,Log!B2:B200)", USD),
        ("Latest date", "=LOOKUP(9.99E+307,Log!A2:A200)", US_DATE),
        ("Entries in the balance column", "=COUNT(Log!B2:B200)", "0"),
        ("Latest balance, safe for an empty log", '=IFERROR(LOOKUP(9.99E+307,Log!B2:B200),"No entries")', USD),
        ("Latest date, safe for an empty log", '=IFERROR(LOOKUP(9.99E+307,Log!A2:A200),"No entries")', US_DATE),
    ]
    for r, (label, f, fmt) in enumerate(items, start=2):
        s[f"A{r}"], s[f"B{r}"] = label, f
        s[f"B{r}"].number_format = fmt
        s[f"B{r}"].fill = FORMULA_FILL
        s[f"B{r}"].alignment = Alignment(horizontal="right")

    notes_sheet(wb, [
        "Tidy Tabs: Get the Last Number in a Column",
        "",
        "1. Log tab: type each date in column A and the cash balance in column B. Blank rows between entries are fine. Column C is for notes.",
        "2. Summary!B2: =LOOKUP(9.99E+307,Log!B2:B200) returns the last number in B2:B200. 9.99E+307 is bigger than any balance you will type, so LOOKUP settles on the last number in the range.",
        "3. Summary!B3 does the same for the dates in column A (dates are numbers in a spreadsheet).",
        "4. Summary!B4: =COUNT(Log!B2:B200) counts the numbers in the column. It ignores blank cells and text.",
        "5. Summary!B5 and B6 wrap the formula in IFERROR. An empty column makes LOOKUP return #N/A, and these show No entries instead.",
        "6. Text is ignored. A note typed in column B, or a number stored as text (it is usually left-aligned), is skipped, so the formula returns the number above it. Retype it as a real number.",
        "7. A 0 as the last balance is still returned. The formula looks at B2:B200 only; use a longer range if your log grows past row 200.",
        "8. This finds numbers and dates only. A text version needs a different formula.",
        "",
        "Dates and balances are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-last-number-in-column.xlsx")


def time_zone_converter():
    wb = Workbook()
    ws = wb.active
    ws.title = "Calls"
    header(ws, ["Client", "Client zone", "Your call time", "Hours difference", "Client time", "Day"],
           [24, 14, 22, 12, 22, 13])
    calls = [
        ("Maple Street Bakery", "Eastern", datetime(2026, 10, 14, 9, 0)),
        ("Harbor Yoga Studio", "Central", datetime(2026, 10, 14, 16, 30)),
        ("Birch Design Co", "Mountain", datetime(2026, 10, 15, 10, 0)),
        ("Lantern Print Shop", "Pacific", datetime(2026, 10, 14, 16, 30)),
        ("Cedar Row Realty", "Pacific", datetime(2026, 10, 15, 1, 0)),
        ("Willow Pet Care", "Central", datetime(2026, 10, 15, 0, 15)),
        ("Granite Fitness", "Mountain", datetime(2026, 10, 16, 14, 45)),
        ("Juniper Books", "Pacific", datetime(2026, 10, 16, 18, 0)),
    ]
    last = 20
    for r in range(2, last + 1):
        ws.append(list(calls[r - 2]) if r - 2 < len(calls) else [])
        ws[f"C{r}"].number_format = US_DATETIME
        ws[f"D{r}"] = (f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},Offsets!$A$2:$B$5,2,FALSE)'
                       f'-VLOOKUP($I$2,Offsets!$A$2:$B$5,2,FALSE),"Check zone"))')
        ws[f"E{r}"] = (f'=IF(OR(C{r}="",NOT(ISNUMBER(D{r}))),"",IF(ISNUMBER(C{r}),C{r}+D{r}/24,"Check time"))')
        ws[f"F{r}"] = (f'=IF(NOT(ISNUMBER(E{r})),"",IF(INT(E{r})>INT(C{r}),"Next day",'
                       f'IF(INT(E{r})<INT(C{r}),"Previous day","Same day")))')
        ws[f"E{r}"].number_format = US_DATETIME
        ws[f"D{r}"].number_format = "0.##;-0.##;0"
        for col in "DEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1="=Offsets!$A$2:$A$5", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"B2:B{last}")
    dv2 = DataValidation(type="list", formula1="=Offsets!$A$2:$A$5", allow_blank=False)
    ws.add_data_validation(dv2)
    dv2.add("I2")
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['OR(F2="Next day",F2="Previous day")'], fill=ALERT_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 26
    ws.column_dimensions["I"].width = 14
    ws["H2"], ws["I2"] = "Your zone", "Eastern"
    ws["H2"].font = Font(bold=True)
    ws["I2"].fill = INPUT_FILL
    ws["H4"] = "Calls on a different day"
    ws["I4"] = f'=COUNTIF(F2:F{last},"Next day")+COUNTIF(F2:F{last},"Previous day")'
    ws["H5"] = "Calls listed"
    ws["I5"] = f"=COUNT(C2:C{last})"
    ws["I4"].fill = ws["I5"].fill = FORMULA_FILL

    o = wb.create_sheet("Offsets")
    header(o, ["Zone", "Hours from Eastern"], [16, 20])
    for r, (zone, h) in enumerate([("Eastern", 0), ("Central", -1), ("Mountain", -2), ("Pacific", -3)], start=2):
        o[f"A{r}"], o[f"B{r}"] = zone, h
        o[f"A{r}"].fill = o[f"B{r}"].fill = INPUT_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Convert Time Zones for Client Calls",
        "",
        "1. Calls tab: type the client, pick the client zone from the dropdown in B, and type your call date and time in C (for example 10/14/2026 4:30 PM). Rows 2 to 20 have the formulas ready.",
        "2. I2 is your own zone (yellow). Pick Eastern, Central, Mountain or Pacific. The Offsets tab lists each zone as hours from Eastern (Central -1, Mountain -2, Pacific -3).",
        "3. D2 subtracts your zone's offset from the client zone's offset, so Eastern to Pacific is -3 and Pacific to Eastern is 3.",
        "4. E2: =C2+D2/24 is the client's time. One hour is 1/24 of a day, so adding hours divided by 24 moves the date and the time together.",
        "5. F2 compares the dates: Next day, Previous day or Same day, so a call at 1:00 AM for you is not missed as the evening before for the client.",
        "6. Offsets are fixed hours. Most of the US changes clocks on the same dates, so the differences between these four zones stay the same through the year. Arizona (most of it), Hawaii and other places that do not change clocks are not handled, and neither are zones outside these four. Check the current rules before you confirm a call.",
        "7. Type the call time as a real date and time. A time stored as text shows Check time.",
        "",
        "Clients and times are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-time-zone-converter.xlsx")


def count_weekdays_in_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Months"
    header(ws, ["First day", "Last day", "Saturdays", "Sundays", "Day picked in I3", "Booth cost (Saturdays)"],
           [13, 13, 11, 11, 17, 20])
    last = 25
    for r in range(2, last + 1):
        ws.append([date(2026, r - 1, 1)] if r <= 13 else [])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"B{r}"] = f'=IF(A{r}="","",EOMONTH(A{r},0))'
        ws[f"C{r}"] = f'=IF(A{r}="","",NETWORKDAYS.INTL(A{r},B{r},"1111101"))'
        ws[f"D{r}"] = f'=IF(A{r}="","",NETWORKDAYS.INTL(A{r},B{r},"1111110"))'
        ws[f"E{r}"] = f'=IF(A{r}="","",NETWORKDAYS.INTL(A{r},B{r},$I$4))'
        ws[f"F{r}"] = f'=IF(A{r}="","",C{r}*$I$2)'
        ws[f"B{r}"].number_format = US_DATE
        ws[f"F{r}"].number_format = USD
        ws[f"A{r}"].fill = INPUT_FILL
        for col in "BCDEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 34
    ws.column_dimensions["I"].width = 14
    ws.column_dimensions["K"].width = 8
    ws["H1"] = "Settings and totals"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["I2"] = "Booth fee per market day", 25
    ws["I2"].number_format = USD
    ws["H3"], ws["I3"] = "Day to count in column E", "Fri"
    ws["H4"] = "Weekend string for that day"
    ws["I4"] = '=REPT("1",MATCH(I3,$K$2:$K$8,0)-1)&"0"&REPT("1",7-MATCH(I3,$K$2:$K$8,0))'
    ws["I4"].number_format = "@"
    for a in ("I2", "I3"):
        ws[a].fill = INPUT_FILL
    ws["I4"].fill = FORMULA_FILL
    ws["K1"] = "Days"
    ws["K1"].font = Font(bold=True)
    for r, d in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"], start=2):
        ws[f"K{r}"] = d
    dv = DataValidation(type="list", formula1="=$K$2:$K$8", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("I3")
    ws["H6"], ws["I6"] = "Saturdays in the year", f"=SUM(C2:C{last})"
    ws["H7"], ws["I7"] = "Sundays in the year", f"=SUM(D2:D{last})"
    ws["H8"], ws["I8"] = "Booth cost for the year", f"=SUM(F2:F{last})"
    ws["I8"].number_format = USD
    ws["H9"] = "Check: Saturdays match one count over the whole span"
    ws["I9"] = (f'=IF(COUNT(A2:A{last})=0,"",IF(I6=NETWORKDAYS.INTL(MIN(A2:A{last}),'
                f'EOMONTH(MAX(A2:A{last}),0),"1111101"),"OK","CHECK"))')
    for a in ("I6", "I7", "I8", "I9"):
        ws[a].fill = FORMULA_FILL
    ws["H9"].alignment = Alignment(wrap_text=True, vertical="top")

    notes_sheet(wb, [
        "Tidy Tabs: Count Market Saturdays in a Month With NETWORKDAYS.INTL",
        "",
        "1. Months tab: column A holds the first day of each month (yellow = input). Rows 2 to 13 are January to December 2026. Rows 14 to 25 are ready for more months.",
        "2. B2: =EOMONTH(A2,0) gives the last day of that month.",
        "3. C2: =NETWORKDAYS.INTL(A2,B2,\"1111101\") counts Saturdays. The string has seven digits for Monday to Sunday. 1 means skip that day, 0 means count it. Only the sixth digit (Saturday) is 0.",
        "4. D2 uses \"1111110\" to count Sundays instead.",
        "5. I3 picks any weekday from a dropdown. I4 builds its string with REPT, and E2 uses it: =NETWORKDAYS.INTL(A2,B2,$I$4).",
        "6. F2: =C2*$I$2 multiplies the Saturdays by the booth fee per market day in I2. It counts calendar Saturdays only. Your real market days, holidays and fees are set by the market.",
        "7. The first date counts too. If the start date is after the end date, the result is negative. The dates must be real dates, not text.",
        "8. I9 checks that the monthly Saturdays add up to one count over the whole span.",
        "",
        "Dates and fees are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-count-weekdays-in-month.xlsx")


def days_to_get_paid():
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    header(ws, ["Invoice", "Client", "Invoice date", "Paid date", "Days to pay", "Status"],
           [12, 24, 13, 13, 12, 12])
    rows = [
        ("INV-2001", "Maple Street Bakery", date(2026, 7, 1), date(2026, 7, 21)),
        ("INV-2002", "Harbor Yoga Studio", date(2026, 7, 8), date(2026, 8, 12)),
        ("INV-2003", "Birch Design Co", date(2026, 7, 15), date(2026, 7, 29)),
        ("INV-2004", "Lantern Print Shop", date(2026, 7, 22), date(2026, 8, 26)),
        ("INV-2005", "Cedar Row Realty", date(2026, 8, 3), date(2026, 8, 10)),
        ("INV-2006", "Willow Pet Care", date(2026, 8, 10), date(2026, 9, 24)),
        ("INV-2007", "Granite Fitness", date(2026, 8, 19), date(2026, 9, 8)),
        ("INV-2008", "Juniper Books", date(2026, 9, 2), date(2026, 9, 30)),
        ("INV-2009", "Maple Street Bakery", date(2026, 9, 16), None),
        ("INV-2010", "Harbor Yoga Studio", date(2026, 10, 1), None),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"E{r}"] = f'=IF(OR(C{r}="",D{r}=""),"",D{r}-C{r})'
        ws[f"F{r}"] = f'=IF(C{r}="","",IF(D{r}="","Unpaid","Paid"))'
        for col in "CD":
            ws[f"{col}{r}"].number_format = US_DATE
        ws[f"E{r}"].number_format = "0"
        for col in "ABCD":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "EF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"E2:E{last}", FormulaRule(formula=['AND(ISNUMBER(E2),E2>$J$2)'], fill=ALERT_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 30
    ws.column_dimensions["J"].width = 12
    ws["H1"] = "Settings and results"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["J2"] = "Your payment terms in days", 30
    ws["J2"].fill = INPUT_FILL
    results = [
        ("Average days to get paid", f'=IFERROR(AVERAGE(E2:E{last}),"")', "0.0"),
        ("Fastest payment (days)", f'=IF(COUNT(E2:E{last})=0,"",MIN(E2:E{last}))', "0"),
        ("Slowest payment (days)", f'=IF(COUNT(E2:E{last})=0,"",MAX(E2:E{last}))', "0"),
        ("Invoices paid", f"=COUNT(E2:E{last})", "0"),
        ("Invoices still unpaid", f'=COUNTIF(F2:F{last},"Unpaid")', "0"),
        ("Paid later than terms", f'=COUNTIF(E2:E{last},">"&J2)', "0"),
        ("Check: paid + unpaid = invoices", f'=IF(J7+J8=COUNTA(A2:A{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(results, start=4):
        ws[f"H{i}"], ws[f"J{i}"] = label, f
        ws[f"J{i}"].number_format = fmt
        ws[f"J{i}"].fill = FORMULA_FILL
        ws[f"J{i}"].alignment = Alignment(horizontal="right")

    notes_sheet(wb, [
        "Tidy Tabs: Calculate Average Days to Get Paid",
        "",
        "1. Invoices tab: type the invoice number, client, invoice date (A to C) and, once the client pays, the paid date in D. Leave D blank while an invoice is unpaid. Yellow cells are inputs; rows 2 to 30 have formulas ready.",
        "2. E2: =D2-C2 is the days between the invoice date and the paid date. Dates are numbers, so subtracting gives days. A blank paid date leaves E blank.",
        "3. J4: =AVERAGE(E2:E30) is the average days to get paid. AVERAGE skips blank cells, so unpaid invoices do not pull it down to zero.",
        "4. J5 and J6 give the fastest and slowest payment with MIN and MAX. J7 and J8 count paid and unpaid invoices.",
        "5. J2 is your payment terms in days (30 for Net 30). E turns red when an invoice took longer, and J9 counts them.",
        "6. J10 checks that paid plus unpaid equals the number of invoices listed.",
        "7. Enter real dates. A date stored as text (usually left-aligned) can give a wrong number or an error. A paid date earlier than the invoice date gives a negative number; fix the typo.",
        "8. This looks back at invoices that were paid. It does not age open invoices; the receivables aging template does that.",
        "",
        "Clients, dates and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-days-to-get-paid.xlsx")


def early_payment_discount():
    wb = Workbook()
    ws = wb.active
    ws.title = "Discounts"
    header(ws, ["Invoice", "Client", "Invoice date", "Amount", "Paid date", "Discount deadline",
                "Discount amount", "Pay by deadline", "Full amount due date", "Status"],
           [11, 22, 12, 12, 12, 13, 12, 13, 13, 16])
    rows = [
        ("INV-3001", "Maple Street Bakery", date(2026, 9, 1), 1250.00, date(2026, 9, 9)),
        ("INV-3002", "Harbor Yoga Studio", date(2026, 9, 3), 480.00, date(2026, 9, 20)),
        ("INV-3003", "Birch Design Co", date(2026, 9, 8), 2200.00, date(2026, 9, 18)),
        ("INV-3004", "Lantern Print Shop", date(2026, 9, 15), 835.50, date(2026, 9, 30)),
        ("INV-3005", "Cedar Row Realty", date(2026, 9, 22), 960.00, date(2026, 9, 30)),
        ("INV-3006", "Willow Pet Care", date(2026, 10, 1), 312.50, None),
        ("INV-3007", "Granite Fitness", date(2026, 10, 5), 1480.00, None),
        ("INV-3008", "Juniper Books", date(2026, 10, 9), 725.40, None),
    ]
    last = 25
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"F{r}"] = f'=IF(C{r}="","",C{r}+$M$3)'
        ws[f"G{r}"] = f'=IF(OR(C{r}="",NOT(ISNUMBER(D{r}))),"",ROUND(D{r}*$M$2,2))'
        ws[f"H{r}"] = f'=IF(G{r}="","",D{r}-G{r})'
        ws[f"I{r}"] = f'=IF(C{r}="","",C{r}+$M$4)'
        ws[f"J{r}"] = f'=IF(OR(C{r}="",G{r}=""),"",IF(E{r}="","Open",IF(E{r}<=F{r},"Discount earned","Full amount")))'
        for col in "CEFI":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "DGH":
            ws[f"{col}{r}"].number_format = USD
        for col in "ABCDE":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "FGHIJ":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"J2:J{last}", FormulaRule(formula=['J2="Full amount"'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"J2:J{last}", FormulaRule(formula=['J2="Discount earned"'], fill=DONE_FILL))
    ws.column_dimensions["K"].width = 4
    ws.column_dimensions["L"].width = 32
    ws.column_dimensions["M"].width = 14
    ws["L1"] = "Terms and totals"
    ws["L1"].font = Font(bold=True)
    ws["L2"], ws["M2"] = "Discount rate", 0.02
    ws["L3"], ws["M3"] = "Discount days", 10
    ws["L4"], ws["M4"] = "Net days (full amount due)", 30
    ws["M2"].number_format = "0.0%"
    for a in ("M2", "M3", "M4"):
        ws[a].fill = INPUT_FILL
    totals = [
        ("Discounts taken by clients", f'=SUMIF(J2:J{last},"Discount earned",G2:G{last})', USD),
        ("Discounts missed", f'=SUMIF(J2:J{last},"Full amount",G2:G{last})', USD),
        ("Discount still available (open)", f'=SUMIF(J2:J{last},"Open",G2:G{last})', USD),
        ("Open invoices", f'=COUNTIF(J2:J{last},"Open")', "0"),
    ]
    for i, (label, f, fmt) in enumerate(totals, start=6):
        ws[f"L{i}"], ws[f"M{i}"] = label, f
        ws[f"M{i}"].number_format = fmt
        ws[f"M{i}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Calculate an Early Payment Discount (2/10 Net 30)",
        "",
        "1. Discounts tab: type the invoice, client, invoice date and amount (A to D). When a client pays, type the paid date in E; leave it blank while open. Rows 2 to 25 have formulas ready.",
        "2. M2, M3 and M4 hold your terms: discount rate (2%), discount days (10) and net days (30). 2/10 net 30 means 2% off if paid within 10 days, otherwise the full amount in 30. Change them for terms like 1/15 net 45.",
        "3. F2: =C2+$M$3 is the last day to pay and still take the discount.",
        "4. G2: =ROUND(D2*$M$2,2) is the discount in dollars. H2: =D2-G2 is what the client pays by the deadline. I2: =C2+$M$4 is the date the full amount is due.",
        "5. J2 compares the paid date with the deadline. A payment on the deadline day still earns the discount. Blank paid date shows Open.",
        "6. M6 to M9 total the discounts taken, missed and still available.",
        "7. Dates must be real dates, not text. The discount is calendar days, not business days.",
        "8. This is arithmetic from the terms you type. The terms that count are the ones in your contract or invoice; the template does not decide what a client owes.",
        "",
        "Clients, dates and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-early-payment-discount.xlsx")


def sell_through_rate():
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    header(ws, ["Product", "Units received", "Units sold", "Units left", "Sell through", "Status"],
           [30, 14, 12, 12, 13, 14])
    rows = [
        ("Lavender soy candle, 8 oz", 48, 41),
        ("Cedar soy candle, 8 oz", 48, 30),
        ("Oatmeal goat milk soap", 120, 96),
        ("Charcoal soap bar", 60, 21),
        ("Brass hoop earrings", 30, 12),
        ("Letterpress birthday card", 100, 78),
        ("Linen tote bag", 25, 25),
        ("Gift tag set", 80, 18),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"D{r}"] = f'=IF(OR(B{r}="",C{r}=""),"",B{r}-C{r})'
        ws[f"E{r}"] = f'=IF(OR(NOT(ISNUMBER(B{r})),NOT(ISNUMBER(C{r}))),"",IF(B{r}<=0,"",C{r}/B{r}))'
        ws[f"F{r}"] = (f'=IF(E{r}="","",IF(E{r}>1,"Check counts",IF(E{r}>=$I$2,"Strong",'
                       f'IF(E{r}<$I$3,"Slow","OK"))))')
        ws[f"E{r}"].number_format = "0.0%"
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['F2="Slow"'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['F2="Strong"'], fill=DONE_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 30
    ws.column_dimensions["I"].width = 12
    ws["H1"] = "Settings and totals"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["I2"] = "Strong at or above", 0.7
    ws["H3"], ws["I3"] = "Slow below", 0.4
    for a in ("I2", "I3"):
        ws[a].number_format = "0%"
        ws[a].fill = INPUT_FILL
    totals = [
        ("Units received", f"=SUM(B2:B{last})", "0"),
        ("Units sold", f"=SUM(C2:C{last})", "0"),
        ("Overall sell through", f'=IF(I5=0,"",I6/I5)', "0.0%"),
        ("Products marked Slow", f'=COUNTIF(F2:F{last},"Slow")', "0"),
        ("Units left on those", f'=SUMIF(F2:F{last},"Slow",D2:D{last})', "0"),
    ]
    for i, (label, f, fmt) in enumerate(totals, start=5):
        ws[f"H{i}"], ws[f"I{i}"] = label, f
        ws[f"I{i}"].number_format = fmt
        ws[f"I{i}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Calculate Sell Through Rate",
        "",
        "1. Products tab: type the product, the units you received or made (B) and the units sold so far (C). Rows 2 to 30 have formulas ready.",
        "2. D2: =B2-C2 is the units left.",
        "3. E2: =C2/B2 is the sell through rate: units sold divided by units received. 41 of 48 candles sold is 85.4%. Format the cell as a percent.",
        "4. E2 stays blank when units received is 0 or blank, so there is no divide by zero error.",
        "5. I2 and I3 are your own cutoffs (yellow). F2 shows Strong at or above I2, Slow below I3, otherwise OK. The 70% and 40% are examples, not industry rules.",
        "6. If units sold is more than units received, F2 shows Check counts. Fix a typo or add the missing stock.",
        "7. I7 is the overall rate: total sold divided by total received, not the average of the row percents.",
        "8. Compare products over the same time period. This sheet counts units, not dollars.",
        "",
        "Products and counts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-sell-through-rate.xlsx")


def boxes_needed():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Item", "Units to ship", "Units per box", "Boxes needed", "Units in last box",
                "Empty slots", "Box cost"],
           [28, 12, 12, 12, 13, 11, 12])
    rows = [
        ("Soy candle 8 oz", 50, 12),
        ("Soap bars", 30, 8),
        ("Earring cards", 17, 20),
        ("Ceramic mugs", 26, 6),
        ("Linen tote bags", 48, 12),
        ("Birthday cards", 130, 25),
        ("Gift tag sets", 0, 40),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ok = f'AND(ISNUMBER(B{r}),ISNUMBER(C{r}),C{r}>0)'
        ws[f"D{r}"] = f'=IF({ok},ROUNDUP(B{r}/C{r},0),"")'
        ws[f"E{r}"] = f'=IF(OR(D{r}="",B{r}=0),"",B{r}-(D{r}-1)*C{r})'
        ws[f"F{r}"] = f'=IF(D{r}="","",D{r}*C{r}-B{r})'
        ws[f"G{r}"] = f'=IF(D{r}="","",D{r}*$J$2)'
        ws[f"G{r}"].number_format = USD
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DEFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.column_dimensions["H"].width = 4
    ws.column_dimensions["I"].width = 28
    ws.column_dimensions["J"].width = 12
    ws["I1"] = "Cost and totals"
    ws["I1"].font = Font(bold=True)
    ws["I2"], ws["J2"] = "Cost per box", 1.45
    ws["J2"].number_format = USD
    ws["J2"].fill = INPUT_FILL
    totals = [
        ("Total boxes", f"=SUM(D2:D{last})", "0"),
        ("Total box cost", f"=SUM(G2:G{last})", USD),
        ("Total empty slots", f"=SUM(F2:F{last})", "0"),
    ]
    for i, (label, f, fmt) in enumerate(totals, start=4):
        ws[f"I{i}"], ws[f"J{i}"] = label, f
        ws[f"J{i}"].number_format = fmt
        ws[f"J{i}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Calculate How Many Boxes You Need",
        "",
        "1. Orders tab: type the item, the units to ship (B) and how many fit in one box (C). Rows 2 to 30 have formulas ready.",
        "2. D2: =ROUNDUP(B2/C2,0) is the boxes needed. ROUNDUP always rounds up, so 50 candles at 12 per box (4.17) needs 5 boxes. A plain ROUND would give 4 and leave candles without a box.",
        "3. E2: =B2-(D2-1)*C2 is how many units go in the last box. When the units divide evenly the last box is simply full.",
        "4. F2: =D2*C2-B2 is the empty slots across all boxes.",
        "5. G2: =D2*$J$2 is the box cost. J2 is your price per box (yellow).",
        "6. If units per box is 0, blank or text, D to G stay blank. If units to ship is 0, D shows 0 boxes.",
        "7. Same size boxes only. It does not check weight, box dimensions or carrier rules, and one box size per row.",
        "",
        "Items, box sizes and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-boxes-needed.xlsx")


def median_order_value():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order", "Date", "Order total"], [12, 13, 14])
    totals_list = [24.50, 38.00, 42.75, 18.00, 55.25, 31.50, 640.00, 27.00, 46.50, 35.25, 29.00, 62.00, 33.75]
    last = 100
    for r in range(2, last + 1):
        i = r - 2
        if i < len(totals_list):
            ws.append([f"ORD-{4001 + i}", date(2026, 9, 1 + 2 * i), totals_list[i]])
        else:
            ws.append([])
        ws[f"B{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 34
    ws.column_dimensions["F"].width = 14
    ws["E1"] = "Settings and results"
    ws["E1"].font = Font(bold=True)
    ws["E2"], ws["F2"] = "Ignore orders above (used in F9, F10)", 500
    ws["F2"].number_format = USD
    ws["F2"].fill = INPUT_FILL
    res = [
        ("Orders counted", f"=COUNT(C2:C{last})", "0"),
        ("Average order value", f'=IFERROR(AVERAGE(C2:C{last}),"")', USD),
        ("Median order value", f'=IFERROR(MEDIAN(C2:C{last}),"")', USD),
        ("Average minus median", f'=IF(F5="","",F5-F6)', USD),
        ("Largest order", f'=IF(F4=0,"",MAX(C2:C{last}))', USD),
        ("Average without orders above F2", f'=IFERROR(AVERAGEIF(C2:C{last},"<="&F2),"")', USD),
        ("Orders above F2", f'=COUNTIF(C2:C{last},">"&F2)', "0"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=4):
        ws[f"E{i}"], ws[f"F{i}"] = label, f
        ws[f"F{i}"].number_format = fmt
        ws[f"F{i}"].fill = FORMULA_FILL
        ws[f"F{i}"].alignment = Alignment(horizontal="right")

    notes_sheet(wb, [
        "Tidy Tabs: Find the Median Order Value",
        "",
        "1. Orders tab: type the order number, date and order total in A to C (yellow). Totals can go down to row 100.",
        "2. F5: =AVERAGE(C2:C100) adds every order and divides by the count. One big wholesale order pulls it up.",
        "3. F6: =MEDIAN(C2:C100) is the middle order when the totals are sorted. With an even count it is the average of the two middle orders. A single large order barely moves it.",
        "4. F7 shows how far the average sits above the median. A large gap means a few big orders are skewing the average.",
        "5. F9 and F10 are optional: type a cutoff in F2 and F9 averages only orders at or below it, F10 counts the orders above it. The cutoff is your call; it is not a rule.",
        "6. Blank cells and text are ignored by AVERAGE and MEDIAN. A total typed as text (usually left-aligned) is skipped, so retype it as a number. With no orders, F5 and F6 stay blank.",
        "7. Order totals only. It does not remove refunds or tax; enter the totals you want to measure.",
        "",
        "Orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-median-order-value.xlsx")

def week_ending_date():
    wb = Workbook()
    ws = wb.active
    ws.title = "Log"
    header(ws, ["Date", "Client", "Hours", "Week ending"], [13, 24, 10, 14])
    rows = [
        (date(2026, 9, 28), "Maple Street Bakery", 3.5),
        (date(2026, 9, 29), "Harbor Yoga Studio", 4.0),
        (date(2026, 9, 30), "Maple Street Bakery", 2.5),
        (date(2026, 10, 1), "Birch Design Co", 6.0),
        (date(2026, 10, 2), "Birch Design Co", 5.5),
        (date(2026, 10, 3), "Harbor Yoga Studio", 2.0),
        (date(2026, 10, 4), "Lantern Print Shop", 1.5),
        (date(2026, 10, 5), "Lantern Print Shop", 7.0),
        (date(2026, 10, 7), "Maple Street Bakery", 4.25),
        (date(2026, 10, 9), "Birch Design Co", 3.0),
        (date(2026, 10, 12), "Harbor Yoga Studio", 5.0),
        (date(2026, 10, 14), "Lantern Print Shop", 2.75),
    ]
    last = 60
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"D{r}"] = f'=IF(ISNUMBER(A{r}),A{r}+MOD($G$2-WEEKDAY(A{r},1),7),"")'
        ws[f"A{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = "0.00"
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        ws[f"D{r}"].fill = FORMULA_FILL
    ws.column_dimensions["E"].width = 4
    ws.column_dimensions["F"].width = 34
    ws.column_dimensions["G"].width = 12
    ws["F1"] = "Setting"
    ws["F1"].font = Font(bold=True)
    ws["F2"], ws["G2"] = "Week ends on (1 Sun, 2 Mon ... 6 Fri, 7 Sat)", 6
    ws["G2"].fill = INPUT_FILL
    ws["F3"] = "Week ends on"
    ws["G3"] = '=IF(AND(ISNUMBER(G2),G2>=1,G2<=7),CHOOSE(G2,"Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"),"Enter 1 to 7")'
    ws["G3"].fill = FORMULA_FILL
    ws["F4"], ws["G4"] = "Total hours in log", f"=SUM(C2:C{last})"
    ws["G4"].number_format = "0.00"
    ws["G4"].fill = FORMULA_FILL
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="7", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("G2")

    wk = wb.create_sheet("Weeks")
    header(wk, ["Week ending", "Hours"], [14, 10])
    wlast = 27
    for r in range(2, wlast + 1):
        if r == 2:
            wk["A2"] = f'=IF(COUNT(Log!D2:D{last})=0,"",MIN(Log!D2:D{last}))'
        else:
            wk[f"A{r}"] = f'=IF(A{r-1}="","",IF(A{r-1}+7>MAX(Log!D2:D{last}),"",A{r-1}+7))'
        wk[f"B{r}"] = f'=IF(A{r}="","",SUMIFS(Log!C2:C{last},Log!D2:D{last},A{r}))'
        wk[f"A{r}"].number_format = US_DATE
        wk[f"B{r}"].number_format = "0.00"
        wk[f"A{r}"].fill = FORMULA_FILL
        wk[f"B{r}"].fill = FORMULA_FILL
    wk.column_dimensions["C"].width = 4
    wk.column_dimensions["D"].width = 34
    wk.column_dimensions["E"].width = 12
    wk["D1"] = "Check"
    wk["D1"].font = Font(bold=True)
    wk["D2"], wk["E2"] = "Hours on this tab", f"=SUM(B2:B{wlast})"
    wk["D3"], wk["E3"] = "Hours in the log", f"=SUM(Log!C2:C{last})"
    wk["D4"], wk["E4"] = "Check: they match", '=IF(ROUND(E2-E3,6)=0,"OK","CHECK")'
    for a in ("E2", "E3"):
        wk[a].number_format = "0.00"
    for a in ("E2", "E3", "E4"):
        wk[a].fill = FORMULA_FILL
        wk[a].alignment = Alignment(horizontal="right")
    wk.conditional_formatting.add("E4", FormulaRule(formula=['E4="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Find the Week Ending Date for Any Date",
        "",
        "1. Log tab: type the date, client and hours in A to C (yellow). Rows 2 to 60 have formulas ready.",
        "2. G2 is the weekday that ends your week: 1 Sunday, 2 Monday, 3 Tuesday, 4 Wednesday, 5 Thursday, 6 Friday, 7 Saturday. It starts at 6 (Friday).",
        "3. D2: =A2+MOD($G$2-WEEKDAY(A2,1),7). WEEKDAY(A2,1) gives 1 for Sunday to 7 for Saturday. MOD counts the days forward to the week ending day, from 0 to 6.",
        "4. A date that already falls on the week ending day stays the same. The day after it jumps ahead 6 days.",
        "5. Weeks tab: A2 is the earliest week ending in the log and each row below adds 7 days. B uses SUMIFS to add the hours whose week ending matches.",
        "6. Weeks!E4 checks that the hours on the Weeks tab equal the hours in the log. CHECK means a row was left out, for example a date typed as text, or more weeks than the 26 rows listed.",
        "7. Enter real dates. A date stored as text (usually left-aligned) gets no week ending, and its hours show up as CHECK. A blank date row stays blank.",
        "8. Date math only. It does not apply pay period rules, overtime or time zones.",
        "",
        "Clients, dates and hours are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-week-ending-date.xlsx")


def free_shipping_threshold():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order", "Subtotal", "Delivery", "Status", "Amount to add"], [12, 12, 12, 12, 14])
    rows = [
        ("ORD-5001", 32.50, "Ship"),
        ("ORD-5002", 50.00, "Ship"),
        ("ORD-5003", 49.99, "Ship"),
        ("ORD-5004", 78.25, "Ship"),
        ("ORD-5005", 24.00, "Pickup"),
        ("ORD-5006", 45.00, "Ship"),
        ("ORD-5007", 112.40, "Ship"),
        ("ORD-5008", 18.75, "Ship"),
        ("ORD-5009", 61.00, "Pickup"),
        ("ORD-5010", 41.50, "Ship"),
    ]
    last = 31
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"D{r}"] = (f'=IF(OR(NOT(ISNUMBER(B{r})),C{r}=""),"",IF(C{r}="Pickup","Pickup",'
                       f'IF(AND(C{r}="Ship",B{r}>=$I$2),"Free","Pay")))')
        ws[f"E{r}"] = f'=IF(OR(D{r}="Pay",D{r}="Free"),MAX(0,ROUND($I$2-B{r},2)),"")'
        ws[f"B{r}"].number_format = USD
        ws[f"E{r}"].number_format = USD
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1='"Ship,Pickup"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C2:C{last}")
    ws.conditional_formatting.add(f"D2:D{last}", FormulaRule(formula=['D2="Free"'], fill=DONE_FILL))
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 34
    ws.column_dimensions["I"].width = 12
    ws["G1"] = "Settings and results"
    ws["G1"].font = Font(bold=True)
    ws["G2"], ws["I2"] = "Free shipping at or above", 50
    ws["G3"], ws["I3"] = "Close = within this much", 10
    for a in ("I2", "I3"):
        ws[a].number_format = USD
        ws[a].fill = INPUT_FILL
    res = [
        ("Orders that qualify (Free)", f'=COUNTIF(D2:D{last},"Free")', "0"),
        ("Orders that pay shipping", f'=COUNTIF(D2:D{last},"Pay")', "0"),
        ("Pay orders within the close gap", f'=COUNTIFS(D2:D{last},"Pay",E2:E{last},"<="&I3)', "0"),
        ("Pickup orders", f'=COUNTIF(D2:D{last},"Pickup")', "0"),
        ("Check: all rows counted", f'=IF(I4+I5+I7=COUNTA(A2:A{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=4):
        ws[f"G{i}"], ws[f"I{i}"] = label, f
        ws[f"I{i}"].number_format = fmt
        ws[f"I{i}"].fill = FORMULA_FILL
        ws[f"I{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("I8", FormulaRule(formula=['I8="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Check a Free Shipping Threshold",
        "",
        "1. Orders tab: type the order number, subtotal and delivery (Ship or Pickup) in A to C (yellow). Rows 2 to 31 have formulas ready.",
        "2. I2 is your free shipping threshold. I3 is how close counts as close. Both are examples; type your own.",
        "3. D2: =IF(AND(C2=\"Ship\",B2>=$I$2),\"Free\",\"Pay\"), wrapped so Pickup rows show Pickup. A subtotal exactly equal to I2 qualifies (>=); one cent under does not.",
        "4. E2: =MAX(0,$I$2-B2) is how much more the customer must add to reach free shipping. It is 0 when the order already qualifies. ROUND(...,2) keeps cents exact.",
        "5. I4 to I7 count Free, Pay, Pay orders within the close gap (the customers a nudge could convert) and Pickup rows.",
        "6. I8 checks that Free + Pay + Pickup equals the number of order numbers listed. CHECK usually means a subtotal typed as text or a blank delivery cell.",
        "7. A subtotal typed as text (usually left-aligned) or a blank subtotal gets no status.",
        "8. Your own threshold only. It does not know carrier rates or marketplace rules.",
        "",
        "Orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-free-shipping-threshold.xlsx")


def forecast_next_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Month number", "Month", "Sales"], [14, 12, 13])
    sales = [1200, 1350, 1500, 1480, 1620, 1710, 1690, 1850]
    last = 25
    for r in range(2, last + 1):
        i = r - 2
        ws.append([i + 1, date(2026, 1 + i, 1), sales[i]] if i < len(sales) else [])
        ws[f"B{r}"].number_format = "mmm yyyy"
        ws[f"C{r}"].number_format = USD
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
    ws.column_dimensions["D"].width = 4
    for col, w in (("E", 22), ("F", 14), ("G", 14)):
        ws.column_dimensions[col].width = w
    ws["E1"], ws["F1"], ws["G1"] = "Forecast", "Month", "Sales estimate"
    for a in ("E1", "F1", "G1"):
        ws[a].font = Font(bold=True)
    X, Y = f"$A$2:$A${last}", f"$C$2:$C${last}"
    for k in (1, 2, 3):
        r = 1 + k
        ws[f"E{r}"] = f'=IF(COUNT({X})=0,"",MAX({X})+{k})'
        ws[f"F{r}"] = f'=IF(OR(E{r}="",COUNT($B$2:$B${last})=0),"",EDATE(MAX($B$2:$B${last}),E{r}-MAX({X})))'
        ws[f"G{r}"] = f'=IF(OR(E{r}="",COUNT({Y})<2),"",FORECAST(E{r},{Y},{X}))'
        ws[f"F{r}"].number_format = "mmm yyyy"
        ws[f"G{r}"].number_format = USD
        for col in "EFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"E{r}"].number_format = "0"
    ws["E6"], ws["G6"] = "Trend per month (SLOPE)", f'=IF(COUNT({Y})<2,"",SLOPE({Y},{X}))'
    ws["E7"], ws["G7"] = "Months of data", f"=COUNT({Y})"
    ws["E8"], ws["G8"] = "Hand calculation, month 1", (
        f'=IF(COUNT({Y})<2,"",(AVERAGE({Y})+(SUMPRODUCT({X},{Y})-COUNT({Y})*AVERAGE({X})*AVERAGE({Y}))'
        f'/(SUMSQ({X})-COUNT({X})*AVERAGE({X})^2)*(E2-AVERAGE({X}))))')
    ws["E9"], ws["G9"] = "Check: FORECAST matches", '=IF(G2="","",IF(ROUND(G2-G8,6)=0,"OK","CHECK"))'
    ws["G6"].number_format = USD
    ws["G8"].number_format = USD
    ws["G7"].number_format = "0"
    for r in range(6, 10):
        ws[f"G{r}"].fill = FORMULA_FILL
        ws[f"G{r}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("G9", FormulaRule(formula=['G9="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Forecast Next Month's Sales With FORECAST",
        "",
        "1. Sales tab: type a month number (1, 2, 3 ...), the month and the sales in A to C (yellow). Rows 2 to 25 are ready for up to 24 months. Keep the month numbers in order, with no gaps.",
        "2. G2: =FORECAST(E2,$C$2:$C$25,$A$2:$A$25). E2 is the next month number (the last one plus 1). FORECAST fits a straight line through your past months and reads the sales value off that line. E3 and E4 give the next two months.",
        "3. G6: =SLOPE($C$2:$C$25,$A$2:$A$25) is how many dollars the line rises or falls each month.",
        "4. G8 repeats the first projection by the least squares formula typed out by hand, and G9 shows OK when FORECAST matches it. CHECK means the month numbers have gaps or text.",
        "5. A straight line is only an estimate drawn from the past. It does not know about seasons, sales or promotions, and it is not a prediction. Use it as one input.",
        "6. Rows with a blank or text month number or sales amount are skipped as a pair by LibreOffice Calc. Retype text as a number.",
        "7. FORECAST needs at least 2 months; with fewer the results stay blank. If the two ranges are different sizes the formula returns an error (#VALUE! in LibreOffice Calc); make both ranges the same length.",
        "8. FORECAST.LINEAR is the newer name in Excel and Google Sheets. This template uses the older FORECAST, which works in both.",
        "",
        "Months and sales are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-forecast-next-month.xlsx")


def client_tenure():
    wb = Workbook()
    ws = wb.active
    ws.title = "Clients"
    header(ws, ["Client", "First project date", "Years", "Months", "Time with you", "Total months"],
           [26, 16, 8, 9, 14, 13])
    rows = [
        ("Maple Street Bakery", date(2021, 3, 15)),
        ("Harbor Yoga Studio", date(2022, 11, 2)),
        ("Birch Design Co", date(2024, 2, 29)),
        ("Lantern Print Shop", date(2025, 8, 20)),
        ("Cedar Row Realty", date(2023, 10, 16)),
        ("Willow Pet Care", date(2019, 6, 30)),
        ("Granite Fitness", date(2026, 1, 5)),
        ("Juniper Books", date(2020, 12, 31)),
        ("Copper Kettle Cafe", date(2026, 10, 1)),
        ("Pine Hill Dental", date(2018, 4, 9)),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"C{r}"] = (f'=IF(B{r}="","",IF(NOT(ISNUMBER(B{r})),"Check date",'
                       f'IF(B{r}>$I$2,"Check date",DATEDIF(B{r},$I$2,"y"))))')
        ws[f"D{r}"] = f'=IF(ISNUMBER(C{r}),DATEDIF(B{r},$I$2,"ym"),"")'
        ws[f"E{r}"] = f'=IF(ISNUMBER(C{r}),C{r}&" yr "&D{r}&" mo","")'
        ws[f"F{r}"] = f'=IF(ISNUMBER(C{r}),C{r}*12+D{r},"")'
        ws[f"B{r}"].number_format = US_DATE
        for col in "AB":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "CDEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"C2:C{last}", FormulaRule(formula=['C2="Check date"'], fill=ALERT_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 34
    ws.column_dimensions["I"].width = 12
    ws["H1"] = "Settings and results"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["I2"] = "As of date", date(2026, 10, 16)
    ws["H3"], ws["I3"] = "Count clients with at least (years)", 3
    ws["I2"].number_format = US_DATE
    for a in ("I2", "I3"):
        ws[a].fill = INPUT_FILL
    res = [
        ("Clients with a valid date", f"=COUNT(C2:C{last})", "0"),
        ("Clients at or over I3 years", f'=COUNTIF(C2:C{last},">="&I3)', "0"),
        ("Longest time (months)", f'=IF(COUNT(F2:F{last})=0,"",MAX(F2:F{last}))', "0"),
        ("Rows to fix (Check date)", f'=COUNTIF(C2:C{last},"Check date")', "0"),
        ("Check: all clients counted", f'=IF(I4+I7=COUNTA(A2:A{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=4):
        ws[f"H{i}"], ws[f"I{i}"] = label, f
        ws[f"I{i}"].number_format = fmt
        ws[f"I{i}"].fill = FORMULA_FILL
        ws[f"I{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("I8", FormulaRule(formula=['I8="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: How Long Has a Client Been With You",
        "",
        "1. Clients tab: type the client and the first project date in A and B (yellow). Rows 2 to 30 have formulas ready.",
        "2. I2 is the as of date. It is typed, not TODAY(), so the results do not change when you reopen the file. Replace it with =TODAY() if you want them to.",
        "3. C2: =DATEDIF(B2,$I$2,\"y\") counts complete years. D2: =DATEDIF(B2,$I$2,\"ym\") counts the months left over after those years.",
        "4. E2 joins them into text like 5 yr 7 mo, and F2 is the total months (years times 12 plus months).",
        "5. DATEDIF counts complete months only: a start on the 20th is not a full month until the 20th of the next month.",
        "6. DATEDIF is not listed in the Excel function picker, but it works when typed. Google Sheets lists it.",
        "7. DATEDIF returns #NUM! when the start date is after the as of date. This sheet shows Check date instead, and so for a date typed as text. A blank date row stays blank.",
        "8. I3 is a number of years you type; I5 counts the clients at or over it.",
        "",
        "Clients and dates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-client-tenure.xlsx")


def incomplete_rows():
    wb = Workbook()
    ws = wb.active
    ws.title = "Customers"
    header(ws, ["Name", "Email", "Phone", "Street", "ZIP", "Blank fields", "Status", "First missing"],
           [22, 28, 15, 24, 8, 9, 11, 14])
    rows = [
        ("Ava Thompson", "ava.t@example.com", "617-555-0101", "12 Elm St", "02108"),
        ("Ben Ortiz", "", "201-555-0142", "88 Pine Ave", "07030"),
        ("Carla Nguyen", "carla.n@example.com", "", "5 Harbor Rd", "02210"),
        ("Dev Patel", "dev.p@example.com", "973-555-0188", "", "07102"),
        ("Erin Walsh", "erin.w@example.com", "617-555-0133", "41 Cedar Ln", ""),
        ("Frank Lee", "", "", "9 Mill St", "01002"),
        ("Gina Rossi", "gina.r@example.com", "201-555-0177", "230 Oak Dr", "07601"),
        ("", "hal.k@example.com", "973-555-0120", "17 Birch Ct", "07042"),
        ("Iris Chen", "iris.c@example.com", "617-555-0165", "63 Maple St", "02139"),
        ("Jon Baker", "", "", "", ""),
        ("Kim Alvarez", "kim.a@example.com", "201-555-0111", "70 River Rd", "07024"),
        ("Liam Burke", "liam.b@example.com", "", "3 Lake Ave", "01851"),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append([v if v != "" else None for v in rows[r - 2]] if r - 2 < len(rows) else [])
        ws[f"F{r}"] = f'=IF(COUNTA(A{r}:E{r})=0,"",COUNTBLANK(A{r}:E{r}))'
        ws[f"G{r}"] = f'=IF(F{r}="","",IF(F{r}=0,"Complete","Missing"))'
        ws[f"H{r}"] = (f'=IF(OR(F{r}="",F{r}=0),"",IF(A{r}="","Name",IF(B{r}="","Email",'
                       f'IF(C{r}="","Phone",IF(D{r}="","Street","ZIP")))))')
        ws[f"E{r}"].number_format = "@"
        for col in "ABCDE":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "FGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"G2:G{last}", FormulaRule(formula=['G2="Missing"'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"G2:G{last}", FormulaRule(formula=['G2="Complete"'], fill=DONE_FILL))
    ws.column_dimensions["I"].width = 4
    ws.column_dimensions["J"].width = 30
    ws.column_dimensions["K"].width = 12
    ws["J1"] = "Totals"
    ws["J1"].font = Font(bold=True)
    res = [
        ("Rows with data", f"=COUNT(F2:F{last})", "0"),
        ("Complete rows", f'=COUNTIF(G2:G{last},"Complete")', "0"),
        ("Rows missing something", f'=COUNTIF(G2:G{last},"Missing")', "0"),
        ("Rows missing a ZIP", f'=SUMPRODUCT((F2:F{last}<>"")*(E2:E{last}=""))', "0"),
        ("Check: blanks add up", f'=IF(SUM(F2:F{last})=COUNT(F2:F{last})*5-COUNTA(A2:E{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=2):
        ws[f"J{i}"], ws[f"K{i}"] = label, f
        ws[f"K{i}"].number_format = fmt
        ws[f"K{i}"].fill = FORMULA_FILL
        ws[f"K{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("K6", FormulaRule(formula=['K6="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Find Rows With Missing Information",
        "",
        "1. Customers tab: paste or type name, email, phone, street and ZIP in A to E (yellow). Rows 2 to 30 have formulas ready. The ZIP column is formatted as text so 02108 keeps its leading zero.",
        "2. F2: =COUNTBLANK(A2:E2) counts the empty cells in the row. A row with no data at all stays blank here, so unused rows do not show 5.",
        "3. G2 shows Complete when F2 is 0 and Missing otherwise. H2 names the first empty field, checked in the order Name, Email, Phone, Street, ZIP.",
        "4. K3 to K5 count complete rows, rows missing something and rows with no ZIP. K6 checks that the blanks counted per row add up to the cells actually filled.",
        "5. COUNTBLANK counts cells that are really empty. A cell holding only a space is not empty, so it counts as filled. Use the TRIM and clean list template to remove stray spaces first.",
        "6. A cell with a formula that returns \"\" counts as blank in LibreOffice Calc and Excel; K6 shows CHECK in that case because COUNTA treats the cell as filled.",
        "7. This counts empty cells only. It does not check that an email or address is correct.",
        "",
        "Names, emails, phone numbers and addresses are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-incomplete-rows.xlsx")


def overlapping_bookings():
    wb = Workbook()
    ws = wb.active
    ws.title = "Bookings"
    header(ws, ["Booking ID", "Space", "Client", "Start date", "End date", "Conflicts", "Status", "Cross-check"],
           [12, 11, 24, 12, 12, 10, 11, 11])
    rows = [
        ("BK-101", "Booth A", "Maple Street Bakery", date(2026, 10, 3), date(2026, 10, 5)),
        ("BK-102", "Booth A", "Harbor Yoga Studio", date(2026, 10, 5), date(2026, 10, 7)),
        ("BK-103", "Booth A", "Birch Design Co", date(2026, 10, 12), date(2026, 10, 20)),
        ("BK-104", "Booth A", "Lantern Print Shop", date(2026, 10, 14), date(2026, 10, 16)),
        ("BK-105", "Booth B", "Cedar Row Realty", date(2026, 10, 3), date(2026, 10, 5)),
        ("BK-106", "Booth B", "Willow Pet Care", date(2026, 10, 10), date(2026, 10, 14)),
        ("BK-107", "Booth B", "Granite Fitness", date(2026, 10, 13), date(2026, 10, 17)),
        ("BK-108", "Booth B", "Juniper Books", date(2026, 10, 16), date(2026, 10, 19)),
        ("BK-109", "Studio", "Copper Kettle Cafe", date(2026, 10, 8), date(2026, 10, 9)),
        ("BK-110", "Studio", "Pine Hill Dental", date(2026, 10, 8), date(2026, 10, 9)),
        ("BK-111", "Studio", "Maple Street Bakery", date(2026, 10, 21), date(2026, 10, 23)),
        ("BK-112", "Studio", "Harbor Yoga Studio", date(2026, 10, 24), date(2026, 10, 24)),
    ]
    last = 40
    sp, st, en = f"$B$2:$B${last}", f"$D$2:$D${last}", f"$E$2:$E${last}"
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"F{r}"] = (
            f'=IF(A{r}="","",IF(OR(B{r}="",NOT(ISNUMBER(D{r})),NOT(ISNUMBER(E{r}))),"",IF(E{r}<D{r},"",'
            f'IF($K$2="Yes",COUNTIFS({sp},B{r},{st},"<="&E{r},{en},">="&D{r})-1,'
            f'COUNTIFS({sp},B{r},{st},"<"&E{r},{en},">"&D{r})-IF(E{r}>D{r},1,0)))))')
        ws[f"G{r}"] = f'=IF(A{r}="","",IF(F{r}="","Check",IF(F{r}>0,"Conflict","OK")))'
        ws[f"H{r}"] = (
            f'=IF(F{r}="","",IF($K$2="Yes",SUMPRODUCT(({sp}=B{r})*({st}<=E{r})*({en}>=D{r}))-1,'
            f'SUMPRODUCT(({sp}=B{r})*({st}<E{r})*({en}>D{r}))-IF(E{r}>D{r},1,0)))')
        for col in "DE":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "ABCDE":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "FGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1='"Booth A,Booth B,Studio"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"B2:B{last}")
    dv2 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    ws.add_data_validation(dv2)
    dv2.add("K2")
    ws.conditional_formatting.add(f"G2:G{last}", FormulaRule(formula=['OR(G2="Conflict",G2="Check")'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"G2:G{last}", FormulaRule(formula=['G2="OK"'], fill=DONE_FILL))
    ws.column_dimensions["I"].width = 4
    ws.column_dimensions["J"].width = 42
    ws.column_dimensions["K"].width = 10
    ws["J1"] = "Settings and results"
    ws["J1"].font = Font(bold=True)
    ws["J2"], ws["K2"] = "Same-day handoff counts as a conflict (Yes/No)", "Yes"
    ws["K2"].fill = INPUT_FILL
    res = [
        ("Bookings with a conflict", f'=COUNTIF(G2:G{last},"Conflict")', "0"),
        ("Total bookings", f"=COUNTA(A2:A{last})", "0"),
        ("Rows to fix (Check)", f'=COUNTIF(G2:G{last},"Check")', "0"),
        ("Check: both counts agree", f'=IF(SUM(F2:F{last})=SUM(H2:H{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=3):
        ws[f"J{i}"], ws[f"K{i}"] = label, f
        ws[f"K{i}"].number_format = fmt
        ws[f"K{i}"].fill = FORMULA_FILL
        ws[f"K{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("K6", FormulaRule(formula=['K6="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Find Overlapping Dates and Double Bookings",
        "",
        "1. Bookings tab: type the booking, the space, the client and the start and end dates in A to E (yellow). Rows 2 to 40 have formulas ready. Pick the space from the dropdown in B.",
        "2. F2 counts the other bookings in the same space whose dates overlap this one: =COUNTIFS($B$2:$B$40,B2,$D$2:$D$40,\"<=\"&E2,$E$2:$E$40,\">=\"&D2)-1. Two ranges overlap when each one starts on or before the day the other ends. The -1 removes the row itself.",
        "3. G2 shows Conflict when F2 is above 0, OK when it is 0, and Check when a date is missing, typed as text, or the end date is before the start date.",
        "4. K2 switches how a same-day handoff is treated. Yes means an end date equal to the next start date is a conflict. No means it is allowed (the formula then uses < and > instead of <= and >=).",
        "5. H2 counts the same thing with SUMPRODUCT. K6 shows OK when the two counts agree across all rows.",
        "6. A date typed as text (left aligned, no date format) is ignored by COUNTIFS, so the other bookings do not see it. That row itself shows Check. Retype it as a date.",
        "7. A booking with a conflict is flagged on every row involved. Three bookings that overlap at once show 2 on the middle one.",
        "",
        "Bookings and clients are fictional. This checks dates only: it does not know about leases, deposits or contract terms. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-overlapping-bookings.xlsx")


def cash_drawer_count():
    wb = Workbook()
    ws = wb.active
    ws.title = "Count"
    header(ws, ["Denomination", "Value", "Count", "Amount"], [20, 11, 9, 13])
    dens = [
        ("$100 bill", 100, 0), ("$50 bill", 50, 0), ("$20 bill", 20, 6), ("$10 bill", 10, 9),
        ("$5 bill", 5, 14), ("$1 bill", 1, 23), ("Quarter", 0.25, 18), ("Dime", 0.10, 12),
        ("Nickel", 0.05, 9), ("Penny", 0.01, 20),
    ]
    for r, (name, val, cnt) in enumerate(dens, start=2):
        ws.append([name, val, cnt])
        ws[f"D{r}"] = f'=IF(C{r}="","",IF(NOT(ISNUMBER(C{r})),"Check",IF(AND(C{r}>=0,C{r}=INT(C{r})),B{r}*C{r},"Check")))'
        ws[f"B{r}"].number_format = USD
        ws[f"D{r}"].number_format = USD
        ws[f"C{r}"].fill = INPUT_FILL
        ws[f"D{r}"].fill = FORMULA_FILL
    ws["A12"], ws["C12"], ws["D12"] = "Total counted", "=SUM(C2:C11)", "=SUM(D2:D11)"
    ws["D12"].number_format = USD
    for a in ("A12", "C12", "D12"):
        ws[a].font = Font(bold=True)
        ws[a].fill = FORMULA_FILL
    ws.conditional_formatting.add("D2:D11", FormulaRule(formula=['D2="Check"'], fill=ALERT_FILL))
    ws.column_dimensions["E"].width = 30
    ws.column_dimensions["F"].width = 14
    ws["E1"] = "Closing numbers"
    ws["E1"].font = Font(bold=True)
    ws["E2"], ws["F2"] = "Starting float", 150
    ws["E3"], ws["F3"] = "Cash sales for the day", 161.85
    for a in ("F2", "F3"):
        ws[a].fill = INPUT_FILL
        ws[a].number_format = USD
    res = [
        ("Expected cash", "=F2+F3", USD),
        ("Over (+) or short (-)", "=ROUND(D12-F4,2)", USD),
        ("Status", '=IF(COUNTIF(D2:D11,"Check")>0,"Check counts",IF(F5=0,"Balanced",IF(F5>0,"Over","Short")))', "General"),
        ("Check: SUMPRODUCT total", "=SUMPRODUCT(B2:B11,C2:C11)", USD),
        ("Check: both totals agree", '=IF(ROUND(F7-D12,2)=0,"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=4):
        ws[f"E{i}"], ws[f"F{i}"] = label, f
        ws[f"F{i}"].number_format = fmt
        ws[f"F{i}"].fill = FORMULA_FILL
        ws[f"F{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("F6", FormulaRule(formula=['OR(F6="Over",F6="Short",F6="Check counts")'], fill=ALERT_FILL))
    ws.conditional_formatting.add("F6", FormulaRule(formula=['F6="Balanced"'], fill=DONE_FILL))
    ws.conditional_formatting.add("F8", FormulaRule(formula=['F8="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Cash Drawer Count Over or Short",
        "",
        "1. Count tab: type how many of each bill and coin are in the drawer in C (yellow). A blank count is treated as none.",
        "2. D2: =B2*C2 gives the amount for each row. D12 adds them: =SUM(D2:D11). C12 is the number of bills and coins.",
        "3. F2 is the cash you started the day with (the float). F3 is the cash sales for the day from your own records. F4 adds them to give the cash you expect to find.",
        "4. F5: =ROUND(D12-F4,2) is the difference. ROUND is there because decimals like 0.10 are stored in binary: a difference of one cent can otherwise come out as 0.00999999999999091, which is not equal to 0.01.",
        "5. F6 shows Balanced, Over or Short. A count that is negative, a fraction or typed as text shows Check in D, and F6 says Check counts.",
        "6. F7 totals the drawer a second way with SUMPRODUCT. F8 shows OK when it matches D12.",
        "7. To count another drawer, save a copy of the file or clear column C and F3.",
        "",
        "Counts and sales are fictional. This is arithmetic only, not accounting advice. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-cash-drawer-count.xlsx")


def nth_weekday_of_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Dates"
    header(ws, ["Item", "Value"], [34, 18])
    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    ws["A2"], ws["B2"] = "Year", 2026
    ws["A3"], ws["B3"] = "Month (1 to 12)", 11
    ws["A4"], ws["B4"] = "Weekday", "Saturday"
    ws["A5"], ws["B5"] = "Which one (1 to 5)", 2
    for a in ("B2", "B3", "B4", "B5"):
        ws[a].fill = INPUT_FILL
        ws[a].alignment = Alignment(horizontal="right")
    ws["A6"], ws["B6"] = "Weekday number (1 = Sunday)", '=IFERROR(MATCH(B4,$E$2:$E$8,0),"Check weekday")'
    ws["A7"] = "Inputs OK?"
    ws["B7"] = ('=IF(AND(ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B5),ISNUMBER(B6)),'
                'IF(AND(B2>=1900,B2<=9999,B2=INT(B2),B3>=1,B3<=12,B3=INT(B3),B5>=1,B5<=5,B5=INT(B5)),"Yes","No"),"No")')
    first = "DATE(B2,B3,1)"
    nth = f"{first}+MOD(B6-WEEKDAY({first},1),7)+7*(B5-1)"
    ws["A8"] = "Date of the Nth weekday"
    ws["B8"] = f'=IF(B7<>"Yes","Check inputs",IF(MONTH({nth})=B3,{nth},"No such date"))'
    eom = f"EOMONTH({first},0)"
    ws["A9"] = "Date of the last such weekday"
    ws["B9"] = f'=IF(B7<>"Yes","Check inputs",{eom}-MOD(WEEKDAY({eom},1)-B6,7))'
    ws["A10"] = "Check: result falls on that weekday"
    ws["B10"] = ('=IF(AND(ISNUMBER(B8),ISNUMBER(B9)),IF(AND(WEEKDAY(B8,1)=B6,WEEKDAY(B9,1)=B6),"OK","CHECK"),'
                 'IF(ISNUMBER(B9),IF(WEEKDAY(B9,1)=B6,"OK","CHECK"),""))')
    for a in ("B6", "B7", "B8", "B9", "B10"):
        ws[a].fill = FORMULA_FILL
        ws[a].alignment = Alignment(horizontal="right")
    ws["B8"].number_format = US_DATE
    ws["B9"].number_format = US_DATE
    ws.conditional_formatting.add("B7:B9", FormulaRule(formula=['OR(B7="No",B7="Check inputs",B7="No such date")'], fill=ALERT_FILL))
    ws.conditional_formatting.add("B10", FormulaRule(formula=['B10="CHECK"'], fill=ALERT_FILL))
    dv = DataValidation(type="list", formula1="=$E$2:$E$8", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B4")
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 14
    ws["E1"] = "Weekday list"
    ws["E1"].font = Font(bold=True)
    for i, d in enumerate(days, start=2):
        ws[f"E{i}"] = d

    sc = wb.create_sheet("Schedule")
    header(sc, ["Month", "Date", "Weekday"], [16, 16, 14])
    ok = ('AND(ISNUMBER(Dates!$B$2),ISNUMBER(Dates!$B$5),ISNUMBER(Dates!$B$6),Dates!$B$2>=1900,Dates!$B$2<=9999,'
          'Dates!$B$2=INT(Dates!$B$2),Dates!$B$5>=1,Dates!$B$5<=5,Dates!$B$5=INT(Dates!$B$5))')
    for r in range(2, 14):
        sc[f"A{r}"] = f"=IF({ok},DATE(Dates!$B$2,1,1),\"\")" if r == 2 else f'=IF(A{r-1}="","",EDATE(A{r-1},1))'
        x = f"A{r}+MOD(Dates!$B$6-WEEKDAY(A{r},1),7)+7*(Dates!$B$5-1)"
        sc[f"B{r}"] = f'=IF(A{r}="","",IF(MONTH({x})=MONTH(A{r}),{x},"No such date"))'
        sc[f"C{r}"] = f'=IF(ISNUMBER(B{r}),Dates!$B$4,"")'
        sc[f"A{r}"].number_format = "mmmm yyyy"
        sc[f"B{r}"].number_format = US_DATE
        for col in "ABC":
            sc[f"{col}{r}"].fill = FORMULA_FILL
        sc[f"B{r}"].alignment = Alignment(horizontal="right")
    sc.conditional_formatting.add("B2:B13", FormulaRule(formula=['B2="No such date"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Date of the Nth Weekday of a Month",
        "",
        "1. Dates tab: type the year, the month number, pick the weekday from the dropdown and type which one you want (1 to 5) in B2 to B5 (yellow).",
        "2. B6 turns the weekday name into a number with MATCH (1 = Sunday to 7 = Saturday, the WEEKDAY default).",
        "3. B8: =DATE(B2,B3,1)+MOD(B6-WEEKDAY(DATE(B2,B3,1),1),7)+7*(B5-1). The first part finds the first wanted weekday on or after the 1st, the last part adds 7 days for each extra week.",
        "4. B8 is wrapped in a MONTH check. If the result lands in the next month (for example a 5th Saturday that does not exist), it shows No such date.",
        "5. B9 is the last such weekday of the month: =EOMONTH(DATE(B2,B3,1),0)-MOD(WEEKDAY(EOMONTH(DATE(B2,B3,1),0),1)-B6,7). It steps back from the last day of the month.",
        "6. A month typed as 13 would roll into the next year inside DATE, so B7 checks the inputs first and the results show Check inputs. A year typed as text does the same.",
        "7. Schedule tab: the same weekday and number for all 12 months of the year in B2. Change the yellow cells on Dates and it follows.",
        "",
        "This is date math only. It does not know about holidays or market rules. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-nth-weekday-of-month.xlsx")


def format_phone_numbers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Phones"
    header(ws, ["Phone as typed", "Digits only", "Without leading 1", "Formatted"], [24, 16, 18, 18])
    raw = ["(212) 555-0101", "212.555.0102", "+1 212-555-0103", 2125550104, "212 555 0105",
           "1-212-555-0106", "(415)555-0107", "415.555.0108", "212-555-010",
           "212-555-0111 x12", "21255501130", "call after 5"]
    last = 40
    for r in range(2, last + 1):
        ws.append([raw[r - 2]] if r - 2 < len(raw) else [])
        s = f'A{r}'
        for ch in (" ", "-", "(", ")", ".", "+"):
            s = f'SUBSTITUTE({s},"{ch}","")'
        ws[f"B{r}"] = f'=IF(A{r}="","",{s})'
        ws[f"C{r}"] = f'=IF(B{r}="","",IF(AND(LEN(B{r})=11,LEFT(B{r},1)="1"),MID(B{r},2,10),B{r}))'
        digits = f'SUMPRODUCT(--ISNUMBER(FIND(MID(C{r},{{1,2,3,4,5,6,7,8,9,10}},1),"0123456789")))=10'
        ws[f"D{r}"] = (f'=IF(A{r}="","",IF(AND(LEN(C{r})=10,{digits}),'
                       f'"("&LEFT(C{r},3)&") "&MID(C{r},4,3)&"-"&RIGHT(C{r},4),"Check"))')
        for col in "ABC":
            ws[f"{col}{r}"].number_format = "@"
        ws[f"A{r}"].fill = INPUT_FILL
        for col in "BCD":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"D2:D{last}", FormulaRule(formula=['D2="Check"'], fill=ALERT_FILL))
    ws.column_dimensions["E"].width = 4
    ws.column_dimensions["F"].width = 30
    ws.column_dimensions["G"].width = 10
    ws["F1"] = "Results"
    ws["F1"].font = Font(bold=True)
    res = [
        ("Rows with a phone entry", f"=COUNTA(A2:A{last})", "0"),
        ("Formatted", f"=SUMPRODUCT(--(LEN(D2:D{last})=14))", "0"),
        ("Rows to check", f'=COUNTIF(D2:D{last},"Check")', "0"),
        ("Check: formatted + check = rows", f'=IF(G3+G4=G2,"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=2):
        ws[f"F{i}"], ws[f"G{i}"] = label, f
        ws[f"G{i}"].number_format = fmt
        ws[f"G{i}"].fill = FORMULA_FILL
        ws[f"G{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("G5", FormulaRule(formula=['G5="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Format Phone Numbers as (555) 123-4567",
        "",
        "1. Phones tab: paste or type the phone numbers in A (yellow). Rows 2 to 40 have formulas ready. Columns A to C are formatted as text so a leading zero is kept.",
        "2. B2 strips the spaces, hyphens, parentheses, dots and plus signs with nested SUBSTITUTE: =SUBSTITUTE(SUBSTITUTE(...(A2,\" \",\"\"),\"-\",\"\"),...).",
        "3. C2: =IF(AND(LEN(B2)=11,LEFT(B2,1)=\"1\"),MID(B2,2,10),B2) drops a leading 1 (the US country code) from an 11 digit number.",
        "4. D2 rebuilds a 10 digit number as (555) 123-4567 with LEFT, MID and RIGHT. Anything else shows Check: too few or too many digits, letters, or an extension such as x12.",
        "5. G2 to G4 count the rows with an entry, the formatted rows and the Check rows. G5 shows OK when formatted plus Check equals the rows with an entry.",
        "6. Copy D and use Paste values only when you need the text without formulas, for example before importing contacts or printing labels.",
        "7. This only fixes the format. It does not check that a number really exists or can be called. Extensions and numbers from other countries are left as Check.",
        "",
        "Phone numbers are fictional (555-01xx style). Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-format-phone-numbers.xlsx")


def scale_recipe_batch():
    wb = Workbook()
    ws = wb.active
    ws.title = "Recipe"
    for col, w in zip("ABCDE", [30, 14, 18, 14, 18]):
        ws.column_dimensions[col].width = w
    ws["A1"], ws["C1"] = "Base batch yield (jars)", 12
    ws["A2"], ws["C2"] = "Target yield (jars)", 30
    ws["A3"], ws["C3"] = "Scale factor", '=IF(AND(ISNUMBER(C1),C1>0),IF(ISNUMBER(C2),C2/C1,0),"")'
    ws["A4"], ws["C4"] = "Check", ('=IF(C3="","Check base yield",IF(C3=0,"Check target yield",'
                                  'IF(SUMPRODUCT((A6:A30<>"")*(1-ISNUMBER(B6:B30)))>0,"Check quantities","OK")))')
    for a in ("C1", "C2"):
        ws[a].fill = INPUT_FILL
    for a in ("C3", "C4"):
        ws[a].fill = FORMULA_FILL
    ws["C3"].number_format = '0.00"x"'
    for a in ("C1", "C2", "C3", "C4"):
        ws[a].alignment = Alignment(horizontal="right")
    for c in ("A1", "A2", "A3", "A4"):
        ws[c].font = Font(bold=True)
    names = ["Ingredient", "Base quantity", "Unit", "Scaled quantity", "Buy whole units"]
    for col, n in zip("ABCDE", names):
        c = ws[f"{col}5"]
        c.value, c.fill, c.font = n, HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    items = [
        ("Strawberries", 64, "oz"), ("Granulated sugar", 6, "cup"), ("Lemon juice", 0.25, "cup"),
        ("Powdered pectin", 4, "tsp"), ("Lemons", 3, "each"), ("Vanilla extract", 1.5, "tsp"),
        ("Salt", 0.25, "tsp"), ("Jar lids", 12, "each"),
    ]
    last = 30
    for r in range(6, last + 1):
        if r - 6 < len(items):
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = items[r - 6]
        ws[f"D{r}"] = f'=IF(OR(C$3="",NOT(ISNUMBER(B{r}))),"",ROUND(B{r}*$C$3,2))'
        ws[f"E{r}"] = (f'=IF(D{r}="","",IF(C{r}="each",ROUNDUP(ROUND(B{r}*$C$3,6),0),ROUND(B{r}*$C$3,2)))')
        ws[f"B{r}"].number_format = "0.##"
        ws[f"D{r}"].number_format = "0.00"
        ws[f"E{r}"].number_format = "0.00"
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1='"oz,cup,tsp,each"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C6:C{last}")
    ws.conditional_formatting.add("C4", FormulaRule(formula=['C4<>"OK"'], fill=ALERT_FILL))
    ws.freeze_panes = "A6"

    notes_sheet(wb, [
        "Tidy Tabs: Scale a Recipe or Batch Up or Down",
        "",
        "1. Recipe tab: type how many jars (or loaves, bars, candles) the recipe makes in C1 and how many you want in C2 (yellow).",
        "2. C3: =C2/C1 is the scale factor, shown as a multiplier. 30 jars from a 12 jar recipe is 2.50x.",
        "3. Type each ingredient, its base quantity and its unit in A to C. Write fractions as decimals: 1/3 cup is 0.33 and 1/4 cup is 0.25.",
        "4. D6: =ROUND(B6*$C$3,2). The $ signs lock the factor to C3, so the formula can be copied down and still points at the same cell. B6 has no $ and moves to B7, B8 and so on.",
        "5. E6 is for things you cannot buy in parts. When the unit is each, it rounds up to a whole number with ROUNDUP (3 lemons times 2.5 is 7.5, so 8). Other units are rounded to 2 decimals like D.",
        "6. If C2 is blank or 0 the factor is 0 and every quantity becomes 0: C4 says Check target yield. If C1 is 0 or blank, C3 stays empty and C4 says Check base yield. A quantity typed as text is left out and C4 says Check quantities.",
        "7. Rows 6 to 30 have formulas ready; copy the formula cells down for more.",
        "",
        "The jam recipe and quantities are fictional. Cooking times, baking chemistry, set points and food safety can change at larger sizes, and this sheet does not cover them. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Recipe"].print_title_rows = "5:5"
    wb.save(OUT / "tidy-tabs-scale-recipe-batch.xlsx")


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
    quantity_price_tiers()
    weighted_average_cost()
    last_order_date_maxifs()
    count_orders_by_month()
    average_order_value_by_channel()
    subtotal_filtered_sales()
    rolling_average_sales()
    prorate_partial_month()
    two_way_rate_card()
    in_cell_bars()
    quarterly_sales_totals()
    missing_invoice_numbers()
    sumif_contains_text()
    invoice_installments()
    marketplace_net_payout()
    invoice_due_dates()
    sales_by_weekday()
    repeat_customers()
    supplier_price_compare()
    split_sku_code()
    listing_length_check()
    ounces_to_pounds_ounces()
    last_number_in_column()
    time_zone_converter()
    count_weekdays_in_month()
    days_to_get_paid()
    early_payment_discount()
    sell_through_rate()
    boxes_needed()
    median_order_value()
    week_ending_date()
    free_shipping_threshold()
    forecast_next_month()
    client_tenure()
    incomplete_rows()
    overlapping_bookings()
    cash_drawer_count()
    nth_weekday_of_month()
    format_phone_numbers()
    scale_recipe_batch()
    print("\n".join(sorted(p.name for p in OUT.iterdir())))
