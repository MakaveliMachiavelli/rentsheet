#!/usr/bin/env python3
"""RentSheet generator — landlord rental income/expense tracker workbook.
Usage: python3 build_sheet.py full|demo"""
import sys, datetime, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

MODE = sys.argv[1] if len(sys.argv) > 1 else "full"
DIST = os.path.join(os.path.dirname(__file__), "dist")
os.makedirs(DIST, exist_ok=True)

NAVY = "1D4ED8"; DARK = "101828"; GREY = "667085"
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
HEAD_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=16, color=DARK)
THIN = Border(*[Side(style="thin", color="E4E7EC")]*4)
MONEY = '#,##0.00'
DATEF = 'yyyy-mm-dd'

wb = Workbook()

def style_header(ws, row, cols):
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL; cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

# ================= README =================
ws = wb.active; ws.title = "README"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 100
rows = [
    ("title", "RENTSHEET — Rental Property Income & Expense Tracker"),
    ("sub", "For landlords with 1–20 units. Works in Excel and Google Sheets."),
    ("blank", ""),
    ("h", "HOW TO USE (3 steps):"),
    ("li", "1. List your units in PROPERTIES (rent, tenant, lease dates — status is automatic)."),
    ("li", "2. Log rent payments in RENT LOG and costs in EXPENSES (property dropdowns auto-fill)."),
    ("li", "3. DASHBOARD shows collection rate, unpaid rent, expenses per unit, and net income."),
    ("blank", ""),
    ("h", "SHEETS:"),
    ("li", "• DASHBOARD — collected vs expected, collection rate, late payments, net operating income, per-unit summary."),
    ("li", "• PROPERTIES — units with tenant, rent, lease start/end; EXPIRED lease alert automatically."),
    ("li", "• RENT LOG — one row per rent payment; auto balance (due − paid) and PAID/PARTIAL/LATE status."),
    ("li", "• EXPENSES — repairs, taxes, insurance, management — rolled into net income per unit."),
    ("blank", ""),
    ("h", "NOTES:"),
    ("li", "• Don't edit grey formula cells; type only in white cells."),
    ("li", "• 'LATE' = rent month's row still has a balance and the month has ended (based on your computer's date)."),
    ("li", "• Use one RENT LOG row per unit per month, even for partial payments."),
    ("blank", ""),
    ("sub2", f"RentSheet v1.0 · {datetime.date.today().isoformat()}"),
]
r = 2
for kind, text in rows:
    cell = ws.cell(row=r, column=2, value=text)
    if kind == "title": cell.font = TITLE_FONT
    elif kind == "sub": cell.font = Font(size=12, color=GREY)
    elif kind == "sub2": cell.font = Font(size=9, color=GREY)
    elif kind == "h": cell.font = Font(bold=True, size=12, color=DARK)
    elif kind == "li": cell.font = Font(size=11)
    r += 1
if MODE == "demo":
    ws.cell(row=r+1, column=2, value="*** DEMO VERSION — limited rows. Full version = 200 rows + free updates. ***").font = Font(bold=True, color="DC2626")

N = 5 if MODE == "demo" else 200
last = N + 1

# ================= PROPERTIES =================
ws = wb.create_sheet("PROPERTIES")
headers = ["Unit / Property", "Tenant", "Monthly Rent (₱)", "Lease Start", "Lease End", "Lease Status", "Notes"]
for i, h in enumerate(headers, 1): ws.cell(row=1, column=i, value=h)
style_header(ws, 1, range(1, len(headers)+1))
for i, w in enumerate([24, 20, 15, 12, 12, 13, 28], 1): ws.column_dimensions[get_column_letter(i)].width = w
today = datetime.date.today()
month1 = today.replace(day=1)
samples = [
    ("Unit 2A — Maple St.", "Santos family", 12000, month1, month1 + datetime.timedelta(days=180), "Corner unit"),
    ("Unit 5B — Maple St.", "R. Cruz", 9500, month1, month1 - datetime.timedelta(days=3), "Lease renewal talk started"),
]
for j, s in enumerate(samples, start=2):
    for i, v in enumerate(s, 1):
        if i not in (6,): ws.cell(row=j, column=i, value=v)
for r in range(2, last+1):
    ws.cell(row=r, column=6, value=f'=IF(A{r}="","",IF(E{r}="","—",IF(E{r}<TODAY(),"⚠️ EXPIRED","ACTIVE")))')
    for c in range(1, 8):
        cell = ws.cell(row=r, column=c); cell.border = THIN
        if c == 3: cell.number_format = MONEY
        if c in (4, 5): cell.number_format = DATEF
ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:G{last}"
ws.conditional_formatting.add(f"F2:F{last}",
    FormulaRule(formula=['$F2="⚠️ EXPIRED"'], fill=PatternFill("solid", fgColor="FEE2E2"), font=Font(color="DC2626", bold=True)))

# ================= RENT LOG =================
ws = wb.create_sheet("RENT LOG")
headers = ["Rent Month", "Property", "Rent Due (₱)", "Amount Paid (₱)", "Date Paid", "Balance (₱)", "Status"]
for i, h in enumerate(headers, 1): ws.cell(row=1, column=i, value=h)
style_header(ws, 1, range(1, len(headers)+1))
for i, w in enumerate([12, 24, 14, 15, 12, 13, 12], 1): ws.column_dimensions[get_column_letter(i)].width = w
month_start = month1
prev_month = (month1 - datetime.timedelta(days=1)).replace(day=1)
samples = [
    (prev_month, "Unit 2A — Maple St.", 12000, 12000, month1),
    (month_start, "Unit 2A — Maple St.", 12000, 6000, None),
    (prev_month, "Unit 5B — Maple St.", 9500, 9500, None),
]
for j, s in enumerate(samples, start=2):
    for i, v in enumerate(s, 1):
        if i not in (6, 7) and v is not None: ws.cell(row=j, column=i, value=v)
for r in range(2, last+1):
    ws.cell(row=r, column=6, value=f'=IF(B{r}="","",C{r}-D{r})')
    ws.cell(row=r, column=7, value=f'=IF(B{r}="","",IF(F{r}<=0,"✅ PAID",IF(D{r}>0,"🟡 PARTIAL","🔴 LATE")))')
    for c in range(1, 8):
        cell = ws.cell(row=r, column=c); cell.border = THIN
        if c in (3, 4, 6): cell.number_format = MONEY
        if c in (1, 5): cell.number_format = DATEF
ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:G{last}"
dv = DataValidation(type="list", formula1=f"=PROPERTIES!$A$2:$A${last}", allow_blank=True, showDropDown=False)
ws.add_data_validation(dv); dv.add(f"B2:B{last}")
ws.conditional_formatting.add(f"G2:G{last}",
    FormulaRule(formula=['$G2="🔴 LATE"'], fill=PatternFill("solid", fgColor="FEE2E2"), font=Font(color="DC2626", bold=True)))
ws.conditional_formatting.add(f"G2:G{last}",
    FormulaRule(formula=['$G2="🟡 PARTIAL"'], fill=PatternFill("solid", fgColor="FEF7E6"), font=Font(color="B45309", bold=True)))

# ================= EXPENSES =================
ws = wb.create_sheet("EXPENSES")
headers = ["Date", "Property", "Category", "Description", "Amount (₱)"]
for i, h in enumerate(headers, 1): ws.cell(row=1, column=i, value=h)
style_header(ws, 1, range(1, len(headers)+1))
for i, w in enumerate([12, 24, 16, 28, 13], 1): ws.column_dimensions[get_column_letter(i)].width = w
dv1 = DataValidation(type="list", formula1=f"=PROPERTIES!$A$2:$A${last}", allow_blank=True, showDropDown=False)
dv2 = DataValidation(type="list", formula1='"Repairs,Maintenance,Property Tax,Insurance,Management Fee,Utilities,Mortgage,Other"', allow_blank=True, showDropDown=False)
ws.add_data_validation(dv1); dv1.add(f"B2:B{last}")
ws.add_data_validation(dv2); dv2.add(f"C2:C{last}")
for j, s in enumerate([(today.replace(day=4), "Unit 2A — Maple St.", "Repairs", "Leaky faucet", 850)], start=2):
    for i, v in enumerate(s, 1): ws.cell(row=j, column=i, value=v)
for r in range(2, last+1):
    for c in range(1, 6):
        cell = ws.cell(row=r, column=c); cell.border = THIN
        if c == 5: cell.number_format = MONEY
        if c == 1: cell.number_format = DATEF
ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:E{last}"

# ================= DASHBOARD =================
ws = wb.create_sheet("DASHBOARD", 0)
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
for col, w in zip("BCDEFG", [30, 18, 4, 30, 18, 4]): ws.column_dimensions[col].width = w
ws.cell(row=2, column=2, value="🏠 RENTAL DASHBOARD").font = TITLE_FONT
ws.cell(row=3, column=2, value="Auto-computed — do not edit").font = Font(color=GREY, size=10)

def card(ws, row, col, label, formula, fmt='#,##0.00', fill="F5F8FF"):
    ws.cell(row=row, column=col, value=label).font = Font(bold=True, size=11, color=GREY)
    c = ws.cell(row=row+1, column=col, value=formula)
    c.font = Font(bold=True, size=18, color=DARK); c.number_format = fmt
    for cc in (col, col+1):
        for rr in (row, row+1):
            ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=fill)

card(ws, 5, 2, "RENT COLLECTED (all time)", f"=SUM('RENT LOG'!D2:D{last})")
card(ws, 5, 5, "RENT EXPECTED (all time)", f"=SUM('RENT LOG'!C2:C{last})")
card(ws, 8, 2, "UNPAID BALANCE", f"=SUM('RENT LOG'!F2:F{last})", fill="FEF7E6")
card(ws, 8, 5, "COLLECTION RATE", f"=IF(SUM('RENT LOG'!C2:C{last})=0,0,SUM('RENT LOG'!D2:D{last})/SUM('RENT LOG'!C2:C{last}))", fmt='0.0%', fill="E9F9F1")
card(ws, 11, 2, "TOTAL EXPENSES", f"=SUM(EXPENSES!E2:E{last})")
card(ws, 11, 5, "NET INCOME (rent − expenses)", f"=SUM('RENT LOG'!D2:D{last})-SUM(EXPENSES!E2:E{last})", fill="E9F9F1")
card(ws, 14, 2, "UNITS WITH LATE/PARTIAL RENT", f'=COUNTIF(\'RENT LOG\'!G2:G{last},"🔴 LATE")+COUNTIF(\'RENT LOG\'!G2:G{last},"🟡 PARTIAL")', fmt='0', fill="FEE2E2")
card(ws, 14, 5, "EXPIRED LEASES", f'=COUNTIF(PROPERTIES!F2:F{last},"⚠️ EXPIRED")', fmt='0', fill="FEE2E2")

ws.cell(row=17, column=2, value="PER-UNIT SUMMARY").font = Font(bold=True, size=11)
ws.cell(row=18, column=2, value="Unit").font = Font(bold=True, size=9, color=GREY)
ws.cell(row=18, column=3, value="Collected (₱)").font = Font(bold=True, size=9, color=GREY)
ws.cell(row=18, column=4, value="Expenses (₱)").font = Font(bold=True, size=9, color=GREY)
ws.cell(row=18, column=5, value="Rent due").font = Font(bold=True, size=9, color=GREY)
for r in range(2, min(12, last+1)):
    rr = 17 + r - 1
    ws.cell(row=rr+1, column=2, value=f'=IF(PROPERTIES!A{r}="","",PROPERTIES!A{r})').font = Font(size=10)
    c1 = ws.cell(row=rr+1, column=3, value=f"=IF(PROPERTIES!A{r}=\"\",\"\",SUMIF('RENT LOG'!B:B,PROPERTIES!A{r},'RENT LOG'!D:D))")
    c2 = ws.cell(row=rr+1, column=4, value=f"=IF(PROPERTIES!A{r}=\"\",\"\",SUMIF(EXPENSES!B:B,PROPERTIES!A{r},EXPENSES!E:E))")
    c3 = ws.cell(row=rr+1, column=5, value=f"=IF(PROPERTIES!A{r}=\"\",\"\",SUMIF('RENT LOG'!B:B,PROPERTIES!A{r},'RENT LOG'!C:C))")
    for c in (c1, c2, c3): c.number_format = MONEY; c.font = Font(size=10)

if MODE == "demo":
    ws.cell(row=30, column=2, value="DEMO VERSION — limited rows. Full version: 200 rows per sheet.").font = Font(bold=True, color="DC2626")

out = os.path.join(DIST, "RentSheet.xlsx" if MODE == "full" else "RentSheet-Demo.xlsx")
wb.save(out)
print("saved", out, os.path.getsize(out), "bytes")
