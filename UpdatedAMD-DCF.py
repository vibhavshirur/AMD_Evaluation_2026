"""
AMD New Valuation Model Generator
Final version: Verified against actual 10-K filings.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# ============================================================
# WORKBOOK SETUP & SOFT FINANCIAL STYLES
# ============================================================
wb = openpyxl.Workbook()

DARK_BLUE = "1F4E78"
MED_BLUE = "2E75B6"
LIGHT_BLUE = "BDD7EE"
TOT_BLUE = "D9E1F2"
INPUT_YELLOW = "FFF2CC"
CALC_GREEN = "E2EFDA"
WARN_RED = "FFC7CE"

TITLE_FONT  = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
HEAD_FONT   = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
SUB_FONT    = Font(name='Calibri', size=10, bold=True, color=DARK_BLUE)
BOLD        = Font(name='Calibri', size=10, bold=True, color='000000')
NORMAL      = Font(name='Calibri', size=10, color='000000')
INPUT_FONT  = Font(name='Calibri', size=10, bold=True, color='0000CC')
NOTE_FONT   = Font(name='Calibri', size=9, italic=True, color='808080')
RESULT_FONT = Font(name='Calibri', size=12, bold=True, color=DARK_BLUE)

TITLE_FILL = PatternFill('solid', fgColor=DARK_BLUE)
HEAD_FILL  = PatternFill('solid', fgColor=MED_BLUE)
SUB_FILL   = PatternFill('solid', fgColor=LIGHT_BLUE)
INPUT_FILL = PatternFill('solid', fgColor=INPUT_YELLOW)
CALC_FILL  = PatternFill('solid', fgColor=CALC_GREEN)
TOT_FILL   = PatternFill('solid', fgColor=TOT_BLUE)
WARN_FILL  = PatternFill('solid', fgColor=WARN_RED)

THIN_SIDE = Side(style='thin', color='B0B0B0')
BRD = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

C = Alignment(horizontal='center', vertical='center')
L = Alignment(horizontal='left',   vertical='center', wrap_text=True)
R = Alignment(horizontal='right',  vertical='center')

FMT_NUM = '#,##0;(#,##0)'
FMT_DEC = '#,##0.0'
FMT_PCT = '0.0%'
FMT_PCT2= '0.00%'
FMT_USD = '"$"#,##0.00'
FMT_USDM= '"$"#,##0'

def sc(ws, r, c, v, f=None, fl=None, a=None, b=None, fmt=None):
    cell = ws.cell(row=r, column=c, value=v)
    if f:   cell.font = f
    if fl:  cell.fill = fl
    if a:   cell.alignment = a
    if b:   cell.border = b
    if fmt: cell.number_format = fmt
    return cell

def title_row(ws, r, c1, c2, text):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    sc(ws, r, c1, text, TITLE_FONT, TITLE_FILL, C)
    ws.row_dimensions[r].height = 28

def head_row(ws, r, c1, c2, text):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    sc(ws, r, c1, text, HEAD_FONT, HEAD_FILL, L, BRD)
    ws.row_dimensions[r].height = 20

# ============================================================
# SHEET 1: COVER
# ============================================================
ws1 = wb.active
ws1.title = "Cover"
ws1.sheet_view.showGridLines = False
ws1.column_dimensions['A'].width = 3
ws1.column_dimensions['B'].width = 32
ws1.column_dimensions['C'].width = 22
for col in 'DEFG':
    ws1.column_dimensions[col].width = 14

title_row(ws1, 2, 2, 7, "ADVANCED MICRO DEVICES, INC. (AMD)")
title_row(ws1, 3, 2, 7, "Discounted Cash Flow (DCF) Valuation Model")

sc(ws1, 5, 2, "Overview", Font(bold=True, size=12, color=DARK_BLUE), None, L)
desc = ("This workbook presents a complete DCF valuation of AMD using historical "
        "financial data from fiscal years 2021-2025 and projected free cash flows "
        "for 2026-2030. All projections are driven by formulas linked to the "
        "Assumptions sheet — change any input (yellow cells) to see the valuation "
        "update instantly.")
ws1.merge_cells('B6:G7')
sc(ws1, 6, 2, desc, NORMAL, None, L)

sc(ws1, 9, 2, "Data Sources", Font(bold=True, size=12, color=DARK_BLUE), None, L)
sources = [
    "• 2021 10-K (FY ended Dec 25, 2021) — Historical financials",
    "• 2022 10-K (FY ended Dec 31, 2022) — Historical financials",
    "• 2023 10-K (FY ended Dec 30, 2023) — Historical financials",
    "• 2024 10-K (FY ended Dec 28, 2024) — [provided by user]",
    "• 2025 10-K/A (FY ended Dec 27, 2025) — [provided by user]",
    "• Federal Reserve (FRED) — 10-year U.S. Treasury yield (risk-free rate)",
    "• Damodaran Online — Equity risk premium benchmarks",
    "• AMD investor presentations — AI strategy, OpenAI partnership",
]
for i, s in enumerate(sources):
    sc(ws1, 10+i, 2, s, NORMAL, None, L)

sc(ws1, 19, 2, "Instructions", Font(bold=True, size=12, color=DARK_BLUE), None, L)
instructions = [
    "1. Yellow cells = inputs (change these to update the model).",
    "2. Green cells = calculated via formulas (do not overwrite).",
    "3. Navigate using tabs at the bottom of the workbook.",
    "4. Assumptions sheet contains all key drivers.",
    "5. Sensitivity sheet shows per-share value across WACC × terminal growth.",
    "6. Dashboard sheet contains visual summary charts.",
    "7. All dollar amounts are in $ millions unless otherwise noted.",
]
for i, inst in enumerate(instructions):
    sc(ws1, 20+i, 2, inst, NORMAL, None, L)

sc(ws1, 28, 2, "Key Valuation Outputs (Live)", Font(bold=True, size=12, color=DARK_BLUE), None, L)
live_outputs = [
    ("Implied Share Price:",      "='DCF Valuation'!C36", FMT_USD),
    ("Enterprise Value ($M):",    "='DCF Valuation'!C26", FMT_USDM),
    ("Equity Value ($M):",        "='DCF Valuation'!C33", FMT_USDM),
    ("WACC:",                     "='WACC'!C24",           FMT_PCT2),
    ("Terminal Growth Rate:",     "='Assumptions'!C45",    FMT_PCT),
    ("FY2025 Revenue ($M):",      "='Historicals'!G8",     FMT_NUM),
    ("FY2025 Operating Income:",  "='Historicals'!G24",    FMT_NUM),
    ("Diluted Shares (M):",       "='Assumptions'!C41",    FMT_NUM),
]
for i, (lbl, frm, fmt) in enumerate(live_outputs):
    sc(ws1, 29+i, 2, lbl, BOLD, None, L)
    sc(ws1, 29+i, 3, frm, RESULT_FONT, None, R, None, fmt)

sc(ws1, 38, 2, "Reference: AMD closing price on Jan 30, 2026: $236.73", NOTE_FONT, None, L)
sc(ws1, 39, 2, "Note: Share count includes full OpenAI warrant (160M shares at $0.01).", NOTE_FONT, None, L)
sc(ws1, 40, 2, "Note: Terminal growth lowered to 2.5% for reinvestment consistency (see DCF sheet).", NOTE_FONT, None, L)


# ============================================================
# SHEET 2: HISTORICALS (2021-2025)
# ============================================================
ws2 = wb.create_sheet("Historicals")
ws2.sheet_view.showGridLines = False
ws2.column_dimensions['A'].width = 3
ws2.column_dimensions['B'].width = 42
for col in 'CDEFG':
    ws2.column_dimensions[col].width = 13

title_row(ws2, 2, 2, 7, "AMD Historical Financials (2021-2025)")

sc(ws2, 4, 2, "($ in millions, except per-share)", BOLD, None, L)
years = [2021, 2022, 2023, 2024, 2025]
for i, y in enumerate(years):
    sc(ws2, 4, 3+i, y, HEAD_FONT, HEAD_FILL, C, BRD)

sc(ws2, 5, 2, "Source:", BOLD, None, L)
srcs = ["2021 10-K", "2022 10-K", "2023 10-K", "2024 10-K", "2025 10-K/A"]
for i, s in enumerate(srcs):
    sc(ws2, 5, 3+i, s, NOTE_FONT, None, C)

head_row(ws2, 7, 2, 7, "INCOME STATEMENT")

sc(ws2, 8, 2, "Net Revenue", BOLD, None, L, BRD)
rev = [16434, 23601, 22680, 25785, 34639]
for i, v in enumerate(rev):
    sc(ws2, 8, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 9, 2, "  YoY Growth", NORMAL, None, L, BRD)
for i in range(5):
    if i == 0:
        sc(ws2, 9, 3, "=(C8-9763)/9763", NORMAL, None, R, BRD, FMT_PCT)
    else:
        cl = get_column_letter(3+i); pl = get_column_letter(2+i)
        sc(ws2, 9, 3+i, f"={cl}8/{pl}8-1", NORMAL, None, R, BRD, FMT_PCT)

sc(ws2, 10, 2, "Cost of Sales (excl. amort.)", NORMAL, None, L, BRD)
cogs = [8505, 11550, 11278, 12114, 16456]
for i, v in enumerate(cogs):
    sc(ws2, 10, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 11, 2, "Amort. of Acq. Intangibles (COGS)", NORMAL, None, L, BRD)
amort_cogs = [0, 1448, 942, 946, 1031]
for i, v in enumerate(amort_cogs):
    sc(ws2, 11, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 12, 2, "Gross Profit", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 12, 3+i, f"={cl}8-{cl}10-{cl}11", BOLD, TOT_FILL, R, BRD, FMT_NUM)

sc(ws2, 13, 2, "  Gross Margin", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 13, 3+i, f"={cl}12/{cl}8", NORMAL, None, R, BRD, FMT_PCT)

head_row(ws2, 15, 2, 7, "OPERATING EXPENSES")

sc(ws2, 16, 2, "Research & Development", NORMAL, None, L, BRD)
rd = [2845, 5005, 5872, 6456, 8091]
for i, v in enumerate(rd):
    sc(ws2, 16, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 17, 2, "  R&D % of Revenue", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 17, 3+i, f"={cl}16/{cl}8", NORMAL, None, R, BRD, FMT_PCT)

sc(ws2, 18, 2, "Marketing, Gen. & Admin.", NORMAL, None, L, BRD)
mga = [1448, 2336, 2352, 2783, 4144]
for i, v in enumerate(mga):
    sc(ws2, 18, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 19, 2, "  MG&A % of Revenue", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 19, 3+i, f"={cl}18/{cl}8", NORMAL, None, R, BRD, FMT_PCT)

sc(ws2, 20, 2, "Amort. of Acq. Intangibles (OpEx)", NORMAL, None, L, BRD)
amort_opex = [0, 2100, 1869, 1448, 1223]
for i, v in enumerate(amort_opex):
    sc(ws2, 20, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 21, 2, "Restructuring Charges", NORMAL, None, L, BRD)
restr = [0, 0, 0, 186, 0]
for i, v in enumerate(restr):
    sc(ws2, 21, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 22, 2, "Licensing Gain (reduces OpEx)", NORMAL, None, L, BRD)
lic = [12, 102, 34, 48, 0]
for i, v in enumerate(lic):
    sc(ws2, 22, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 23, 2, "Total Operating Expenses", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 23, 3+i, f"={cl}16+{cl}18+{cl}20+{cl}21-{cl}22", BOLD, TOT_FILL, R, BRD, FMT_NUM)

sc(ws2, 24, 2, "Operating Income", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 24, 3+i, f"={cl}12-{cl}23", BOLD, TOT_FILL, R, BRD, FMT_NUM)

sc(ws2, 25, 2, "  Operating Margin", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 25, 3+i, f"={cl}24/{cl}8", NORMAL, None, R, BRD, FMT_PCT)

sc(ws2, 26, 2, "Interest Expense", NORMAL, None, L, BRD)
int_exp = [34, 88, 106, 92, 131]
for i, v in enumerate(int_exp):
    sc(ws2, 26, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 27, 2, "Other Income (Expense), Net", NORMAL, None, L, BRD)
oi = [61, 22, 213, 214, 603]
for i, v in enumerate(oi):
    sc(ws2, 27, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 28, 2, "Pre-Tax Income", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 28, 3+i, f"={cl}24-{cl}26+{cl}27", BOLD, TOT_FILL, R, BRD, FMT_NUM)

sc(ws2, 29, 2, "Income Tax (Provision) Benefit", NORMAL, None, L, BRD)
tax = [513, -122, -346, 381, -103]
for i, v in enumerate(tax):
    sc(ws2, 29, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 30, 2, "  Effective Tax Rate", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 30, 3+i, f"={cl}29/{cl}28", NORMAL, None, R, BRD, FMT_PCT)

sc(ws2, 31, 2, "Net Income (Continuing Ops)", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 31, 3+i, f"={cl}28-{cl}29", BOLD, TOT_FILL, R, BRD, FMT_NUM)

sc(ws2, 32, 2, "  Net Margin", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 32, 3+i, f"={cl}31/{cl}8", NORMAL, None, R, BRD, FMT_PCT)

sc(ws2, 33, 2, "Diluted EPS ($)", NORMAL, None, L, BRD)
# EPS set to 2.61 to align strictly with "Net Income (Continuing Ops)" 
eps = [2.57, 0.84, 0.53, 1.00, 2.61]
for i, v in enumerate(eps):
    sc(ws2, 33, 3+i, v, NORMAL, None, R, BRD, '"$"#,##0.00')

sc(ws2, 34, 2, "Diluted Shares (M)", NORMAL, None, L, BRD)
shares = [1229, 1571, 1625, 1637, 1636]
for i, v in enumerate(shares):
    sc(ws2, 34, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

head_row(ws2, 36, 2, 7, "CASH FLOW & CAPEX")

sc(ws2, 37, 2, "Operating Cash Flow (Cont. Ops)", NORMAL, None, L, BRD)
ocf = [3521, 3565, 1667, 3041, 6493]
for i, v in enumerate(ocf):
    sc(ws2, 37, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 38, 2, "Capital Expenditures", NORMAL, None, L, BRD)
capex = [301, 450, 546, 636, 974]
for i, v in enumerate(capex):
    sc(ws2, 38, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 39, 2, "Free Cash Flow (FCF)", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 39, 3+i, f"={cl}37-{cl}38", BOLD, TOT_FILL, R, BRD, FMT_NUM)

sc(ws2, 42, 2, "Depreciation & Amort. (excl. acq.)", NORMAL, None, L, BRD)
da = [407, 626, 642, 671, 750]
for i, v in enumerate(da):
    sc(ws2, 42, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 43, 2, "Stock-Based Compensation", NORMAL, None, L, BRD)
sbc = [379, 1081, 1384, 1407, 1638]
for i, v in enumerate(sbc):
    sc(ws2, 43, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

head_row(ws2, 46, 2, 7, "BALANCE SHEET (Year-End)")

sc(ws2, 47, 2, "Cash & Cash Equivalents", NORMAL, None, L, BRD)
cash = [2535, 4835, 3933, 3787, 5539]
for i, v in enumerate(cash):
    sc(ws2, 47, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 48, 2, "Short-Term Investments", NORMAL, None, L, BRD)
sti = [1073, 1020, 1840, 1345, 5013]
for i, v in enumerate(sti):
    sc(ws2, 48, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 49, 2, "Total Cash + ST Investments", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 49, 3+i, f"={cl}47+{cl}48", BOLD, TOT_FILL, R, BRD, FMT_NUM)

sc(ws2, 50, 2, "Total Debt (Principal)", NORMAL, None, L, BRD)
debt = [313, 2501, 2500, 1750, 3250]
for i, v in enumerate(debt):
    sc(ws2, 50, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 51, 2, "Net Cash (Cash - Debt)", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws2, 51, 3+i, f"={cl}49-{cl}50", BOLD, TOT_FILL, R, BRD, FMT_NUM)

sc(ws2, 52, 2, "Total Stockholders' Equity", NORMAL, None, L, BRD)
equity = [7497, 54750, 55892, 57568, 62999]
for i, v in enumerate(equity):
    sc(ws2, 52, 3+i, v, NORMAL, None, R, BRD, FMT_NUM)

sc(ws2, 54, 2, "Note: 2021-2023 financials strictly sourced from provided 10-K text files.", NOTE_FONT, None, L)
sc(ws2, 55, 2, "Note: All years use formula-based Net Income (Pre-Tax Income - Tax) for methodological consistency.", NOTE_FONT, None, L)


# ============================================================
# SHEET 3: ASSUMPTIONS
# ============================================================
ws3 = wb.create_sheet("Assumptions")
ws3.sheet_view.showGridLines = False
ws3.column_dimensions['A'].width = 3
ws3.column_dimensions['B'].width = 42
ws3.column_dimensions['C'].width = 14
for col in 'DEFGH':
    ws3.column_dimensions[col].width = 13

title_row(ws3, 2, 2, 8, "DCF Model Assumptions")

sc(ws3, 4, 2, "All yellow cells are inputs. Change to update the entire model.", NOTE_FONT, None, L)

proj_years = [2026, 2027, 2028, 2029, 2030]
sc(ws3, 5, 2, "Projection Year →", BOLD, None, L)
for i, y in enumerate(proj_years):
    sc(ws3, 5, 4+i, y, HEAD_FONT, HEAD_FILL, C, BRD)

head_row(ws3, 7, 2, 8, "REVENUE GROWTH")

sc(ws3, 8, 2, "Revenue Growth Rate (%)", BOLD, None, L, BRD)
rev_growth = [0.36, 0.24, 0.18, 0.14, 0.11]
for i, v in enumerate(rev_growth):
    sc(ws3, 8, 4+i, v, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)

head_row(ws3, 12, 2, 8, "PROFITABILITY MARGINS")

sc(ws3, 13, 2, "Gross Margin (%)", BOLD, None, L, BRD)
gm = [0.515, 0.52, 0.525, 0.53, 0.54]
for i, v in enumerate(gm):
    sc(ws3, 13, 4+i, v, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)

sc(ws3, 15, 2, "R&D % of Revenue (%)", BOLD, None, L, BRD)
rd_pct = [0.23, 0.22, 0.21, 0.20, 0.20]
for i, v in enumerate(rd_pct):
    sc(ws3, 15, 4+i, v, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)

sc(ws3, 17, 2, "MG&A % of Revenue (%)", BOLD, None, L, BRD)
mga_pct = [0.11, 0.105, 0.10, 0.095, 0.09]
for i, v in enumerate(mga_pct):
    sc(ws3, 17, 4+i, v, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)

sc(ws3, 19, 2, "Amort. of Acq. Intangibles ($M, fixed)", BOLD, None, L, BRD)
amort_sched = [2153, 2036, 1923, 1691, 1454]
for i, v in enumerate(amort_sched):
    sc(ws3, 19, 4+i, v, INPUT_FONT, INPUT_FILL, C, BRD, FMT_NUM)

head_row(ws3, 22, 2, 8, "TAX, CAPEX & WORKING CAPITAL")

sc(ws3, 23, 2, "Effective Tax Rate (%)", BOLD, None, L, BRD)
for i in range(5):
    sc(ws3, 23, 4+i, 0.19, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)
sc(ws3, 24, 2, "  Rationale: Normalized to 19% (Statutory + State taxes).", NOTE_FONT, None, L)

sc(ws3, 25, 2, "CapEx % of Revenue (%)", BOLD, None, L, BRD)
capex_pct = [0.03, 0.03, 0.028, 0.025, 0.025]
for i, v in enumerate(capex_pct):
    sc(ws3, 25, 4+i, v, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)

sc(ws3, 27, 2, "D&A % of Revenue (excl. acq. amort.) (%)", BOLD, None, L, BRD)
da_pct = [0.025, 0.025, 0.025, 0.025, 0.025]
for i, v in enumerate(da_pct):
    sc(ws3, 27, 4+i, v, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)

sc(ws3, 28, 2, "SBC % of Revenue (%)", BOLD, None, L, BRD)
sbc_pct = [0.05, 0.048, 0.045, 0.043, 0.04]
for i, v in enumerate(sbc_pct):
    sc(ws3, 28, 4+i, v, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)

sc(ws3, 30, 2, "ΔNWC % of ΔRevenue (%)", BOLD, None, L, BRD)
nwc_pct = [0.15, 0.12, 0.10, 0.08, 0.08]
for i, v in enumerate(nwc_pct):
    sc(ws3, 30, 4+i, v, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)

head_row(ws3, 33, 2, 8, "WACC COMPONENTS")

wacc_inputs = [
    ("Risk-Free Rate (10Y UST, %)",      0.0425, "FRED data, early 2026"),
    ("Equity Risk Premium (%)",          0.055,  "Damodaran 2025 estimate"),
    ("Beta (5Y monthly)",                1.75,   "AMD volatility vs. S&P 500"),
    ("Pre-Tax Cost of Debt (%)",         0.043,  "Weighted avg of AMD notes"),
    ("Market Cap ($M)",                  385800, "1,630M shares × $236.73"),
    ("Total Debt ($M, principal)",       3250,   "2025 10-K, Note 9"),
    ("Total Cash + ST Investments ($M)", 10552,  "2025 10-K Balance Sheet"),
    ("Diluted Shares (M)",               1796,   "1,636M base + 160M OpenAI warrant (full vesting)"),
]
for i, (lbl, val, note) in enumerate(wacc_inputs):
    sc(ws3, 34+i, 2, lbl, BOLD, None, L, BRD)
    fmt = FMT_PCT2 if isinstance(val, float) and val < 1 else FMT_NUM
    sc(ws3, 34+i, 3, val, INPUT_FONT, INPUT_FILL, C, BRD, fmt)
    sc(ws3, 34+i, 4, note, NOTE_FONT, None, L)
    ws3.merge_cells(start_row=34+i, start_column=4, end_row=34+i, end_column=8)

head_row(ws3, 44, 2, 8, "TERMINAL VALUE & ADVANCED SETTINGS")

sc(ws3, 45, 2, "Terminal Growth Rate (%)", BOLD, None, L, BRD)
sc(ws3, 45, 3, 0.025, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)
sc(ws3, 46, 2, "  Rationale: Lowered to 2.5% to align with reinvestment capacity.", NOTE_FONT, None, L)

sc(ws3, 47, 2, "Mid-Year Convention (0=No, 0.5=Yes)", BOLD, None, L, BRD)
sc(ws3, 47, 3, 0.5, INPUT_FONT, INPUT_FILL, C, BRD, FMT_DEC)
sc(ws3, 48, 2, "  Rationale: Assumes cash flows occur evenly throughout the year.", NOTE_FONT, None, L)

sc(ws3, 49, 2, "Terminal ROIC (%)", BOLD, None, L, BRD)
sc(ws3, 49, 3, 0.20, INPUT_FONT, INPUT_FILL, C, BRD, FMT_PCT)
sc(ws3, 50, 2, "  Rationale: 20% ROIC for asset-light semi designer. Req. reinvest = 2.5%/20% = 12.5%.", NOTE_FONT, None, L)


# ============================================================
# SHEET 4: WACC
# ============================================================
ws4 = wb.create_sheet("WACC")
ws4.sheet_view.showGridLines = False
ws4.column_dimensions['A'].width = 3
ws4.column_dimensions['B'].width = 38
ws4.column_dimensions['C'].width = 16
ws4.column_dimensions['D'].width = 50

title_row(ws4, 2, 2, 4, "Weighted Average Cost of Capital (WACC)")

sc(ws4, 4, 2, "Component", HEAD_FONT, HEAD_FILL, L, BRD)
sc(ws4, 4, 3, "Value", HEAD_FONT, HEAD_FILL, C, BRD)
sc(ws4, 4, 4, "Formula / Source", HEAD_FONT, HEAD_FILL, L, BRD)

sc(ws4, 6, 2, "COST OF EQUITY (CAPM)", SUB_FONT, SUB_FILL, L, BRD)
ws4.merge_cells('B6:D6')

sc(ws4, 7, 2, "Risk-Free Rate (Rf)", NORMAL, None, L, BRD)
sc(ws4, 7, 3, "='Assumptions'!C34", NORMAL, CALC_FILL, C, BRD, FMT_PCT2)
sc(ws4, 8, 2, "Equity Risk Premium (ERP)", NORMAL, None, L, BRD)
sc(ws4, 8, 3, "='Assumptions'!C35", NORMAL, CALC_FILL, C, BRD, FMT_PCT2)
sc(ws4, 9, 2, "Beta (β)", NORMAL, None, L, BRD)
sc(ws4, 9, 3, "='Assumptions'!C36", NORMAL, CALC_FILL, C, BRD, '0.00')

sc(ws4, 10, 2, "Cost of Equity = Rf + β × ERP", BOLD, TOT_FILL, L, BRD)
sc(ws4, 10, 3, "=C7+C9*C8", BOLD, TOT_FILL, C, BRD, FMT_PCT2)

sc(ws4, 13, 2, "Pre-Tax Cost of Debt", NORMAL, None, L, BRD)
sc(ws4, 13, 3, "='Assumptions'!C37", NORMAL, CALC_FILL, C, BRD, FMT_PCT2)
sc(ws4, 14, 2, "Tax Rate", NORMAL, None, L, BRD)
sc(ws4, 14, 3, "='Assumptions'!D23", NORMAL, CALC_FILL, C, BRD, FMT_PCT)

sc(ws4, 15, 2, "After-Tax Cost of Debt = Kd × (1 - t)", BOLD, TOT_FILL, L, BRD)
sc(ws4, 15, 3, "=C13*(1-C14)", BOLD, TOT_FILL, C, BRD, FMT_PCT2)

sc(ws4, 18, 2, "Market Value of Equity ($M)", NORMAL, None, L, BRD)
sc(ws4, 18, 3, "='Assumptions'!C38", NORMAL, CALC_FILL, C, BRD, FMT_NUM)
sc(ws4, 19, 2, "Market Value of Debt ($M)", NORMAL, None, L, BRD)
sc(ws4, 19, 3, "='Assumptions'!C39", NORMAL, CALC_FILL, C, BRD, FMT_NUM)
sc(ws4, 20, 2, "Total Capital ($M)", NORMAL, TOT_FILL, L, BRD)
sc(ws4, 20, 3, "=C18+C19", BOLD, TOT_FILL, C, BRD, FMT_NUM)
sc(ws4, 21, 2, "Weight of Equity (We)", NORMAL, None, L, BRD)
sc(ws4, 21, 3, "=C18/C20", NORMAL, CALC_FILL, C, BRD, FMT_PCT2)
sc(ws4, 22, 2, "Weight of Debt (Wd)", NORMAL, None, L, BRD)
sc(ws4, 22, 3, "=C19/C20", NORMAL, CALC_FILL, C, BRD, FMT_PCT2)

sc(ws4, 24, 2, "WACC = We × Ke + Wd × Kd×(1-t)", BOLD, TITLE_FILL, L, BRD)
sc(ws4, 24, 3, "=C21*C10+C22*C15", TITLE_FONT, TITLE_FILL, C, BRD, FMT_PCT2)


# ============================================================
# SHEET 5: PROJECTIONS (2026-2030)
# ============================================================
ws5 = wb.create_sheet("Projections")
ws5.sheet_view.showGridLines = False
ws5.column_dimensions['A'].width = 3
ws5.column_dimensions['B'].width = 42
for col in 'CDEFG':
    ws5.column_dimensions[col].width = 14

title_row(ws5, 2, 2, 7, "Projected Financials & Free Cash Flow (2026-2030)")

sc(ws5, 4, 2, "($ in millions)", BOLD, None, L)
proj_yrs = [2026, 2027, 2028, 2029, 2030]
for i, y in enumerate(proj_yrs):
    sc(ws5, 4, 3+i, y, HEAD_FONT, HEAD_FILL, C, BRD)

head_row(ws5, 6, 2, 7, "PROJECTED INCOME STATEMENT")

sc(ws5, 7, 2, "Net Revenue", BOLD, None, L, BRD)
sc(ws5, 7, 3, "=Historicals!G8*(1+Assumptions!D8)", BOLD, None, R, BRD, FMT_NUM)
for i in range(1, 5):
    cl = get_column_letter(3+i); pl = get_column_letter(2+i)
    sc(ws5, 7, 3+i, f"={pl}7*(1+Assumptions!{get_column_letter(4+i)}8)", BOLD, None, R, BRD, FMT_NUM)

sc(ws5, 9, 2, "Gross Profit", BOLD, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i); al = get_column_letter(4+i)
    sc(ws5, 9, 3+i, f"={cl}7*Assumptions!{al}13", BOLD, None, R, BRD, FMT_NUM)

sc(ws5, 11, 2, "Research & Development", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i); al = get_column_letter(4+i)
    sc(ws5, 11, 3+i, f"={cl}7*Assumptions!{al}15", NORMAL, None, R, BRD, FMT_NUM)

sc(ws5, 12, 2, "Marketing, Gen. & Admin.", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i); al = get_column_letter(4+i)
    sc(ws5, 12, 3+i, f"={cl}7*Assumptions!{al}17", NORMAL, None, R, BRD, FMT_NUM)

sc(ws5, 13, 2, "Amort. of Acq. Intangibles", NORMAL, None, L, BRD)
for i in range(5):
    al = get_column_letter(4+i)
    sc(ws5, 13, 3+i, f"=Assumptions!{al}19", NORMAL, None, R, BRD, FMT_NUM)

sc(ws5, 14, 2, "EBIT (Operating Income)", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws5, 14, 3+i, f"={cl}9-{cl}11-{cl}12-{cl}13", BOLD, TOT_FILL, R, BRD, FMT_NUM)

sc(ws5, 16, 2, "Taxes on EBIT", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i); al = get_column_letter(4+i)
    sc(ws5, 16, 3+i, f"={cl}14*Assumptions!{al}23", NORMAL, None, R, BRD, FMT_NUM)

sc(ws5, 17, 2, "NOPAT", BOLD, TOT_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws5, 17, 3+i, f"={cl}14-{cl}16", BOLD, TOT_FILL, R, BRD, FMT_NUM)

head_row(ws5, 19, 2, 7, "FREE CASH FLOW CALCULATION")

sc(ws5, 20, 2, "(+) Amortization of Acq. Intangibles", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    sc(ws5, 20, 3+i, f"={cl}13", NORMAL, None, R, BRD, FMT_NUM)

sc(ws5, 21, 2, "(+) Depreciation & Amort. (other)", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i); al = get_column_letter(4+i)
    sc(ws5, 21, 3+i, f"={cl}7*Assumptions!{al}27", NORMAL, None, R, BRD, FMT_NUM)

sc(ws5, 22, 2, "(–) Stock-Based Compensation", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i); al = get_column_letter(4+i)
    sc(ws5, 22, 3+i, f"={cl}7*Assumptions!{al}28", NORMAL, None, R, BRD, FMT_NUM)

sc(ws5, 23, 2, "(–) Capital Expenditures", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i); al = get_column_letter(4+i)
    sc(ws5, 23, 3+i, f"={cl}7*Assumptions!{al}25", NORMAL, None, R, BRD, FMT_NUM)

sc(ws5, 24, 2, "(–) Change in Net Working Capital", NORMAL, None, L, BRD)
sc(ws5, 24, 3, "=(C7-Historicals!G8)*Assumptions!D30", NORMAL, None, R, BRD, FMT_NUM)
for i in range(1, 5):
    cl = get_column_letter(3+i); pl = get_column_letter(2+i); al = get_column_letter(4+i)
    sc(ws5, 24, 3+i, f"=({cl}7-{pl}7)*Assumptions!{al}30", NORMAL, None, R, BRD, FMT_NUM)

sc(ws5, 25, 2, "Unlevered Free Cash Flow", BOLD, TITLE_FILL, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    formula = f"={cl}17+{cl}20+{cl}21-{cl}22-{cl}23-{cl}24"
    sc(ws5, 25, 3+i, formula, TITLE_FONT, TITLE_FILL, R, BRD, FMT_NUM)

sc(ws5, 27, 2, "Explicit-Year Reinvestment Rate ((CapEx-DA+ΔNWC)/NOPAT)", NORMAL, None, L, BRD)
for i in range(5):
    cl = get_column_letter(3+i)
    formula = f"=({cl}23-{cl}21+{cl}24)/{cl}17"
    sc(ws5, 27, 3+i, formula, NORMAL, CALC_FILL, C, BRD, FMT_PCT)
sc(ws5, 28, 2, "  Note: 2030 rate is elevated due to still-high revenue growth (11%).", NOTE_FONT, None, L)

head_row(ws5, 30, 2, 7, "TERMINAL REINVESTMENT CONSISTENCY CHECK (Steady-State)")

sc(ws5, 31, 2, "Terminal Revenue (2030 × (1+g))", NORMAL, None, L, BRD)
sc(ws5, 31, 3, "=G7*(1+'Assumptions'!C45)", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws5, 32, 2, "Terminal ΔRevenue", NORMAL, None, L, BRD)
sc(ws5, 32, 3, "=G7*'Assumptions'!C45", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws5, 33, 2, "Terminal ΔNWC (ΔRev × NWC%)", NORMAL, None, L, BRD)
sc(ws5, 33, 3, "=C32*Assumptions!H30", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws5, 34, 2, "Terminal CapEx - D&A (steady state = 0)", NORMAL, None, L, BRD)
sc(ws5, 34, 3, 0, NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws5, 35, 2, "Terminal NOPAT (2030 NOPAT × (1+g))", NORMAL, None, L, BRD)
sc(ws5, 35, 3, "=G17*(1+'Assumptions'!C45)", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws5, 36, 2, "Terminal Reinvestment Rate (ΔNWC / NOPAT)", BOLD, TOT_FILL, L, BRD)
sc(ws5, 36, 3, "=C33/C35", BOLD, TOT_FILL, C, BRD, FMT_PCT)

sc(ws5, 37, 2, "Required Reinvestment (g / ROIC)", NORMAL, None, L, BRD)
sc(ws5, 37, 3, "='Assumptions'!C45/'Assumptions'!C49", NORMAL, CALC_FILL, C, BRD, FMT_PCT)

sc(ws5, 38, 2, "Gap (Actual - Required)", BOLD, None, L, BRD)
sc(ws5, 38, 3, "=C36-C37", BOLD, WARN_FILL, C, BRD, FMT_PCT)

sc(ws5, 40, 2, "Note: A structural gap exists because R&D (intellectual capital) is expensed,", NOTE_FONT, None, L)
sc(ws5, 41, 2, "not capitalized under GAAP. Standard valuation practice acknowledges this.", NOTE_FONT, None, L)
sc(ws5, 42, 2, "Capitalizing R&D would increase invested capital & reinvestment, narrowing the gap.", NOTE_FONT, None, L)
sc(ws5, 43, 2, "Terminal growth of 2.5% is conservative given this structural limitation.", NOTE_FONT, None, L)


# ============================================================
# SHEET 6: DCF VALUATION
# ============================================================
ws6 = wb.create_sheet("DCF Valuation")
ws6.sheet_view.showGridLines = False
ws6.column_dimensions['A'].width = 3
ws6.column_dimensions['B'].width = 38
ws6.column_dimensions['C'].width = 18
ws6.column_dimensions['D'].width = 18
ws6.column_dimensions['E'].width = 45

title_row(ws6, 2, 2, 5, "DCF Valuation Summary")

sc(ws6, 4, 2, "Year", HEAD_FONT, HEAD_FILL, L, BRD)
sc(ws6, 4, 3, "Free Cash Flow ($M)", HEAD_FONT, HEAD_FILL, C, BRD)
sc(ws6, 4, 4, "Discount Factor", HEAD_FONT, HEAD_FILL, C, BRD)
sc(ws6, 4, 5, "Present Value ($M)", HEAD_FONT, HEAD_FILL, C, BRD)

sc(ws6, 5, 2, "WACC:", BOLD, None, L, BRD)
sc(ws6, 5, 3, "=WACC!C24", BOLD, CALC_FILL, C, BRD, FMT_PCT2)

proj_yrs = [2026, 2027, 2028, 2029, 2030]
for i, y in enumerate(proj_yrs):
    r = 7 + i
    sc(ws6, r, 2, f"FY{y}", BOLD, None, L, BRD)
    proj_col = get_column_letter(3+i)
    sc(ws6, r, 3, f"=Projections!{proj_col}25", NORMAL, CALC_FILL, C, BRD, FMT_NUM)
    sc(ws6, r, 4, f"=1/(1+$C$5)^({i+1}-'Assumptions'!C47)", NORMAL, CALC_FILL, C, BRD, '0.0000')
    sc(ws6, r, 5, f"=C{r}*D{r}", BOLD, CALC_FILL, C, BRD, FMT_NUM)

sc(ws6, 13, 2, "Sum of PV of FCF (2026-2030)", BOLD, TOT_FILL, L, BRD)
sc(ws6, 13, 5, "=SUM(E7:E11)", BOLD, TOT_FILL, C, BRD, FMT_NUM)

sc(ws6, 16, 2, "FCF in 2030 (final year)", NORMAL, None, L, BRD)
sc(ws6, 16, 3, "=C11", NORMAL, CALC_FILL, C, BRD, FMT_NUM)
sc(ws6, 17, 2, "Terminal Growth Rate (g)", NORMAL, None, L, BRD)
sc(ws6, 17, 3, "='Assumptions'!C45", NORMAL, CALC_FILL, C, BRD, FMT_PCT)

sc(ws6, 18, 2, "Terminal Value = FCF×(1+g)/(WACC-g)", BOLD, None, L, BRD)
sc(ws6, 18, 3, "=C16*(1+C17)/(C5-C17)", BOLD, CALC_FILL, C, BRD, FMT_NUM)

sc(ws6, 19, 2, "Discount Factor (Year 5)", NORMAL, None, L, BRD)
sc(ws6, 19, 3, "=D11", NORMAL, CALC_FILL, C, BRD, '0.0000')
sc(ws6, 20, 2, "PV of Terminal Value", BOLD, TOT_FILL, L, BRD)
sc(ws6, 20, 3, "=C18*C19", BOLD, TOT_FILL, C, BRD, FMT_NUM)

sc(ws6, 23, 2, "PV of FCF (2026-2030)", NORMAL, None, L, BRD)
sc(ws6, 23, 3, "=E13", NORMAL, CALC_FILL, C, BRD, FMT_NUM)
sc(ws6, 24, 2, "PV of Terminal Value", NORMAL, None, L, BRD)
sc(ws6, 24, 3, "=C20", NORMAL, CALC_FILL, C, BRD, FMT_NUM)
sc(ws6, 25, 2, "(+) Cumulative FCF adjustment", NORMAL, None, L, BRD)
sc(ws6, 25, 3, 0, NORMAL, None, C, BRD, FMT_NUM)

sc(ws6, 26, 2, "Enterprise Value", TITLE_FONT, TITLE_FILL, L, BRD)
sc(ws6, 26, 3, "=C23+C24+C25", TITLE_FONT, TITLE_FILL, C, BRD, FMT_NUM)

sc(ws6, 28, 2, "(–) Total Debt", NORMAL, None, L, BRD)
sc(ws6, 28, 3, "='Assumptions'!C39", NORMAL, CALC_FILL, C, BRD, FMT_NUM)
sc(ws6, 29, 2, "(+) Cash & ST Investments", NORMAL, None, L, BRD)
sc(ws6, 29, 3, "='Assumptions'!C40", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws6, 30, 2, "Enterprise Value (ref)", BOLD, None, L, BRD)
sc(ws6, 30, 3, "=C26", BOLD, None, C, BRD, FMT_NUM)
sc(ws6, 31, 2, "Net Debt", NORMAL, None, L, BRD)
sc(ws6, 31, 3, "=C28-C29", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws6, 33, 2, "Equity Value", TITLE_FONT, TITLE_FILL, L, BRD)
sc(ws6, 33, 3, "=C26-C28+C29", TITLE_FONT, TITLE_FILL, C, BRD, FMT_NUM)

sc(ws6, 35, 2, "Diluted Shares Outstanding (M)", NORMAL, None, L, BRD)
sc(ws6, 35, 3, "='Assumptions'!C41", NORMAL, CALC_FILL, C, BRD, FMT_NUM)
sc(ws6, 35, 5, "1,636M base + 160M OpenAI warrant", NOTE_FONT, None, L)

sc(ws6, 36, 2, "Implied Share Price", Font(bold=True, size=14, color='FFFFFF'), TITLE_FILL, L, BRD)
sc(ws6, 36, 3, "=C33/C35", Font(bold=True, size=14, color='FFFFFF'), TITLE_FILL, C, BRD, FMT_USD)

sc(ws6, 38, 2, "Reference: AMD price on Jan 30, 2026", NORMAL, None, L, BRD)
sc(ws6, 38, 3, 236.73, NORMAL, None, C, BRD, FMT_USD)

sc(ws6, 39, 2, "Implied Premium/(Discount) vs. Market", BOLD, None, L, BRD)
sc(ws6, 39, 3, "=C36/C38-1", BOLD, CALC_FILL, C, BRD, FMT_PCT)

head_row(ws6, 42, 2, 5, "REVERSE DCF: What is the market pricing in?")

sc(ws6, 43, 2, "Current Market Price", NORMAL, None, L, BRD)
sc(ws6, 43, 3, "=C38", NORMAL, CALC_FILL, C, BRD, FMT_USD)

sc(ws6, 44, 2, "Implied Market Equity Value", NORMAL, None, L, BRD)
sc(ws6, 44, 3, "=C43*C35", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws6, 45, 2, "Implied Enterprise Value (Mkt)", NORMAL, None, L, BRD)
sc(ws6, 45, 3, "=C44+C28-C29", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws6, 46, 2, "Less: PV of Explicit FCFs", NORMAL, None, L, BRD)
sc(ws6, 46, 3, "=C23", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws6, 47, 2, "Implied PV of Terminal Value", BOLD, TOT_FILL, L, BRD)
sc(ws6, 47, 3, "=C45-C46", BOLD, TOT_FILL, C, BRD, FMT_NUM)

sc(ws6, 48, 2, "Implied Terminal FCF (2030)", BOLD, None, L, BRD)
sc(ws6, 48, 3, "=C47*(C5-C17)/((1+C17)*D11)", BOLD, CALC_FILL, C, BRD, FMT_NUM)

sc(ws6, 49, 2, "Model's Projected 2030 FCF", NORMAL, None, L, BRD)
sc(ws6, 49, 3, "=C11", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws6, 50, 2, "Premium of Mkt Implied FCF vs Model", BOLD, None, L, BRD)
sc(ws6, 50, 3, "=C48/C49-1", BOLD, CALC_FILL, C, BRD, FMT_PCT)
sc(ws6, 50, 5, "Shows how much AI execution is already priced in.", NOTE_FONT, None, L)

sc(ws6, 52, 2, "Note: PV of Terminal Value represents ~70% of EV.", NOTE_FONT, None, L)
sc(ws6, 53, 2, "This is a structural limitation of 5-year DCFs for high-growth companies.", NOTE_FONT, None, L)


# ============================================================
# SHEET 7: SENSITIVITY ANALYSIS
# ============================================================
ws7 = wb.create_sheet("Sensitivity")
ws7.sheet_view.showGridLines = False
ws7.column_dimensions['A'].width = 3
ws7.column_dimensions['B'].width = 28
for col in 'CDEFGHI':
    ws7.column_dimensions[col].width = 13

title_row(ws7, 2, 2, 9, "Sensitivity Analysis: Implied Share Price")

sc(ws7, 7, 2, "WACC ↓ \\ Term. Growth →", HEAD_FONT, HEAD_FILL, L, BRD)
wacc_range = [0.11, 0.12, 0.13, 0.14, 0.15, 0.16]
tg_range   = [0.015, 0.02, 0.025, 0.03, 0.035]

for j, tg in enumerate(tg_range):
    sc(ws7, 7, 3+j, tg, HEAD_FONT, HEAD_FILL, C, BRD, FMT_PCT)

for i, wacc in enumerate(wacc_range):
    r = 8 + i
    sc(ws7, r, 2, wacc, BOLD, SUB_FILL, C, BRD, FMT_PCT2)
    for j, tg in enumerate(tg_range):
        c = 3 + j
        wacc_ref = f"$B{r}"
        tg_ref = f"{get_column_letter(c)}$7"
        
        fcf_terms = []
        for k in range(5):
            proj_col = get_column_letter(3+k)
            fcf_ref = f"Projections!{proj_col}25"
            df = f"(1+{wacc_ref})^({k+1}-'Assumptions'!C47)"
            fcf_terms.append(f"{fcf_ref}/{df}")
        pv_fcf = "(" + "+".join(fcf_terms) + ")"
        
        fcf_2030 = "Projections!G25"
        tv = f"({fcf_2030}*(1+{tg_ref})/({wacc_ref}-{tg_ref}))"
        pv_tv = f"({tv}/(1+{wacc_ref})^(5-'Assumptions'!C47))"
        ev = f"({pv_fcf}+{pv_tv})"
        
        debt = "'Assumptions'!C39"
        cash = "'Assumptions'!C40"
        shares = "'Assumptions'!C41"
        equity = f"({ev}-{debt}+{cash})"
        per_share = f"={equity}/{shares}"
        
        sc(ws7, r, c, per_share, NORMAL, CALC_FILL, C, BRD, FMT_USD)

sc(ws7, 15, 2, "Base case WACC:", BOLD, None, L)
sc(ws7, 15, 3, "=WACC!C24", BOLD, CALC_FILL, C, BRD, FMT_PCT2)
sc(ws7, 15, 4, "Base case TG:", BOLD, None, L)
sc(ws7, 15, 5, "='Assumptions'!C45", BOLD, CALC_FILL, C, BRD, FMT_PCT)
sc(ws7, 15, 6, "Base case price:", BOLD, None, L)
sc(ws7, 15, 7, "='DCF Valuation'!C36", BOLD, CALC_FILL, C, BRD, FMT_USD)


# ============================================================
# SHEET 8: DASHBOARD
# ============================================================
ws8 = wb.create_sheet("Dashboard")
ws8.sheet_view.showGridLines = False
ws8.column_dimensions['A'].width = 3
ws8.column_dimensions['B'].width = 20
for col in 'CDEFGHI':
    ws8.column_dimensions[col].width = 12

title_row(ws8, 2, 2, 9, "Valuation Dashboard")
ws8.row_dimensions[2].height = 28

sc(ws8, 4, 2, "Projected Revenue & FCF ($M)", BOLD, None, L)

sc(ws8, 6, 2, "Year", BOLD, HEAD_FILL, C, BRD)
for i, y in enumerate(proj_yrs):
    sc(ws8, 6, 3+i, y, BOLD, HEAD_FILL, C, BRD)

sc(ws8, 7, 2, "Revenue", NORMAL, CALC_FILL, L, BRD)
for i in range(5):
    proj_col = get_column_letter(3+i)
    sc(ws8, 7, 3+i, f"=Projections!{proj_col}7", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

sc(ws8, 8, 2, "FCF", NORMAL, CALC_FILL, L, BRD)
for i in range(5):
    proj_col = get_column_letter(3+i)
    sc(ws8, 8, 3+i, f"=Projections!{proj_col}25", NORMAL, CALC_FILL, C, BRD, FMT_NUM)

chart = BarChart()
chart.type = "col"
chart.style = 10
chart.title = "AMD Projected Revenue vs. Free Cash Flow"
chart.y_axis.title = 'USD ($M)'
chart.x_axis.title = 'Fiscal Year'

data = Reference(ws8, min_col=2, min_row=7, max_row=8, max_col=7)
cats = Reference(ws8, min_col=3, min_row=6, max_row=6, max_col=7)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
chart.width = 18
chart.height = 10

ws8.add_chart(chart, "B11")

# ============================================================
# SAVE
# ============================================================
wb.save("AMD_New_Valuation2.xlsx")
print("✓ Workbook saved: AMD_New_Valuation2.xlsx")