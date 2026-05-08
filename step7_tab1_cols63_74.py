"""Step 7: Tab 1 — cols 63-74 (BK-BV), Market Tracking + Variance + STATUS."""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule

# ── Colors ────────────────────────────────────────────────────────────────────
NAVY        = "1A3A5C"
NAVY_LIGHT  = "2E5F8A"
INSTRUCT_BG = "EAF0F6"
MED_GREY    = "CCCCCC"
DARK_GREY   = "888888"
WHITE       = "FFFFFF"
GREEN_BG    = "E2EFDA"
GREEN_TXT   = "375623"
RED_BG      = "FCE4D6"
RED_TXT     = "9C0006"
YELLOW_BG   = "FFF2CC"
ORANGE_BG   = "FCE9D6"

# ── Helpers ───────────────────────────────────────────────────────────────────
def solid(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin", color=MED_GREY)
    return Border(left=s, right=s, top=s, bottom=s)

def apply(cell, value=None, fill=None, font=None, alignment=None):
    if value is not None:
        cell.value = value
    if fill is not None:
        cell.fill = fill
    if font is not None:
        cell.font = font
    if alignment is not None:
        cell.alignment = alignment
    cell.border = thin_border()

def center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def left_wrap():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def hdr_font(color=WHITE):
    return Font(bold=True, size=9, color=color, name="Arial")

def instr_font():
    return Font(italic=True, size=7, color=DARK_GREY, name="Arial")

# ── Load workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook("SCG_CMA_System_v7.xlsx")
ws = wb["Property Submissions"]

# ── Column map ────────────────────────────────────────────────────────────────
# Col 63=BK, 64=BL, 65=BM, 66=BN, 67=BO, 68=BP, 69=BQ (wait — let me recount)
# A=1..Z=26, AA=27..AZ=52, BA=53..BJ=62
# BK=63, BL=64, BM=65, BN=66, BO=67, BP=68, BQ=69, BR=70
# BS=71, BT=72, BU=73, BV=74

# ── ROW 2 — Section Banners ───────────────────────────────────────────────────
ws.merge_cells("BK2:BR2")
apply(ws["BK2"],
      value="MARKET TRACKING — Auto-populated by Apify",
      fill=solid(GREEN_BG),
      font=hdr_font(color=GREEN_TXT),
      alignment=center())

ws.merge_cells("BS2:BU2")
apply(ws["BS2"],
      value="VARIANCE",
      fill=solid(RED_BG),
      font=hdr_font(color=RED_TXT),
      alignment=center())

# STATUS — single col, no merge needed
apply(ws["BV2"],
      value="STATUS",
      fill=solid(NAVY),
      font=hdr_font(),
      alignment=center())

# ── MARKET TRACKING (BK-BR) ───────────────────────────────────────────────────
# (col, label, instruction, width, num_fmt, formula)
market_cols = [
    ("BK", "Original List Price",
     "Auto — captured first time Apify detects listing — NEVER overwritten",
     16, "$#,##0", None),
    ("BL", "List Date",
     "Auto — Apify",
     12, None, None),
    ("BM", "Current List Price",
     "Auto — updated each monitoring check",
     16, "$#,##0", None),
    ("BN", "Price Reduction Count",
     "Auto-calculated",
     12, None,
     '=IFERROR(IF(AND(BK5<>"",BM5<>""),IF(BM5<BK5,1,0),""),"")'),
    ("BO", "DOM at Check",
     "Auto — Apify",
     10, None, None),
    ("BP", "Sold Price",
     "Auto — Apify",
     13, "$#,##0", None),
    ("BQ", "Sold Date",
     "Auto — Apify",
     12, None, None),
    ("BR", "Final DOM",
     "Auto — Apify",
     10, None, None),
]

for col, label, instruction, width, num_fmt, formula in market_cols:
    apply(ws[f"{col}3"],
          value=label,
          fill=solid(GREEN_BG),
          font=hdr_font(color=GREEN_TXT),
          alignment=center(wrap=True))
    apply(ws[f"{col}4"],
          value=instruction,
          fill=solid(INSTRUCT_BG),
          font=instr_font(),
          alignment=left_wrap())
    ws.column_dimensions[col].width = width
    if num_fmt:
        for row in range(5, 2001):
            ws[f"{col}{row}"].number_format = num_fmt
    if formula:
        c = ws[f"{col}5"]
        c.value = formula
        c.border = thin_border()

# ── VARIANCE (BS-BU) ──────────────────────────────────────────────────────────
# Variance formulas reference:
#   BH = Claude Midpoint (col 60)
#   BK = Original List Price (col 63)
#   BM = Current List Price (col 65)  — spec says BQ5 for "list", using BK (original)
#   BP = Sold Price (col 68)
# Per spec:
#   Claude vs List %  = (BH - BQ) / BQ  — spec ref BQ5 = Original List Price col
#   Actually re-reading the spec: BQ5 refers to col 65 in the original spec
#   But with our mapping BK=63(OrigList), BM=65(CurrList), BP=68(Sold)
#   The spec formulas used: BQ5=OrigList, BS5=CurrList, BV5=Sold
#   In our mapping those land on: BK, BM, BP
#   So:
#     Claude vs List %  = (BH5 - BK5) / BK5
#     Claude vs Sold %  = (BH5 - BP5) / BP5
#     List vs Sold %    = (BP5 - BK5) / BK5

variance_cols = [
    ("BS", "Claude vs List %",
     "Auto-formula",
     13, "0.0%",
     '=IFERROR(IF(AND(BH5<>"",BK5<>""),(BH5-BK5)/BK5,""),"—")'),
    ("BT", "Claude vs Sold %",
     "Auto-formula",
     13, "0.0%",
     '=IFERROR(IF(AND(BH5<>"",BP5<>""),(BH5-BP5)/BP5,""),"—")'),
    ("BU", "List vs Sold %",
     "Auto-formula",
     13, "0.0%",
     '=IFERROR(IF(AND(BK5<>"",BP5<>""),(BP5-BK5)/BK5,""),"—")'),
]

for col, label, instruction, width, num_fmt, formula in variance_cols:
    apply(ws[f"{col}3"],
          value=label,
          fill=solid(RED_BG),
          font=hdr_font(color=RED_TXT),
          alignment=center(wrap=True))
    apply(ws[f"{col}4"],
          value=instruction,
          fill=solid(INSTRUCT_BG),
          font=instr_font(),
          alignment=left_wrap())
    ws.column_dimensions[col].width = width
    for row in range(5, 2001):
        ws[f"{col}{row}"].number_format = num_fmt
    c = ws[f"{col}5"]
    c.value = formula
    c.number_format = num_fmt
    c.border = thin_border()

# ── STATUS (BV) ───────────────────────────────────────────────────────────────
apply(ws["BV3"],
      value="STATUS",
      fill=solid(NAVY),
      font=hdr_font(),
      alignment=center(wrap=True))
apply(ws["BV4"],
      value="Change to SUBMITTED when complete — triggers Make.com",
      fill=solid(INSTRUCT_BG),
      font=instr_font(),
      alignment=left_wrap())
ws.column_dimensions["BV"].width = 14

dv_status = DataValidation(
    type="list",
    formula1='"Draft,Submitted,Monitoring,Listed,Pending,Fell Through,Sold,Archived"',
    allow_blank=True
)
dv_status.sqref = "BV5:BV2000"
ws.add_data_validation(dv_status)

# ── Conditional Formatting — Full-row STATUS rules ────────────────────────────
# Applies to A:BV (cols 1-74), rows 5-2000
DATA_RANGE = "A5:BV2000"

# STATUS = Submitted OR Pending → GREEN_BG
ws.conditional_formatting.add(
    DATA_RANGE,
    FormulaRule(formula=['OR($BV5="Submitted",$BV5="Pending")'],
                fill=solid(GREEN_BG))
)
# STATUS = Draft → YELLOW_BG
ws.conditional_formatting.add(
    DATA_RANGE,
    FormulaRule(formula=['$BV5="Draft"'],
                fill=solid(YELLOW_BG))
)
# STATUS = Fell Through → RED_BG
ws.conditional_formatting.add(
    DATA_RANGE,
    FormulaRule(formula=['$BV5="Fell Through"'],
                fill=solid(RED_BG))
)

# ── Conditional Formatting — Variance cols > 10% ABS → RED_BG bold RED_TXT ───
red_fill = solid(RED_BG)
red_font = Font(bold=True, color=RED_TXT, name="Arial")

for col in ["BS", "BT", "BU"]:
    rng = f"{col}5:{col}2000"
    ws.conditional_formatting.add(
        rng,
        FormulaRule(formula=[f'ABS({col}5)>0.10'],
                    fill=red_fill, font=red_font)
    )

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save("SCG_CMA_System_v7.xlsx")
print("Step 7 complete — Tab 1 cols 63-74 (BK-BV) built.")
print("  MARKET TRACKING banner BK2:BR2 (GREEN_BG / GREEN_TXT)")
print("  VARIANCE banner BS2:BU2 (RED_BG / RED_TXT)")
print("  STATUS col BV — dropdown + NAVY header")
print("  Variance formulas seeded in BS5, BT5, BU5 — 0.0% format")
print("  Full-row CF: Submitted/Pending=GREEN, Draft=YELLOW, Fell Through=RED")
print("  Variance CF: ABS > 10% → RED_BG bold RED_TXT")
