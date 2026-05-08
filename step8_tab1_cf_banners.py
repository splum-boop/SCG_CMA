"""Step 8: Tab 1 — Add font-color CF rules and two instruction banner rows (7 & 8)."""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

# ── Colors ────────────────────────────────────────────────────────────────────
NAVY       = "1A3A5C"
MED_GREY   = "CCCCCC"
WHITE      = "FFFFFF"
GREEN_BG   = "E2EFDA"
GREEN_TXT  = "375623"
YELLOW_BG  = "FFF2CC"
RED_BG     = "FCE4D6"
RED_TXT    = "9C0006"
LIGHT_BLUE = "D5E8F4"

# ── Helpers ───────────────────────────────────────────────────────────────────
def solid(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin", color=MED_GREY)
    return Border(left=s, right=s, top=s, bottom=s)

def apply_banner(cell, value, fill_color, font_color, bold, italic, size):
    cell.value = value
    cell.fill = solid(fill_color)
    cell.font = Font(bold=bold, italic=italic, size=size, color=font_color, name="Arial")
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell.border = thin_border()

# ── Load workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook("SCG_CMA_System_v7.xlsx")
ws = wb["Property Submissions"]

LAST_COL = get_column_letter(74)   # BV
DATA_RANGE  = f"A5:{LAST_COL}2000"

# ── Conditional Formatting — STATUS full-row rules (with font colors) ─────────
# Priority: highest added first.  We want Submitted/Pending > Draft > Fell Through.
# Add in reverse priority so the first-added ends up with highest priority in XML.

# Fell Through → RED_BG (no font override needed — default is fine)
ws.conditional_formatting.add(
    DATA_RANGE,
    FormulaRule(formula=['$BV5="Fell Through"'],
                fill=solid(RED_BG))
)
# Draft → YELLOW_BG
ws.conditional_formatting.add(
    DATA_RANGE,
    FormulaRule(formula=['$BV5="Draft"'],
                fill=solid(YELLOW_BG))
)
# Submitted OR Pending → GREEN_BG + GREEN_TXT font
ws.conditional_formatting.add(
    DATA_RANGE,
    FormulaRule(formula=['OR($BV5="Submitted",$BV5="Pending")'],
                fill=solid(GREEN_BG),
                font=Font(color=GREEN_TXT, name="Arial"))
)

# ── Conditional Formatting — Variance cols BS, BT, BU (ABS > 10%) ─────────────
red_fill = solid(RED_BG)
red_bold = Font(bold=True, color=RED_TXT, name="Arial")
for col in ["BS", "BT", "BU"]:
    ws.conditional_formatting.add(
        f"{col}5:{col}2000",
        FormulaRule(formula=[f'ABS({col}5)>0.10'],
                    fill=red_fill, font=red_bold)
    )

# ── Conditional Formatting — State col AS always LIGHT_BLUE ──────────────────
ws.conditional_formatting.add(
    "AS5:AS2000",
    FormulaRule(formula=["TRUE"], fill=solid(LIGHT_BLUE))
)

# ── Banner Row 7 — HOW TO USE (LIGHT_BLUE, bold NAVY, size 8) ─────────────────
ws.merge_cells(f"A7:{LAST_COL}7")
apply_banner(
    ws["A7"],
    value=(
        "HOW TO USE:  1) Fill Submission Info + Property Details  "
        "2) Enter year last updated for each system — leave blank if unknown  "
        "3) Fill Structure Features  "
        "4) Lot Size — ONE option only: L×W cols AG+AH  OR  Sq Ft col AI  OR  Acreage col AJ  "
        "5) Rural/Acreage only — fill yellow cols AN:AR  "
        "6) Select starting radius or Let System Decide  "
        "7) Add Local Knowledge notes  "
        "8) Change STATUS to SUBMITTED — this triggers Make.com"
    ),
    fill_color=LIGHT_BLUE,
    font_color=NAVY,
    bold=True,
    italic=False,
    size=8,
)
ws.row_dimensions[7].height = 30

# ── Banner Row 8 — VALUE LOGIC (YELLOW_BG, italic NAVY, size 8) ──────────────
ws.merge_cells(f"A8:{LAST_COL}8")
apply_banner(
    ws["A8"],
    value=(
        "VALUE LOGIC:  Year cols GREEN = 2015+   YELLOW = pre-2000  |  "
        "Age of home does NOT penalize value — condition and updates override  |  "
        "Attached Garage ~$5,000 default adjustment  |  "
        "Fully Finished Basement = material value add — agent adjusts  |  "
        "Acreage properties — land value first, structure second  |  "
        "Manufactured homes never comped against stick-built"
    ),
    fill_color=YELLOW_BG,
    font_color=NAVY,
    bold=False,
    italic=True,
    size=8,
)
ws.row_dimensions[8].height = 30

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save("SCG_CMA_System_v7.xlsx")
print("Step 8 complete — CF rules + banner rows added to Tab 1.")
print("  STATUS CF: Submitted/Pending=GREEN+font, Draft=YELLOW, Fell Through=RED")
print("  Variance CF: ABS>10% → RED_BG bold RED_TXT on BS:BU")
print("  State col AS: LIGHT_BLUE always")
print("  Row 7: HOW TO USE banner (LIGHT_BLUE, bold NAVY)")
print("  Row 8: VALUE LOGIC banner (YELLOW_BG, italic NAVY)")
