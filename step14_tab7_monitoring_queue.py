"""Step 14: Tab 7 — Monitoring Queue (16 columns, A-P)."""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

# ── Colors ────────────────────────────────────────────────────────────────────
NAVY        = "1A3A5C"
NAVY_LIGHT  = "2E5F8A"
INSTRUCT_BG = "EAF0F6"
LIGHT_BLUE  = "D5E8F4"
LIGHT_GREY  = "F2F2F2"
GREEN_BG    = "E2EFDA"
GREEN_TXT   = "375623"
YELLOW_BG   = "FFF2CC"
ORANGE_BG   = "FCE9D6"
RED_BG      = "FCE4D6"
MED_GREY    = "CCCCCC"
DARK_GREY   = "888888"
WHITE       = "FFFFFF"

def solid(hex_color): return PatternFill("solid", fgColor=hex_color)
def thin():
    s = Side(style="thin", color=MED_GREY)
    return Border(left=s, right=s, top=s, bottom=s)
def apply(cell, value=None, fill=None, font=None, alignment=None):
    if value is not None: cell.value = value
    if fill is not None:  cell.fill = fill
    if font is not None:  cell.font = font
    if alignment is not None: cell.alignment = alignment
    cell.border = thin()
def center(wrap=False): return Alignment(horizontal="center", vertical="center", wrap_text=wrap)
def left_wrap():        return Alignment(horizontal="left",   vertical="center", wrap_text=True)
def hf(color=WHITE, size=9, bold=True): return Font(bold=bold, size=size, color=color, name="Arial")
def instr_font(): return Font(italic=True, size=7, color=DARK_GREY, name="Arial")

# ── Load workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook("SCG_CMA_System_v7.xlsx")
ws = wb["Monitoring Queue"]

LAST_COL = "P"   # 16 columns A-P

# ── ROW 1 — Title Banner ──────────────────────────────────────────────────────
ws.merge_cells(f"A1:{LAST_COL}1")
apply(ws["A1"],
      value="SCG CMA SYSTEM — MONITORING QUEUE | Make.com checks daily | Zero manual entry after submission | v7",
      fill=solid(NAVY), font=hf(size=12), alignment=center())
ws.row_dimensions[1].height = 22

# ── ROW 2 — Section Banners ───────────────────────────────────────────────────
ws.merge_cells("A2:I2")
apply(ws["A2"],
      value="SUBMISSION TRACKING",
      fill=solid(NAVY_LIGHT), font=hf(), alignment=center())

ws.merge_cells("J2:P2")
apply(ws["J2"],
      value="AUTO-POPULATED BY APIFY — ZERO MANUAL ENTRY",
      fill=solid(GREEN_BG), font=hf(color=GREEN_TXT), alignment=center())
ws.row_dimensions[2].height = 18

# ── Column definitions ────────────────────────────────────────────────────────
# (col, header, instruction, header_bg, hdr_txt, width, num_fmt)
columns = [
    # SUBMISSION TRACKING — NAVY_LIGHT
    ("A", "Submission ID",     "From Property Submissions",                      NAVY_LIGHT, WHITE,     13,  None),
    ("B", "Agent Name",        "From Property Submissions",                      NAVY_LIGHT, WHITE,     18,  None),
    ("C", "Office/Group",      "From Agent Roster",                              NAVY_LIGHT, WHITE,     22,  None),
    ("D", "Property Address",  "Subject property address",                       NAVY_LIGHT, WHITE,     32,  None),
    ("E", "Zip Code",          "For Apify address match",                        NAVY_LIGHT, WHITE,     10,  None),
    ("F", "Property Class",    "In-Town or Rural/Acreage",                       NAVY_LIGHT, WHITE,     14,  None),
    ("G", "Status",            "System updates automatically",                   NAVY_LIGHT, WHITE,     14,  None),
    ("H", "CMA Report Date",   "Date report was generated",                      NAVY_LIGHT, WHITE,     14,  None),
    ("I", "Claude Midpoint",   "Claude recommended midpoint",                    NAVY_LIGHT, WHITE,     14,  "$#,##0"),
    # AUTO-POPULATED BY APIFY — GREEN
    ("J", "Last Apify Check",  "Last Make.com check date",                       GREEN_BG,  GREEN_TXT, 14,  None),
    ("K", "Original List Price","Auto — first time detected — never overwritten",GREEN_BG,  GREEN_TXT, 16,  "$#,##0"),
    ("L", "Current List Price","Auto — updated each check",                      GREEN_BG,  GREEN_TXT, 14,  "$#,##0"),
    ("M", "Price Reductions",  "Auto — count of price drops detected",           GREEN_BG,  GREEN_TXT, 12,  None),
    ("N", "Contract Price",    "From submission when agent marks Pending",        GREEN_BG,  GREEN_TXT, 14,  "$#,##0"),
    ("O", "Sold Price",        "Auto — Apify detects closed sale",               GREEN_BG,  GREEN_TXT, 13,  "$#,##0"),
    ("P", "Sold Date",         "Auto — Apify",                                   GREEN_BG,  GREEN_TXT, 12,  None),
]

# ── ROW 3 — Headers ───────────────────────────────────────────────────────────
for col, label, _, bg, txt, width, _ in columns:
    apply(ws[f"{col}3"],
          value=label,
          fill=solid(bg),
          font=hf(color=txt),
          alignment=center(wrap=True))
    ws.column_dimensions[col].width = width
ws.row_dimensions[3].height = 44

# ── ROW 4 — Instructions ──────────────────────────────────────────────────────
for col, _, instruction, _, _, _, _ in columns:
    apply(ws[f"{col}4"],
          value=instruction,
          fill=solid(INSTRUCT_BG),
          font=instr_font(),
          alignment=left_wrap())
ws.row_dimensions[4].height = 30

# ── Number formats on data rows ───────────────────────────────────────────────
for col, _, _, _, _, _, num_fmt in columns:
    if num_fmt:
        for row in range(5, 2001):
            ws[f"{col}{row}"].number_format = num_fmt

# ── Status Dropdown ───────────────────────────────────────────────────────────
dv = DataValidation(
    type="list",
    formula1='"Monitoring,Listed,Pending,Fell Through,Sold,Archived"',
    allow_blank=True
)
dv.sqref = "G5:G2000"
ws.add_data_validation(dv)

# ── ROW 5 — Sample Data ───────────────────────────────────────────────────────
sample_fills = {
    "A": LIGHT_BLUE, "B": LIGHT_BLUE, "C": LIGHT_BLUE,
    "D": LIGHT_GREY, "E": LIGHT_GREY, "F": LIGHT_GREY,
    "G": LIGHT_GREY, "H": LIGHT_GREY, "I": LIGHT_GREY,
    "J": GREEN_BG,   "K": GREEN_BG,   "L": GREEN_BG,
    "M": GREEN_BG,   "N": GREEN_BG,   "O": GREEN_BG,
    "P": GREEN_BG,
}
sample_values = [
    ("A", "CMA-001"),
    ("B", "Jeff Fleeharty"),
    ("C", "RE/MAX Jeff Fleeharty"),
    ("D", "202 N Main St, Saint Elmo IL"),
    ("E", "62458"),
    ("F", "In-Town Residential"),
    ("G", "Monitoring"),
    ("H", "2026-05-07"),
    ("I", 102500),
    ("J", "2026-05-07"),
    ("K", "Pending"),
    ("L", "Pending"),
    ("M", 0),
    ("N", None),
    ("O", "Pending"),
    ("P", "Pending"),
]
for col, val in sample_values:
    c = ws[f"{col}5"]
    if val is not None:
        c.value = val
    c.fill      = solid(sample_fills[col])
    c.font      = Font(size=9, color="000000", name="Arial")
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border    = thin()
    for ccol, _, _, _, _, _, num_fmt in columns:
        if ccol == col and num_fmt and isinstance(val, (int, float)):
            c.number_format = num_fmt

# ── ROW 7 — How It Works Note ─────────────────────────────────────────────────
ws.merge_cells(f"A7:{LAST_COL}7")
apply(ws["A7"],
      value=("HOW THIS WORKS: Make.com runs daily on Status = Monitoring or Listed. "
             "Apify scrapes Realtor.com by exact address. "
             "Listing found → Original List Price captured (locked forever) → Current List Price and date update → Status → Listed. "
             "Price change detected → Price Reduction Count increments. "
             "Agent marks Pending → Contract Price captured → Velocity Metrics updated. "
             "Sold detected → Sold Price and Date populate → Triggers accuracy analysis → Scorecard and Velocity Metrics update. "
             "Jeff never enters anything after initial submission."),
      fill=solid(LIGHT_BLUE),
      font=Font(italic=True, size=8, color=NAVY, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[7].height = 40

# ── Conditional Formatting — STATUS (full row A:P) ────────────────────────────
DATA_RANGE = f"A5:{LAST_COL}2000"

# Priority order: most specific first (added last = lowest priority in openpyxl append)
# so add lowest-priority first, highest last
ws.conditional_formatting.add(DATA_RANGE,
    FormulaRule(formula=['$G5="Fell Through"'], fill=solid(RED_BG)))
ws.conditional_formatting.add(DATA_RANGE,
    FormulaRule(formula=['$G5="Pending"'],      fill=solid(ORANGE_BG)))
ws.conditional_formatting.add(DATA_RANGE,
    FormulaRule(formula=['$G5="Listed"'],        fill=solid(YELLOW_BG)))
ws.conditional_formatting.add(DATA_RANGE,
    FormulaRule(formula=['$G5="Sold"'],
                fill=solid(GREEN_BG),
                font=Font(bold=True, color=GREEN_TXT, name="Arial")))

# ── Freeze Panes ──────────────────────────────────────────────────────────────
ws.freeze_panes = "A5"

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save("SCG_CMA_System_v7.xlsx")
print("Step 14 complete — Tab 7 Monitoring Queue built.")
print("  16 columns A-P, freeze at A5.")
print("  Section banners: NAVY_LIGHT (A-I), GREEN_BG/GREEN_TXT (J-P).")
print("  Sample row 5 with 3-color section fills.")
print("  Status dropdown G5:G2000.")
print("  CF: Sold=GREEN+bold, Listed=YELLOW, Pending=ORANGE, Fell Through=RED.")
print("  How It Works note row 7 (LIGHT_BLUE italic).")
