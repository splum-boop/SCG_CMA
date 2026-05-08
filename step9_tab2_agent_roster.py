"""Step 9: Tab 2 — Agent Roster."""
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
YELLOW_BG   = "FFF2CC"
SKIP_GREY   = "D9D9D9"
MED_GREY    = "CCCCCC"
DARK_GREY   = "888888"
WHITE       = "FFFFFF"

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

# ── Load workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook("SCG_CMA_System_v7.xlsx")
ws = wb["Agent Roster"]

LAST_COL = "G"   # 7 columns total

# ── ROW 1 — Title Banner ──────────────────────────────────────────────────────
ws.merge_cells(f"A1:{LAST_COL}1")
apply(ws["A1"],
      value="SCG CMA SYSTEM — AGENT ROSTER | Broker manages cols A-F | Sam owns col G | v7",
      fill=solid(NAVY),
      font=Font(bold=True, size=12, color=WHITE, name="Arial"),
      alignment=center())
ws.row_dimensions[1].height = 22

# ── ROW 2 — How It Works Banner ───────────────────────────────────────────────
ws.merge_cells(f"A2:{LAST_COL}2")
apply(ws["A2"],
      value=(
          "HOW THIS WORKS: Add one row per agent. Broker/responsible party manages columns A-F — "
          "add agents, deactivate agents. Column G is managed by SCG only. "
          "Agent Name must match EXACTLY what agent types on submission form — case sensitive. "
          "Make.com routes each report based on Agent Name matching Office/Group Name in Office Profiles tab."
      ),
      fill=solid(LIGHT_BLUE),
      font=Font(italic=True, size=8, color=NAVY, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[2].height = 30

# ── ROW 3 — Column Headers ────────────────────────────────────────────────────
headers = [
    ("A", "Agent Name",        NAVY_LIGHT, 22),
    ("B", "Agent Email",       NAVY_LIGHT, 24),
    ("C", "License Number",    NAVY_LIGHT, 16),
    ("D", "Active",            NAVY_LIGHT, 10),
    ("E", "Office/Group Name", NAVY_LIGHT, 28),
    ("F", "Onboard Date",      NAVY_LIGHT, 14),
    ("G", "Notes",             LIGHT_GREY, 24),
]
for col, label, bg, width in headers:
    apply(ws[f"{col}3"],
          value=label,
          fill=solid(bg),
          font=Font(bold=True, size=9,
                    color=WHITE if bg == NAVY_LIGHT else DARK_GREY,
                    name="Arial"),
          alignment=center(wrap=True))
    ws.column_dimensions[col].width = width
ws.row_dimensions[3].height = 44

# ── ROW 4 — Instructions ──────────────────────────────────────────────────────
instructions = [
    ("A", "Must match EXACTLY what agent types on submission — case sensitive"),
    ("B", "Report delivered to this address"),
    ("C", "State license number"),
    ("D", "YES = active, NO = deactivated — no reports generated"),
    ("E", "Must match exactly an office name in Office Profiles tab"),
    ("F", "Date agent was added"),
    ("G", "Sam only — broker cannot edit"),
]
for col, text in instructions:
    apply(ws[f"{col}4"],
          value=text,
          fill=solid(INSTRUCT_BG),
          font=Font(italic=True, size=7, color=DARK_GREY, name="Arial"),
          alignment=left_wrap())
ws.row_dimensions[4].height = 30

# ── ROW 5 — Sample Data (LIGHT_BLUE) ─────────────────────────────────────────
sample = [
    ("A", "Jeff Fleeharty"),
    ("B", "jeff@remax.com"),
    ("C", "IL License #"),
    ("D", "YES"),
    ("E", "RE/MAX Jeff Fleeharty"),
    ("F", "2026-05-05"),
    ("G", "Beta agent — rural IL market"),
]
for col, val in sample:
    apply(ws[f"{col}5"],
          value=val,
          fill=solid(LIGHT_BLUE),
          font=Font(size=9, name="Arial", color="000000"),
          alignment=Alignment(horizontal="left", vertical="center"))

# ── ROW 6 — Placeholder Row (LIGHT_GREY, italic) ─────────────────────────────
placeholder = [
    ("A", "Agent 2 Name"),
    ("B", "agent2@office.com"),
    ("C", "License #"),
    ("D", "NO"),
    ("E", "Office Name"),
    ("F", ""),
    ("G", "Add agent here"),
]
for col, val in placeholder:
    apply(ws[f"{col}6"],
          value=val if val else None,
          fill=solid(LIGHT_GREY),
          font=Font(italic=True, size=9, color=DARK_GREY, name="Arial"),
          alignment=Alignment(horizontal="left", vertical="center"))

# ── ROW 8 — Broker Instructions (YELLOW_BG) ───────────────────────────────────
ws.merge_cells(f"A8:{LAST_COL}8")
apply(ws["A8"],
      value=(
          "BROKER INSTRUCTIONS: You may add, edit, or delete agent rows in columns A-F only. "
          "Column G (Notes) is managed by Sidwell Consulting Group. "
          "To deactivate an agent set Active = NO — they will no longer receive reports. "
          "Agent Name must match exactly what the agent types on their submission form including capitalization."
      ),
      fill=solid(YELLOW_BG),
      font=Font(size=8, color=NAVY, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[8].height = 30

# ── Active Dropdown ───────────────────────────────────────────────────────────
dv = DataValidation(type="list", formula1='"YES,NO"', allow_blank=True)
dv.sqref = "D5:D2000"
ws.add_data_validation(dv)

# ── Conditional Formatting ────────────────────────────────────────────────────
DATA_RANGE = "A5:G2000"

# Active = YES → LIGHT_BLUE full row
ws.conditional_formatting.add(
    DATA_RANGE,
    FormulaRule(formula=['$D5="YES"'],
                fill=solid(LIGHT_BLUE))
)
# Active = NO → SKIP_GREY full row + italic dark grey font
ws.conditional_formatting.add(
    DATA_RANGE,
    FormulaRule(formula=['$D5="NO"'],
                fill=solid(SKIP_GREY),
                font=Font(italic=True, color=DARK_GREY, name="Arial"))
)

# ── Freeze Panes ──────────────────────────────────────────────────────────────
ws.freeze_panes = "A5"

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save("SCG_CMA_System_v7.xlsx")
print("Step 9 complete — Tab 2 Agent Roster built.")
print("  7 columns A-G, freeze at A5.")
print("  Sample row 5 (LIGHT_BLUE), placeholder row 6 (LIGHT_GREY italic).")
print("  Broker instructions row 8 (YELLOW_BG).")
print("  CF: Active=YES → LIGHT_BLUE, Active=NO → SKIP_GREY italic.")
