"""Step 10: Tab 3 — Office Profiles."""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# ── Colors ────────────────────────────────────────────────────────────────────
NAVY        = "1A3A5C"
NAVY_LIGHT  = "2E5F8A"
GOLD        = "C9A84C"
TEAL        = "1F6B75"
INSTRUCT_BG = "EAF0F6"
LIGHT_BLUE  = "D5E8F4"
ORANGE_BG   = "FCE9D6"
RED_BG      = "FCE4D6"
RED_TXT     = "9C0006"
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

def hdr_font(color=WHITE, size=9, bold=True):
    return Font(bold=bold, size=size, color=color, name="Arial")

def instr_font():
    return Font(italic=True, size=7, color=DARK_GREY, name="Arial")

# ── Load workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook("SCG_CMA_System_v7.xlsx")
ws = wb["Office Profiles"]

LAST_COL = "V"   # 22 columns A-V

# ── ROW 1 — Title Banner ──────────────────────────────────────────────────────
ws.merge_cells(f"A1:{LAST_COL}1")
apply(ws["A1"],
      value="SCG CMA SYSTEM — OFFICE PROFILES | Sam manages ALL fields | Fully locked — contact SCG to make changes | v7",
      fill=solid(NAVY),
      font=hdr_font(size=12),
      alignment=center())
ws.row_dimensions[1].height = 22

# ── ROW 2 — Warning Banner ────────────────────────────────────────────────────
ws.merge_cells(f"A2:{LAST_COL}2")
apply(ws["A2"],
      value="WARNING: THIS TAB IS MANAGED EXCLUSIVELY BY SIDWELL CONSULTING GROUP — DO NOT EDIT — Contact splum@sidwellconsulting.com for any profile changes",
      fill=solid(RED_BG),
      font=Font(bold=True, size=9, color=RED_TXT, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[2].height = 22

# ── ROW 3 — Section Banner ────────────────────────────────────────────────────
ws.merge_cells(f"A3:{LAST_COL}3")
apply(ws["A3"],
      value="OFFICE / GROUP PROFILES — One row per office | All fields configured by SCG during onboarding",
      fill=solid(NAVY),
      font=hdr_font(),
      alignment=left_wrap())
ws.row_dimensions[3].height = 18

# ── Column definitions ────────────────────────────────────────────────────────
# (col, header, instruction, bg, width, num_fmt, dv_formula)
columns = [
    ("A", "Office/Group Name",
     "Unique name — must match Agent Roster Office column exactly",
     NAVY, 28, None, None),
    ("B", "State",
     "State where office operates",
     NAVY, 12, None,
     '"AL,AK,AZ,AR,CA,CO,CT,DE,FL,GA,HI,ID,IL,IN,IA,KS,KY,LA,ME,MD,MA,MI,MN,MS,MO,MT,NE,NV,NH,NJ,NM,NY,NC,ND,OH,OK,OR,PA,RI,SC,SD,TN,TX,UT,VT,VA,WA,WV,WI,WY,DC"'),
    ("C", "Market Type",
     "Market classification for radius logic",
     NAVY_LIGHT, 16, None,
     '"Metro/Urban,Suburban,Small Town,Rural,Extreme Rural"'),
    ("D", "Primary Counties",
     "Standard counties this office works — comma separated",
     NAVY_LIGHT, 36, None, None),
    ("E", "Standard City Exclusions",
     "Cities excluded from all comp searches for this office",
     NAVY_LIGHT, 36, None, None),
    ("F", "Default Comp Age (days)",
     "90 standard",
     NAVY_LIGHT, 14, None, None),
    ("G", "Min Comps Required",
     "3 standard",
     NAVY_LIGHT, 10, None, None),
    ("H", "Attached Garage Value ($)",
     "Dollar value adjustment for attached garage in this market",
     GOLD, 16, "$#,##0", None),
    ("I", "Finished Basement Value ($)",
     "Dollar value for fully finished basement",
     GOLD, 16, "$#,##0", None),
    ("J", "Per Sqft Adjustment Rate ($/sqft)",
     "Rate for sqft difference adjustments",
     GOLD, 18, None, None),
    ("K", "Bedroom Adjustment Value ($)",
     "Dollar value per bedroom difference",
     GOLD, 16, "$#,##0", None),
    ("L", "Central Air Adjustment ($)",
     "Downward adjustment when lacking central air",
     GOLD, 16, "$#,##0", None),
    ("M", "Max Search Radius (mi)",
     "Maximum radius before LOW confidence triggered",
     NAVY_LIGHT, 14, None, None),
    ("N", "Property Types Handled",
     "Types this office prices",
     NAVY_LIGHT, 28, None, None),
    ("O", "Never Cross-Comp",
     "Types never comped against each other",
     NAVY_LIGHT, 24, None, None),
    ("P", "Foreclosures in Comp Pool",
     "How to handle distressed sales",
     NAVY_LIGHT, 16, None,
     '"Exclude,Flag Only,Include"'),
    ("Q", "Primary Value Driver",
     "What drives value in this market",
     NAVY_LIGHT, 16, None,
     '"Structure First,Land First,Equal"'),
    ("R", "Local Market Rules",
     "Free text — Claude reads this directly before every report",
     TEAL, 40, None, None),
    ("S", "Local Market Quirks",
     "Anything unique about this market Claude should know",
     TEAL, 40, None, None),
    ("T", "Active",
     "YES = active office, NO = deactivated",
     NAVY, 8, None,
     '"YES,NO"'),
    ("U", "Setup Date",
     "Date SCG configured this profile",
     NAVY, 14, None, None),
    ("V", "Last Updated",
     "Date SCG last modified this profile",
     NAVY, 14, None, None),
]

# ── ROW 4 — Headers ───────────────────────────────────────────────────────────
for col, label, _, bg, width, _, _ in columns:
    txt_color = "000000" if bg == GOLD else WHITE
    apply(ws[f"{col}4"],
          value=label,
          fill=solid(bg),
          font=hdr_font(color=txt_color),
          alignment=center(wrap=True))
    ws.column_dimensions[col].width = width
ws.row_dimensions[4].height = 44

# ── ROW 5 — Instructions ──────────────────────────────────────────────────────
for col, _, instruction, _, _, _, _ in columns:
    apply(ws[f"{col}5"],
          value=instruction,
          fill=solid(INSTRUCT_BG),
          font=instr_font(),
          alignment=left_wrap())
ws.row_dimensions[5].height = 30

# ── Dropdowns ─────────────────────────────────────────────────────────────────
for col, _, _, _, _, _, dv_formula in columns:
    if dv_formula:
        dv = DataValidation(type="list", formula1=dv_formula, allow_blank=True)
        dv.sqref = f"{col}6:{col}500"
        ws.add_data_validation(dv)

# ── Dollar format on data rows ─────────────────────────────────────────────────
for col, _, _, _, _, num_fmt, _ in columns:
    if num_fmt:
        for row in range(6, 501):
            ws[f"{col}{row}"].number_format = num_fmt

# ── ROW 6 — Sample Data (LIGHT_BLUE) ─────────────────────────────────────────
sample = [
    ("A", "RE/MAX Jeff Fleeharty"),
    ("B", "IL"),
    ("C", "Small Town"),
    ("D", "Fayette, Marion, Shelby, Effingham (Rural), Bond, Montgomery, Clay"),
    ("E", "City of Effingham, Centralia, Mount Vernon, Mattoon, Charleston, Decatur"),
    ("F", 90),
    ("G", 3),
    ("H", 5000),
    ("I", 0),
    ("J", 45),
    ("K", 5000),
    ("L", 3000),
    ("M", 35),
    ("N", "Single Family - Stick Built, Single Family - Manufactured, Multi-Family"),
    ("O", "Manufactured never against Stick-Built"),
    ("P", "Exclude"),
    ("Q", "Land First"),
    ("R", "Land value assessed FIRST. Structure secondary. Crappy house on 20ac outvalues nice house on 1ac. "
          "Age does NOT drive value — condition and updates override. Split level common — minimal adjustment. "
          "Comp by coordinates + radius NEVER by zip code."),
    ("S", "Rural IL market — thin comps common. Radius expansion to county level expected. "
          "Brownstown and Altamont are valid Saint Elmo comps despite different zip codes."),
    ("T", "YES"),
    ("U", "2026-05-05"),
    ("V", "2026-05-05"),
]
for col, val in sample:
    c = ws[f"{col}6"]
    c.value = val
    c.fill = solid(LIGHT_BLUE)
    c.font = Font(size=9, color="000000", name="Arial")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = thin_border()
    # Apply dollar format to financial cols
    for ccol, _, _, _, _, num_fmt, _ in columns:
        if ccol == col and num_fmt:
            c.number_format = num_fmt

ws.row_dimensions[6].height = 60

# ── ROW 9 — Radius Expansion Section Banner ───────────────────────────────────
ws.merge_cells(f"A9:{LAST_COL}9")
apply(ws["A9"],
      value="RADIUS EXPANSION LOGIC — Applied automatically by Make.com based on Market Type",
      fill=solid(NAVY_LIGHT),
      font=hdr_font(),
      alignment=left_wrap())
ws.row_dimensions[9].height = 18

# ── ROW 10 — Radius Table Headers ────────────────────────────────────────────
radius_headers = ["Market Type", "Start", "Expand 1", "Expand 2",
                  "Expand 3", "Max", "Stop Threshold", "Notes"]
radius_cols = ["A", "B", "C", "D", "E", "F", "G", "H"]
for col, label in zip(radius_cols, radius_headers):
    apply(ws[f"{col}10"],
          value=label,
          fill=solid(NAVY_LIGHT),
          font=hdr_font(),
          alignment=center(wrap=True))
ws.row_dimensions[10].height = 30

# ── ROWS 11-14 — Radius Data ──────────────────────────────────────────────────
radius_data = [
    ["Metro/Urban",   "1 mile",   "2 miles",  "5 miles",  "10 miles", "15 miles",  "5+ comps", "Dense transaction volume — tight radius first"],
    ["Small Town",    "2 miles*", "5 miles",  "10 miles", "20 miles", "35 miles",  "5+ comps", "Agent can select 1 mile override for dense town centers"],
    ["Rural",         "15 miles", "25 miles", "35 miles", "50 miles", "75 miles",  "3+ comps", "Unique/large properties may start at 35+ miles"],
    ["Extreme Rural", "25 miles", "50 miles", "75 miles", "100 miles","150 miles", "2+ comps", "Very thin markets — LOW confidence expected"],
]
row_colors = [LIGHT_BLUE, "FFFFFF", LIGHT_BLUE, "FFFFFF"]
for i, (row_data, bg) in enumerate(zip(radius_data, row_colors), start=11):
    for col, val in zip(radius_cols, row_data):
        apply(ws[f"{col}{i}"],
              value=val,
              fill=solid(bg),
              font=Font(size=9, color="000000", name="Arial"),
              alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))
    ws.row_dimensions[i].height = 18

# ── ROW 15 — Override Note ────────────────────────────────────────────────────
ws.merge_cells(f"A15:{LAST_COL}15")
apply(ws["A15"],
      value=(
          "AGENT OVERRIDE: Agent can select any starting radius on submission including 1 mile for dense areas "
          "like Effingham city proper. Let System Decide uses Market Type default above. Expansion logic always "
          "applies from whatever starting radius is used. Zip codes are NEVER used to filter comps — "
          "radius from coordinates only."
      ),
      fill=solid(ORANGE_BG),
      font=Font(italic=True, size=8, color="000000", name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[15].height = 30

# ── Freeze Panes ──────────────────────────────────────────────────────────────
ws.freeze_panes = "A6"

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save("SCG_CMA_System_v7.xlsx")
print("Step 10 complete — Tab 3 Office Profiles built.")
print("  22 columns A-V, freeze at A6.")
print("  Title/warning/section banners rows 1-3.")
print("  Headers row 4, instructions row 5, sample data row 6.")
print("  Radius expansion table rows 9-15.")
print("  Dropdowns: State, Market Type, Foreclosures, Primary Value Driver, Active.")
