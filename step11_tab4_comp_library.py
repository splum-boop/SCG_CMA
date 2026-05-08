"""Step 11: Tab 4 — Comp Library (30 columns, A-AD)."""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ── Colors ────────────────────────────────────────────────────────────────────
NAVY        = "1A3A5C"
NAVY_LIGHT  = "2E5F8A"
GOLD        = "C9A84C"
INSTRUCT_BG = "EAF0F6"
LIGHT_BLUE  = "D5E8F4"
LIGHT_GREY  = "F2F2F2"
GREEN_BG    = "E2EFDA"
GREEN_TXT   = "375623"
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
ws = wb["Comp Library"]

# 30 columns: A-AD
LAST_COL = "AD"

# ── ROW 1 — Title Banner ──────────────────────────────────────────────────────
ws.merge_cells(f"A1:{LAST_COL}1")
apply(ws["A1"],
      value="SCG CMA SYSTEM — COMP LIBRARY | Auto-populated by Make.com + Apify | Zero manual entry required | v7",
      fill=solid(NAVY),
      font=hdr_font(size=12),
      alignment=center())
ws.row_dimensions[1].height = 22

# ── ROW 2 — Warning Banner ────────────────────────────────────────────────────
ws.merge_cells(f"A2:{LAST_COL}2")
apply(ws["A2"],
      value=("CRITICAL: Comps are pulled by RADIUS FROM COORDINATES — never by zip code. "
             "A comp in Brownstown (62418) or Altamont (62411) is valid for Saint Elmo (62458) if within radius. "
             "Zip code is metadata only — never used to accept or reject a comp. "
             "City exclusions are coordinate-based, not city name filtering."),
      fill=solid(RED_BG),
      font=Font(bold=True, size=8, color=RED_TXT, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[2].height = 26

# ── ROW 3 — Section Banners ───────────────────────────────────────────────────
banners = [
    ("A3", "D3",  "COMP SOURCE",                                         NAVY,      WHITE),
    ("E3", "T3",  "PROPERTY DETAILS — Zip is metadata only, never used for filtering", NAVY_LIGHT, WHITE),
    ("U3", "V3",  "LOT / ACREAGE",                                       GOLD,      "000000"),
    ("W3", "AA3", "PRICING + DOM",                                        GREEN_BG,  GREEN_TXT),
    ("AB3","AD3", "DISTANCE + QUALITY",                                   NAVY,      WHITE),
]
for start, end, label, bg, txt in banners:
    ws.merge_cells(f"{start}:{end}")
    apply(ws[start],
          value=label,
          fill=solid(bg),
          font=hdr_font(color=txt),
          alignment=center())
ws.row_dimensions[3].height = 18

# ── Column definitions ────────────────────────────────────────────────────────
# (col, header, instruction, header_bg, hdr_txt, width, num_fmt)
columns = [
    # COMP SOURCE
    ("A",  "Comp ID",               "Auto",                                               NAVY,       WHITE,     10,  None),
    ("B",  "Source Report ID",      "Links to Property Submissions",                      NAVY,       WHITE,     14,  None),
    ("C",  "Pull Date",             "Date Apify pulled this comp",                        NAVY,       WHITE,     12,  None),
    ("D",  "Source",                "Realtor.com or Redfin",                              NAVY,       WHITE,     12,  None),
    # PROPERTY DETAILS
    ("E",  "Address / URL",         "Comp property URL from Apify",                       NAVY_LIGHT, WHITE,     30,  None),
    ("F",  "City",                  "Metadata only — not used for filtering",             NAVY_LIGHT, WHITE,     14,  None),
    ("G",  "County",                "Metadata only",                                      NAVY_LIGHT, WHITE,     14,  None),
    ("H",  "Zip",                   "Metadata only — NEVER used for comp acceptance",     NAVY_LIGHT, WHITE,     10,  None),
    ("I",  "Property Type",         "Type",                                               NAVY_LIGHT, WHITE,     18,  None),
    ("J",  "Style",                 "Ranch, Colonial, Split Level, etc.",                 NAVY_LIGHT, WHITE,     14,  None),
    ("K",  "Beds",                  "#",                                                  NAVY_LIGHT, WHITE,      6,  None),
    ("L",  "Baths",                 "#",                                                  NAVY_LIGHT, WHITE,      6,  None),
    ("M",  "Sqft",                  "Living area",                                        NAVY_LIGHT, WHITE,      9,  None),
    ("N",  "Year Built",            "Year",                                               NAVY_LIGHT, WHITE,      9,  None),
    ("O",  "Condition (if known)",  "Excellent / Good / Fair / Poor",                     NAVY_LIGHT, WHITE,     12,  None),
    ("P",  "Basement",              "Yes / No",                                           NAVY_LIGHT, WHITE,      8,  None),
    ("Q",  "Basement Type",         "Type if Yes",                                        NAVY_LIGHT, WHITE,     14,  None),
    ("R",  "Garage",                "Type and spaces",                                    NAVY_LIGHT, WHITE,     14,  None),
    ("S",  "Stories",               "Single / Multi / Split",                             NAVY_LIGHT, WHITE,     10,  None),
    ("T",  "Central Air",           "Yes / No",                                           NAVY_LIGHT, WHITE,      8,  None),
    # LOT/ACREAGE
    ("U",  "Lot Sq Ft",             "Lot size sq ft",                                     GOLD,       "000000",  10,  None),
    ("V",  "Lot Acres",             "Lot size acres",                                     GOLD,       "000000",  10,  None),
    # PRICING + DOM
    ("W",  "Original List Price",   "First list price — never overwritten",               GREEN_BG,   GREEN_TXT, 14,  "$#,##0"),
    ("X",  "Final List Price",      "Price at time of sale",                              GREEN_BG,   GREEN_TXT, 14,  "$#,##0"),
    ("Y",  "Price Reduction Count", "Number of reductions before sale",                  GREEN_BG,   GREEN_TXT, 12,  None),
    ("Z",  "Sold Price",            "Final sold price",                                   GREEN_BG,   GREEN_TXT, 13,  "$#,##0"),
    ("AA", "List-to-Sale %",        "Auto-formula — Sold / Final List Price",             GREEN_BG,   GREEN_TXT, 13,  "0.0%"),
    # DISTANCE + QUALITY
    ("AB", "Days on Market",        "Total DOM",                                          NAVY,       WHITE,     12,  None),
    ("AC", "Distance (mi)",         "Miles from subject property — radius based",         NAVY,       WHITE,     14,  None),
    ("AD", "Comp Quality Score",    "Claude assigns: A=Strong B=Good C=Acceptable D=Weak context only", NAVY, WHITE, 16, None),
]

# ── ROW 4 — Headers ───────────────────────────────────────────────────────────
for col, label, _, bg, txt, width, _ in columns:
    apply(ws[f"{col}4"],
          value=label,
          fill=solid(bg),
          font=hdr_font(color=txt),
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

# ── Number formats on data rows ───────────────────────────────────────────────
for col, _, _, _, _, _, num_fmt in columns:
    if num_fmt:
        for row in range(6, 2001):
            ws[f"{col}{row}"].number_format = num_fmt

# ── ROW 6 — Sample Data ───────────────────────────────────────────────────────
sample_fills = {
    "A": LIGHT_BLUE, "B": LIGHT_BLUE, "C": LIGHT_BLUE, "D": LIGHT_BLUE,
    "E": LIGHT_GREY, "F": LIGHT_GREY, "G": LIGHT_GREY, "H": LIGHT_GREY,
    "I": LIGHT_GREY, "J": LIGHT_GREY, "K": LIGHT_GREY, "L": LIGHT_GREY,
    "M": LIGHT_GREY, "N": LIGHT_GREY, "O": LIGHT_GREY, "P": LIGHT_GREY,
    "Q": LIGHT_GREY, "R": LIGHT_GREY, "S": LIGHT_GREY, "T": LIGHT_GREY,
    "U": LIGHT_GREY, "V": LIGHT_GREY,
    "W": GREEN_BG,   "X": GREEN_BG,   "Y": GREEN_BG,   "Z": GREEN_BG,
    "AA": GREEN_BG,  "AB": GREEN_BG,
    "AC": WHITE,     "AD": WHITE,
}
sample_values = [
    ("A",  "COMP-001"),
    ("B",  "CMA-001"),
    ("C",  "2026-05-07"),
    ("D",  "Realtor.com"),
    ("E",  "https://realtor.com/sample"),
    ("F",  "Vandalia"),
    ("G",  "Fayette"),
    ("H",  "62471"),
    ("I",  "Single Family - Stick Built"),
    ("J",  "Ranch"),
    ("K",  3),
    ("L",  2),
    ("M",  1380),
    ("N",  1995),
    ("O",  "Good"),
    ("P",  "Yes"),
    ("Q",  "Unfinished"),
    ("R",  "2 Car Attached"),
    ("S",  "Single"),
    ("T",  "Yes"),
    ("U",  9000),
    ("V",  0.207),
    ("W",  185000),
    ("X",  179000),
    ("Y",  0),
    ("Z",  179000),
    ("AA", '=IFERROR(Z6/X6,"")'),   # formula
    ("AB", 32),
    ("AC", 12.4),
    ("AD", "B"),
]
for col, val in sample_values:
    c = ws[f"{col}6"]
    c.value = val
    c.fill  = solid(sample_fills.get(col, WHITE))
    c.font  = Font(size=9, color="000000", name="Arial")
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border = thin_border()
    # Apply number format
    for ccol, _, _, _, _, _, num_fmt in columns:
        if ccol == col and num_fmt:
            c.number_format = num_fmt

# AA6 gets 0.0% format
ws["AA6"].number_format = "0.0%"

# ── ROW 8 — Comp Quality Note ─────────────────────────────────────────────────
ws.merge_cells(f"A8:{LAST_COL}8")
apply(ws["A8"],
      value=("COMP QUALITY:  A = Strong match — same type, similar sqft, similar acreage, "
             "within starting radius, sold < 90 days  |  "
             "B = Good match — minor differences in 1-2 variables  |  "
             "C = Acceptable — used when stronger comps unavailable, flagged in report  |  "
             "D = Weak — context only, not used in pricing calculation"),
      fill=solid(LIGHT_BLUE),
      font=Font(italic=True, size=8, color=NAVY, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[8].height = 30

# ── Freeze Panes ──────────────────────────────────────────────────────────────
ws.freeze_panes = "A6"

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save("SCG_CMA_System_v7.xlsx")
print("Step 11 complete — Tab 4 Comp Library built.")
print("  30 columns A-AD, freeze at A6.")
print("  5 section banners row 3.")
print("  Sample row 6 with mixed section background colors.")
print("  List-to-Sale % formula seeded in AA6.")
print("  Comp quality note row 8 (LIGHT_BLUE italic).")
