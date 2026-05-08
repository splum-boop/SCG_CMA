"""Step 13: Tab 6 — Velocity Metrics."""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ── Colors ────────────────────────────────────────────────────────────────────
NAVY        = "1A3A5C"
NAVY_LIGHT  = "2E5F8A"
TEAL        = "1F6B75"
TEAL_LIGHT  = "D6EEF1"
PURPLE_LIGHT= "E8E0F5"
PURPLE      = "4B2E83"
LIGHT_BLUE  = "D5E8F4"
LIGHT_GREY  = "F2F2F2"
GREEN_BG    = "E2EFDA"
YELLOW_BG   = "FFF2CC"
ORANGE_BG   = "FCE9D6"
RED_BG      = "FCE4D6"
RED_TXT     = "9C0006"
MED_GREY    = "CCCCCC"
DARK_GREY   = "888888"
WHITE       = "FFFFFF"

PLACEHOLDER = "Auto"

def solid(hex): return PatternFill("solid", fgColor=hex)
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
def pf(): return Font(italic=True, size=8, color=DARK_GREY, name="Arial")

def section_banner(ws, start, end, text, bg, txt=WHITE, height=18):
    ws.merge_cells(f"{start}:{end}")
    apply(ws[start], value=text, fill=solid(bg), font=hf(color=txt), alignment=left_wrap())
    row = int(''.join(c for c in start if c.isdigit()))
    ws.row_dimensions[row].height = height

def write_headers(ws, row_num, cols, headers, bg, txt=WHITE, width=20):
    for col, lbl in zip(cols, headers):
        apply(ws[f"{col}{row_num}"], value=lbl, fill=solid(bg),
              font=hf(color=txt), alignment=center(wrap=True))
        ws.column_dimensions[col].width = width
    ws.row_dimensions[row_num].height = 36

# ── Load workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook("SCG_CMA_System_v7.xlsx")
ws = wb["Velocity Metrics"]

COLS8  = ["A","B","C","D","E","F","G","H"]
COLS11 = ["A","B","C","D","E","F","G","H","I","J","K"]
LAST   = "K"

# ── ROW 1 — Title ─────────────────────────────────────────────────────────────
ws.merge_cells(f"A1:{LAST}1")
apply(ws["A1"],
      value="SCG CMA SYSTEM — VELOCITY METRICS | Auto-populated by Make.com | Zero agent input required | v7",
      fill=solid(NAVY), font=hf(size=12), alignment=center())
ws.row_dimensions[1].height = 22

# ── ROW 2 — How It Works ──────────────────────────────────────────────────────
ws.merge_cells(f"A2:{LAST}2")
apply(ws["A2"],
      value=("HOW THIS WORKS: This tab is written entirely by Make.com — no agent input required. "
             "LOCAL SIGNAL: Calculated from Comp Library list-to-sale ratio trend and DOM trend from validated transactions. "
             "PENDING SIGNAL: Contract prices entered by agents when marking Pending provide leading indicators without waiting for closing. "
             "NATIONAL SIGNAL: Make.com pulls FRED API daily. "
             "2008 LESSON: National signal warns rural IL markets 60-90 days before local comps reflect the shift."),
      fill=solid(LIGHT_BLUE), font=Font(italic=True, size=8, color=NAVY, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[2].height = 36

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Local Market Signal
# ═══════════════════════════════════════════════════════════════════════════════
section_banner(ws, "A3", f"{LAST}3",
               "LOCAL MARKET SIGNAL — From Comp Library and Pending submissions", TEAL)

s1_hdrs = ["Market Area","Transactions 30d","Median Sold 30d","Median Sold 60d",
           "Median Sold 90d","Price Trend","Avg List-to-Sale %","Avg DOM",
           "Pending Contracts","Avg Contract Price","Signal"]
write_headers(ws, 4, COLS11, s1_hdrs, TEAL)

markets = ["Fayette County In-Town","Fayette County Rural","Marion County In-Town",
           "Marion County Rural","Effingham County Rural",
           "St. Clair County Suburban (Metro East)","Madison County Suburban (Metro East)"]
for i, mkt in enumerate(markets, start=5):
    apply(ws[f"A{i}"], value=mkt, fill=solid(TEAL_LIGHT),
          font=Font(bold=True, size=9, color=TEAL, name="Arial"), alignment=left_wrap())
    for col in COLS11[1:]:
        apply(ws[f"{col}{i}"], value=PLACEHOLDER, fill=solid(WHITE),
              font=pf(), alignment=center())
    ws.row_dimensions[i].height = 16

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Signal Matrix
# ═══════════════════════════════════════════════════════════════════════════════
section_banner(ws, "A13", f"{LAST}13",
               "SIGNAL MATRIX — Claude reads Combined Signal column before every report", NAVY_LIGHT)

s2_cols  = ["A","B","C","D","E","F"]
s2_hdrs  = ["List-to-Sale","DOM Trend","Pending Activity","Combined Signal","Claude Price Weighting","Notes"]
write_headers(ws, 14, s2_cols, s2_hdrs, NAVY_LIGHT, width=22)

signal_rows = [
    (15, GREEN_BG,   ["Above list","Decreasing","Multiple contracts","Hot — accelerating","Weight HIGH end +3-5%","Sellers market — multiple offers"]),
    (16, LIGHT_GREY, ["At list","Stable","Normal","Balanced","Midpoint pricing","Normal market conditions"]),
    (17, YELLOW_BG,  ["At list","Increasing","Slowing","Cooling","Weight LOW end","Buyer hesitation increasing"]),
    (18, ORANGE_BG,  ["Below list","Increasing","Low","Soft","Weight LOW end -3-5%","Price reductions common"]),
    (19, RED_BG,     ["Below list","Decreasing","None","Mixed — unusual","Flag for agent review","Unusual pattern — agent judgment needed"]),
]
for row_num, bg, vals in signal_rows:
    for col, val in zip(s2_cols, vals):
        apply(ws[f"{col}{row_num}"], value=val, fill=solid(bg),
              font=Font(size=9, color="000000", name="Arial"), alignment=left_wrap())
    ws.row_dimensions[row_num].height = 18

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — Pending Contract Tracker
# ═══════════════════════════════════════════════════════════════════════════════
section_banner(ws, "A21", f"{LAST}21",
               "PENDING CONTRACT TRACKER — Agent enters contract price when marking Submitted to Pending — leading indicator only",
               ORANGE_BG, txt="000000")

ws.merge_cells(f"A22:{LAST}22")
apply(ws["A22"],
      value=("PENDING DATA IS NEVER USED AS A CLOSED COMP. Leading indicator only. "
             "A deal can fall through. Claude notes pending activity as market context — never as pricing evidence."),
      fill=solid(RED_BG),
      font=Font(bold=True, size=8, color=RED_TXT, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[22].height = 26

s3_hdrs = ["Property Address","Original List Price","Contract Price","Contract Date",
           "List-to-Contract %","Days to Contract","Status","Notes"]
write_headers(ws, 23, COLS8, s3_hdrs, NAVY_LIGHT, width=20)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — Price Reduction Intelligence
# ═══════════════════════════════════════════════════════════════════════════════
section_banner(ws, "A25", f"{LAST}25",
               "PRICE REDUCTION INTELLIGENCE — Auto-calculated from Original vs Current List Price", NAVY)

ws.merge_cells(f"A26:{LAST}26")
apply(ws["A26"],
      value=("Original list price is captured the first time Apify detects a listing and is never overwritten. "
             "Current list price updates on every monitoring check. "
             "The difference reveals market resistance patterns over time."),
      fill=solid(LIGHT_BLUE),
      font=Font(italic=True, size=8, color=NAVY, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[26].height = 26

s4_hdrs = ["Market Area","Avg Original List","Avg Final List","Avg Reduction %",
           "Avg Reductions Per Property","Avg Days Before First Reduction","Signal"]
s4_cols  = ["A","B","C","D","E","F","G"]
write_headers(ws, 27, s4_cols, s4_hdrs, NAVY_LIGHT, width=20)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — National Leading Indicator
# ═══════════════════════════════════════════════════════════════════════════════
section_banner(ws, "A29", f"{LAST}29",
               "NATIONAL LEADING INDICATOR — Make.com pulls FRED API daily — Rural IL lags national by 60-90 days",
               NAVY)

s5_hdrs = ["Metric","Current","Prior Month","Change","Direction","Signal","Last Updated","FRED Series ID"]
write_headers(ws, 30, COLS8, s5_hdrs, NAVY, width=20)

fred_rows = [
    ("National Median Sale Price",   "MSPUS"),
    ("National HPI",                 "USSTHPI"),
    ("30-Year Mortgage Rate",        "MORTGAGE30US"),
    ("National Active Inventory",    "ACTLISCOUUS"),
    ("National Months of Supply",    "MSACSR"),
    ("Midwest HPI",                  "MWXRSA"),
    ("National Median DOM",          "Redfin"),
    ("National List-to-Sale Ratio",  "Redfin"),
]
for i, (metric, series_id) in enumerate(fred_rows, start=31):
    apply(ws[f"A{i}"], value=metric, fill=solid(LIGHT_GREY),
          font=Font(bold=True, size=9, color="000000", name="Arial"), alignment=left_wrap())
    for col in ["B","C","D","E","F","G"]:
        apply(ws[f"{col}{i}"], value="Auto", fill=solid(WHITE), font=pf(), alignment=center())
    apply(ws[f"H{i}"], value=series_id, fill=solid(LIGHT_BLUE),
          font=Font(size=9, color=NAVY, name="Arial"), alignment=center())
    ws.row_dimensions[i].height = 16

# ── ROW 40 — National Signal Summary ─────────────────────────────────────────
s5_summary = ["Overall Signal","Mortgage Direction","Inventory Direction","Price Direction",
              "Rural IL Lag Warning","Leading Indicator Alert","Last Updated","Notes"]
for col, lbl in zip(COLS8, s5_summary):
    apply(ws[f"{col}40"], value=lbl, fill=solid(PURPLE_LIGHT),
          font=Font(bold=True, size=9, color=PURPLE, name="Arial"), alignment=center(wrap=True))
ws.row_dimensions[40].height = 30

# ── ROW 42 — 2008 Warning ─────────────────────────────────────────────────────
ws.merge_cells(f"A42:{LAST}42")
apply(ws["A42"],
      value=("2008 LESSON BAKED IN: When FRED shows accelerating deterioration — rising inventory, "
             "falling HPI, rising mortgage rates — Claude applies forward-looking caution flag "
             "EVEN IF local comps have not yet reflected the shift. "
             "Rural IL historically lags national trends by 60-90 days. "
             "Claude recommends pricing toward conservative end and advising seller of potential near-term market shift."),
      fill=solid(RED_BG),
      font=Font(bold=True, size=8, color=RED_TXT, name="Arial"),
      alignment=left_wrap())
ws.row_dimensions[42].height = 40

# ── Save ──────────────────────────────────────────────────────────────────────
wb.save("SCG_CMA_System_v7.xlsx")
print("Step 13 complete — Tab 6 Velocity Metrics built.")
print("  Section 1: Local Market Signal (7 markets, rows 5-11)")
print("  Section 2: Signal Matrix (5 color-coded rows 15-19)")
print("  Section 3: Pending Contract Tracker (headers row 23)")
print("  Section 4: Price Reduction Intelligence (headers row 27)")
print("  Section 5: National FRED metrics (8 rows 31-38) + summary row 40 + 2008 warning row 42")
