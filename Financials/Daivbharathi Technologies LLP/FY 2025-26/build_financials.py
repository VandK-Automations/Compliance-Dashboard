#!/usr/bin/env python3
"""
Daivbharathi Technologies LLP - Financial Statements FY 2025-26
Rebuilt in the exact presentation format of the signed FY 2024-25 financial statements.

All figures are taken verbatim from "Daivbharathi_Techn_LLP_Financials_FY_2025-26_Final.pdf"
(current year) and from the signed FY 2024-25 statements (previous-year comparatives).
No figure is computed, rounded or altered by this script; totals are written as provided
and cross-checked in verify() below.
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.page import PageMargins
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

FONT = "Century Gothic"
CY = "March 31, 2026"          # current year column caption
PY = "March 31, 2025"          # previous year column caption
SHADE = "D9D5E9"              # header-row shading, sampled from the reference scan
GREY = SHADE                  # kept as an alias: existing call sites pass fill=GREY
HIGHLIGHT = False              # draft highlighting of open points - off in the final statements
HL = "FFFF99"                  # light-yellow used when HIGHLIGHT is on

# Indian grouping with accounting-style padding, negatives in brackets (values >= 1 lakh get the lakh comma)
NUM = r'[>=100000]_(* ##\,##\,##0_);[<=-100000]_(* \(##\,##\,##0\);_(* #,##0_)'
# nil cells: shown as "-" with the accounting offset, exactly as in the reference
NIL = r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_)'

THIN = Side(style="thin", color="000000")
MED = Side(style="medium", color="000000")


def wu(inches):
    """Excel column-width units for a column of the given printed width (Normal style Calibri 11)."""
    px = inches * 96.0
    return round((px - 5) / 7.0, 2)


class Page:
    """One printed page = one worksheet, with widths/margins measured from the reference scan."""

    def __init__(self, wb, name, widths_in, margins, landscape=False):
        self.ws = wb.create_sheet(name)
        self.widths = widths_in
        for i, w in enumerate(widths_in, start=1):
            self.ws.column_dimensions[get_column_letter(i)].width = wu(w)
        self.merges = []
        self.borders = {}
        self.heights = {}
        self.landscape = landscape
        l, r, t, b = margins
        self.ws.page_margins = PageMargins(left=l, right=r, top=t, bottom=b, header=0.2, footer=0.2)
        self.ws.sheet_view.showGridLines = False
        self.ws.sheet_view.zoomScale = 150
        self.maxrow = 1
        self.maxcol = len(widths_in)

    # ---- content -------------------------------------------------------------------------------
    def cell(self, r, c, v=None, size=7.5, bold=False, italic=False, h="left", v_="center",
             wrap=False, indent=0, fill=None, span=1, fmt=None, hl=False, underline=None, rows=1):
        ws = self.ws
        cell = ws.cell(row=r, column=c)
        if v is not None:
            cell.value = v
        cell.font = Font(name=FONT, size=size, bold=bold, italic=italic, underline=underline)
        cell.alignment = Alignment(horizontal=h, vertical=v_, wrap_text=wrap, indent=indent)
        if hl and HIGHLIGHT:
            cell.fill = PatternFill("solid", fgColor=HL)
        elif fill:
            cell.fill = PatternFill("solid", fgColor=fill)
        if fmt:
            cell.number_format = fmt
        if span > 1 or rows > 1:
            self.merges.append((r, c, r + rows - 1, c + span - 1))
        self.maxrow = max(self.maxrow, r + rows - 1)
        return cell

    def amt(self, r, c, v, bold=False, hl=False, size=7.5, span=1):
        """Amount cell. v=None -> nil shown as '-' ; otherwise the exact integer provided."""
        if v is None:
            return self.cell(r, c, 0, size=size, bold=bold, h="right", fmt=NIL, hl=hl, span=span)
        assert isinstance(v, int), v
        return self.cell(r, c, v, size=size, bold=bold, h="right", fmt=NUM, hl=hl, span=span)

    def blank(self, r, c, hl=False):
        return self.cell(r, c, None, hl=hl)

    # ---- borders ---------------------------------------------------------------------------------
    def _side(self, r, c, side, style):
        self.borders.setdefault((r, c), {})[side] = style

    def hline(self, r, c1, c2, where="bottom", style=THIN):
        for c in range(c1, c2 + 1):
            self._side(r, c, where, style)

    def vline(self, c, r1, r2, where="right", style=THIN):
        for r in range(r1, r2 + 1):
            self._side(r, c, where, style)

    def box(self, r1, c1, r2, c2, style=THIN):
        self.hline(r1, c1, c2, "top", style)
        self.hline(r2, c1, c2, "bottom", style)
        self.vline(c1, r1, r2, "left", style)
        self.vline(c2, r1, r2, "right", style)

    def grid(self, r1, c1, r2, c2, style=THIN):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                for s in ("top", "bottom", "left", "right"):
                    self._side(r, c, s, style)

    def height(self, r, pts):
        self.heights[r] = pts

    # ---- finish ------------------------------------------------------------------------------------
    def finish(self, print_rows=None, fit_width=True, default_height=9.75, row_breaks=()):
        ws = self.ws
        last_row = print_rows or self.maxrow
        for r in range(1, last_row + 1):
            ws.row_dimensions[r].height = self.heights.get(r, default_height)
        for (r1, c1, r2, c2) in self.merges:
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        for (r, c), sides in self.borders.items():
            cell = ws.cell(row=r, column=c)
            cell.border = Border(left=sides.get("left"), right=sides.get("right"),
                                 top=sides.get("top"), bottom=sides.get("bottom"))
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "landscape" if self.landscape else "portrait"
        if fit_width:
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
        ws.print_area = f"A1:{get_column_letter(self.maxcol)}{last_row}"
        for rb in row_breaks:
            ws.row_breaks.append(Break(id=rb))
        return ws


# =====================================================================================================
#  FIGURES (verbatim)
# =====================================================================================================
# FY 2025-26 (source: Daivbharathi_Techn_LLP_Financials_FY_2025-26_Final.pdf)
S = dict(
    capital=90000, reserves=-3613505, partners_funds=-3523505,
    ltb=3047564, noncurrent_total=3047564,
    tp=456392, ocl=46506, current_total=502898, bs_total=26957,
    cash=26957, cash_bank=5102, cash_hand=21837, cash_others=18,
    revenue=1115272, other_exp=3607863, loss=-2492591,
    res_opening=-1120914, res_loss=-2492591, res_closing=-3613505,
    tp_other=456392, ocl_other=46506, rev_sale=1115272,
)
EXP_CY = [  # Note 2.8 of the source, captions verbatim, in source order
    ("Accommodation and hotel expenses", 17510),
    ("Advertisement", 410685),
    ("Bank charges", 1453),
    ("Branding expenses", 19100),
    ("Commission", 28817),
    ("Conveyance", 219790),
    ("Data infrastructure expenses", 28583),
    ("Design and development expenses", 75000),
    ("GST late filing fees", 20),
    ("Interest on TDS", 2164),
    ("Internship charges", 36000),
    ("Office expenses", 77518),
    ("Pooja expenses", 1345327),
    ("Pooja items", 50000),
    ("Postage and courier expenses", 27286),
    ("Pre-incorporation expenses", 36000),
    ("Printing charges", 44706),
    ("Professional / consultancy / technical fees", 1068252),
    ("Staff welfare", 8021),
    ("Studio expenses", 24000),
    ("Subscription expenses", 15000),
    ("Telephone expenses", 9871),
    ("Translation expenses", 52760),
    ("Website development expenses", 10000),
]
# FY 2024-25 comparatives (source: signed FY 2024-25 statements)
P = dict(
    contribution=-1020912, partners_funds=-1020912, stb=780680, tp=228000, ocl=35000,
    current_total=1043680, bs_total=22768, cash=12768, oca=10000,
    other_exp=1120912, loss=-1120912,
    n3a_intro_total=100000, n3a_share_total=-1120912, n3a_close_total=-1020912,
    n3a_s_agreed=90000, n3a_s_intro=90000, n3a_s_share=-1008820, n3a_s_close=-918820,
    n3a_m_agreed=10000, n3a_m_intro=10000, n3a_m_share=-112091, n3a_m_close=-102091,
    exp_tds=780, exp_rates=2500, exp_legal=485000, exp_adv=125000, exp_bank=312, exp_bp=28000,
    exp_misc=479320,
)


def verify():
    """Arithmetic cross-check of the provided figures (raises if anything does not tie)."""
    assert S["capital"] + S["reserves"] == S["partners_funds"]
    assert S["tp"] + S["ocl"] == S["current_total"]
    assert S["partners_funds"] + S["noncurrent_total"] + S["current_total"] == S["bs_total"]
    assert S["cash"] == S["bs_total"]
    assert S["cash_bank"] + S["cash_hand"] + S["cash_others"] == S["cash"]
    assert S["revenue"] - S["other_exp"] == S["loss"]
    assert S["res_opening"] + S["res_loss"] == S["res_closing"] == S["reserves"]
    assert sum(v for _, v in EXP_CY) == S["other_exp"]
    assert P["partners_funds"] + P["stb"] + P["tp"] + P["ocl"] == P["bs_total"]
    assert P["stb"] + P["tp"] + P["ocl"] == P["current_total"]
    assert P["cash"] + P["oca"] == P["bs_total"]
    assert (P["exp_tds"] + P["exp_rates"] + P["exp_legal"] + P["exp_adv"] + P["exp_bank"]
            + P["exp_bp"] + P["exp_misc"]) == P["other_exp"]


# =====================================================================================================
#  PAGE 1 - BALANCE SHEET
# =====================================================================================================
def statement_header(pg, title, amount_caption, amount_h, name_size=8):
    """Header block inside the outer box: name / LLPIN / address / email / title / amount caption."""
    W = pg.maxcol
    pg.height(1, 6)
    pg.cell(2, 1, "Daivbharathi Technologies LLP", size=name_size, bold=True, h="center", span=W)
    pg.cell(3, 1, "LLPIN: ACL-2422", h="center", span=W)
    pg.cell(4, 1, "63, 1st Floor, 1st Main ,7th Block , Koramangala, Koramangala VI Bk, Bangalore, "
                  "Bangalore South, Karnataka, India, 560095", h="center", span=W)
    pg.cell(5, 1, "email: samartharaghava@gmail.com;", h="center", span=W)
    pg.cell(7, 1, title, size=8, bold=True, h="center", span=W)
    pg.cell(8, W, amount_caption, h=amount_h)
    return 9  # header row


def signature_block(pg, r, col_left, col_right, size=7.5):
    pg.cell(r, col_left, "For and on behalf of the Board", size=size)
    pg.cell(r + 1, col_left, "M/s. Daivbharathi Technologies LLP", size=size, bold=True)
    r2 = r + 6
    pg.cell(r2, col_left, "Samartha Raghava Nagabhushanam I", size=size)
    pg.cell(r2, col_right, "Manasa Nagabhushanam", size=size)
    pg.cell(r2 + 1, col_left, "Designated Partner", size=size)
    pg.cell(r2 + 1, col_right, "Designated Partner", size=size)
    pg.cell(r2 + 2, col_left, "DIN:00329885", size=size)
    pg.cell(r2 + 2, col_right, "DIN:03566928", size=size)
    return r2 + 2


def build_balance_sheet(wb):
    pg = Page(wb, "Balance Sheet", [0.56, 1.62, 2.74, 0.64, 1.06, 1.03], (0.22, 0.35, 0.55, 0.4))
    hr = statement_header(pg, "Balance Sheet as at March 31, 2026", "Amount in Rs.", "center")
    # column header row
    pg.height(hr, 24)
    for c, t in ((1, "Sl. No"), (2, "Particulars"), (4, "Note No."), (5, f"As at\n{CY}"), (6, f"As at\n{PY}")):
        pg.cell(hr, c, t, bold=True, h="center", wrap=True, fill=GREY, span=(2 if c == 2 else 1))
    pg.grid(hr, 1, hr, 6)

    r = hr + 1
    rows = []  # (sl, particulars, note, cy, py, style)

    def line(sl=None, text=None, note=None, cy="", py="", bold=False, kind="item", indent=0):
        rows.append(dict(sl=sl, text=text, note=note, cy=cy, py=py, bold=bold, kind=kind, indent=indent))

    line(kind="spacer")
    line("I.", "EQUITY AND LIABILITIES", bold=True, kind="head")
    line(kind="spacer")
    line("1", "Partners' Funds", bold=True, kind="head")
    line("A.", "Partners' Capital Account", kind="head")
    line(None, "(a) Partners Contribution", "3A", S["capital"], P["contribution"])
    line(None, "(b) Partners Current Account", "3B", None, None)
    line("B.", "Reserves and surplus", "4", S["reserves"], None)
    line(kind="subtotal", cy=S["partners_funds"], py=P["partners_funds"])
    line(kind="spacer")
    line("3", "Non-current liabilities", bold=True, kind="head")
    line(None, "(a) Long-term borrowings", "5", S["ltb"], None)
    line(None, "(b) Deferred Tax Liabilities (Net)", None, None, None)
    line(None, "(c) Other Long term liabilities", None, None, None)
    line(None, "(d) Long term provisions", None, None, None)
    line(kind="subtotal", cy=S["noncurrent_total"], py=None)
    line(kind="spacer")
    line("4", "Current liabilities", bold=True, kind="head")
    line(None, "(a) Short-term borrowings", "5A", None, P["stb"])
    line(None, "(b) Trade payables", "6", S["tp"], P["tp"])
    line(None, "(c) Other current liabilities", "7", S["ocl"], P["ocl"])
    line(None, "(d) Short-term provisions", None, None, None)
    line(kind="subtotal", cy=S["current_total"], py=P["current_total"])
    line(kind="total", cy=S["bs_total"], py=P["bs_total"])
    line(kind="spacer")
    line("II.", "ASSETS", bold=True, kind="head")
    line("1", "Non-current assets", bold=True, kind="head")
    line(None, "(a) Property, Plant, Equipment and Intangible Assets", kind="head")
    line(None, "(i) Property, Plant and Equipment", None, None, None, indent=2)
    line(None, "(ii) Intangible assets", None, None, None, indent=2)
    line(None, "(iii) Capital work-in-progress", None, None, None, indent=2)
    line(None, "(iv) Intangible assets under development", None, None, None, indent=2)
    line(None, "(b) Non-current investments", None, None, None)
    line(None, "(c) Deferred tax assets (Net)", None, None, None)
    line(None, "(d) Long term loans and advances", None, None, None)
    line(None, "(e) Other non-current assets", None, None, None)
    line(kind="subtotal", cy=None, py=None)
    line(kind="spacer")
    line("2", "Current assets", bold=True, kind="head")
    line(None, "(a) Current investments", None, None, None)
    line(None, "(b) Inventories", None, None, None)
    line(None, "(c) Trade receivables", "8", None, None)
    line(None, "(d) Cash and Cash Equivalents", "9", S["cash"], P["cash"])
    line(None, "(e) Short-term loans and advances", None, None, None)
    line(None, "(f) Other current assets", "10", None, P["oca"])
    line(kind="subtotal", cy=S["cash"], py=P["bs_total"])
    line(kind="total", cy=S["bs_total"], py=P["bs_total"])

    first_body = r
    for it in rows:
        k = it["kind"]
        if k == "spacer":
            r += 1
            continue
        if k == "subtotal":
            pg.amt(r, 5, it["cy"])
            pg.amt(r, 6, it["py"])
            pg.hline(r, 5, 6, "top")
            pg.hline(r, 5, 6, "bottom")
            r += 1
            continue
        if k == "total":
            pg.cell(r, 1, "TOTAL", bold=True, h="center", span=3)
            pg.amt(r, 5, it["cy"], bold=True)
            pg.amt(r, 6, it["py"], bold=True)
            pg.hline(r, 1, 6, "top")
            pg.hline(r, 1, 6, "bottom")
            pg.height(r, 11.25)
            r += 1
            continue
        if it["sl"]:
            pg.cell(r, 1, it["sl"], bold=it["bold"], h="center")
        pg.cell(r, 2, it["text"], bold=it["bold"], indent=it["indent"], span=2)
        if it["note"]:
            pg.cell(r, 4, it["note"], h="center")
        if k == "item":
            pg.amt(r, 5, it["cy"])
            pg.amt(r, 6, it["py"])
        r += 1
    last_body = r - 1
    # notes reference rows inside the box
    pg.cell(r, 1, "Brief about the Firm", span=3)
    pg.cell(r, 4, "1", h="center")
    pg.cell(r + 1, 1, "Summary of significant accounting policies", span=3)
    pg.cell(r + 1, 4, "2", h="center")
    end_box = r + 1
    # vertical rules of the table body and the outer box
    pg.vline(1, hr + 1, last_body, "right")
    pg.vline(3, hr + 1, end_box, "right")
    pg.vline(4, hr + 1, end_box, "right")
    pg.vline(5, hr + 1, end_box, "right")
    pg.box(1, 1, end_box, 6)
    r = end_box + 1
    pg.cell(r, 1, "The accompanying notes form an integral part of the financial statements")
    r += 2
    last = signature_block(pg, r, 1, 3)
    pg.finish(print_rows=last + 1)


# =====================================================================================================
#  PAGE 2 - STATEMENT OF PROFIT AND LOSS
# =====================================================================================================
def build_pl(wb):
    pg = Page(wb, "Profit and Loss", [0.475, 2.0, 1.725, 0.675, 0.975, 0.875], (0.72, 0.8, 0.48, 0.4))
    hr = statement_header(pg, "Statement of Profit and Loss for the Period ended March 31, 2026",
                          "Amount in INR (Rs.)", "right")
    pg.height(hr, 36)
    for c, t in ((1, "Sl. No."), (2, "Particulars"), (4, "Note No."), (5, f"As at\n{CY}"), (6, f"As at\n{PY}")):
        pg.cell(hr, c, t, bold=True, h="center", wrap=True, fill=GREY, span=(2 if c == 2 else 1))
    pg.grid(hr, 1, hr, 6)

    r = hr + 1
    items = [
        # (sl, text, note, cy, py, bold, height, rule, two_line)
        ("spacer",),
        ("I", "Revenue from operations", "11", S["revenue"], None, False, 14.5, False),
        ("II", "Other income", None, None, None, False, 14.5, False),
        ("spacer",),
        ("III", "Total revenue (I+II)", None, S["revenue"], None, True, 11.25, True),
        ("spacer",),
        ("IV", "Expenses", None, "", "", True, 12, False),
        (None, "Cost of goods sold", None, None, None, False, 9.75, False),
        (None, "Employee benefit expense", None, None, None, False, 9.75, False),
        (None, "Finance costs", None, None, None, False, 9.75, False),
        (None, "Depreciation and amortization expense", None, None, None, False, 9.75, False),
        (None, "Other expenses", "12", S["other_exp"], P["other_exp"], False, 9.75, False),
        ("spacer",),
        (None, "Total expenses", None, S["other_exp"], P["other_exp"], True, 11.25, True),
        ("spacer",),
        ("V", "Profit / (Loss) before exceptional and extraordinary items,partners' remuneration and tax (III - IV)",
         None, S["loss"], P["loss"], True, 20, False),
        ("spacer",),
        ("VI", "Exceptional items (specify nature & provide note/delete if none)", None, None, None, False, 14.5, False),
        ("spacer",),
        ("VII", "Profit / (Loss) before extraordinary items,partners' remuneration and tax (V - VI)",
         None, S["loss"], P["loss"], True, 20, False),
        ("spacer",),
        ("VIII", "Extraordinary Items (specify nature & provide note/delete if none)", None, None, None, False, 14.5, False),
        ("spacer",),
        ("IX", "Profit / (Loss) before Partners' Remuneration tax  (VII-VIII)", None, S["loss"], P["loss"], True, 14.5, False),
        ("spacer",),
        ("X", "Partners' Remuneration", None, "", "", True, 14.5, False),
        ("spacer",),
        ("XI", "Profit / (Loss) before tax  (IX-X)", None, S["loss"], P["loss"], True, 14.5, False),
        ("spacer",),
        ("XII", "Less: Tax expense:", None, "", "", True, 12, False),
        (None, "(1) Current tax", None, None, None, False, 9.75, False, 2),
        (None, "(2) Deferred tax", None, None, None, False, 9.75, False, 2),
        ("spacer",),
        ("XIII", "Profit / (Loss) from continuing operations (XI - XII)", "A", S["loss"], P["loss"], True, 14.5, False),
        ("spacer",),
        ("XIV", "Profit / (Loss) from discontinuing operations (before tax)", None, None, None, False, 9.75, False),
        ("XV", "Add / (Less) Tax expense of discontinuing operations", None, None, None, False, 9.75, False),
        ("spacer",),
        ("XVI", "Profit / (Loss) from discontinuing operations ( after tax )(XII - XIII)", "B", None, None, True, 14.5, False),
        ("spacer",),
        ("XVII", "Profit / (Loss) for the year (XI + XIV)", "A+B", S["loss"], P["loss"], True, 11.25, True),
    ]
    for it in items:
        if it[0] == "spacer":
            pg.height(r, 9.75)
            r += 1
            continue
        sl, text, note, cy, py, bold, ht, rule = it[:8]
        indent = it[8] if len(it) > 8 else 0
        pg.height(r, ht)
        if sl:
            pg.cell(r, 1, sl, bold=True, h="center")
        pg.cell(r, 2, text, bold=bold, span=2, wrap=(ht >= 20), indent=indent)
        if note:
            pg.cell(r, 4, note, h="center")
        if cy != "":
            pg.amt(r, 5, cy, bold=bold)
            pg.amt(r, 6, py, bold=bold)
        if rule:
            pg.hline(r, 1, 6, "top")
            pg.hline(r, 1, 6, "bottom")
        r += 1
    last_body = r - 1
    pg.vline(1, hr + 1, last_body, "right")
    pg.vline(3, hr + 1, last_body, "right")
    pg.vline(4, hr + 1, last_body, "right")
    pg.vline(5, hr + 1, last_body, "right")
    pg.box(1, 1, last_body, 6)
    r = last_body + 1
    pg.cell(r, 1, "The accompanying notes form an integral part of the financial statements")
    r += 2
    last = signature_block(pg, r, 1, 3)
    pg.finish(print_rows=last + 1)


# =====================================================================================================
#  PAGES 3-4 - BRIEF ABOUT THE FIRM AND SIGNIFICANT ACCOUNTING POLICIES (Notes 1 & 2)
# =====================================================================================================
POLICY_LINE_CHARS = 96   # characters per wrapped line at 8pt across the 5.685in text block
POLICY_LINE_PT = 10.5


def para_height(text, chars=POLICY_LINE_CHARS):
    lines = 0
    for p in text.split("\n"):
        lines += max(1, -(-len(p) // chars))
    return lines * POLICY_LINE_PT + 3


def build_policies(wb):
    pg = Page(wb, "Notes 1-2 Policies", [0.54, 0.91, 0.615, 0.795, 2.565, 0.8], (0.925, 1.0, 0.75, 0.5))
    ws = pg.ws
    W = 6
    r = 1

    def header(r):
        pg.cell(r, 1, "Daivbharathi Technologies LLP", size=8.5, bold=True, span=W)
        pg.cell(r + 1, 1, "Brief about the Firm and Significant Accounting policies for the year ended 31.03.2026",
                size=8.5, bold=True, span=W)
        pg.height(r, 12)
        pg.height(r + 1, 12)
        return r + 2

    def heading(r, text, size=8.5):
        pg.cell(r, 2, text, size=size, bold=True, span=W - 1)
        pg.height(r, 12)
        return r + 1

    def para(r, text, hl=False):
        pg.cell(r, 2, text, size=8, wrap=True, v_="top", span=W - 1, hl=hl)
        pg.height(r, para_height(text))
        return r + 1

    r = header(r)
    r += 1
    r = heading(r, "1.Brief about the Firm")
    rich = CellRichText(
        TextBlock(InlineFont(rFont=FONT, sz=8), "The  "),
        TextBlock(InlineFont(rFont=FONT, sz=8, b=True),
                  "M/S Daivbharathi Technologies LLP (LLPIN: ACL-2422) incorporated as on 08th January,2025"),
        TextBlock(InlineFont(rFont=FONT, sz=8),
                  " having registered office at  63, 1st Floor, 1st Main ,7th Block , Koramangala, Koramangala VI Bk, "
                  "Bangalore, Bangalore South, Karnataka, India, 560095"),
    )
    c = pg.cell(r, 2, None, size=8, wrap=True, v_="top", span=W - 1)
    c.value = rich
    pg.height(r, 3 * POLICY_LINE_PT + 3)
    r += 2
    pg.cell(r, 2, "The profit/loss sharing ratio as per the partnership deed is as follows:", size=8, span=W - 1)
    r += 1
    # partner table: SR. NO | NAME OF PARTNER | PERCENTAGE OF / PROFIT/LOSS
    pg.cell(r, 2, "SR. NO", size=7.5, bold=True, h="left", rows=2)
    pg.cell(r, 3, "NAME OF PARTNER", size=7.5, bold=True, h="center", span=3, rows=2)
    pg.cell(r, 6, "PERCENTAGE OF", size=7, bold=True, h="center")
    pg.cell(r + 1, 6, "PROFIT/LOSS", size=7, bold=True, h="center")
    pg.height(r, 20)
    pg.height(r + 1, 18)
    pg.grid(r, 2, r + 1, 6)
    r += 2
    for i, (name, pct) in enumerate((("Samartha Raghava Nagabhushanam", "90%"), ("Manasa Nagabhushanam", "10%")), 1):
        pg.cell(r, 2, str(i), size=8)
        pg.cell(r, 3, name, size=8, span=3)
        pg.cell(r, 6, pct, size=8, bold=True, h="center")
        pg.grid(r, 2, r, 6)
        pg.height(r, 10.5)
        r += 1
    pg.cell(r, 2, "Nature of Business:", size=8, span=W - 1)
    r += 2
    r = para(r, " To design and deploy innovative technological solutions that foster the spiritual growth and "
                "well-being of individuals, enabling them to lead more meaningful and fulfilling lives.")
    r += 2
    r = para(r, " To provide technology platforms and tools that facilitate the performance, preservation, and "
                "propagation of Hindu religious, cultural, and spiritual practices in alignment with traditional "
                "principles.")
    r += 3
    r = heading(r, "2.SIGNIFICANT ACCOUNTING POLICIES")
    r += 1
    r = heading(r, "a.Basis of preparation The Financial Statements")
    r += 2
    r = para(r, "The Financial Statements have been prepared on accrual basis under historical cost convention and "
                "in accordance with the applicable accounting standards prescribed by the Institute of Chartered "
                "Accountants of India (ICAI). The accounting policies are consistently applied unless otherwise "
                "stated.")
    r += 2
    r = heading(r, "b.Use of estimates")
    r += 2
    r = para(r, "The preparation of financial statements in conformity with generally accepted accounting "
                "principles requires management to make estimates and assumptions that affect the reported amount "
                "of revenue, expenses, assets and liabilities and the disclosure of contingent liabilities at the "
                "date of the financial statements and the results of operations during the reporting period end. "
                "Although these estimates are based upon management's best knowledge of current events and "
                "actions, actual results could differ from these estimates.")
    r += 2
    r = heading(r, "c.Property, Plant & Equipment Properties,")
    r += 2
    r = para(r, "Plant & Equipment's are stated at cost less accumulated depreciation, amortization and impairment "
                "losses if any. Cost comprises the purchase price and any attributable cost of bringing the asset "
                "to its working condition for its intended use. Borrowing costs relating to acquisition of "
                "property, plant and equipment which takes substantial period of time to get ready for its "
                "intended use are also included to the extent they relate to the period till such assets are "
                "ready to be put to use.")
    r += 3
    r = para(r, "Depreciation on Property, Plant & Equipment Depreciation on property, plant and equipment is "
                "provided on straight line method over the useful lives of assets except for leasehold "
                "improvements. Leasehold improvements are amortized over a period of 10 years.")
    r += 1
    pg.cell(r, 2, "Class of Assets", size=8, bold=True, h="center", span=2)
    pg.cell(r, 4, "Useful lives", size=8, bold=True)
    pg.grid(r, 2, r, 4)
    pg.height(r, 10.5)
    r += 1
    for name, yrs in (("Office Equipment", 5), ("Computer", 3)):
        pg.cell(r, 2, name, size=8, span=2)
        pg.cell(r, 4, yrs, size=8, h="right")
        pg.grid(r, 2, r, 4)
        pg.height(r, 10.5)
        r += 1
    r += 1
    page2 = r  # ---- page break: page 4 -------------------------------------------------------------------
    r = header(r)
    r = heading(r, "d. Revenue recognition")
    r += 1
    r = para(r, "Revenue is recognised to the extent that it is probable that the economic benefits will flow to "
                "the Firm and the amount of the revenue can be reliably measured with no uncertainty as regards to "
                "ultimate collection.\nIncome from services:\nRevenue is recognised and accounted on rendering of "
                "services in accordance with the terms of arrangement by reference to the stage of completion of "
                "the contract.")
    r += 2
    r = heading(r, "e.Investments")
    r += 1
    r = para(r, "Investments that are readily realizable and intended to be held for not more than a year are "
                "classified as current investments. All other investments are classified as long-term "
                "investments. Current investments are carried at lower of cost and fair value determined on an "
                "individual investment basis. Long-term investments are carried at cost. However, provision for "
                "diminution in value is made to recognize a decline, other than temporary, in the value of the "
                "investments.")
    r += 2
    r = heading(r, "f. Employee benefits")
    r += 1
    r = para(r, "Defined contribution plans: Retirement benefits in the form of Provident Fund are a defined "
                "contribution scheme and the contributions are charged to the Statement of Profit and Loss of the "
                "year when the contributions to the fund is due. There are no other obligations other than the "
                "contribution payable to the fund. Defined benefit plans: Under Payment of Gratuity Act,1972 "
                "Gratuity liability is a defined benefit obligation and is provided for on the basis of an "
                "actuarial valuation on Projected Unit Credit Method made at the end of the financial year.  The "
                "Company records its gratuity liability based on an actuarial valuation made by an independent "
                "actuary as at year end. Compensated absences: Long term compensated absences are provided for "
                "based on actuarial valuation. The actuarial valuation is done as per Projected Unit Credit "
                "Method. All actuarial gains/losses are immediately taken to the Profit and Loss account and are "
                "not deferred.")
    r += 3
    r = heading(r, "g.Cash and cash equivalents")
    r += 1
    r = para(r, " Cash and cash equivalents in the Cash Flow Statement comprise cash at bank and in hand and short "
                "term investments with an original maturity of three months or less.")
    r += 3
    r = heading(r, "h.Borrowing costs")
    r += 1
    r = para(r, "Borrowing costs are recognized as an expense in the period in which these are incurred. Borrowing "
                "Costs that are attributable to the acquisition or construction of qualifying assets are "
                "capitalized as part of the cost of such assets.")
    r += 1
    r = heading(r, "i.Provisions")
    r += 1
    r = para(r, " A provision is recognized when an enterprise has a present obligation as a result of past event; "
                "it is probable that an outflow of resources will be required to settle the obligation, in respect "
                "of which a reliable estimate can be made. Provisions are not discounted to its present value and "
                "are determined based on best estimate required to settle the obligation at the balance sheet "
                "date. These are reviewed at each balance sheet date and adjusted to reflect the current best "
                "estimates.")
    pg.finish(print_rows=r, row_breaks=(page2 - 1,))


# =====================================================================================================
#  PAGE 5 - NOTE 3A / 3B PARTNERS' ACCOUNTS
# =====================================================================================================
def build_note3(wb):
    widths = [0.24, 1.32, 0.7, 0.5, 0.58, 0.66, 0.66, 0.49, 0.6, 0.68, 0.7]
    pg = Page(wb, "Note 3A-3B", widths, (0.48, 0.3, 0.9, 0.4))
    W = 11
    r = 1
    pg.height(r, 5)
    r += 1
    pg.cell(r, 1, " DAIVBHARATHI TECHNOLOGIES LLP", size=8, bold=True, span=W)
    pg.cell(r + 1, 1, "LLPIN: ACL-2422", size=8, bold=True, span=W)
    pg.cell(r + 2, 1, " Notes forming part of the Financial Statements as at 31st March 2026", size=8, bold=True, span=W)
    for k in range(3):
        pg.height(r + k, 11)
    r += 4
    pg.cell(r, 1, "Note - 3A Partners Contribution Account", size=8, bold=True, span=6)
    pg.cell(r + 1, W - 1, "(Amount in Rs.)", h="center", span=2)
    pg.height(r, 11)
    r += 2
    pg.cell(r, 1, "3A", bold=True, h="center")
    pg.cell(r, 2, "Partners Contribution Account", bold=True, span=2)
    pg.hline(r, 1, W, "top")
    pg.hline(r, 1, W, "bottom")
    pg.vline(1, r, r, "right")
    r += 1
    heads = ["Sr. No.", "Name of Partner", "Agreed contribution", "Share of profit/ (loss) (%)",
             "As at 1st April 2025 (Opening Balance)", "Introduced/ contributed during the year",
             "Remuneration for the year", "Interest for the year", "Withdrawals during the year",
             "Share of Profit / Loss for the year", "As at 31st March 2026 (Closing Balance)"]

    def header_row(r, heads, merge_name=False):
        for c, t in enumerate(heads, 1):
            if merge_name and c == 2:
                pg.cell(r, c, t, bold=True, h="center", wrap=True, fill=GREY, span=2)
            elif merge_name and c == 3:
                continue
            else:
                pg.cell(r, c, t, bold=True, h="center", wrap=True, fill=GREY)
        pg.height(r, 43)
        pg.grid(r, 1, r, W)

    header_row(r, heads)
    r += 1
    top3a = r
    # --- Note 3A rows (current year): per-partner split of the 90,000 contribution is not given in the
    #     FY 2025-26 statements (Note 2.1 D is blank) -> cells left open and highlighted for confirmation.
    pg.cell(r, 1, "1", h="center")
    pg.cell(r, 2, "Samartha R N")
    pg.amt(r, 3, P["n3a_s_agreed"])
    pg.cell(r, 4, "90%", h="right")
    pg.blank(r, 5, hl=True)
    pg.blank(r, 6, hl=True)
    for c in (7, 8, 9):
        pg.amt(r, c, None)
    pg.amt(r, 10, None, hl=True)
    pg.blank(r, 11, hl=True)
    r += 1
    pg.cell(r, 1, "2", h="center")
    pg.cell(r, 2, "Manasa Nagabhushnam")
    pg.amt(r, 3, P["n3a_m_agreed"])
    pg.cell(r, 4, "10%", h="right")
    pg.blank(r, 5, hl=True)
    pg.blank(r, 6, hl=True)
    for c in (7, 8, 9):
        pg.amt(r, c, None)
    pg.amt(r, 10, None, hl=True)
    pg.blank(r, 11, hl=True)
    r += 1
    # total row
    pg.blank(r, 5, hl=True)
    pg.blank(r, 6, hl=True)
    for c in (7, 8, 9, 10):
        pg.amt(r, c, None, bold=True)
    pg.amt(r, 11, S["capital"], bold=True)
    pg.hline(r, 5, W, "top")
    pg.hline(r, 5, W, "bottom")
    pg.height(r, 10.5)
    r += 1
    pg.cell(r, 1, "Previous Year (PY)", bold=True, span=4)
    pg.amt(r, 5, None, bold=True)
    pg.amt(r, 6, P["n3a_intro_total"], bold=True)
    for c in (7, 8, 9):
        pg.amt(r, c, None, bold=True)
    pg.amt(r, 10, P["n3a_share_total"], bold=True)
    pg.amt(r, 11, P["n3a_close_total"], bold=True)
    pg.hline(r, 1, W, "bottom")
    pg.height(r, 10.5)
    for c in range(1, W):
        pg.vline(c, top3a, r, "right")
    pg.box(1, 1, r, W)
    r += 3
    # --- Note 3B ---------------------------------------------------------------------------------------
    pg.cell(r, 1, "Note - 3B Partners Current Account", size=8, bold=True, span=6)
    pg.cell(r + 1, W - 1, "(Amount in Rs.)", h="center", span=2)
    pg.height(r, 11)
    r += 2
    header_row(r, heads, merge_name=True)
    r += 1
    top3b = r
    for i, name in enumerate(("Samartha R N", "Manasa Nagabhushnam"), 1):
        pg.cell(r, 1, str(i), h="center")
        pg.cell(r, 2, name)
        for c in range(5, W + 1):
            pg.amt(r, c, None)
        r += 1
    r += 1  # spacer row as in the reference
    for c in range(5, W + 1):
        pg.amt(r, c, None, bold=True)
    pg.hline(r, 5, W, "top")
    pg.hline(r, 5, W, "bottom")
    r += 1
    pg.cell(r, 1, "Previous Year (PY)", bold=True, span=4)
    for c in range(5, W + 1):
        pg.amt(r, c, None, bold=True)
    pg.hline(r, 1, W, "bottom")
    for c in range(1, W):
        pg.vline(c, top3b, r, "right")
    pg.box(top3b - 1, 1, r, W)
    pg.finish(print_rows=r + 1)


# =====================================================================================================
#  NOTES 4-14 (two-column note tables)
# =====================================================================================================
class NotePage:
    """Notes page: Particulars | (label) | current year | previous year."""

    def __init__(self, wb, name, part_w, cy_w, py_w, margins, header_lines, label_w=0.0, extra_cols=()):
        cols = [part_w - label_w, label_w] if label_w else [part_w]
        cols += list(extra_cols) if extra_cols else [cy_w, py_w]
        self.has_label = bool(label_w)
        self.pg = Page(wb, name, cols, margins)
        self.W = len(cols)
        self.cP = 1                                # particulars column
        self.cL = 2 if self.has_label else None    # label column (Note 9)
        self.cCY = self.W - 1
        self.cPY = self.W
        r = 1
        for i, (text, size) in enumerate(header_lines):
            self.pg.cell(r + i, 1, text, size=size, bold=True, span=self.W)
            self.pg.height(r + i, 11.5)
        self.r = r + len(header_lines) + 1

    def part_span(self):
        return 2 if self.has_label else 1

    def title(self, text, amount="(Amount in Rs.)", size=8):
        pg = self.pg
        pg.cell(self.r, 1, text, size=size, bold=True, span=self.part_span())
        pg.cell(self.r, self.cPY, amount, h="center")
        pg.height(self.r, 11.5)
        self.r += 1

    def header(self, cy=f"As at\n{CY}", py=f"As at\n{PY}", height=22):
        pg = self.pg
        pg.cell(self.r, 1, "Particulars", bold=True, h="center", wrap=True, fill=GREY, span=self.part_span())
        pg.cell(self.r, self.cCY, cy, bold=True, h="center", wrap=True, fill=GREY)
        pg.cell(self.r, self.cPY, py, bold=True, h="center", wrap=True, fill=GREY)
        pg.grid(self.r, 1, self.r, self.W)
        pg.height(self.r, height)
        self.r += 1
        self.top = self.r

    def row(self, text=None, cy="", py="", bold=False, indent=1, label=None, height=None, wrap=False,
            hl=False, hl_text=False, rule=None, total=False):
        pg = self.pg
        r = self.r
        if text is not None:
            pg.cell(r, 1, text, bold=bold, indent=indent, wrap=wrap, span=self.part_span() if label is None else 1,
                    v_="top" if wrap else "center", hl=hl_text)
        if label is not None:
            pg.cell(r, self.cL, label, bold=bold, h="center")
        if cy != "":
            pg.amt(r, self.cCY, cy, bold=bold, hl=hl)
            pg.amt(r, self.cPY, py, bold=bold, hl=hl)
        if height:
            pg.height(r, height)
        if rule == "num":
            pg.hline(r, self.cCY, self.cPY, "top")
            pg.hline(r, self.cCY, self.cPY, "bottom")
        if rule == "top":
            pg.hline(r, self.cCY, self.cPY, "top")
        if total:
            pg.hline(r, 1, self.W, "top")
            pg.hline(r, 1, self.W, "bottom")
        self.r += 1
        return r

    def close(self, gap=2):
        """Draw the vertical rules + outer box of the table just written."""
        pg = self.pg
        bottom = self.r - 1
        for c in range(1, self.W):
            if c == 1 and self.has_label:
                continue  # no rule between particulars and its label column
            pg.vline(c, self.top, bottom, "right")
        pg.box(self.top - 1, 1, bottom, self.W)
        self.r += gap

    def finish(self):
        self.pg.finish(print_rows=self.r)


def notes_header(company_case, date_text):
    return [(company_case, 8.5), ("LLPIN: ACL-2422", 8.5),
            (f"Notes forming part of the Financial Statements as at {date_text}", 8.5)]


def build_notes_4_7(wb):
    n = NotePage(wb, "Notes 4-5", 4.15, 1.025, 1.15, (0.7, 0.55, 0.8, 0.4),
                 notes_header("DAIVBHARATHI TECHNOLOGIES LLP", "31st March 2026"))
    # ---- Note 4 ---------------------------------------------------------------------------------------
    n.title("Note 4 : Reserves and Surplus")
    n.header()
    n.row("Capital Reserve", None, None)
    n.row("Revaluation Reserve", None, None)
    n.row("Other Reserve", None, None)
    n.row("Undistributed Surplus (Balance from statement of profit and loss)")
    n.row("Opening Balance", S["res_opening"], None, indent=3, hl=True)
    n.row("Add: Profit/(Loss) for the year", S["res_loss"], None, indent=3, hl=True)
    n.row("Less: Amount utilised", None, None, indent=3, hl=True)
    n.row("Less: Transfer to reserves", None, None, indent=3, hl=True)
    n.row("Closing Balance", S["res_closing"], None, indent=3, hl=True)
    n.row("Total", S["reserves"], None, bold=True, indent=0, total=True)
    n.close()
    # ---- Note 5 ---------------------------------------------------------------------------------------
    n.title("Note 5 : Long-term borrowings")
    n.header()
    n.row("Unsecured", None, None, bold=True)
    n.row("Term loans", bold=True)
    n.row("from banks")
    n.row("from other parties")
    n.row(None)
    n.row("Loans repayable on demand", bold=True)
    n.row("from banks", None, None)
    n.row("from other parties", None, None)
    n.row("Deferred payment liabilities", None, None)
    n.row("Loans and advances from related parties", None, None)
    n.row("Loans and advances - Others", S["ltb"], None, hl=True, hl_text=True)
    n.row("Long term/current maturitites of finance lease obligation", None, None)
    n.row("Total", S["ltb"], None, bold=True, total=True)
    n.close()
    # ---- Note 5A --------------------------------------------------------------------------------------
    n.title("Note 5A : Short-term borrowings")
    n.header()
    n.row("Unsecured", None, None, bold=True)
    n.row("Term loans", bold=True)
    n.row("from banks")
    n.row("from other parties")
    n.row(None)
    n.row("Loans repayable on demand", bold=True)
    n.row("from banks", None, None)
    n.row("from other parties", None, None)
    n.row("Deferred payment liabilities", None, None)
    n.row("Loans and advances from related parties", None, P["stb"], hl=True)
    n.row("Long term/current maturitites of finance lease obligation", None, None)
    n.row("Total", None, P["stb"], bold=True, total=True)
    n.close()
    n.finish()


def build_notes_6_7(wb):
    n = NotePage(wb, "Notes 6-7", 4.15, 1.025, 1.15, (0.7, 0.55, 0.8, 0.4),
                 notes_header("DAIVBHARATHI TECHNOLOGIES LLP", "31st March 2026"))
    # ---- Note 6 ---------------------------------------------------------------------------------------
    n.title("Note 6 : Trade Payables")
    n.header()
    n.row("Total outstanding dues of micro, small and medium enterprises", None, None)
    n.row("Total outstanding dues of creditors other than micro, small and medium enterprises",
          S["tp_other"], P["tp"], wrap=True, height=20)
    n.row("Total", S["tp"], P["tp"], bold=True, total=True)
    n.close(gap=1)
    n.pg.cell(n.r, 1, "Disclosure relating to suppliers registered under MSMED Act based on the information "
                      "available with the entity Company:", bold=True, span=n.W)
    n.pg.height(n.r, 14)
    n.r += 2
    n.header()
    n.row("(a) Amount remaining unpaid to any supplier at the end of each accounting year:", wrap=True, height=20)
    n.row("-Principal", None, None)
    n.row("-Interest", None, None)
    n.row("Total", None, None, bold=True, rule="num")
    n.row("(b) The amount of interest paid by the buyer in terms of section 16 of the MSMED Act, along with "
          "the amount of the payment made to the supplier beyond the appointed day during each accounting year.",
          None, None, wrap=True, height=31)
    n.row("(c) The amount of interest due and payable for the period of delay in making payment (which have "
          "been paid but beyond the appointed day during the year) but without adding the interest specified "
          "under the MSMED Act.", None, None, wrap=True, height=31)
    n.row("(d) The amount of interest accrued and remaining unpaid at the end of each accounting year.",
          None, None, wrap=True, height=21)
    n.row("(e) The amount of further interest remaining due and payable even in the succeeding years, until "
          "such date when the interest dues above are actually paid to the small enterprise, for the purpose "
          "of disallowance of a deductible expenditure under section 23 of the MSMED Act.",
          None, None, wrap=True, height=41)
    n.close(gap=1)
    n.pg.cell(n.r, 1, "There are no micro and small enterprises to which the LLP owes dues which are outstanding "
                      "for more than 45 days as at the balance sheet date. Dues to Micro and Small Enterprises "
                      "have been determined to the extent such parties have been identified on the basis of "
                      "information available with the LLP.", wrap=True, v_="top", span=n.W, hl=True)
    n.pg.height(n.r, 21)
    n.r += 2
    # ---- Note 7 ---------------------------------------------------------------------------------------
    n.title("Note 7 : Other current liabilities")
    n.header()
    n.row("Other payables", S["ocl_other"], P["ocl"], hl=True)
    n.row(None)
    n.row("Total", S["ocl"], P["ocl"], bold=True, total=True)
    n.close()
    n.finish()


def build_notes_8_10(wb):
    n = NotePage(wb, "Notes 8-10", 4.125, 0.9, 0.9, (0.82, 0.6, 0.78, 0.4),
                 notes_header("Daivbharathi Technologies LLP", "March 31, 2026"), label_w=0.55)
    # ---- Note 8 ---------------------------------------------------------------------------------------
    n.title("Note 8 : Trade receivables")
    n.header()
    n.row("Outstanding for a period less than 6 months from the date they are due for receipt", wrap=True, height=20)
    n.row("Secured Considered good", None, None)
    n.row("Unsecured Considered good", None, None)
    n.row("Doubtful", None, None)
    n.row("Less: Provision for doubtful receivables", None, None)
    n.row(None, None, None, rule="num")
    n.row("Outstanding for a period exceeding 6 months from the date they are due for receipt", None, None,
          wrap=True, height=20)
    n.row("Secured Considered good", None, None)
    n.row("Unsecured Considered good", None, None)
    n.row("Doubtful", None, None)
    n.row("Less: Provision for doubtful receivables", None, None)
    n.row("Unbilled receivables", None, None)
    n.row(None, None, None, rule="num")
    n.row("Total", None, None, bold=True, total=True)
    n.close()
    # ---- Note 9 ---------------------------------------------------------------------------------------
    n.title("Note 9 : Cash and Bank Balances")
    n.header()
    n.row("Cash and cash equivalents", bold=True)
    n.row("On current accounts", None, P["cash"], indent=2)
    n.row("Balances with banks", S["cash_bank"], None, indent=2, hl=True, hl_text=True)
    n.row("Cash credit account (Debit balance)", None, None)
    n.row("Fixed Deposits", bold=True)
    n.row("Deposits with original maturity of less than three months", None, None, indent=2)
    n.row("Cheques, drafts on hand", None, None)
    n.row("Cash on hand", S["cash_hand"], None)
    n.row("Others", S["cash_others"], None, hl=True, hl_text=True)
    n.row("Total", S["cash"], P["cash"], bold=True, label="(I)", rule="num")
    n.row(None)
    n.row("Other bank balances", bold=True)
    n.row("Bank Deposits")
    n.row("Earmarked Bank Deposits", None, None)
    n.row("Deposits with original maturity for more than 3 months but less than 12 months from reporting date",
          None, None, wrap=True, height=20)
    n.row("Margin money or deposits under lien", None, None)
    n.row("Others (specify nature)", None, None)
    n.row("Total other bank balances", None, None, bold=True, label="(II)", total=True)
    n.row("Total Cash and bank balances", S["cash"], P["cash"], bold=True, label="(I+II)", total=True)
    n.close()
    # ---- Note 10 --------------------------------------------------------------------------------------
    n.title("Note 10 Other Current Assets", amount="")
    n.header()
    n.row("Receivable From Partner", None, P["oca"])
    n.row(None)
    n.row(None, None, P["oca"], bold=True, rule="top")
    n.close()
    n.finish()


def build_notes_11_12(wb):
    n = NotePage(wb, "Notes 11-12", 3.925, 1.5, 1.235, (0.9, 0.6, 0.8, 0.4),
                 notes_header("Daivbharathi Technologies LLP", "31st March 2026"))
    pe_cy, pe_py = f"Period ended\n{CY}", f"Period ended\n{PY}"
    # ---- Note 11 --------------------------------------------------------------------------------------
    n.title("Note 11: Revenue from operations")
    n.header(pe_cy, pe_py)
    n.row("Sale of products", S["rev_sale"], None)
    n.row("Sale of services", None, None)
    n.row("Grants or donations received", None, None)
    n.row("Other operating revenue", None, None)
    n.row("Revenue from operations (Gross)", S["revenue"], None, bold=True, total=True)
    n.row("Less: Excise duty", None, None)
    n.row("Revenue from operations (Net)", S["revenue"], None, bold=True, total=True)
    n.close()
    # ---- Note 12 --------------------------------------------------------------------------------------
    n.title("Note 12: Other Expenses")
    n.header(pe_cy, pe_py)
    cy_map = {k.lower(): v for k, v in EXP_CY}
    ref_rows = [  # reference captions in reference order, with the FY 2024-25 figure where one exists
        ("Consumption of stores and spare parts", None),
        ("Interest on TDS", P["exp_tds"]),
        ("Rent", None),
        ("Repairs and maintenance - Buildings", None),
        ("Repairs and maintenance - Machinery", None),
        ("Insurance", None),
        ("Rent, Rates and taxes, excluding, taxes on income", P["exp_rates"]),
        ("Labour charges", None),
        ("Travelling expenses", None),
        ("Auditor's remuneration (Refer note below)", None),
        ("Printing and stationery", None),
        ("Communication expenses", None),
        ("Legal and professional charges", P["exp_legal"]),
        ("Advertisement and publicity", P["exp_adv"]),
        ("Bank Charges", P["exp_bank"]),
        ("Business promotion expenses", P["exp_bp"]),
        ("Commission", None),
        ("Clearing and forwarding charges", None),
        ("Loss on sale of Property, Plant and Equipment", None),
        ("Loss on foreign exchange transactions (net)", None),
        ("Loss on cancellation of forward contracts", None),
        ("Loss on sale of investments (net)", None),
        ("Provision for diminution in value of investments", None),
        ("Provision for doubtful debts", None),
    ]
    used = set()
    for cap, py in ref_rows:
        cy = cy_map.get(cap.lower())
        if cy is not None:
            used.add(cap.lower())
        n.row(cap, cy, py)
    # captions that exist only in the FY 2025-26 statements, verbatim, in source order (highlighted)
    for cap, cy in EXP_CY:
        if cap.lower() in used:
            continue
        n.row(cap, cy, None, hl=True, hl_text=True)
    n.row("Miscellaneous expenses", None, P["exp_misc"])
    n.row("Total Other Expenses", S["other_exp"], P["other_exp"], bold=True, total=True)
    n.close()
    n.finish()


def build_notes_13_14(wb):
    """Disclosures present in the FY 2025-26 statements that have no counterpart in the reference format."""
    pg = Page(wb, "Notes 13-14", [2.2, 0.9, 0.9, 0.9, 0.9, 0.9], (0.9, 0.6, 0.8, 0.4))
    W = 6
    r = 1
    for i, (text, size) in enumerate(notes_header("Daivbharathi Technologies LLP", "31st March 2026")):
        pg.cell(r + i, 1, text, size=size, bold=True, span=W)
        pg.height(r + i, 11.5)
    r += 4
    pg.cell(r, 1, "Note 13 : Trade Payables - Ageing schedule (Figures for the current reporting period)",
            size=8, bold=True, span=5, hl=True)
    pg.cell(r, W, "(Amount in Rs.)", h="center")
    pg.height(r, 11.5)
    r += 1
    pg.cell(r, 1, "Particulars", bold=True, h="center", fill=GREY, rows=2)
    pg.cell(r, 2, "Outstanding for following periods from due date of payment", bold=True, h="center", wrap=True,
            fill=GREY, span=4)
    pg.cell(r, W, "Total", bold=True, h="center", fill=GREY, rows=2)
    pg.height(r, 12)
    r += 1
    for c, t in zip(range(2, 6), ("< 1 Year", "1-2 Years", "2-3 Years", "> 3 Years")):
        pg.cell(r, c, t, bold=True, h="center", fill=GREY)
    pg.grid(r - 1, 1, r, W)
    pg.height(r, 12)
    r += 1
    top = r
    ageing = [("MSME", [None] * 5), ("Others", [S["tp_other"], None, None, None, S["tp_other"]]),
              ("Disputed dues - MSME", [None] * 5), ("Disputed dues - Others", [None] * 5)]
    for cap, vals in ageing:
        pg.cell(r, 1, cap, indent=1)
        for c, v in enumerate(vals, 2):
            pg.amt(r, c, v)
        r += 1
    pg.cell(r, 1, "Total", bold=True, indent=1)
    for c, v in enumerate([S["tp"], None, None, None, S["tp"]], 2):
        pg.amt(r, c, v, bold=True)
    pg.hline(r, 1, W, "top")
    pg.hline(r, 1, W, "bottom")
    for c in range(1, W):
        pg.vline(c, top, r, "right")
    pg.box(top - 2, 1, r, W)
    r += 3
    pg.cell(r, 1, "Note 14 : Related Party Disclosures", size=8, bold=True, span=4, hl=True)
    pg.height(r, 11.5)
    r += 1
    pg.cell(r, 1, "a. Key Management Personnel:", bold=True, span=4)
    r += 1
    pg.cell(r, 1, "Sl No", bold=True, h="center", fill=GREY)
    pg.cell(r, 2, "Name of the Party", bold=True, h="center", fill=GREY, span=3)
    pg.cell(r, 5, "Name of relationship", bold=True, h="center", fill=GREY, span=2)
    pg.grid(r, 1, r, W)
    pg.height(r, 12)
    r += 1
    for i, name in enumerate(("Samartha Raghava Nagabhushanam", "Manasa Nagabhushnam"), 1):
        pg.cell(r, 1, str(i), h="center")
        pg.cell(r, 2, name, span=3)
        pg.cell(r, 5, "Designated Partner", span=2)
        pg.grid(r, 1, r, W)
        r += 1
    r += 1
    pg.cell(r, 1, "b. Related Party Balances and Transactions:", bold=True, span=4)
    pg.cell(r, W, "(Amount in Rs.)", h="center")
    r += 1
    pg.cell(r, 1, "Sl.No.", bold=True, h="center", fill=GREY)
    pg.cell(r, 2, "Nature of transaction", bold=True, h="center", fill=GREY, span=3)
    pg.cell(r, 5, f"For the year ended\n{CY}", bold=True, h="center", wrap=True, fill=GREY, span=2)
    pg.grid(r, 1, r, W)
    pg.height(r, 22)
    r += 1
    for i, (cap, v) in enumerate((("Partners' capital contribution (closing balance)", S["capital"]),
                                  ("Unsecured loans - Others (closing balance)", S["ltb"])), 1):
        pg.cell(r, 1, str(i), h="center")
        pg.cell(r, 2, cap, span=3)
        pg.amt(r, 5, v, span=2)
        pg.grid(r, 1, r, W)
        r += 1
    pg.finish(print_rows=r)


# =====================================================================================================
#  OPEN POINTS (draft review sheet)
# =====================================================================================================
OPEN_POINTS = [
    ("1", "Partners' funds presentation",
     "FY 2024-25 (signed) shows the loss allocated to the partners inside 'Partners Contribution' (Note 3A, 90:10) "
     "with Reserves & Surplus nil. The FY 2025-26 statements show Partners' capital 90,000 and the accumulated loss "
     "(36,13,505) under Reserves & Surplus (Note 2.2). The draft follows the FY 2025-26 presentation for the current "
     "year and shows the previous year exactly as signed. Please confirm which presentation to use for FY 2025-26. "
     "If the loss is to be allocated to partners (FY 2024-25 style), the partner-wise share of (24,92,591) at 90:10 "
     "does not divide into whole rupees - please provide the partner-wise figures."),
    ("2", "Opening balance of Reserves & Surplus",
     "Note 2.2 of the FY 2025-26 statements shows an opening balance of (11,20,914), whereas the FY 2024-25 loss "
     "as signed is (11,20,912) - a difference of Rs. 2. Both figures are retained as provided; please confirm."),
    ("3", "Partners' contribution 90,000 vs 1,00,000",
     "FY 2024-25 shows contribution introduced 1,00,000 (Samartha 90,000 / Manasa 10,000) with 'Receivable From "
     "Partner' 10,000 in other current assets. FY 2025-26 shows contribution received 90,000 and no receivable. "
     "The FY 2025-26 statements do not give the partner-wise split, opening balance or amount introduced during "
     "the year (Note 2.1 B and D are blank) - these cells in Note 3A are left open (highlighted)."),
    ("4", "Classification of the loan (30,47,564)",
     "FY 2024-25 classified the loan as Short-term borrowings - 'Loans and advances from related parties' "
     "(7,80,680). FY 2025-26 classifies 30,47,564 as Long-term borrowings - Unsecured - 'Loans and advances - "
     "Others' (Note 2.3), while Note 2.9 lists the same balance as a related-party balance. The draft keeps each "
     "year as provided (Note 5 long-term for FY 2025-26; Note 5A short-term for FY 2024-25). Please confirm the "
     "classification and caption."),
    ("5", "Previous-year column",
     "The FY 2025-26 statements have a single column. The draft fills the 'As at March 31, 2025' column from the "
     "signed FY 2024-25 statements. Confirm that comparatives are to be shown."),
    ("6", "Note 12 - expense heads",
     "The reference note uses fixed heads (Rent, Legal and professional charges, Printing and stationery, "
     "Communication expenses ...). The FY 2025-26 statements use 24 different heads. To avoid re-classifying, the "
     "draft keeps every FY 2025-26 head verbatim (highlighted) and only merges captions that are identical "
     "(Interest on TDS, Bank charges, Commission). Please confirm, or tell me which heads to merge (e.g. "
     "Advertisement -> Advertisement and publicity; Printing charges -> Printing and stationery; Telephone "
     "expenses -> Communication expenses; Conveyance -> Travelling expenses; Professional / consultancy / "
     "technical fees -> Legal and professional charges)."),
    ("7", "Note 9 - cash captions",
     "FY 2025-26 shows 'Balances with banks' 5,102, 'Cash on hand' 21,837 and 'Others' 18. The reference format "
     "has 'On current accounts'. The draft keeps the FY 2025-26 captions (highlighted). Confirm whether 'Balances "
     "with banks' should be shown as 'On current accounts', and the nature of 'Others' (Rs. 18)."),
    ("8", "Note 7 - Other current liabilities",
     "The reference scan has the Note 7 heading but the table itself is missing (page 6). The draft uses the "
     "FY 2025-26 caption 'Other payables' 46,506 and shows the FY 2024-25 figure 35,000 on the same line - please "
     "confirm the FY 2024-25 break-up/caption."),
    ("9", "Note 4 - movement rows",
     "The reference shows only the closing balance. The draft adds the opening/loss/closing movement from Note 2.2 "
     "of the FY 2025-26 statements (highlighted). Confirm whether to keep or drop."),
    ("10", "Additional disclosures (Notes 13-14)",
     "The FY 2025-26 statements contain a trade payables ageing schedule, a trade receivables ageing schedule (all "
     "nil), related-party disclosures, the partners' rights/duties paragraph and the MSME 45-day statement, none of "
     "which appear in the reference format. The draft includes the payables ageing and related-party notes as "
     "Notes 13-14 and the MSME statement under Note 6 (highlighted); the nil receivables ageing and the "
     "rights/duties paragraph are omitted. Confirm."),
    ("11", "Names and identifiers",
     "The reference uses 'Samartha Raghava Nagabhushanam I' (trailing 'I'), 'Samartha R N', 'Manasa Nagabhushanam' "
     "and 'Manasa Nagabhushnam' in different places, and 'DIN'. The FY 2025-26 statements use 'MANASA NAGABHUSHNAM' "
     "and 'DPIN'. The draft replicates the reference text in each place. Please confirm the correct spellings."),
    ("12", "Text corrections",
     "The reference text is reproduced verbatim except for obvious typographical errors / duplicated words "
     "(incoporated, as flows, purcnase, thee, onon, duue, the garbled sentences in policies e, h and i, etc.). "
     "The full list is in the accompanying note. Tell me if you prefer the text exactly as in the reference."),
    ("13", "Signing date",
     "Place is kept as Bengaluru; the date on the Balance Sheet and P&L is left blank."),
]


def build_review(wb):
    pg = Page(wb, "Open points - delete before use", [0.5, 2.0, 5.5], (0.5, 0.5, 0.5, 0.5))
    pg.cell(1, 1, "Draft for review - open points requiring confirmation (yellow cells in the statements refer to "
                  "these points). Delete this sheet and remove the highlights once confirmed.",
            size=9, bold=True, span=3, wrap=True, v_="top")
    pg.height(1, 28)
    r = 3
    for c, t in ((1, "No."), (2, "Item"), (3, "Point")):
        pg.cell(r, c, t, size=9, bold=True, fill=GREY)
    pg.grid(r, 1, r, 3)
    pg.height(r, 14)
    r += 1
    for no, item, text in OPEN_POINTS:
        pg.cell(r, 1, no, size=9, h="center", v_="top")
        pg.cell(r, 2, item, size=9, bold=True, wrap=True, v_="top")
        pg.cell(r, 3, text, size=9, wrap=True, v_="top")
        pg.grid(r, 1, r, 3)
        pg.height(r, 12.5 * max(2, -(-len(text) // 105)) + 4)
        r += 1
    pg.finish(print_rows=r)


def main(out):
    verify()
    wb = Workbook()
    wb.remove(wb.active)
    build_balance_sheet(wb)
    build_pl(wb)
    build_policies(wb)
    build_note3(wb)
    build_notes_4_7(wb)
    build_notes_6_7(wb)
    build_notes_8_10(wb)
    build_notes_11_12(wb)
    build_notes_13_14(wb)
    wb.save(out)
    print("written", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Daivbharathi_Technologies_LLP_Financials_FY_2025-26.xlsx")
