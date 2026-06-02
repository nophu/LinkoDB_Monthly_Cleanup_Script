import json, os, re
from datetime import date
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from config import REPORT_CONFIG

# ── style constants (matching their sheets exactly) ───────────────────────────
GREEN   = "92D050"
YELLOW  = "FFFF00"
GREY    = "D9D9D9"
RED     = "FF0000"
BLUE    = "0070C0"   # label colour (Report Name:, Fields Checked:, Issues Found:)
FONT    = "Aptos Narrow"

def _f(size=11, bold=False, color="000000"):  return Font(name=FONT, size=size, bold=bold, color=color)

def _fill(hex_color): return PatternFill("solid", fgColor=hex_color)

def _border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def _cell(ws, row, col, value, bold=False, color="000000", fill=None,  halign="left", wrap=False, size=11):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _f(size, bold, color)
    c.alignment = Alignment(horizontal=halign, vertical="center",  wrap_text=wrap, indent=1 if halign == "left" else 0)
    if fill:  c.fill = _fill(fill)
    c.border = _border()
    return c


# ── extract a suggested "Changed To" value from a change entry ────────────────
def _suggestion(change):
    if change["status"] == "fixed": return change["cleaned_value"]
    note = change.get("note", "")
    # "should it be changed to 'X'?"

    m = re.search(r"should it be changed to '([^']+)'", note)
    if m: return m.group(1)
    # "suggested: 'X gpm'"

    m = re.search(r"suggested: '([^']+)'", note)
    if m: return m.group(1)

    # "should be deleted — leave this field blank"
    if "leave this field blank" in note or "should be deleted" in note: return "Blank"
    return ""


# group changes by facility, dedup same field per facility
def _group(changes, field_order):
    """
    Returns an OrderedDict:
      { (facility, permit_no): [ (field, original, suggestion), ... ] }
    One entry per unique (facility, field, original_value) — so a facility
    with two different bad values for the same field gets both shown.
    """
    seen  = {}   # (facility, permit, field, original) -> True
    order = {}   # (facility, permit)                  -> list of issues

    for chg in changes:
        key   = (chg["facility"], chg["permit_no"])
        fkey  = key + (chg["field"], chg["original"])   # include value in dedup key
        if fkey in seen:  continue
        seen[fkey] = True
        entry = (chg["field"], chg["original"], _suggestion(chg))
        if key not in order:   order[key] = []
        order[key].append(entry)

    # re-sort each facility's issues by the rubric field order
    def sort_key(item):
        try:   return field_order.index(item[0])
        except ValueError: return 99
    for key in order: order[key].sort(key=sort_key)
    return order


# layout constants
#  A = label | B = facility | C = permit | D,E = pair1 | F,G = pair2 | …
COL_LABEL    = 1   # A
COL_FACILITY = 2   # B
COL_PERMIT   = 3   # C
COL_PAIRS_START = 4   # D

def _pair_cols(pair_idx):
    """Return (current_col, changed_col) for pair_idx (0-based)."""
    base = COL_PAIRS_START + pair_idx * 2
    return base, base + 1


# write one report block
def _write_block(ws, row, cfg, changes, write_col_headers, max_pairs=4):
    report_name = cfg["name"]
    fields      = cfg["fields"]
    n_pairs     = max(len(fields), 1)

    # row: "Report Name:"
    _cell(ws, row, COL_LABEL, "Report Name:", bold=True, color=BLUE)
    _cell(ws, row, COL_FACILITY, report_name, bold=True)

    if write_col_headers:
        _cell(ws, row, COL_PERMIT, "Permit No", bold=True,  halign="center", fill=GREY)
        # write headers for ALL pairs (max across reports) so columns are labelled fully
        for i in range(max_pairs):
            cur, chg = _pair_cols(i)
            _cell(ws, row, cur, "Current",    bold=True, halign="center", fill=GREY)
            _cell(ws, row, chg, "Changed To", bold=True, halign="center", fill=GREY)

    ws.row_dimensions[row].height = 18
    row += 1

    # row: "Fields Checked:"
    _cell(ws, row, COL_LABEL, "Fields Checked:", bold=True, color=BLUE)
    _cell(ws, row, COL_FACILITY,  ",  ".join(fields), wrap=True)
    ws.row_dimensions[row].height = 16
    row += 1

    # row(s): "Issues Found:"
    if not changes:
        _cell(ws, row, COL_LABEL, "Issues Found:", bold=True, color=BLUE)
        _cell(ws, row, COL_FACILITY, "None")
        ws.row_dimensions[row].height = 16
        row += 1
        return row

    by_fac = _group(changes, fields)

    for idx, ((facility, permit), issues) in enumerate(by_fac.items()):
        label = "Issues Found:" if idx == 0 else ""

        # decide row fill: any flagged issue → yellow, all fixed → green
        has_flagged = any(
            chg["status"] == "flagged"
            for chg in changes
            if chg["facility"] == facility and chg["permit_no"] == permit
        )
        row_fill = YELLOW if has_flagged else GREEN
        _cell(ws, row, COL_LABEL,    label,    bold=(idx == 0), color=BLUE)
        _cell(ws, row, COL_FACILITY, facility, fill=row_fill)
        _cell(ws, row, COL_PERMIT,   permit,   fill=row_fill, halign="center")

        for pair_idx, (field, original, suggestion) in enumerate(issues):
            cur_col, chg_col = _pair_cols(pair_idx)

            # current value → red text (needs changing)
            _cell(ws, row, cur_col, original,   color=RED, fill=row_fill, halign="center")

            # changed to → red if flagged, black if auto-fixed
            chg_color = RED if has_flagged else "000000"
            _cell(ws, row, chg_col, suggestion, color=chg_color, fill=row_fill,  halign="center")
        ws.row_dimensions[row].height = 16
        row += 1
    return row


# main
def build_report(changes_path="output/all_changes.json",output_path="output/Monthly_Quality_Check_Report.xlsx"):
    with open(changes_path) as f:  all_changes = json.load(f)
    by_file = defaultdict(list)
    for c in all_changes:   by_file[c["source_file"]].append(c)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Quality Check {date.today().strftime('%b %Y')}"

    # title
    ws.merge_cells("A1:K1")
    t = ws["A1"]
    t.value     = f"Linko DB Monthly Quality Check  —  {date.today().strftime('%B %d, %Y')}"
    t.font      = _f(13, bold=True)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24
    current_row = 2

    # calculate the max number of Current/Changed To pairs needed
    # (based on the report with the most fields to check)
    max_pairs = max(len(cfg["fields"]) for cfg in REPORT_CONFIG.values())

    # one block per report
    for report_idx, (filename, cfg) in enumerate(REPORT_CONFIG.items()):
        changes = by_file.get(filename, [])
        include_headers = (report_idx == 0)   # column headers only in first block
        current_row = _write_block( ws, current_row, cfg, changes, include_headers, max_pairs)
        current_row += 1   # blank separator row

    # column widths
    ws.column_dimensions["A"].width = 16   # label
    ws.column_dimensions["B"].width = 40   # facility
    ws.column_dimensions["C"].width = 10   # permit
    # Current / Changed To pairs
    for i in range(4):
        cur_col, chg_col = _pair_cols(i)
        ws.column_dimensions[get_column_letter(cur_col)].width = 20
        ws.column_dimensions[get_column_letter(chg_col)].width = 22

    ws.freeze_panes = "B2"
    wb.save(output_path)
    print(f"Saved: {output_path}")
    return output_path

if __name__ == "__main__":
    os.makedirs("output", exist_ok=True)
    build_report()