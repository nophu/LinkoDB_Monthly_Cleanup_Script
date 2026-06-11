import json, os, re
from datetime import date
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from config import REPORT_CONFIG

# ── style ─────────────────────────────────────────────────────────────────────
GREEN  = "92D050"
YELLOW = "FFFF00"
ONEOFF = "BDD7EE"
GREY   = "D9D9D9"
RED    = "FF0000"
BLUE   = "0070C0"
FONT   = "Aptos Narrow"

# friendly display names for the new "Field" column
FIELD_DISPLAY = {
    "Extractor ID":        "Extractor ID",
    "Extractor Type":      "Extractor Type",
    "Trap Size and Units": "Trap Size & Units",
    "Cleaning Frequency":  "Cleaning Frequency",
    "ReceivingPlant":      "Receiving Plant",
    "ClassCode":           "Class Code",
    "SecondClass":         "Second Class",
    "TrunkLine":           "Trunk Line",
    "MapCategory":         "Map Category",
    "EventTypeAbbrv":      "Event Type",
}

def _f(size=11, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)

def _fill(c): return PatternFill("solid", fgColor=c)

def _border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def _cell(ws, row, col, value, bold=False, color="000000", fill=None,
          halign="left", italic=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _f(11, bold, color, italic)
    c.alignment = Alignment(horizontal=halign, vertical="center",
                            indent=1 if halign == "left" else 0, wrap_text=False)
    if fill: c.fill = _fill(fill)
    c.border = _border()
    return c


# columns: A=label  B=facility  C=permit  D=field  E=current  F=changed to
COL_LABEL, COL_FAC, COL_PERMIT, COL_FIELD, COL_CUR, COL_CHG = 1, 2, 3, 4, 5, 6


def _suggestion(change):
    if change["status"] == "fixed":
        return change["cleaned_value"]
    note = change.get("note", "")
    m = re.search(r"should it be changed to '([^']+)'", note)
    if m: return m.group(1)
    m = re.search(r"suggested: '([^']+)'", note)
    if m: return m.group(1)
    if "leave this field blank" in note or "should be deleted" in note:
        return ""   # actual blank cell instead of the word "Blank"
    return "Needs Manual Review"


# one issue per row, deduplicated by (facility, field, original value)
def _issue_rows(changes, field_order):
    seen = set()
    rows = []
    for chg in changes:
        key = (chg["facility"], chg["permit_no"], chg["field"], chg["original"])
        if key in seen: continue
        seen.add(key)
        rows.append({
            "facility": chg["facility"],
            "permit":   chg["permit_no"],
            "field":    chg["field"],
            "current":  chg["original"],
            "changed":  _suggestion(chg),
            "status":   chg["status"],
        })
    # sort by facility, then rubric field order
    def sk(r):
        try: fi = field_order.index(r["field"])
        except ValueError: fi = 99
        return (r["facility"].lower(), fi)
    rows.sort(key=sk)
    return rows


def _write_block(ws, row, cfg, changes, write_headers):
    fields = cfg["fields"]

    # Report Name row (+ column headers on the first block)
    _cell(ws, row, COL_LABEL, "Report Name:", bold=True, color=BLUE)
    _cell(ws, row, COL_FAC,   cfg["name"], bold=True)
    if write_headers:
        _cell(ws, row, COL_PERMIT, "Permit No",  bold=True, halign="center", fill=GREY)
        _cell(ws, row, COL_FIELD,  "Field",      bold=True, halign="center", fill=GREY)
        _cell(ws, row, COL_CUR,    "Current",    bold=True, halign="center", fill=GREY)
        _cell(ws, row, COL_CHG,    "Changed To", bold=True, halign="center", fill=GREY)
    ws.row_dimensions[row].height = 18
    row += 1

    # Fields Checked row
    _cell(ws, row, COL_LABEL, "Fields Checked:", bold=True, color=BLUE)
    friendly = [FIELD_DISPLAY.get(f, f) for f in fields]
    _cell(ws, row, COL_FAC, ",  ".join(friendly))
    ws.row_dimensions[row].height = 16
    row += 1

    # Issues Found rows
    rows = _issue_rows(changes, fields)
    if not rows:
        _cell(ws, row, COL_LABEL, "Issues Found:", bold=True, color=BLUE)
        _cell(ws, row, COL_FAC, "None")
        ws.row_dimensions[row].height = 16
        return row + 1

    for i, r in enumerate(rows):
        if r["status"] == "fixed":             fill = GREEN
        elif r["changed"] == "Needs Manual Review": fill = ONEOFF   # one-off, no suggestion
        else:                                  fill = YELLOW
        _cell(ws, row, COL_LABEL, "Issues Found:" if i == 0 else "",
              bold=(i == 0), color=BLUE)
        _cell(ws, row, COL_FAC,    r["facility"], fill=fill)
        _cell(ws, row, COL_PERMIT, r["permit"],   fill=fill, halign="center")
        _cell(ws, row, COL_FIELD,  FIELD_DISPLAY.get(r["field"], r["field"]),
              fill=fill, bold=True)
        _cell(ws, row, COL_CUR,    r["current"], fill=fill, color=RED)
        nmr = (r["changed"] == "Needs Manual Review")
        _cell(ws, row, COL_CHG,    r["changed"], fill=fill,
              color=(RED if r["status"] == "flagged" else "000000"),
              italic=nmr)
        ws.row_dimensions[row].height = 15
        row += 1

    return row


def build_report(changes_path="output/all_changes.json",
                 output_path="output/Monthly_Quality_Check_Report.xlsx",
                 only_reports=None):
    with open(changes_path) as f:
        all_changes = json.load(f)

    by_file = defaultdict(list)
    for c in all_changes:
        by_file[c["source_file"]].append(c)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Quality Check {date.today().strftime('%b %Y')}"

    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = f"Linko DB Monthly Quality Check  —  {date.today().strftime('%B %d, %Y')}"
    t.font = _f(13, bold=True)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24

    # ── color key / legend ──────────────────────────────────────────────────
    row = 2
    _cell(ws, row, COL_LABEL, "Color Key:", bold=True)
    row += 1

    _cell(ws, row, COL_FAC, "Green", bold=True, fill=GREEN, halign="center")
    ws.merge_cells(start_row=row, start_column=COL_PERMIT, end_row=row, end_column=COL_CHG)
    _cell(ws, row, COL_PERMIT,
          "Auto-fixed — the tool applied a clear correction (casing, EX prefix, trap unit). Apply it as-is in Linko.")
    ws.row_dimensions[row].height = 15
    row += 1

    _cell(ws, row, COL_FAC, "Yellow", bold=True, fill=YELLOW, halign="center")
    ws.merge_cells(start_row=row, start_column=COL_PERMIT, end_row=row, end_column=COL_CHG)
    _cell(ws, row, COL_PERMIT,
          "Needs review — the tool has a suggested fix for you to confirm.")
    ws.row_dimensions[row].height = 15
    row += 1

    _cell(ws, row, COL_FAC, "Blue", bold=True, fill=ONEOFF, halign="center")
    ws.merge_cells(start_row=row, start_column=COL_PERMIT, end_row=row, end_column=COL_CHG)
    _cell(ws, row, COL_PERMIT,
          "One-off — a unique value with no match (e.g. Quasar, HSW - Station 1); decide individually.")
    ws.row_dimensions[row].height = 15
    row += 2  # blank separator before first report

    items = list(REPORT_CONFIG.items())
    if only_reports:
        items = [(k, c) for k, c in items if k in only_reports]

    for idx, (key, cfg) in enumerate(items):
        row = _write_block(ws, row, cfg, by_file.get(key, []), write_headers=(idx == 0))
        row += 1  # blank separator

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 11
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 30
    ws.column_dimensions["F"].width = 26
    ws.freeze_panes = "B2"

    wb.save(output_path)
    print(f"Saved: {output_path}")
    return output_path


if __name__ == "__main__":
    os.makedirs("output", exist_ok=True)
    build_report()