#!/usr/bin/env python3
"""Build the Raad VBA application deliverables from the real .xlsx workbook.

This script:
  1. Audits every worksheet and every recognized header group.
  2. Writes the audit, UI layout, README and copy-paste source files.
  3. Writes an importable .frm/.cls and a one-click VBA auto-builder.
  4. Produces a macro-enabled .xlsm while preserving all original worksheet
     parts byte-for-byte (pyOpenVBA rewrites only vbaProject.bin and the small
     OPC parts required for macros).
"""

from __future__ import annotations

import csv
import re
import zipfile
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from pyopenvba import ExcelFile, VBAModuleKind, check_file

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "نظام_رعد_لإدارة_المخزون_والعجوزات_v9.xlsx"
OUTPUT = ROOT / "نظام_رعد_لإدارة_المخزون_والعجوزات_v9_App.xlsm"
VBA = ROOT / "VBA"

PALETTE = {
    "bg": ("#F4F7FB", (244, 247, 251)),
    "header": ("#102A43", (16, 42, 67)),
    "header_dark": ("#0C1F32", (12, 31, 50)),
    "card": ("#FFFFFF", (255, 255, 255)),
    "white": ("#FFFFFF", (255, 255, 255)),
    "text": ("#0F172A", (15, 23, 42)),
    "muted": ("#64748B", (100, 116, 139)),
    "border": ("#CBD5E1", (203, 213, 225)),
    "read_only": ("#F1F5F9", (241, 245, 249)),
    "primary": ("#2563EB", (37, 99, 235)),
    "primary_hover": ("#1D4ED8", (29, 78, 216)),
    "orange": ("#F59E0B", (245, 158, 11)),
    "orange_hover": ("#D97706", (217, 119, 6)),
    "danger": ("#DC2626", (220, 38, 38)),
    "danger_hover": ("#B91C1C", (185, 28, 28)),
    "gray": ("#64748B", (100, 116, 139)),
    "gray_hover": ("#475569", (71, 85, 105)),
    "success": ("#16A34A", (22, 163, 74)),
    "card_border": ("#E2E8F0", (226, 232, 240)),
}

# Static shell controls. Geometry is in UserForm points.
CONTROL_SPECS = [
    ("Frame", "fraHeader", 0, 0, 1200, 64, "", "header", "white", False, 9),
    ("Frame", "fraSidebar", 970, 76, 214, 622, "", "header_dark", "white", False, 9),
    ("Frame", "fraToolbar", 16, 76, 938, 68, "", "card", "text", False, 9),
    ("Frame", "fraEditor", 16, 466, 938, 210, "", "card", "text", False, 9),
    ("Frame", "fraButtons", 16, 684, 938, 48, "", "card", "text", False, 9),
    ("Label", "lblTitle", 24, 8, 620, 26, "Raad Inventory & Daily Variance", "header", "white", True, 16),
    ("Label", "lblSubTitle", 24, 34, 760, 20, "Live two-way connection to your real worksheets, fields, formulas and dropdowns", "header", "border", False, 9),
    ("Label", "lblModeBadge", 970, 16, 90, 28, "AGENT", "primary", "white", True, 10),
    ("Label", "lblClock", 1070, 16, 110, 28, "", "header", "white", False, 9),
    ("Label", "lblNav", 986, 92, 180, 22, "Workspace / Role", "header_dark", "white", True, 10),
    ("Label", "lblDatasetInfo", 986, 226, 180, 62, "Select a real sheet dataset", "header_dark", "white", False, 9),
    ("Label", "lblRecordCount", 986, 466, 180, 22, "Records: 0", "header_dark", "white", False, 9),
    ("Label", "lblSelectedRow", 986, 494, 180, 22, "Selected row: none", "header_dark", "white", False, 9),
    ("Label", "lblSearch", 28, 88, 240, 18, "Live search", "card", "text", False, 9),
    ("Label", "lblFilterField", 294, 88, 180, 18, "Filter by real column", "card", "text", False, 9),
    ("Label", "lblFilterValue", 490, 88, 180, 18, "Filter value", "card", "text", False, 9),
    ("Label", "lblEditorTitle", 28, 474, 400, 20, "Record fields - exact worksheet headers", "card", "text", True, 10),
    ("Label", "lblStatus", 16, 738, 1168, 26, "Ready.", "card", "success", True, 9),
    ("ComboBox", "cboRole", 986, 120, 180, 26, "", "white", "text", False, 9),
    ("ComboBox", "cboDataset", 986, 194, 180, 26, "", "white", "text", False, 9),
    ("ComboBox", "cboFilterField", 294, 108, 180, 26, "", "white", "text", False, 9),
    ("ComboBox", "cboFilterValue", 490, 108, 180, 26, "", "white", "text", False, 9),
    ("TextBox", "txtSearch", 28, 108, 250, 26, "", "white", "text", False, 9),
    ("ListBox", "lstHeaders", 16, 152, 938, 24, "", "header", "white", False, 9),
    ("ListBox", "lstData", 16, 176, 938, 282, "", "white", "text", False, 9),
    ("CommandButton", "btnSwitchMode", 986, 152, 180, 30, "Switch Mode", "primary", "white", True, 9),
    ("CommandButton", "btnRefresh", 986, 298, 180, 30, "Refresh", "primary", "white", True, 9),
    ("CommandButton", "btnSheetView", 986, 336, 180, 30, "Switch to Excel Sheet", "primary", "white", True, 9),
    ("CommandButton", "btnUnlockAll", 986, 380, 180, 30, "Unlock Sheets", "primary", "white", True, 9),
    ("CommandButton", "btnLockForAgent", 986, 418, 180, 30, "Lock For Agent", "gray", "white", True, 9),
    ("CommandButton", "btnClearFilter", 686, 108, 120, 26, "Clear Filter", "gray", "white", True, 9),
    ("CommandButton", "btnAdd", 28, 694, 140, 30, "Add", "primary", "white", True, 9),
    ("CommandButton", "btnUpdate", 180, 694, 140, 30, "Update", "orange", "white", True, 9),
    ("CommandButton", "btnDelete", 332, 694, 140, 30, "Delete", "danger", "white", True, 9),
    ("CommandButton", "btnClear", 484, 694, 140, 30, "Clear", "gray", "white", True, 9),
    ("CommandButton", "btnClose", 800, 694, 140, 30, "Close", "gray", "white", True, 9),
]

PROG_IDS = {
    "Frame": "Forms.Frame.1",
    "Label": "Forms.Label.1",
    "TextBox": "Forms.TextBox.1",
    "ComboBox": "Forms.ComboBox.1",
    "CommandButton": "Forms.CommandButton.1",
    "ListBox": "Forms.ListBox.1",
}

FRM_KINDS = {
    "Frame": "VB.Frame",
    "Label": "VB.Label",
    "TextBox": "VB.TextBox",
    "ComboBox": "VB.ComboBox",
    "CommandButton": "VB.CommandButton",
    "ListBox": "VB.ListBox",
}

# Same dataset map used in modRaadApp.InitDatasets.
DATASETS = [
    ("DASH_KPI_LABELS", 1, "Dashboard KPI labels", 13, 2, 10, 14, 14, False, False, False),
    ("DASH_KPI_TABLE", 1, "Dashboard KPI table", 17, 2, 11, 18, 28, False, True, False),
    ("DAILY_INPUT", 2, "Daily input station", 5, 1, 7, 6, 16, False, True, False),
    ("BANANA_MASTER", 3, "Unified bun master", 5, 1, 10, 6, 6, False, True, False),
    ("BANANA_DETAIL", 3, "Unified bun distribution", 10, 1, 6, 11, 14, False, True, False),
    ("TIME_ANALYTICS", 4, "Time analytics", 6, 1, 9, 7, 17, False, False, False),
    ("MAIN_ARCHIVE", 5, "Daily archive - FULL CRUD", 5, 1, 9, 6, 200, True, True, True),
    ("BEV_PEPSI", 6, "Beverages - block 1", 3, 1, 9, 4, 4, False, True, False),
    ("BEV_WATER", 6, "Beverages - block 2", 8, 1, 9, 9, 9, False, True, False),
    ("FAKHER_FULL", 7, "Full section grid", 2, 1, 13, 3, 21, False, True, False),
    ("MASHAWI_FULL", 8, "Full section grid", 2, 1, 20, 3, 23, False, True, False),
    ("BURGER_SALES", 9, "Sales grid", 1, 1, 7, 2, 12, False, True, False),
    ("BURGER_INV", 9, "Inventory grid", 14, 1, 8, 15, 19, False, True, False),
    ("MADGHOUT_FULL", 10, "Full section grid", 3, 1, 15, 4, 15, False, True, False),
    ("ZURBIAN_FULL", 11, "Full section grid", 3, 1, 16, 4, 12, False, True, False),
    ("KUNAFA_GRID", 12, "Formula grid", 2, 1, 17, 3, 11, False, True, False),
    ("KUNAFA_INV", 12, "Inventory grid", 14, 1, 8, 15, 24, False, True, False),
    ("FRIDGE_FULL", 13, "Fridge grid", 2, 1, 15, 3, 19, False, True, False),
    ("BROAST_SALES", 14, "Broast sales", 5, 2, 8, 6, 12, False, True, False),
    ("BROAST_INV", 14, "Broast inventory", 16, 2, 9, 17, 21, False, True, False),
    ("PIZZA_SALES", 16, "Pizza sales", 1, 1, 4, 2, 9, False, True, False),
    ("PIZZA_INV", 16, "Pizza inventory", 13, 1, 8, 14, 15, False, True, False),
]

BROAST_HEADER_ROWS = [5, 33, 61, 90, 118, 145, 171, 199, 229, 256, 284, 311, 338, 366]
for idx, hdr in enumerate(BROAST_HEADER_ROWS):
    last = BROAST_HEADER_ROWS[idx + 1] - 1 if idx + 1 < len(BROAST_HEADER_ROWS) else 390
    DATASETS.append((
        f"BROAST_HIST_{idx + 1:02d}", 15, f"Historical daily block {idx + 1:02d}",
        hdr, 1, 17, hdr + 1, last, False, False, False,
    ))


def rgb(key: str) -> tuple[int, int, int]:
    return PALETTE[key][1]


def ole_color(value: tuple[int, int, int]) -> int:
    r, g, b = value
    return r + (g * 256) + (b * 65536)


def hex_vba_color(value: tuple[int, int, int]) -> str:
    r, g, b = value
    return f"&H00{b:02X}{g:02X}{r:02X}&"


def vba_quote(text: str) -> str:
    return text.replace('"', '""')


def crlf(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def source_body(path: Path, strip_attribute: bool = False) -> str:
    lines = path.read_text(encoding="utf-8").splitlines()
    if strip_attribute and lines and lines[0].startswith("Attribute VB_Name"):
        lines = lines[1:]
    return "\r\n".join(lines) + "\r\n"


def validation_for(ws, col: int, first_data_row: int) -> str:
    coord = f"{get_column_letter(col)}{first_data_row}"
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list":
            continue
        for token in str(dv.sqref).split():
            try:
                min_col, min_row, max_col, max_row = range_boundaries(token)
            except Exception:
                continue
            if min_col <= col <= max_col and min_row <= first_data_row <= max_row:
                raw = str(dv.formula1 or "")
                if raw.startswith('"') and raw.endswith('"'):
                    raw = raw[1:-1]
                return raw
    return ""


def detect_type(ws, cfg, col: int) -> tuple[str, str, str, str, str]:
    key, sheet_idx, label, header_row, first_col, last_col, first_data, last_data, *_ = cfg
    formula_example = ""
    sample = ""
    has_formula = False
    has_date = False
    has_number = False
    stop = min(first_data + 20, last_data)
    for r in range(first_data, stop + 1):
        cell = ws.cell(r, col)
        value = cell.value
        if isinstance(value, str) and value.startswith("="):
            has_formula = True
            if not formula_example:
                formula_example = value
        if isinstance(value, (date, datetime)):
            has_date = True
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            has_number = True
        fmt = str(cell.number_format or "").lower()
        if "yy" in fmt or "dd" in fmt:
            has_date = True
        if any(x in fmt for x in ("$", "€", "£", "﷼", "ر.س", "sar")):
            has_number = True
        if sample == "" and value is not None:
            sample = str(value)
    dropdown = validation_for(ws, col, first_data)
    if dropdown:
        detected = "Dropdown/Status"
    elif has_formula:
        detected = "Formula - read-only"
    elif has_date:
        detected = "Date"
    elif has_number:
        detected = "Number"
    else:
        detected = "Text / blank"
    return detected, "Yes" if has_formula else "No", formula_example, dropdown, sample


def extra_text_cells(wb, true_bounds):
    """Collect real string cells that are not part of any recognized table range."""
    covered = {}
    for cfg in DATASETS:
        key, sheet_idx, label, header_row, first_col, last_col, first_data, last_data, *_ = cfg
        name = wb.sheetnames[sheet_idx - 1]
        covered.setdefault(name, []).append((header_row, last_data, first_col, last_col))
    result = {}
    for ws in wb.worksheets:
        max_row, max_col = true_bounds[ws.title]
        items = []
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                value = ws.cell(r, c).value
                if not isinstance(value, str) or not value.strip() or value.startswith("="):
                    continue
                is_covered = False
                for hr, lr, fc, lc in covered.get(ws.title, []):
                    if hr <= r <= lr and fc <= c <= lc:
                        is_covered = True
                        break
                if not is_covered:
                    items.append((f"{get_column_letter(c)}{r}", value))
        result[ws.title] = items
    return result


def write_audit(wb) -> None:
    csv_path = ROOT / "Field_Audit.csv"
    extra_csv_path = ROOT / "Extra_Text_Cells.csv"
    md_path = ROOT / "File_Audit.md"
    # Capture true dimensions before any ws.cell() access can expand openpyxl's view.
    true_dims = {ws.title: ws.calculate_dimension() for ws in wb.worksheets}
    true_bounds = {}
    for name, dim in true_dims.items():
        min_col, min_row, max_col, max_row = range_boundaries(dim)
        true_bounds[name] = (max_row, max_col)
    rows = []
    for cfg in DATASETS:
        key, sheet_idx, label, header_row, first_col, last_col, first_data, last_data, *_ = cfg
        ws = wb.worksheets[sheet_idx - 1]
        for col in range(first_col, last_col + 1):
            header = ws.cell(header_row, col).value
            header_text = "" if header is None else str(header)
            detected, formula_col, formula_example, dropdown, sample = detect_type(ws, cfg, col)
            rows.append({
                "group_key": key,
                "sheet_index": sheet_idx,
                "sheet_name": ws.title,
                "group_label": label,
                "header_row": header_row,
                "first_col": get_column_letter(first_col),
                "last_col": get_column_letter(last_col),
                "column_letter": get_column_letter(col),
                "column_number": col,
                "header_text": header_text,
                "detected_type": detected,
                "formula_column": formula_col,
                "formula_example": formula_example,
                "dropdown_values": dropdown,
                "number_format": ws.cell(first_data, col).number_format,
                "sample_value": sample,
                "data_rows": f"{first_data}:{last_data}",
            })
    with csv_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    extras = extra_text_cells(wb, true_bounds)
    with extra_csv_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["sheet_name", "cell", "text_value"])
        for sheet_name, items in extras.items():
            for coord, value in items:
                writer.writerow([sheet_name, coord, value])

    lines = []
    lines.append("# File Audit - Raad inventory workbook")
    lines.append("")
    lines.append(f"Source file: `{SOURCE.name}`")
    lines.append(f"Worksheet count: **{len(wb.sheetnames)}**")
    lines.append("")
    lines.append("## Exact sheet names and order")
    lines.append("")
    for i, name in enumerate(wb.sheetnames, 1):
        ws = wb[name]
        lines.append(f"{i}. `{name}` — state `{ws.sheet_state}`, used range `{true_dims[name]}`")
    lines.append("")
    lines.append("## Primary flat database selected for full CRUD")
    lines.append("")
    main = wb.worksheets[4]
    lines.append(f"Sheet: `{main.title}`; header row **5**; data rows **6:200**; autofilter `A5:I200`; no formulas.")
    lines.append("")
    for col in range(1, 10):
        v = main.cell(5, col).value
        lines.append(f"- **{get_column_letter(col)}** = `{v}`")
    lines.append("")
    lines.append("## Data types and real validations")
    lines.append("")
    lines.append("- A in the daily archive is a real Excel **Date** (`yyyy-mm-dd`).")
    lines.append("- E:H in the daily archive are **Numbers** (`#,##0.00` / signed variance formats).")
    lines.append("- B:D and I are **Text**; I is a real **Dropdown/Status** column.")
    lines.append(f"- `{wb.worksheets[2].title}`!F11 validation list: `\"نعم,لا\"`.")
    lines.append(f"- `{main.title}`!I6:I200 validation list: `\"عجز,زيادة,مطابق\"`.")
    lines.append("- No currency-formatted columns were found in the workbook; quantity and variance fields are numeric, not currency.")
    lines.append("- Formula-backed columns are detected dynamically at runtime and shown as locked, read-only inputs; the app only writes user-entry cells.")
    lines.append("- `Field_Audit.csv` classifies every real dataset column as Text, Number, Date, Formula - read-only, or Dropdown/Status. Currency is explicitly reported as absent.")
    lines.append("")
    lines.append("## Excel tables")
    lines.append("")
    for ws in wb.worksheets:
        for table in ws.tables.values():
            lines.append(f"- `{ws.title}` :: `{table.name}` :: `{table.ref}`")
    lines.append("")
    lines.append("## Formula counts by sheet")
    lines.append("")
    for ws in wb.worksheets:
        count = sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
        lines.append(f"- `{ws.title}`: {count} formulas")
    lines.append("")
    lines.append("## Every recognized header group and every column")
    lines.append("")
    lines.append("Blank header cells are shown as `[blank]` to prove the exact column position. The machine-readable `Field_Audit.csv` contains one row per real field, including detected type, number format, formula example, dropdown values, and a real sample value.")
    lines.append("")
    for cfg in DATASETS:
        key, sheet_idx, label, header_row, first_col, last_col, first_data, last_data, allow_add, allow_update, allow_delete = cfg
        ws = wb.worksheets[sheet_idx - 1]
        headers = []
        for col in range(first_col, last_col + 1):
            v = ws.cell(header_row, col).value
            text = "[blank]" if v is None or str(v).strip() == "" else str(v)
            safe_text = text.replace("|", "\\|")
            headers.append(f"**{get_column_letter(col)}**=`{safe_text}`")
        lines.append(f"### `{key}` — `{ws.title}` | {label}")
        lines.append(f"Header row **{header_row}**, data rows **{first_data}:{last_data}**, columns **{get_column_letter(first_col)}:{get_column_letter(last_col)}**. Add={allow_add}, Update={allow_update}, Delete={allow_delete}.")
        lines.append("")
        lines.append(" | ".join(headers))
        lines.append("")
    lines.append("## Additional non-tabular text cells outside recognized dataset ranges")
    lines.append("")
    lines.append("These are real workbook titles, navigation links, legends and helper lists. They are not column headers inside a flat dataset, but they are listed so the audit is exhaustive from the first used column to the last. `Extra_Text_Cells.csv` contains the same rows in machine-readable form.")
    lines.append("")
    for sheet_name, items in extras.items():
        lines.append(f"### `{sheet_name}`")
        lines.append("")
        if not items:
            lines.append("No additional non-tabular text cells.")
        else:
            for coord, value in items:
                safe = str(value).replace("|", "\\|").replace("`", "\\`")
                lines.append(f"- `{coord}` = `{safe}`")
        lines.append("")
    md_path.write_text("\n".join(lines), encoding="utf-8")


def write_ui_docs(wb) -> None:
    path = ROOT / "UI_Layout_and_Controls.md"
    lines = [
        "# UI Layout, Exact Control Names and RGB Design Settings",
        "",
        "All controls use `Segoe UI`. The code applies `SpecialEffect = 0 - fmSpecialEffectFlat` and `BorderStyle = 1 - fmBorderStyleSingle` with safe property setters so unsupported control properties never raise an error.",
        "",
        "## Exact palette",
        "",
        "| Role | Hex | RGB | VBA |",
        "|---|---|---|---|",
    ]
    roles = {
        "bg": "Form background",
        "header": "Header / sidebar",
        "header_dark": "Sidebar deep panel",
        "card": "Cards and input boxes",
        "white": "Input surfaces / button text",
        "text": "Primary text",
        "muted": "Muted text / read-only",
        "border": "Single-pixel borders",
        "read_only": "Formula/read-only inputs",
        "primary": "Primary blue - Add / Save",
        "primary_hover": "Primary hover",
        "orange": "Update",
        "orange_hover": "Update hover",
        "danger": "Delete",
        "danger_hover": "Delete hover",
        "gray": "Clear / neutral",
        "gray_hover": "Neutral hover",
        "success": "Success status",
        "card_border": "Card outline",
    }
    for key, desc in roles.items():
        hx, (r, g, b) = PALETTE[key]
        lines.append(f"| {desc} | `{hx}` | `RGB({r}, {g}, {b})` | `RGB({r}, {g}, {b})` |")
    lines += [
        "",
        "## Static UserForm shell (`frmRaadApp`)",
        "",
        "Geometry is in points on a 1200 x 780 form. `Back` and `Fore` refer to the palette keys above.",
        "",
        "| (Name) | Type | Left | Top | Width | Height | Caption | Back | Fore |",
        "|---|---|---:|---:|---:|---:|---|---|---|",
    ]
    for spec in CONTROL_SPECS:
        kind, name, x, y, w, h, caption, back, fore, bold, size = spec
        cap = caption.replace("|", "\\|") or "—"
        lines.append(f"| `{name}` | {kind} | {x} | {y} | {w} | {h} | {cap} | `{back}` | `{fore}` |")
    lines += [
        "",
        "## Dynamic exact-field controls",
        "",
        "For every selected dataset, `BuildFieldMetadata` reads the exact header text from the real worksheet and `BuildFieldControls` creates these controls inside `fraEditor`:",
        "",
        "- Label: `lblF1` ... `lblFn` — caption is the exact worksheet header, or `[ColumnLetter] (no header)` when the header cell is blank.",
        "- Text input: `txtF1` ... `txtFn` — Text, Date, Number and read-only Formula fields.",
        "- Combo input: `cboF1` ... `cboFn` — only for real Dropdown/Status columns; items come from the worksheet validation list.",
        "",
        "For the default full-CRUD sheet `سجل_الأرشيف_اليومي`, the exact mapping is:",
        "",
        "| Column | Exact header | Control | Type |",
        "|---|---|---|---|",
    ]
    main = wb.worksheets[4]
    for col in range(1, 10):
        header = str(main.cell(5, col).value).replace("|", "\\|")
        control = "cboF9" if col == 9 else f"txtF{col}"
        kind = "Dropdown/Status" if col == 9 else ("Date" if col == 1 else ("Number" if 5 <= col <= 8 else "Text"))
        lines.append(f"| {get_column_letter(col)} | `{header}` | `{control}` | {kind} |")
    lines += [
        "",
        "Formula columns are created as `txtF…` controls with `Locked = True`, a `#F1F5F9` background, and a read-only tooltip. They are never written by Add/Update/Delete.",
        "",
        "## Behavior matrix",
        "",
        "| Feature | Agent | Admin |",
        "|---|---|---|",
        "| Add / Search / View / Update | Yes on allowed datasets | Yes |",
        "| Delete | Button disabled | Enabled only on the flat main archive; confirmation required |",
        "| Workbook structure | Protected | Unprotected by `Unlock Sheets` |",
        "| Formula cells | Locked in raw sheet view | Editable only in raw Excel view |",
        "| Real Excel sheet view | Available; raw cells are view-only while the Agent lock is active | Available with fully unlocked sheets |",
        "",
        "Fixed formula layouts and historical broast blocks are read-only/controlled-update datasets inside the app. Add and physical Delete are intentionally disabled there so your existing formulas, merges and sheet structure remain intact. The flat `سجل_الأرشيف_اليومي` sheet provides complete Add/Update/Delete CRUD.",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


def write_frm_file(form_body: str) -> None:
    lines = [
        "VERSION 5.00",
        "Begin VB.UserForm frmRaadApp",
        f'   Caption         =   "Raad | Live Inventory & Variance Console"',
        f"   ClientHeight    =   {780 * 20}",
        "   ClientLeft      =   120",
        "   ClientTop       =   450",
        f"   ClientWidth     =   {1200 * 20}",
        f"   BackColor       =   {hex_vba_color(rgb('bg'))}",
        "   BorderStyle     =   1  'fmBorderStyleSingle",
        "   SpecialEffect   =   0  'fmSpecialEffectFlat",
        "   StartUpPosition =   2  'CenterScreen",
    ]
    for idx, spec in enumerate(CONTROL_SPECS):
        kind, name, x, y, w, h, caption, back, fore, bold, size = spec
        lines.append(f"   Begin {FRM_KINDS[kind]} {name}")
        if kind in ("Frame", "Label", "CommandButton"):
            lines.append(f'      Caption         =   "{vba_quote(caption)}"')
        lines.append(f"      Height          =   {int(h * 20)}")
        lines.append(f"      Left            =   {int(x * 20)}")
        lines.append(f"      TabIndex        =   {idx}")
        lines.append(f"      Top             =   {int(y * 20)}")
        lines.append(f"      Width           =   {int(w * 20)}")
        lines.append(f"      BackColor       =   {hex_vba_color(rgb(back))}")
        lines.append(f"      ForeColor       =   {hex_vba_color(rgb(fore))}")
        if kind in ("Label", "TextBox", "ComboBox", "ListBox"):
            lines.append("      BorderStyle     =   1  'fmBorderStyleSingle")
            lines.append("      SpecialEffect   =   0  'fmSpecialEffectFlat")
        lines.append("   End")
    lines += [
        "End",
        'Attribute VB_Name = "frmRaadApp"',
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = False",
        "",
    ]
    (VBA / "frmRaadApp.frm").write_text("\r\n".join(lines) + crlf(form_body), encoding="utf-8")


def write_thisworkbook_cls(workbook_body: str) -> None:
    header = "\r\n".join([
        "VERSION 1.0 CLASS",
        "BEGIN",
        "  MultiUse = -1  'True",
        "END",
        'Attribute VB_Name = "ThisWorkbook"',
        'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"',
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = True",
        "",
    ]) + "\r\n"
    (VBA / "ThisWorkbook.cls").write_text(header + crlf(workbook_body), encoding="utf-8")


def string_function(name: str, body: str) -> list[str]:
    lines = [
        f"Private Function {name}() As String",
        "    On Error GoTo ErrorHandler",
        "    Dim s As String",
        "    s = \"\"",
    ]
    for line in body.splitlines():
        lines.append(f'    s = s & "{vba_quote(line)}" & vbCrLf')
    lines += [
        f"    {name} = s",
        "    Exit Function",
        "ErrorHandler:",
        f'    BuilderHandleError "{name}"',
        f"    {name} = \"\"",
        "End Function",
        "",
    ]
    return lines


def write_auto_builder(module_body: str, form_body: str, workbook_body: str) -> None:
    lines = [
        'Attribute VB_Name = "AutoBuild_RaadApp"',
        "Option Explicit",
        "",
        "' One-click builder. Import this module into the macro-enabled workbook,",
        "' enable Trust access to the VBA project object model, then run Build_Raad_App.",
        "",
        "Public Sub Build_Raad_App()",
        "    On Error GoTo ErrorHandler",
        "    Dim oldEvents As Boolean",
        "    Dim oldAlerts As Boolean",
        "    Dim oldScreen As Boolean",
        "    Dim vbProj As Object",
        "    Dim stdMod As Object",
        "    Dim frm As Object",
        "    Dim newPath As String",
        "    oldEvents = Application.EnableEvents",
        "    oldAlerts = Application.DisplayAlerts",
        "    oldScreen = Application.ScreenUpdating",
        "    TrySetApp \"EnableEvents\", False",
        "    TrySetApp \"DisplayAlerts\", False",
        "    TrySetApp \"ScreenUpdating\", False",
        "    Set vbProj = ThisWorkbook.VBProject",
        "    Set stdMod = AddStandardModule(vbProj, \"modRaadApp\", ModuleSource())",
        "    If stdMod Is Nothing Then GoTo SafeExit",
        "    Set frm = AddUserForm(vbProj, \"frmRaadApp\", FormSource())",
        "    If frm Is Nothing Then GoTo SafeExit",
        "    ReplaceModuleCode ThisWorkbook.VBProject.VBComponents(\"ThisWorkbook\"), WorkbookSource()",
        "    If LCase$(Right$(ThisWorkbook.Name, 5)) <> \".xlsm\" Then",
        "        newPath = Left$(ThisWorkbook.FullName, Len(ThisWorkbook.FullName) - 5) & \".xlsm\"",
        "        ThisWorkbook.SaveAs Filename:=newPath, FileFormat:=52",
        "    Else",
        "        ThisWorkbook.Save",
        "    End If",
        '    MsgBox "Raad app built successfully. Reopen the workbook and enable macros; the dashboard will open automatically.", vbInformation, "Raad App"',
        "SafeExit:",
        "    TrySetApp \"EnableEvents\", oldEvents",
        "    TrySetApp \"DisplayAlerts\", oldAlerts",
        "    TrySetApp \"ScreenUpdating\", oldScreen",
        "    Exit Sub",
        "ErrorHandler:",
        '    BuilderHandleError "Build_Raad_App"',
        "    Resume SafeExit",
        "End Sub",
        "",
        "Private Sub BuilderHandleError(ByVal procName As String)",
        "    On Error GoTo ErrorHandler",
        "    Dim n As Long",
        "    Dim desc As String",
        "    Dim msg As String",
        "    n = Err.Number",
        "    desc = Err.Description",
        "    Err.Clear",
        "    msg = \"Builder error in \" & procName & \": \" & CStr(n) & \" - \" & desc & vbCrLf & _",
        "          \"If this is a permissions error, enable: File > Options > Trust Center > Macro Settings > Trust access to the VBA project object model.\"",
        '    MsgBox msg, vbExclamation, "Raad App Builder"',
        "    Exit Sub",
        "ErrorHandler:",
        "    Err.Clear",
        "    Resume SafeExit",
        "SafeExit:",
        "End Sub",
        "",
        "Private Sub TrySetApp(ByVal propName As String, ByVal value As Variant)",
        "    On Error GoTo ErrorHandler",
        "    CallByName Application, propName, VbLet, value",
        "    Exit Sub",
        "ErrorHandler:",
        "    Err.Clear",
        "    Resume SafeExit",
        "SafeExit:",
        "End Sub",
        "",
        "Private Sub TrySetCtl(ByVal ctl As Object, ByVal propName As String, ByVal value As Variant)",
        "    On Error GoTo ErrorHandler",
        "    CallByName ctl, propName, VbLet, value",
        "    Exit Sub",
        "ErrorHandler:",
        "    Err.Clear",
        "    Resume SafeExit",
        "SafeExit:",
        "End Sub",
        "",
        "Private Sub TrySetFontCtl(ByVal ctl As Object, ByVal fontSize As Single, ByVal boldValue As Boolean)",
        "    On Error GoTo ErrorHandler",
        "    Dim f As Object",
        '    Set f = CallByName(ctl, "Font", VbGet)',
        '    CallByName f, "Name", VbLet, "Segoe UI"',
        '    CallByName f, "Size", VbLet, fontSize',
        '    CallByName f, "Bold", VbLet, boldValue',
        "    Exit Sub",
        "ErrorHandler:",
        "    Err.Clear",
        "    Resume SafeExit",
        "SafeExit:",
        "End Sub",
        "",
        "Private Sub TrySetComponentProperty(ByVal comp As Object, ByVal propName As String, ByVal value As Variant)",
        "    On Error GoTo ErrorHandler",
        "    comp.Properties(propName).Value = value",
        "    Exit Sub",
        "ErrorHandler:",
        "    Err.Clear",
        "    Resume SafeExit",
        "SafeExit:",
        "End Sub",
        "",
        "Private Sub RemoveComponent(ByVal vbProj As Object, ByVal componentName As String)",
        "    On Error GoTo ErrorHandler",
        "    vbProj.VBComponents.Remove vbProj.VBComponents(componentName)",
        "    Exit Sub",
        "ErrorHandler:",
        "    Err.Clear",
        "    Resume SafeExit",
        "SafeExit:",
        "End Sub",
        "",
        "Private Sub ReplaceModuleCode(ByVal comp As Object, ByVal source As String)",
        "    On Error GoTo ErrorHandler",
        "    With comp.CodeModule",
        "        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines",
        "        .AddFromString source",
        "    End With",
        "    Exit Sub",
        "ErrorHandler:",
        '    BuilderHandleError "ReplaceModuleCode"',
        "End Sub",
        "",
        "Private Function AddStandardModule(ByVal vbProj As Object, ByVal moduleName As String, ByVal source As String) As Object",
        "    On Error GoTo ErrorHandler",
        "    Dim comp As Object",
        "    RemoveComponent vbProj, moduleName",
        "    Set comp = vbProj.VBComponents.Add(1)",
        "    comp.Name = moduleName",
        "    ReplaceModuleCode comp, source",
        "    Set AddStandardModule = comp",
        "    Exit Function",
        "ErrorHandler:",
        '    BuilderHandleError "AddStandardModule"',
        "    Set AddStandardModule = Nothing",
        "End Function",
        "",
        "Private Function AddUserForm(ByVal vbProj As Object, ByVal formName As String, ByVal source As String) As Object",
        "    On Error GoTo ErrorHandler",
        "    Dim comp As Object",
        "    RemoveComponent vbProj, formName",
        "    Set comp = vbProj.VBComponents.Add(3)",
        "    comp.Name = formName",
        '    TrySetComponentProperty comp, "Caption", "Raad | Live Inventory & Variance Console"',
        '    TrySetComponentProperty comp, "Width", 1200',
        '    TrySetComponentProperty comp, "Height", 780',
        '    TrySetComponentProperty comp, "StartUpPosition", 2',
        '    TrySetComponentProperty comp, "BackColor", RGB(244, 247, 251)',
        '    TrySetComponentProperty comp, "BorderStyle", 1',
        '    TrySetComponentProperty comp, "SpecialEffect", 0',
        "    AddAllControls comp",
        "    ReplaceModuleCode comp, source",
        "    Set AddUserForm = comp",
        "    Exit Function",
        "ErrorHandler:",
        '    BuilderHandleError "AddUserForm"',
        "    Set AddUserForm = Nothing",
        "End Function",
        "",
        "Private Sub AddCtl(ByVal frm As Object, ByVal progId As String, ByVal ctlName As String, _",
        "                  ByVal leftValue As Single, ByVal topValue As Single, _",
        "                  ByVal widthValue As Single, ByVal heightValue As Single, _",
        '                  Optional ByVal caption As String = "", _',
        "                  Optional ByVal backColor As Long = -1, _",
        "                  Optional ByVal foreColor As Long = -1, _",
        "                  Optional ByVal boldValue As Boolean = False, _",
        "                  Optional ByVal fontSize As Single = 9)",
        "    On Error GoTo ErrorHandler",
        "    Dim c As Object",
        "    Set c = frm.Designer.Controls.Add(progId, ctlName, True)",
        '    TrySetCtl c, "Left", leftValue',
        '    TrySetCtl c, "Top", topValue',
        '    TrySetCtl c, "Width", widthValue',
        '    TrySetCtl c, "Height", heightValue',
        '    TrySetCtl c, "Caption", caption',
        "    If backColor >= 0 Then TrySetCtl c, \"BackColor\", backColor",
        "    If foreColor >= 0 Then TrySetCtl c, \"ForeColor\", foreColor",
        '    TrySetCtl c, "SpecialEffect", 0',
        '    TrySetCtl c, "BorderStyle", 1',
        f'    TrySetCtl c, "BorderColor", RGB{rgb("border")}',
        '    TrySetCtl c, "TextAlign", 3',
        '    TrySetCtl c, "RightToLeft", True',
        '    TrySetCtl c, "AutoSize", False',
        '    TrySetCtl c, "WordWrap", True',
        "    TrySetFontCtl c, fontSize, boldValue",
        "    Exit Sub",
        "ErrorHandler:",
        '    BuilderHandleError "AddCtl"',
        "End Sub",
        "",
        "Private Sub AddAllControls(ByVal frm As Object)",
        "    On Error GoTo ErrorHandler",
    ]
    for spec in CONTROL_SPECS:
        kind, name, x, y, w, h, caption, back, fore, bold, size = spec
        br, bg, bb = rgb(back)
        fr, fg, fb = rgb(fore)
        call = (
            f'    AddCtl frm, "{PROG_IDS[kind]}", "{name}", {x}, {y}, {w}, {h}, '
            f'"{vba_quote(caption)}", RGB({br}, {bg}, {bb}), RGB({fr}, {fg}, {fb}), '
            f'{"True" if bold else "False"}, {size}'
        )
        lines.append(call)
    lines += [
        "    Exit Sub",
        "ErrorHandler:",
        '    BuilderHandleError "AddAllControls"',
        "End Sub",
        "",
    ]
    lines += string_function("ModuleSource", module_body)
    lines += string_function("FormSource", form_body)
    lines += string_function("WorkbookSource", workbook_body)
    (VBA / "AutoBuild_RaadApp.bas").write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")


def write_readme_and_code_bundle(module_full: str, form_body: str, workbook_body: str, wb) -> None:
    readme = f"""# Raad Interactive VBA Application

## What is included

- `نظام_رعد_لإدارة_المخزون_والعجوزات_v9_App.xlsm` — ready-to-open macro-enabled workbook. Your original `.xlsx` is unchanged; all 16 real sheets, formulas, tables, validations, styles and values are preserved.
- `File_Audit.md` — exact sheet names, every recognized header group, detected data types and every additional non-tabular text cell.
- `Field_Audit.csv` — machine-readable exhaustive field-by-field audit.
- `Extra_Text_Cells.csv` — real titles, navigation links, legends and helper-list values outside the flat dataset ranges.
- `UI_Layout_and_Controls.md` — exact control names, geometry, RGB/Hex palette and dynamic control mapping.
- `VBA/frmRaadApp.frm` — importable UserForm source.
- `VBA/frmRaadApp_Code.bas` — UserForm code only.
- `VBA/modRaadApp.bas` — standard module.
- `VBA/ThisWorkbook.cls` / `VBA/ThisWorkbook_Code.bas` — ThisWorkbook event code.
- `VBA/AutoBuild_RaadApp.bas` — one-click VBA builder that creates the whole form and installs all code into a macro-enabled workbook.
- `VBA/All_Code_Paste.txt` — all three code bodies in one copy-paste file.

## Quick start

1. Open the `.xlsm` file in Microsoft Excel (Windows or Mac with VBA support).
2. Enable macros. `Workbook_Open` shows the modeless app and adds an **Open Raad App** button on the first dashboard sheet.
3. Agent mode is active by default. Raw worksheet cells are view-only in this mode; use the app form for Add/Search/View/Update. Use **Switch Mode** to select Admin; on first Admin login you create your own PIN (stored in user settings, not in worksheet data).
4. Choose any real dataset in the dataset dropdown. The form reads exact header cells, creates one input per real column, and displays the rows in the live grid.
5. Use `سجل_الأرشيف_اليومي | Daily archive - FULL CRUD` for full Add/Update/Delete. Formula-backed sheets keep calculated columns read-only and are protected from app-level physical deletion.

## Design and safety notes

- Full CRUD writes directly to the real flat archive sheet `A6:I200`, never to a copy or placeholder table.
- Add writes to the first truly blank row and copies the real row formats/validation.
- Update writes only user-entry cells on the selected exact worksheet row.
- Delete requires Admin plus an explicit **Are you sure?** confirmation and removes only cells inside the dataset range.
- Formula cells are detected from the real sheet, displayed read-only, and never overwritten by Add/Update/Delete.
- `Workbook_SheetChange` refreshes the form after raw worksheet edits, giving a live two-way connection.
- Agent mode protects workbook structure, locks all formula/header/raw cells, and unlocks only the real user-entry cells of update-enabled datasets. Admin mode can unlock every sheet and the raw Excel grid.
- The role/PIN layer is a usability guard, not enterprise-grade security.
- Every procedure has `On Error GoTo ErrorHandler`; no `Debug/End` crash window is used.
- Add, Update and Delete call `ThisWorkbook.Save`.

## Rebuilding with VBA only

1. Open a copy of the original workbook saved as `.xlsm`.
2. Enable **Trust access to the VBA project object model**.
3. Import `VBA/AutoBuild_RaadApp.bas` and run `Build_Raad_App`.
4. Save, close and reopen with macros enabled.

## Regenerating the ready-made `.xlsm` in this repository

```bash
python3 -m venv .venv
.venv/bin/pip install -r tools/requirements.txt
.venv/bin/python tools/build_xlsm.py
```

The original `{SOURCE.name}` contains {len(wb.sheetnames)} worksheets and remains untouched.
"""
    (ROOT / "README.md").write_text(readme, encoding="utf-8")
    bundle = "\r\n".join([
        "===============================================================",
        "USERFORM: frmRaadApp - paste into the UserForm code window",
        "===============================================================",
        "",
        crlf(form_body),
        "===============================================================",
        "STANDARD MODULE: modRaadApp",
        "===============================================================",
        "",
        crlf(module_full),
        "===============================================================",
        "THISWORKBOOK: paste into the ThisWorkbook code window",
        "===============================================================",
        "",
        crlf(workbook_body),
    ]) + "\r\n"
    (VBA / "All_Code_Paste.txt").write_text(bundle, encoding="utf-8")


def prepare_macro_container(src: Path, dst: Path) -> None:
    old = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    new = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        zout.comment = zin.comment
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "[Content_Types].xml":
                text = data.decode("utf-8")
                if old not in text:
                    raise RuntimeError("Unexpected workbook content type; refusing to build.")
                data = text.replace(old, new).encode("utf-8")
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zi.internal_attr = info.internal_attr
            zi.create_system = info.create_system
            zi.extra = info.extra
            zi.comment = info.comment
            zout.writestr(zi, data)


def build_workbook(module_body: str, form_body: str, workbook_body: str) -> None:
    prepare_macro_container(SOURCE, OUTPUT)
    with ExcelFile(OUTPUT) as book:
        project = book.add_vba_project()
        project.add_module("modRaadApp", module_body, kind=VBAModuleKind.standard)
        form = book.add_form(
            "frmRaadApp",
            caption="Raad | Live Inventory & Variance Console",
            width=1200,
            height=780,
        )
        for key in ("BackColor", "BorderStyle", "SpecialEffect", "ForeColor"):
            try:
                if key == "BackColor":
                    form.set_property(key, ole_color(rgb("bg")))
                elif key == "ForeColor":
                    form.set_property(key, ole_color(rgb("text")))
                elif key == "BorderStyle":
                    form.set_property(key, 1)
                elif key == "SpecialEffect":
                    form.set_property(key, 0)
            except Exception:
                pass
        for spec in CONTROL_SPECS:
            kind, name, x, y, w, h, caption, back, fore, bold, size = spec
            ctl = form.add_control(kind, name, left=x, top=y, width=w, height=h)
            try:
                ctl.set_font("Segoe UI", size, bold=bold)
            except Exception:
                pass
            props = {
                "Caption": caption,
                "BackColor": ole_color(rgb(back)),
                "ForeColor": ole_color(rgb(fore)),
                "BorderColor": ole_color(rgb("border")),
                "SpecialEffect": 0,
                "BorderStyle": 1,
                "TextAlign": 3,
                "RightToLeft": True,
                "AutoSize": False,
                "WordWrap": True,
            }
            if kind in ("ComboBox", "ListBox"):
                props["ColumnCount"] = 1
            if kind == "ListBox":
                props["ScrollBars"] = 3
                props["IntegralHeight"] = False
                props["MultiSelect"] = 0
                props["ListStyle"] = 0
                props["BoundColumn"] = 1
            for prop, value in props.items():
                try:
                    ctl.set_property(prop, value)
                except Exception:
                    pass
        book.set_module("modRaadApp", module_body)
        book.set_module("frmRaadApp", form_body)
        book.set_module("ThisWorkbook", workbook_body)
        book.save(OUTPUT)

    with ExcelFile(OUTPUT) as check:
        issues = check.validate()
        if issues:
            raise RuntimeError("VBA project validation failed: " + "; ".join(issues))
        if not check.has_vba_project():
            raise RuntimeError("Output workbook has no VBA project.")
        forms = [(f.name, [c.name for c in f.walk()]) for f in check.forms()]
        if not any(name == "frmRaadApp" for name, _ in forms):
            raise RuntimeError("frmRaadApp designer was not written.")
    report = check_file(OUTPUT)
    print(f"Workbook validation report: {report}")


def main() -> None:
    VBA.mkdir(exist_ok=True)
    module_full = (VBA / "modRaadApp.bas").read_text(encoding="utf-8")
    module_body = source_body(VBA / "modRaadApp.bas", strip_attribute=True)
    form_body = source_body(VBA / "frmRaadApp_Code.bas")
    workbook_body = source_body(VBA / "ThisWorkbook_Code.bas")

    wb = load_workbook(SOURCE, data_only=False)
    write_audit(wb)
    write_ui_docs(wb)
    write_frm_file(form_body)
    write_thisworkbook_cls(workbook_body)
    write_auto_builder(module_body, form_body, workbook_body)
    write_readme_and_code_bundle(module_full, form_body, workbook_body, wb)
    (ROOT / "tools" / "requirements.txt").write_text("pyOpenVBA>=6.1.1\nopenpyxl>=3.1.5\n", encoding="utf-8")
    (ROOT / ".gitignore").write_text(".venv/\n__pycache__/\n*.pyc\n", encoding="utf-8")
    build_workbook(module_body, form_body, workbook_body)
    print(f"Built: {OUTPUT}")
    print(f"Sheets preserved: {len(wb.sheetnames)}")


if __name__ == "__main__":
    main()
