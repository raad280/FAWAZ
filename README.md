# Raad Interactive VBA Application

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

The original `نظام_رعد_لإدارة_المخزون_والعجوزات_v9.xlsx` contains 16 worksheets and remains untouched.
