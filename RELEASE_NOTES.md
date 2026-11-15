# Raad Interactive VBA Application - v1.0

Built directly from the real `نظام_رعد_لإدارة_المخزون_والعجوزات_v9.xlsx` workbook (16 sheets). No placeholder data.

## Download
- `Raad_App_Deliverables.zip` - complete bundle: macro-enabled workbook, Step 1 audits, Step 2 UI/control/RGB docs, Step 3 VBA sources, one-click AutoBuilder, regenerating script, original workbook and `START_HERE.txt`.
- `Raad_App_Workbook_Only.zip` - just the `_App.xlsm`, `OPEN_ME.txt` and `README.md`.
- `نظام_رعد_لإدارة_المخزون_والعجوزات_v9_App.xlsm` - the ready-to-open Excel app itself.

## Quick start
1. Download and extract a ZIP.
2. Open the file ending in `_App.xlsm` (not the original `.xlsx`).
3. If Windows blocks it: right-click > Properties > Unblock > Apply.
4. In Microsoft Excel click **Enable Macros / Enable Content**.
5. The form opens on `Workbook_Open`; otherwise click **Open Raad App** on sheet 1 or run `ShowRaadApp` with Alt+F8.
6. Agent mode is default. Switch to Admin for Delete/unlock; the first Admin login creates your own PIN.
7. Full Add/Update/Delete dataset: **Daily archive - FULL CRUD** (`سجل_الأرشيف_اليومي`).

## What is preserved
All 16 real sheets, exact headers, 10 Excel tables, 2 validation lists and 790 formulas. The original `.xlsx` is untouched. The `.xlsm` adds only `codeName` attributes to worksheet XML plus `xl/vbaProject.bin` and the required OPC parts.

## Validation performed
- `pyOpenVBA.validate()` returned no issues; `check_file()` reported a complete project.
- `olevba` extracts `ThisWorkbook`, `modRaadApp` and `frmRaadApp`.
- Designer contains all 36 static controls; dynamic exact-header fields are created at runtime.
- `openpyxl keep_vba=True` confirms 16 sheets, exact archive headers, 10 tables, 2 validations, 790 formulas.
- All 127 VBA procedures have `On Error GoTo ErrorHandler`; no `On Error Resume Next`; balanced blocks.

Note: this Linux sandbox has no Excel/VBA compiler, so the VBA runtime was not executed here. Open in Microsoft Excel and enable macros.
