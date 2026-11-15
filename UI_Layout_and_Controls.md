# UI Layout, Exact Control Names and RGB Design Settings

All controls use `Segoe UI`. The code applies `SpecialEffect = 0 - fmSpecialEffectFlat` and `BorderStyle = 1 - fmBorderStyleSingle` with safe property setters so unsupported control properties never raise an error.

## Exact palette

| Role | Hex | RGB | VBA |
|---|---|---|---|
| Form background | `#F4F7FB` | `RGB(244, 247, 251)` | `RGB(244, 247, 251)` |
| Header / sidebar | `#102A43` | `RGB(16, 42, 67)` | `RGB(16, 42, 67)` |
| Sidebar deep panel | `#0C1F32` | `RGB(12, 31, 50)` | `RGB(12, 31, 50)` |
| Cards and input boxes | `#FFFFFF` | `RGB(255, 255, 255)` | `RGB(255, 255, 255)` |
| Input surfaces / button text | `#FFFFFF` | `RGB(255, 255, 255)` | `RGB(255, 255, 255)` |
| Primary text | `#0F172A` | `RGB(15, 23, 42)` | `RGB(15, 23, 42)` |
| Muted text / read-only | `#64748B` | `RGB(100, 116, 139)` | `RGB(100, 116, 139)` |
| Single-pixel borders | `#CBD5E1` | `RGB(203, 213, 225)` | `RGB(203, 213, 225)` |
| Formula/read-only inputs | `#F1F5F9` | `RGB(241, 245, 249)` | `RGB(241, 245, 249)` |
| Primary blue - Add / Save | `#2563EB` | `RGB(37, 99, 235)` | `RGB(37, 99, 235)` |
| Primary hover | `#1D4ED8` | `RGB(29, 78, 216)` | `RGB(29, 78, 216)` |
| Update | `#F59E0B` | `RGB(245, 158, 11)` | `RGB(245, 158, 11)` |
| Update hover | `#D97706` | `RGB(217, 119, 6)` | `RGB(217, 119, 6)` |
| Delete | `#DC2626` | `RGB(220, 38, 38)` | `RGB(220, 38, 38)` |
| Delete hover | `#B91C1C` | `RGB(185, 28, 28)` | `RGB(185, 28, 28)` |
| Clear / neutral | `#64748B` | `RGB(100, 116, 139)` | `RGB(100, 116, 139)` |
| Neutral hover | `#475569` | `RGB(71, 85, 105)` | `RGB(71, 85, 105)` |
| Success status | `#16A34A` | `RGB(22, 163, 74)` | `RGB(22, 163, 74)` |
| Card outline | `#E2E8F0` | `RGB(226, 232, 240)` | `RGB(226, 232, 240)` |

## Static UserForm shell (`frmRaadApp`)

Geometry is in points on a 1200 x 780 form. `Back` and `Fore` refer to the palette keys above.

| (Name) | Type | Left | Top | Width | Height | Caption | Back | Fore |
|---|---|---:|---:|---:|---:|---|---|---|
| `fraHeader` | Frame | 0 | 0 | 1200 | 64 | — | `header` | `white` |
| `fraSidebar` | Frame | 970 | 76 | 214 | 622 | — | `header_dark` | `white` |
| `fraToolbar` | Frame | 16 | 76 | 938 | 68 | — | `card` | `text` |
| `fraEditor` | Frame | 16 | 466 | 938 | 210 | — | `card` | `text` |
| `fraButtons` | Frame | 16 | 684 | 938 | 48 | — | `card` | `text` |
| `lblTitle` | Label | 24 | 8 | 620 | 26 | Raad Inventory & Daily Variance | `header` | `white` |
| `lblSubTitle` | Label | 24 | 34 | 760 | 20 | Live two-way connection to your real worksheets, fields, formulas and dropdowns | `header` | `border` |
| `lblModeBadge` | Label | 970 | 16 | 90 | 28 | AGENT | `primary` | `white` |
| `lblClock` | Label | 1070 | 16 | 110 | 28 | — | `header` | `white` |
| `lblNav` | Label | 986 | 92 | 180 | 22 | Workspace / Role | `header_dark` | `white` |
| `lblDatasetInfo` | Label | 986 | 226 | 180 | 62 | Select a real sheet dataset | `header_dark` | `white` |
| `lblRecordCount` | Label | 986 | 466 | 180 | 22 | Records: 0 | `header_dark` | `white` |
| `lblSelectedRow` | Label | 986 | 494 | 180 | 22 | Selected row: none | `header_dark` | `white` |
| `lblSearch` | Label | 28 | 88 | 240 | 18 | Live search | `card` | `text` |
| `lblFilterField` | Label | 294 | 88 | 180 | 18 | Filter by real column | `card` | `text` |
| `lblFilterValue` | Label | 490 | 88 | 180 | 18 | Filter value | `card` | `text` |
| `lblEditorTitle` | Label | 28 | 474 | 400 | 20 | Record fields - exact worksheet headers | `card` | `text` |
| `lblStatus` | Label | 16 | 738 | 1168 | 26 | Ready. | `card` | `success` |
| `cboRole` | ComboBox | 986 | 120 | 180 | 26 | — | `white` | `text` |
| `cboDataset` | ComboBox | 986 | 194 | 180 | 26 | — | `white` | `text` |
| `cboFilterField` | ComboBox | 294 | 108 | 180 | 26 | — | `white` | `text` |
| `cboFilterValue` | ComboBox | 490 | 108 | 180 | 26 | — | `white` | `text` |
| `txtSearch` | TextBox | 28 | 108 | 250 | 26 | — | `white` | `text` |
| `lstHeaders` | ListBox | 16 | 152 | 938 | 24 | — | `header` | `white` |
| `lstData` | ListBox | 16 | 176 | 938 | 282 | — | `white` | `text` |
| `btnSwitchMode` | CommandButton | 986 | 152 | 180 | 30 | Switch Mode | `primary` | `white` |
| `btnRefresh` | CommandButton | 986 | 298 | 180 | 30 | Refresh | `primary` | `white` |
| `btnSheetView` | CommandButton | 986 | 336 | 180 | 30 | Switch to Excel Sheet | `primary` | `white` |
| `btnUnlockAll` | CommandButton | 986 | 380 | 180 | 30 | Unlock Sheets | `primary` | `white` |
| `btnLockForAgent` | CommandButton | 986 | 418 | 180 | 30 | Lock For Agent | `gray` | `white` |
| `btnClearFilter` | CommandButton | 686 | 108 | 120 | 26 | Clear Filter | `gray` | `white` |
| `btnAdd` | CommandButton | 28 | 694 | 140 | 30 | Add | `primary` | `white` |
| `btnUpdate` | CommandButton | 180 | 694 | 140 | 30 | Update | `orange` | `white` |
| `btnDelete` | CommandButton | 332 | 694 | 140 | 30 | Delete | `danger` | `white` |
| `btnClear` | CommandButton | 484 | 694 | 140 | 30 | Clear | `gray` | `white` |
| `btnClose` | CommandButton | 800 | 694 | 140 | 30 | Close | `gray` | `white` |

## Dynamic exact-field controls

For every selected dataset, `BuildFieldMetadata` reads the exact header text from the real worksheet and `BuildFieldControls` creates these controls inside `fraEditor`:

- Label: `lblF1` ... `lblFn` — caption is the exact worksheet header, or `[ColumnLetter] (no header)` when the header cell is blank.
- Text input: `txtF1` ... `txtFn` — Text, Date, Number and read-only Formula fields.
- Combo input: `cboF1` ... `cboFn` — only for real Dropdown/Status columns; items come from the worksheet validation list.

For the default full-CRUD sheet `سجل_الأرشيف_اليومي`, the exact mapping is:

| Column | Exact header | Control | Type |
|---|---|---|---|
| A | `التاريخ` | `txtF1` | Date |
| B | `القسم` | `txtF2` | Text |
| C | `الصنف` | `txtF3` | Text |
| D | `الوحدة` | `txtF4` | Text |
| E | `المتاح الكلي` | `txtF5` | Number |
| F | `الاستهلاك الفعلي` | `txtF6` | Number |
| G | `المتبقي الفعلي` | `txtF7` | Number |
| H | `صافي الفارق` | `txtF8` | Number |
| I | `الحالة` | `cboF9` | Dropdown/Status |

Formula columns are created as `txtF…` controls with `Locked = True`, a `#F1F5F9` background, and a read-only tooltip. They are never written by Add/Update/Delete.

## Behavior matrix

| Feature | Agent | Admin |
|---|---|---|
| Add / Search / View / Update | Yes on allowed datasets | Yes |
| Delete | Button disabled | Enabled only on the flat main archive; confirmation required |
| Workbook structure | Protected | Unprotected by `Unlock Sheets` |
| Formula cells | Locked in raw sheet view | Editable only in raw Excel view |
| Real Excel sheet view | Available; raw cells are view-only while the Agent lock is active | Available with fully unlocked sheets |

Fixed formula layouts and historical broast blocks are read-only/controlled-update datasets inside the app. Add and physical Delete are intentionally disabled there so your existing formulas, merges and sheet structure remain intact. The flat `سجل_الأرشيف_اليومي` sheet provides complete Add/Update/Delete CRUD.