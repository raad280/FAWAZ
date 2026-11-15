# File Audit - Raad inventory workbook

Source file: `نظام_رعد_لإدارة_المخزون_والعجوزات_v9.xlsx`
Worksheet count: **16**

## Exact sheet names and order

1. `لوحة_التحكم_والرئيسية` — state `visible`, used range `B2:K28`
2. `محطة_الإدخال_اليومي` — state `visible`, used range `A1:G16`
3. `مخزون_البان_الموحد` — state `visible`, used range `A1:J14`
4. `التحليلات_الزمنية` — state `visible`, used range `A1:I17`
5. `سجل_الأرشيف_اليومي` — state `visible`, used range `A1:I16`
6. `المشروبات الغازية` — state `visible`, used range `A1:AE10`
7. `الفاخر` — state `visible`, used range `A1:Q22`
8. `المشاوي` — state `visible`, used range `A1:T24`
9. `برجر` — state `visible`, used range `A1:M26`
10. `المضغوط` — state `visible`, used range `A1:O26`
11. `الزربيان` — state `visible`, used range `A1:P13`
12. `الكنافة` — state `visible`, used range `A1:Q24`
13. `الثلاجة` — state `visible`, used range `A1:O20`
14. `البروست` — state `visible`, used range `A1:I21`
15. `أرشيف البروست` — state `visible`, used range `A1:Q390`
16. `بيتزا` — state `visible`, used range `A1:M15`

## Primary flat database selected for full CRUD

Sheet: `سجل_الأرشيف_اليومي`; header row **5**; data rows **6:200**; autofilter `A5:I200`; no formulas.

- **A** = `التاريخ`
- **B** = `القسم`
- **C** = `الصنف`
- **D** = `الوحدة`
- **E** = `المتاح الكلي`
- **F** = `الاستهلاك الفعلي`
- **G** = `المتبقي الفعلي`
- **H** = `صافي الفارق`
- **I** = `الحالة`

## Data types and real validations

- A in the daily archive is a real Excel **Date** (`yyyy-mm-dd`).
- E:H in the daily archive are **Numbers** (`#,##0.00` / signed variance formats).
- B:D and I are **Text**; I is a real **Dropdown/Status** column.
- `مخزون_البان_الموحد`!F11 validation list: `"نعم,لا"`.
- `سجل_الأرشيف_اليومي`!I6:I200 validation list: `"عجز,زيادة,مطابق"`.
- No currency-formatted columns were found in the workbook; quantity and variance fields are numeric, not currency.
- Formula-backed columns are detected dynamically at runtime and shown as locked, read-only inputs; the app only writes user-entry cells.
- `Field_Audit.csv` classifies every real dataset column as Text, Number, Date, Formula - read-only, or Dropdown/Status. Currency is explicitly reported as absent.

## Excel tables

- `الفاخر` :: `tbl_Fakher` :: `A2:I21`
- `المشاوي` :: `tbl_Mashawi` :: `A2:N23`
- `برجر` :: `tbl_Burger_Inv` :: `A14:G19`
- `المضغوط` :: `tbl_Madghout` :: `A3:I15`
- `الزربيان` :: `tbl_Zurbian` :: `A3:P12`
- `الكنافة` :: `tbl_Kunafa_Inv` :: `A14:G24`
- `الثلاجة` :: `tbl_Fridge` :: `A2:O19`
- `البروست` :: `tbl_Broast_Sales` :: `B5:H12`
- `البروست` :: `tbl_Broast_Inv` :: `B16:H21`
- `بيتزا` :: `tbl_Pizza_Inv` :: `A13:G15`

## Formula counts by sheet

- `لوحة_التحكم_والرئيسية`: 72 formulas
- `محطة_الإدخال_اليومي`: 44 formulas
- `مخزون_البان_الموحد`: 13 formulas
- `التحليلات_الزمنية`: 99 formulas
- `سجل_الأرشيف_اليومي`: 0 formulas
- `المشروبات الغازية`: 12 formulas
- `الفاخر`: 23 formulas
- `المشاوي`: 42 formulas
- `برجر`: 52 formulas
- `المضغوط`: 14 formulas
- `الزربيان`: 12 formulas
- `الكنافة`: 54 formulas
- `الثلاجة`: 74 formulas
- `البروست`: 27 formulas
- `أرشيف البروست`: 234 formulas
- `بيتزا`: 18 formulas

## Every recognized header group and every column

Blank header cells are shown as `[blank]` to prove the exact column position. The machine-readable `Field_Audit.csv` contains one row per real field, including detected type, number format, formula example, dropdown values, and a real sample value.

### `DASH_KPI_LABELS` — `لوحة_التحكم_والرئيسية` | Dashboard KPI labels
Header row **13**, data rows **14:14**, columns **B:J**. Add=False, Update=False, Delete=False.

**B**=`إجمالي استهلاك البان الموحد` | **C**=`[blank]` | **D**=`فارق مخزون البان الموحد` | **E**=`[blank]` | **F**=`عدد الأصناف ذات العجز` | **G**=`[blank]` | **H**=`عدد الأصناف ذات الزيادة` | **I**=`[blank]` | **J**=`عدد الأصناف المطابقة 100%`

### `DASH_KPI_TABLE` — `لوحة_التحكم_والرئيسية` | Dashboard KPI table
Header row **17**, data rows **18:28**, columns **B:K**. Add=False, Update=True, Delete=False.

**B**=`القسم` | **C**=`الصنف / المادة الخام` | **D**=`وحدة القياس` | **E**=`إجمالي المتاح` | **F**=`الاستهلاك الفعلي` | **G**=`المتبقي (جرد / نظري)` | **H**=`صافي الفارق (عجز / زيادة)` | **I**=`نسبة الانحراف %` | **J**=`الحالة الرقابية` | **K**=`ملاحظات التدقيق`

### `DAILY_INPUT` — `محطة_الإدخال_اليومي` | Daily input station
Header row **5**, data rows **6:16**, columns **A:G**. Add=False, Update=True, Delete=False.

**A**=`القسم` | **B**=`المادة الخام / الصنف الرئيسي` | **C**=`وحدة القياس` | **D**=`الرصيد السابق (افتتاحي)` | **E**=`المنصرف الجديد (وارد)` | **F**=`المبيعات / الاستهلاك الفعلي` | **G**=`الجرد الفعلي (نهاية اليوم)`

### `BANANA_MASTER` — `مخزون_البان_الموحد` | Unified bun master
Header row **5**, data rows **6:6**, columns **A:J**. Add=False, Update=True, Delete=False.

**A**=`المادة الخام` | **B**=`الرصيد السابق (حبة)` | **C**=`المنصرف الجديد (حبة)` | **D**=`إجمالي المتاح` | **E**=`تالف / هدر (إدخال)` | **F**=`إجمالي استهلاك الأقسام` | **G**=`المتبقي النظري` | **H**=`الجرد الفعلي (إدخال)` | **I**=`صافي الفارق (عجز / زيادة)` | **J**=`الحالة`

### `BANANA_DETAIL` — `مخزون_البان_الموحد` | Unified bun distribution
Header row **10**, data rows **11:14**, columns **A:F**. Add=False, Update=True, Delete=False.

**A**=`القسم المستهلك` | **B**=`مصدر البيانات` | **C**=`تفصيل الأصناف` | **D**=`الكمية المستهلكة (حبة)` | **E**=`نسبة الاستهلاك %` | **F**=`تفعيل احتساب بان الزنجر 1x1؟`

### `TIME_ANALYTICS` — `التحليلات_الزمنية` | Time analytics
Header row **6**, data rows **7:17**, columns **A:I**. Add=False, Update=False, Delete=False.

**A**=`التاريخ` | **B**=`القسم` | **C**=`الصنف` | **D**=`الوحدة` | **E**=`إجمالي المتاح` | **F**=`الاستهلاك الفعلي` | **G**=`المتبقي الفعلي` | **H**=`صافي الفارق (عجز/زيادة)` | **I**=`الحالة`

### `MAIN_ARCHIVE` — `سجل_الأرشيف_اليومي` | Daily archive - FULL CRUD
Header row **5**, data rows **6:200**, columns **A:I**. Add=True, Update=True, Delete=True.

**A**=`التاريخ` | **B**=`القسم` | **C**=`الصنف` | **D**=`الوحدة` | **E**=`المتاح الكلي` | **F**=`الاستهلاك الفعلي` | **G**=`المتبقي الفعلي` | **H**=`صافي الفارق` | **I**=`الحالة`

### `BEV_PEPSI` — `المشروبات الغازية` | Beverages - block 1
Header row **3**, data rows **4:4**, columns **A:I**. Add=False, Update=True, Delete=False.

**A**=`الرصيد ` | **B**=`الصرف الجديد ` | **C**=`لحقة صرف` | **D**=`اجمالي الرصيد` | **E**=`المتبقي` | **F**=`الكمية المباعة للمقارنه` | **G**=`المباع في النظام` | **H**=`العجز ` | **I**=`الزيادة`

### `BEV_WATER` — `المشروبات الغازية` | Beverages - block 2
Header row **8**, data rows **9:9**, columns **A:I**. Add=False, Update=True, Delete=False.

**A**=`الرصيد ` | **B**=`الصرف الجديد ` | **C**=`لحقة صرف` | **D**=`اجمالي الرصيد` | **E**=`المتبقي` | **F**=`الكمية المباعة للمقارنه` | **G**=`المباع في النظام` | **H**=`العجز ` | **I**=`الزيادة`

### `FAKHER_FULL` — `الفاخر` | Full section grid
Header row **2**, data rows **3:21**, columns **A:M**. Add=False, Update=True, Delete=False.

**A**=`الكمية المباعة` | **B**=`الصنف` | **C**=`اجمالي مبيعات اللحم` | **D**=`اجمالي مبيعات الدجاج` | **E**=`المصروف الدجاج` | **F**=`المصروف اللحم` | **G**=`الرصيد` | **H**=`الرصيد من واقع الجرد` | **I**=`مردود مبيعات` | **J**=`زيادة ` | **K**=`عجز ` | **L**=`زيادة ` | **M**=`عجز `

### `MASHAWI_FULL` — `المشاوي` | Full section grid
Header row **2**, data rows **3:23**, columns **A:T**. Add=False, Update=True, Delete=False.

**A**=`الكمية المباعة` | **B**=`الصنف` | **C**=`الرصيد السابق` | **D**=`اجمالي مبيعات الدجاج` | **E**=`مبيعات المسحب` | **F**=`مبيعات اللحم` | **G**=`مبيعات الفروج` | **H**=`المصروف الدجاج` | **I**=`الرصيد` | **J**=`الرصيد من واقع الجرد` | **K**=`مردود مبيعات` | **L**=`زيادة` | **M**=`عجز` | **N**=`المحول للفروج` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`1` | **R**=`سندوتش كباب بلادي` | **S**=`[blank]` | **T**=`[blank]`

### `BURGER_SALES` — `برجر` | Sales grid
Header row **1**, data rows **2:12**, columns **A:G**. Add=False, Update=True, Delete=False.

**A**=`رئيسي` | **B**=`الصنف` | **C**=`الجبن` | **D**=`زنجر` | **E**=`بان` | **F**=`لحم ` | **G**=`دجاج`

### `BURGER_INV` — `برجر` | Inventory grid
Header row **14**, data rows **15:19**, columns **A:H**. Add=False, Update=True, Delete=False.

**A**=`المكون / الصنف` | **B**=`الرصيد السابق` | **C**=`المصروف الجديد` | **D**=`إجمالي المتاح` | **E**=`الاستهلاك الفعلي (المباع)` | **F**=`المتبقي الفعلي (الجرد)` | **G**=`النتيجة (عجز / زيادة)` | **H**=`صافي الفارق (رقمي)`

### `MADGHOUT_FULL` — `المضغوط` | Full section grid
Header row **3**, data rows **4:15**, columns **A:O**. Add=False, Update=True, Delete=False.

**A**=`الكمية المباعة` | **B**=`الصنف` | **C**=`اجمالي مبيعات الدجاج` | **D**=`المصروف الدجاج` | **E**=`الرصيد` | **F**=`الرصيد من واقع الجرد` | **G**=`مردود مبيعات` | **H**=`زيادة` | **I**=`عجز` | **J**=`عقدة` | **K**=`[blank]` | **L**=`صافي العجز بعد خصم المحروقات` | **M**=`عجز المحروقات ` | **N**=`[blank]` | **O**=`المرحل للمحروقات`

### `ZURBIAN_FULL` — `الزربيان` | Full section grid
Header row **3**, data rows **4:12**, columns **A:P**. Add=False, Update=True, Delete=False.

**A**=`الكمية المباعة` | **B**=`الصنف` | **C**=`اجمالي مبيعات اللحم` | **D**=`ملاحظات` | **E**=`اجمالي مبيعات الدجاج` | **F**=`المصروف الدجاج` | **G**=`المصروف اللحم` | **H**=`الرصيد` | **I**=`المتبقي من الدجاج` | **J**=`المتبقي من اللحم` | **K**=`الرصيد من واقع الجرد` | **L**=`مردود مبيعات` | **M**=`زيادة دجاج` | **N**=`عجز دجاج` | **O**=`زيادة لحم` | **P**=`عجز لحم`

### `KUNAFA_GRID` — `الكنافة` | Formula grid
Header row **2**, data rows **3:11**, columns **A:Q**. Add=False, Update=True, Delete=False.

**A**=`الكنافة` | **B**=`='لوحة_التحكم_والرئيسية'!D4` | **C**=`[blank]` | **D**=`[blank]` | **E**=`الرصيد بعد المبيعات` | **F**=`الرصيد الفعلي من واقع الجرد ` | **G**=`[blank]` | **H**=`عجز ` | **I**=`زيادة` | **J**=`المبيعات` | **K**=`الكمية المباعة من واقع النظام` | **L**=`المبيعات من واقع النظام` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `KUNAFA_INV` — `الكنافة` | Inventory grid
Header row **14**, data rows **15:24**, columns **A:H**. Add=False, Update=True, Delete=False.

**A**=`الصنف` | **B**=`الرصيد السابق` | **C**=`الوارد / المصروف` | **D**=`إجمالي المتاح` | **E**=`المباع في النظام` | **F**=`المتبقي الفعلي (الجرد)` | **G**=`النتيجة (عجز / زيادة)` | **H**=`صافي الفارق (رقمي)`

### `FRIDGE_FULL` — `الثلاجة` | Fridge grid
Header row **2**, data rows **3:19**, columns **A:O**. Add=False, Update=True, Delete=False.

**A**=`رصيد الدجاج في الثلاجة` | **B**=`الرصيد السابق` | **C**=`مورد للثلاجة` | **D**=`البروست` | **E**=`الفاخر` | **F**=`المضغوط` | **G**=`المشاوي` | **H**=`الزربيان` | **I**=`البرجر` | **J**=`الشوارما` | **K**=`المتبقي في الثلاجة` | **L**=`الرصيد من واقع الجرد في الثلاجة` | **M**=`العجز` | **N**=`الزيادة` | **O**=`حالة الجرد`

### `BROAST_SALES` — `البروست` | Broast sales
Header row **5**, data rows **6:12**, columns **B:H**. Add=False, Update=True, Delete=False.

**B**=`الكمية المباعة` | **C**=`اسم الوجبة / الصنف` | **D**=`عدد القطع في الوجبة` | **E**=`إجمالي قطع الدجاج` | **F**=`ملاحظات التشغيل` | **G**=`كود الصنف` | **H**=`التصنيف`

### `BROAST_INV` — `البروست` | Broast inventory
Header row **16**, data rows **17:21**, columns **B:I**. Add=False, Update=True, Delete=False.

**B**=`المكون / الصنف` | **C**=`الرصيد السابق` | **D**=`المصروف الجديد` | **E**=`إجمالي المتاح` | **F**=`الاستهلاك الفعلي (المباع)` | **G**=`المتبقي الفعلي (الجرد)` | **H**=`النتيجة (عجز / زيادة)` | **I**=`صافي الفارق (رقمي)`

### `PIZZA_SALES` — `بيتزا` | Pizza sales
Header row **1**, data rows **2:9**, columns **A:D**. Add=False, Update=True, Delete=False.

**A**=`الكمية` | **B**=`الصنف` | **C**=`الجبن بالحبة` | **D**=`اجمالي الجبن`

### `PIZZA_INV` — `بيتزا` | Pizza inventory
Header row **13**, data rows **14:15**, columns **A:H**. Add=False, Update=True, Delete=False.

**A**=`المادة الخام` | **B**=`الرصيد السابق (جرام/حبة)` | **C**=`المصروف الجديد` | **D**=`إجمالي المتاح` | **E**=`الاستهلاك الفعلي (المباع)` | **F**=`المتبقي الفعلي (الجرد)` | **G**=`النتيجة (عجز / زيادة)` | **H**=`صافي الفارق (رقمي)`

### `BROAST_HIST_01` — `أرشيف البروست` | Historical daily block 01
Header row **5**, data rows **6:32**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`[blank]` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الثلاثاء` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_02` — `أرشيف البروست` | Historical daily block 02
Header row **33**, data rows **34:60**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`20` | **O**=`1500` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_03` — `أرشيف البروست` | Historical daily block 03
Header row **61**, data rows **62:89**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الثلاثاء` | **I**=`[blank]` | **J**=`[blank]` | **K**=`=K59/K60` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_04` — `أرشيف البروست` | Historical daily block 04
Header row **90**, data rows **91:117**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_05` — `أرشيف البروست` | Historical daily block 05
Header row **118**, data rows **119:144**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_06` — `أرشيف البروست` | Historical daily block 06
Header row **145**, data rows **146:170**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_07` — `أرشيف البروست` | Historical daily block 07
Header row **171**, data rows **172:198**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_08` — `أرشيف البروست` | Historical daily block 08
Header row **199**, data rows **200:228**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_09` — `أرشيف البروست` | Historical daily block 09
Header row **229**, data rows **230:255**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_10` — `أرشيف البروست` | Historical daily block 10
Header row **256**, data rows **257:283**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_11` — `أرشيف البروست` | Historical daily block 11
Header row **284**, data rows **285:310**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_12` — `أرشيف البروست` | Historical daily block 12
Header row **311**, data rows **312:337**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_13` — `أرشيف البروست` | Historical daily block 13
Header row **338**, data rows **339:365**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

### `BROAST_HIST_14` — `أرشيف البروست` | Historical daily block 14
Header row **366**, data rows **367:390**, columns **A:Q**. Add=False, Update=False, Delete=False.

**A**=`[blank]` | **B**=`الرصيد بالاكياس` | **C**=`القطع` | **D**=`الرصيد بالقطع` | **E**=`الصرف الجديد` | **F**=`رصيد الصرف الجديد بالقطع` | **G**=`رصيد الصرف الجديد بالاكياس` | **H**=`اجمالي قطع البروست المرحلة ليوم الجمعة` | **I**=`[blank]` | **J**=`[blank]` | **K**=`[blank]` | **L**=`[blank]` | **M**=`[blank]` | **N**=`[blank]` | **O**=`[blank]` | **P**=`[blank]` | **Q**=`[blank]`

## Additional non-tabular text cells outside recognized dataset ranges

These are real workbook titles, navigation links, legends and helper lists. They are not column headers inside a flat dataset, but they are listed so the audit is exhaustive from the first used column to the last. `Extra_Text_Cells.csv` contains the same rows in machine-readable form.

### `لوحة_التحكم_والرئيسية`

- `B2` = `نظام رعد الذكي لإدارة المخزون والعجوزات اليومية (الإصدار التجريبي v8.0)`
- `B5` = `تاريخ التشغيل الموحد:`
- `E5` = `دليل الألوان:  [ الخلايا الصفراء = إدخال يدوي يومي ]   \|   [ الخلايا البيضاء/الملونة = معادلات محمية تلقائية ]`
- `B7` = `أولاً: لوحة التنقل السريع بين الأقسام والمخازن (اضغط للانتقال المباشر)`
- `B8` = `⬅ محطة الإدخال اليومي`
- `D8` = `⬅ مخزون البان الموحد`
- `F8` = `⬅ مخزون الثلاجة المركزي`
- `H8` = `⬅ قسم البروست`
- `J8` = `⬅ قسم البرجر`
- `B9` = `⬅ قسم الفاخر`
- `D9` = `⬅ قسم المشاوي`
- `F9` = `⬅ قسم المضغوط`
- `H9` = `⬅ قسم الزربيان`
- `J9` = `⬅ الكنافة والحلويات`
- `B10` = `⬅ البيتزا والمشروبات`
- `E10` = `📊 التحليلات الزمنية`
- `H10` = `📁 سجل الأرشيف اليومي`
- `B12` = `ثانياً: مؤشرات الأداء والرقابة الفورية (KPIs)`
- `B16` = `ثالثاً: التقرير التفصيلي الموحد للعجوزات والزيادات عبر جميع الأقسام`

### `محطة_الإدخال_اليومي`

- `A1` = `🏠 الرئيسية`
- `B2` = `محطة الإدخال اليومي المركزية (إدخال مبيعات وجرد كافة الأقسام في شاشة واحدة)`

### `مخزون_البان_الموحد`

- `A1` = `🏠 الرئيسية`
- `B2` = `نظام إدارة مخزون خبز البان الموحد وتوزيعه بين الأقسام`
- `A9` = `تفصيل توزيع واستهلاك خبز البان حسب الأقسام (مرتبط تلقائياً بمبيعات الكاشير)`

### `التحليلات_الزمنية`

- `A1` = `🏠 الرئيسية`
- `B2` = `لوحة التحليلات والمقارنات الزمنية المتقدمة (تتبع تطور العجوزات والهدر عبر الأيام)`
- `B5` = `أولاً: جدول ملخص السجلات التاريخية من الأرشيف ومؤشرات التغير الزمني`

### `سجل_الأرشيف_اليومي`

- `A1` = `🏠 الرئيسية`
- `B2` = `سجل الأرشيف اليومي الموحد للعجوزات والجرد (قاعدة البيانات التاريخية)`

### `المشروبات الغازية`

- `M1` = `🏠 الرئيسية`
- `A2` = `البيبسي`
- `A7` = `الماء`
- `AE7` = `كنافه`
- `AE8` = `كنافة نوتيلا`
- `AE9` = `عريكه`
- `AE10` = `كاب سويت`

### `الفاخر`

- `A1` = `🏠 الرئيسية`
- `B1` = `رئيسي`
- `J1` = `الدجاج`
- `L1` = `لحم`

### `المشاوي`

- `A1` = `🏠 الرئيسية`

### `برجر`

- `M1` = `🏠 الرئيسية`
- `A13` = `جدول مطابقة الجرد والعجز / الزيادة - قسم البرجر`

### `المضغوط`

- `A1` = `🏠 الرئيسية`
- `A2` = `رئيسي`
- `B19` = `مضغوط لحم`
- `O19` = `نص دجاج مضغوط مع الرز`
- `B20` = `نص دجاج مضغوط مع الرز`
- `O20` = `رز مضغوط ساده`
- `B21` = `رز مضغوط ساده`
- `O21` = `رز بسمتي`
- `B22` = `رز بسمتي`
- `O22` = `ربع مندي دجاج مع الرز`
- `B23` = `ربع مندي دجاج مع الرز`
- `O23` = `ربع دجاج مضغوط مع الرز`
- `B24` = `ربع دجاج مضغوط مع الرز`
- `O24` = `طباخه علبة`
- `B25` = `نص دجاج مضغوط سادة`
- `O25` = `نص دجاج مضغوط سادة`
- `B26` = `رز مضغوط تغذية موظفين مجاناً`
- `O26` = `رز مضغوط تغذية موظفين مجاناً`

### `الزربيان`

- `A1` = `🏠 الرئيسية`
- `B2` = `رئيسي`

### `الكنافة`

- `A1` = `رئيسي`
- `G1` = `مردود مبيعات`
- `M1` = `🏠 الرئيسية`
- `A13` = `جدول جرد وحساب العجز والزيادة لجميع أصناف الحلويات`

### `الثلاجة`

- `M1` = `🏠 الرئيسية`
- `A20` = `الإجمالي العام`

### `البروست`

- `A1` = `🏠 الرئيسية`
- `B2` = `قسم البروست - نظام حساب استهلاك القطع ومطابقة الجرد اليومي`
- `B4` = `أولاً: مبيعات وجبات البروست وحساب إجمالي القطع المستهلكة`
- `C13` = `إجمالي المبيعات والقطع المستهلكة`
- `D13` = `-`
- `F13` = `مرتبط تلقائياً بجدول الجرد أدناه`
- `B15` = `ثانياً: جدول مطابقة الجرد وحساب العجز والزيادة - قسم البروست (مرتبط بلوحة التحكم)`

### `أرشيف البروست`

- `A1` = `أرشيف سجلات وجرد البروست اليومية السابقة (مرجع تاريخي كامل)`
- `Q1` = `🏠 الرئيسية`
- `A2` = `تحتوي هذه الورقة على كافة السجلات واليوميات القديمة لقسم البروست مرتبة زمنياً للرجوع إليها في أي وقت`
- `A4` = `البروست `

### `بيتزا`

- `M1` = `🏠 الرئيسية`
- `A10` = `إجمالي استهلاك الجبن (جرام)`
- `A12` = `جدول مطابقة الجرد والعجز / الزيادة - قسم البيتزا`
