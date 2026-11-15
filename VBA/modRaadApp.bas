Attribute VB_Name = "modRaadApp"
Option Explicit

'===============================================================================
' Raad Inventory App - shared configuration, startup, role and protection logic.
' All worksheet names and field captions are read from the real workbook at run
' time. This source deliberately contains no hard-coded Arabic field names.
'===============================================================================

Public Type AppDataset
    Key As String
    DisplayName As String
    SheetName As String
    HeaderRow As Long
    FirstCol As Long
    LastCol As Long
    FirstDataRow As Long
    LastDataRow As Long
    AllowAdd As Boolean
    AllowUpdate As Boolean
    AllowDelete As Boolean
    RequiredCols As String
    Description As String
End Type

Public Type AppField
    Index As Long
    ColNumber As Long
    ColLetter As String
    Header As String
    LabelText As String
    Kind As String
    InputName As String
    Editable As Boolean
    ValidationList As String
End Type

Public Const ROLE_AGENT As String = "AGENT"
Public Const ROLE_ADMIN As String = "ADMIN"

Public gDatasets() As AppDataset
Public gDatasetCount As Long
Public gRole As String
Public gAppWriting As Boolean
Public gWorkbookLocked As Boolean

Private Const SETTINGS_APP As String = "RaadInventoryApp"
Private Const SETTINGS_SECTION As String = "Security"
Private Const SETTINGS_KEY As String = "AdminPIN"

Public Sub StartApplication()
    On Error GoTo ErrorHandler
    EnsureDatasets
    gRole = ROLE_AGENT
    AddLauncherButton
    LockForAgent
    ShowRaadApp
    Exit Sub
ErrorHandler:
    HandleGlobalError "StartApplication"
End Sub

Public Sub ShowRaadApp()
    On Error GoTo ErrorHandler
    EnsureDatasets
    frmRaadApp.Show vbModeless
    Exit Sub
ErrorHandler:
    HandleGlobalError "ShowRaadApp"
End Sub

Public Sub RefreshOpenForm(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler
    If gAppWriting Then Exit Sub
    If Not IsFormLoaded("frmRaadApp") Then Exit Sub
    frmRaadApp.ExternalSheetRefresh CStr(Sh.Name), Target
    Exit Sub
ErrorHandler:
    HandleGlobalError "RefreshOpenForm"
End Sub

Public Sub CleanupBeforeClose()
    On Error GoTo ErrorHandler
    gAppWriting = False
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    HandleGlobalError "CleanupBeforeClose"
End Sub

Private Function IsFormLoaded(ByVal formName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim f As Object
    For Each f In UserForms
        If StrComp(f.Name, formName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next f
    IsFormLoaded = False
    Exit Function
ErrorHandler:
    HandleGlobalError "IsFormLoaded"
    IsFormLoaded = False
End Function

Public Sub HandleGlobalError(ByVal procName As String)
    On Error GoTo ErrorHandler
    Dim n As Long
    Dim desc As String
    Dim msg As String
    n = Err.Number
    desc = Err.Description
    Err.Clear
    msg = "A recoverable error occurred in " & procName & "." & vbCrLf & _
          "Error " & CStr(n) & ": " & desc & vbCrLf & _
          "No Debug/End window was shown. Please review the last action."
    MsgBox msg, vbExclamation, "Raad App"
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Public Sub EnsureDatasets()
    On Error GoTo ErrorHandler
    If gDatasetCount = 0 Then InitDatasets
    Exit Sub
ErrorHandler:
    HandleGlobalError "EnsureDatasets"
End Sub

Public Function FindMainDatasetIndex() As Long
    On Error GoTo ErrorHandler
    Dim i As Long
    EnsureDatasets
    For i = 1 To gDatasetCount
        If StrComp(gDatasets(i).Key, "MAIN_ARCHIVE", vbBinaryCompare) = 0 Then
            FindMainDatasetIndex = i
            Exit Function
        End If
    Next i
    FindMainDatasetIndex = 1
    Exit Function
ErrorHandler:
    HandleGlobalError "FindMainDatasetIndex"
    FindMainDatasetIndex = 1
End Function

Private Sub InitDatasets()
    On Error GoTo ErrorHandler
    Dim blockRows As Variant
    Dim i As Long
    Dim hdr As Long
    Dim lastRow As Long
    gDatasetCount = 0
    Erase gDatasets

    AddDataset "DASH_KPI_LABELS", 1, "Dashboard KPI labels", 13, 2, 10, 14, 14, False, False, False, "", "Read-only KPI labels row."
    AddDataset "DASH_KPI_TABLE", 1, "Dashboard KPI table", 17, 2, 11, 18, 28, False, True, False, "", "Formula-backed KPI table; calculated cells are read-only."
    AddDataset "DAILY_INPUT", 2, "Daily input station", 5, 1, 7, 6, 16, False, True, False, "", "A:C editable; linked formula columns stay read-only."
    AddDataset "BANANA_MASTER", 3, "Unified bun master", 5, 1, 10, 6, 6, False, True, False, "", "Single master record with live formula columns."
    AddDataset "BANANA_DETAIL", 3, "Unified bun distribution", 10, 1, 6, 11, 14, False, True, False, "", "Contains the real Yes/No validation column."
    AddDataset "TIME_ANALYTICS", 4, "Time analytics", 6, 1, 9, 7, 17, False, False, False, "", "Read-only linked analytics formulas."
    AddDataset "MAIN_ARCHIVE", 5, "Daily archive - FULL CRUD", 5, 1, 9, 6, 200, True, True, True, "1,2,3,4,5,6,7,8,9", "Flat database table; Add/Update/Delete write directly here."
    AddDataset "BEV_PEPSI", 6, "Beverages - block 1", 3, 1, 9, 4, 4, False, True, False, "", "Single-record beverage block."
    AddDataset "BEV_WATER", 6, "Beverages - block 2", 8, 1, 9, 9, 9, False, True, False, "", "Single-record beverage block."
    AddDataset "FAKHER_FULL", 7, "Full section grid", 2, 1, 13, 3, 21, False, True, False, "", "Full-width grid; formula columns are preserved."
    AddDataset "MASHAWI_FULL", 8, "Full section grid", 2, 1, 20, 3, 23, False, True, False, "", "Full-width grid including helper columns."
    AddDataset "BURGER_SALES", 9, "Sales grid", 1, 1, 7, 2, 12, False, True, False, "", "Sales grid; formula columns are preserved."
    AddDataset "BURGER_INV", 9, "Inventory grid", 14, 1, 8, 15, 19, False, True, False, "", "Inventory grid; formula columns are preserved."
    AddDataset "MADGHOUT_FULL", 10, "Full section grid", 3, 1, 15, 4, 15, False, True, False, "", "Full-width grid; formula columns are preserved."
    AddDataset "ZURBIAN_FULL", 11, "Full section grid", 3, 1, 16, 4, 12, False, True, False, "", "Full-width grid; formula columns are preserved."
    AddDataset "KUNAFA_GRID", 12, "Formula grid", 2, 1, 17, 3, 11, False, True, False, "", "Formula-heavy grid; calculated cells are read-only."
    AddDataset "KUNAFA_INV", 12, "Inventory grid", 14, 1, 8, 15, 24, False, True, False, "", "Inventory grid; formula columns are preserved."
    AddDataset "FRIDGE_FULL", 13, "Fridge grid", 2, 1, 15, 3, 19, False, True, False, "", "Fridge grid; status/formula columns are preserved."
    AddDataset "BROAST_SALES", 14, "Broast sales", 5, 2, 8, 6, 12, False, True, False, "", "Sales table; formula column E is read-only."
    AddDataset "BROAST_INV", 14, "Broast inventory", 16, 2, 9, 17, 21, False, True, False, "", "Inventory table; formula columns are preserved."
    AddDataset "PIZZA_SALES", 16, "Pizza sales", 1, 1, 4, 2, 9, False, True, False, "", "Sales grid; formula column D is read-only."
    AddDataset "PIZZA_INV", 16, "Pizza inventory", 13, 1, 8, 14, 15, False, True, False, "", "Inventory table; formula columns are preserved."

    blockRows = Array(5, 33, 61, 90, 118, 145, 171, 199, 229, 256, 284, 311, 338, 366)
    For i = LBound(blockRows) To UBound(blockRows)
        hdr = CLng(blockRows(i))
        If i < UBound(blockRows) Then
            lastRow = CLng(blockRows(i + 1)) - 1
        Else
            lastRow = 390
        End If
        AddDataset "BROAST_HIST_" & Format(i + 1, "00"), 15, _
                   "Historical daily block " & Format(i + 1, "00"), _
                   hdr, 1, 17, hdr + 1, lastRow, False, False, False, "", _
                   "Read-only historical block; use Excel sheet view for raw edits."
    Next i
    Exit Sub
ErrorHandler:
    HandleGlobalError "InitDatasets"
End Sub

Private Sub AddDataset(ByVal key As String, ByVal sheetIndex As Long, ByVal label As String, _
                       ByVal headerRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, _
                       ByVal firstDataRow As Long, ByVal lastDataRow As Long, _
                       ByVal allowAdd As Boolean, ByVal allowUpdate As Boolean, ByVal allowDelete As Boolean, _
                       ByVal requiredCols As String, ByVal description As String)
    On Error GoTo ErrorHandler
    Dim sheetName As String
    sheetName = ThisWorkbook.Worksheets(sheetIndex).Name
    gDatasetCount = gDatasetCount + 1
    If gDatasetCount = 1 Then
        ReDim gDatasets(1 To 1)
    Else
        ReDim Preserve gDatasets(1 To gDatasetCount)
    End If
    With gDatasets(gDatasetCount)
        .Key = key
        .DisplayName = sheetName & " | " & label
        .SheetName = sheetName
        .HeaderRow = headerRow
        .FirstCol = firstCol
        .LastCol = lastCol
        .FirstDataRow = firstDataRow
        .LastDataRow = lastDataRow
        .AllowAdd = allowAdd
        .AllowUpdate = allowUpdate
        .AllowDelete = allowDelete
        .RequiredCols = requiredCols
        .Description = description
    End With
    Exit Sub
ErrorHandler:
    HandleGlobalError "AddDataset"
End Sub

Public Function ValidateAdminLogin() As Boolean
    On Error GoTo ErrorHandler
    Dim pin As String
    Dim entered As String
    pin = GetSetting(SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, "")
    If Len(pin) = 0 Then
        entered = InputBox("Create an Admin PIN (at least 4 characters). It will be stored in your user settings.", "First Admin Setup")
        If Len(Trim$(entered)) < 4 Then
            MsgBox "Admin setup was not completed. Agent mode remains active.", vbInformation, "Raad App"
            ValidateAdminLogin = False
        Else
            SaveSetting SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, Trim$(entered)
            ValidateAdminLogin = True
        End If
    Else
        entered = InputBox("Enter Admin PIN:", "Admin Login")
        If StrComp(Trim$(entered), pin, vbBinaryCompare) = 0 Then
            ValidateAdminLogin = True
        Else
            MsgBox "Incorrect Admin PIN. Agent mode remains active.", vbExclamation, "Raad App"
            ValidateAdminLogin = False
        End If
    End If
    Exit Function
ErrorHandler:
    HandleGlobalError "ValidateAdminLogin"
    ValidateAdminLogin = False
End Function

Public Sub LockForAgent()
    On Error GoTo ErrorHandler
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldScreen As Boolean
    Dim ws As Worksheet
    EnsureDatasets
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ThisWorkbook.Unprotect
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect
        LockSheetCells ws
        LockFormulaCells ws
        UnlockEditableCells ws
        ws.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                   UserInterfaceOnly:=True, AllowFormattingCells:=False, _
                   AllowFormattingColumns:=False, AllowFormattingRows:=False, _
                   AllowInsertingColumns:=False, AllowInsertingRows:=False, _
                   AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, _
                   AllowDeletingRows:=False, AllowSorting:=True, AllowFiltering:=True, _
                   AllowUsingPivotTables:=True
    Next ws
    ThisWorkbook.Protect Password:="", Structure:=True, Windows:=False
    gWorkbookLocked = True
SafeExit:
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    Exit Sub
ErrorHandler:
    HandleGlobalError "LockForAgent"
    Resume SafeExit
End Sub

Private Sub LockSheetCells(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    ws.UsedRange.Locked = True
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub UnlockEditableCells(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim d As AppDataset
    For i = 1 To gDatasetCount
        d = gDatasets(i)
        If StrComp(d.SheetName, ws.Name, vbTextCompare) = 0 And d.AllowUpdate Then
            For r = d.FirstDataRow To d.LastDataRow
                For c = d.FirstCol To d.LastCol
                    If Not ws.Cells(r, c).HasFormula Then ws.Cells(r, c).Locked = False
                Next c
            Next r
        End If
    Next i
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Public Sub UnlockAllSheets()
    On Error GoTo ErrorHandler
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim ws As Worksheet
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    ThisWorkbook.Unprotect
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect
    Next ws
    gWorkbookLocked = False
SafeExit:
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Exit Sub
ErrorHandler:
    HandleGlobalError "UnlockAllSheets"
    Resume SafeExit
End Sub

Private Sub LockFormulaCells(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    Dim rng As Range
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    rng.Locked = True
    Exit Sub
ErrorHandler:
    If Err.Number <> 1004 Then HandleGlobalError "LockFormulaCells"
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Public Sub AddLauncherButton()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim shp As Shape
    Set ws = ThisWorkbook.Worksheets(1)
    If LauncherShapeExists(ws) Then ws.Shapes("shpRaadLauncher").Delete
    Set shp = ws.Shapes.AddShape(5, 860, 8, 150, 32)
    shp.Name = "shpRaadLauncher"
    shp.OnAction = "ShowRaadApp"
    shp.TextFrame.Characters.Text = "Open Raad App"
    shp.TextFrame.Characters.Font.Name = "Segoe UI"
    shp.TextFrame.Characters.Font.Size = 10
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(37, 99, 235)
    shp.Line.Visible = 0
    Exit Sub
ErrorHandler:
    HandleGlobalError "AddLauncherButton"
End Sub

Private Function LauncherShapeExists(ByVal ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    Dim shp As Shape
    Set shp = ws.Shapes("shpRaadLauncher")
    LauncherShapeExists = True
    Exit Function
ErrorHandler:
    Err.Clear
    LauncherShapeExists = False
    Resume SafeExit
SafeExit:
End Function
