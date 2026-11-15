Attribute VB_Name = "AutoBuild_RaadApp"
Option Explicit

' One-click builder. Import this module into the macro-enabled workbook,
' enable Trust access to the VBA project object model, then run Build_Raad_App.

Public Sub Build_Raad_App()
    On Error GoTo ErrorHandler
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldScreen As Boolean
    Dim vbProj As Object
    Dim stdMod As Object
    Dim frm As Object
    Dim newPath As String
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating
    TrySetApp "EnableEvents", False
    TrySetApp "DisplayAlerts", False
    TrySetApp "ScreenUpdating", False
    Set vbProj = ThisWorkbook.VBProject
    Set stdMod = AddStandardModule(vbProj, "modRaadApp", ModuleSource())
    If stdMod Is Nothing Then GoTo SafeExit
    Set frm = AddUserForm(vbProj, "frmRaadApp", FormSource())
    If frm Is Nothing Then GoTo SafeExit
    ReplaceModuleCode ThisWorkbook.VBProject.VBComponents("ThisWorkbook"), WorkbookSource()
    If LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        newPath = Left$(ThisWorkbook.FullName, Len(ThisWorkbook.FullName) - 5) & ".xlsm"
        ThisWorkbook.SaveAs Filename:=newPath, FileFormat:=52
    Else
        ThisWorkbook.Save
    End If
    MsgBox "Raad app built successfully. Reopen the workbook and enable macros; the dashboard will open automatically.", vbInformation, "Raad App"
SafeExit:
    TrySetApp "EnableEvents", oldEvents
    TrySetApp "DisplayAlerts", oldAlerts
    TrySetApp "ScreenUpdating", oldScreen
    Exit Sub
ErrorHandler:
    BuilderHandleError "Build_Raad_App"
    Resume SafeExit
End Sub

Private Sub BuilderHandleError(ByVal procName As String)
    On Error GoTo ErrorHandler
    Dim n As Long
    Dim desc As String
    Dim msg As String
    n = Err.Number
    desc = Err.Description
    Err.Clear
    msg = "Builder error in " & procName & ": " & CStr(n) & " - " & desc & vbCrLf & _
          "If this is a permissions error, enable: File > Options > Trust Center > Macro Settings > Trust access to the VBA project object model."
    MsgBox msg, vbExclamation, "Raad App Builder"
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub TrySetApp(ByVal propName As String, ByVal value As Variant)
    On Error GoTo ErrorHandler
    CallByName Application, propName, VbLet, value
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub TrySetCtl(ByVal ctl As Object, ByVal propName As String, ByVal value As Variant)
    On Error GoTo ErrorHandler
    CallByName ctl, propName, VbLet, value
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub TrySetFontCtl(ByVal ctl As Object, ByVal fontSize As Single, ByVal boldValue As Boolean)
    On Error GoTo ErrorHandler
    Dim f As Object
    Set f = CallByName(ctl, "Font", VbGet)
    CallByName f, "Name", VbLet, "Segoe UI"
    CallByName f, "Size", VbLet, fontSize
    CallByName f, "Bold", VbLet, boldValue
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub TrySetComponentProperty(ByVal comp As Object, ByVal propName As String, ByVal value As Variant)
    On Error GoTo ErrorHandler
    comp.Properties(propName).Value = value
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub RemoveComponent(ByVal vbProj As Object, ByVal componentName As String)
    On Error GoTo ErrorHandler
    vbProj.VBComponents.Remove vbProj.VBComponents(componentName)
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub ReplaceModuleCode(ByVal comp As Object, ByVal source As String)
    On Error GoTo ErrorHandler
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString source
    End With
    Exit Sub
ErrorHandler:
    BuilderHandleError "ReplaceModuleCode"
End Sub

Private Function AddStandardModule(ByVal vbProj As Object, ByVal moduleName As String, ByVal source As String) As Object
    On Error GoTo ErrorHandler
    Dim comp As Object
    RemoveComponent vbProj, moduleName
    Set comp = vbProj.VBComponents.Add(1)
    comp.Name = moduleName
    ReplaceModuleCode comp, source
    Set AddStandardModule = comp
    Exit Function
ErrorHandler:
    BuilderHandleError "AddStandardModule"
    Set AddStandardModule = Nothing
End Function

Private Function AddUserForm(ByVal vbProj As Object, ByVal formName As String, ByVal source As String) As Object
    On Error GoTo ErrorHandler
    Dim comp As Object
    RemoveComponent vbProj, formName
    Set comp = vbProj.VBComponents.Add(3)
    comp.Name = formName
    TrySetComponentProperty comp, "Caption", "Raad | Live Inventory & Variance Console"
    TrySetComponentProperty comp, "Width", 1200
    TrySetComponentProperty comp, "Height", 780
    TrySetComponentProperty comp, "StartUpPosition", 2
    TrySetComponentProperty comp, "BackColor", RGB(244, 247, 251)
    TrySetComponentProperty comp, "BorderStyle", 1
    TrySetComponentProperty comp, "SpecialEffect", 0
    AddAllControls comp
    ReplaceModuleCode comp, source
    Set AddUserForm = comp
    Exit Function
ErrorHandler:
    BuilderHandleError "AddUserForm"
    Set AddUserForm = Nothing
End Function

Private Sub AddCtl(ByVal frm As Object, ByVal progId As String, ByVal ctlName As String, _
                  ByVal leftValue As Single, ByVal topValue As Single, _
                  ByVal widthValue As Single, ByVal heightValue As Single, _
                  Optional ByVal caption As String = "", _
                  Optional ByVal backColor As Long = -1, _
                  Optional ByVal foreColor As Long = -1, _
                  Optional ByVal boldValue As Boolean = False, _
                  Optional ByVal fontSize As Single = 9)
    On Error GoTo ErrorHandler
    Dim c As Object
    Set c = frm.Designer.Controls.Add(progId, ctlName, True)
    TrySetCtl c, "Left", leftValue
    TrySetCtl c, "Top", topValue
    TrySetCtl c, "Width", widthValue
    TrySetCtl c, "Height", heightValue
    TrySetCtl c, "Caption", caption
    If backColor >= 0 Then TrySetCtl c, "BackColor", backColor
    If foreColor >= 0 Then TrySetCtl c, "ForeColor", foreColor
    TrySetCtl c, "SpecialEffect", 0
    TrySetCtl c, "BorderStyle", 1
    TrySetCtl c, "BorderColor", RGB(203, 213, 225)
    TrySetCtl c, "TextAlign", 3
    TrySetCtl c, "RightToLeft", True
    TrySetCtl c, "AutoSize", False
    TrySetCtl c, "WordWrap", True
    TrySetFontCtl c, fontSize, boldValue
    Exit Sub
ErrorHandler:
    BuilderHandleError "AddCtl"
End Sub

Private Sub AddAllControls(ByVal frm As Object)
    On Error GoTo ErrorHandler
    AddCtl frm, "Forms.Frame.1", "fraHeader", 0, 0, 1200, 64, "", RGB(16, 42, 67), RGB(255, 255, 255), False, 9
    AddCtl frm, "Forms.Frame.1", "fraSidebar", 970, 76, 214, 622, "", RGB(12, 31, 50), RGB(255, 255, 255), False, 9
    AddCtl frm, "Forms.Frame.1", "fraToolbar", 16, 76, 938, 68, "", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.Frame.1", "fraEditor", 16, 466, 938, 210, "", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.Frame.1", "fraButtons", 16, 684, 938, 48, "", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.Label.1", "lblTitle", 24, 8, 620, 26, "Raad Inventory & Daily Variance", RGB(16, 42, 67), RGB(255, 255, 255), True, 16
    AddCtl frm, "Forms.Label.1", "lblSubTitle", 24, 34, 760, 20, "Live two-way connection to your real worksheets, fields, formulas and dropdowns", RGB(16, 42, 67), RGB(203, 213, 225), False, 9
    AddCtl frm, "Forms.Label.1", "lblModeBadge", 970, 16, 90, 28, "AGENT", RGB(37, 99, 235), RGB(255, 255, 255), True, 10
    AddCtl frm, "Forms.Label.1", "lblClock", 1070, 16, 110, 28, "", RGB(16, 42, 67), RGB(255, 255, 255), False, 9
    AddCtl frm, "Forms.Label.1", "lblNav", 986, 92, 180, 22, "Workspace / Role", RGB(12, 31, 50), RGB(255, 255, 255), True, 10
    AddCtl frm, "Forms.Label.1", "lblDatasetInfo", 986, 226, 180, 62, "Select a real sheet dataset", RGB(12, 31, 50), RGB(255, 255, 255), False, 9
    AddCtl frm, "Forms.Label.1", "lblRecordCount", 986, 466, 180, 22, "Records: 0", RGB(12, 31, 50), RGB(255, 255, 255), False, 9
    AddCtl frm, "Forms.Label.1", "lblSelectedRow", 986, 494, 180, 22, "Selected row: none", RGB(12, 31, 50), RGB(255, 255, 255), False, 9
    AddCtl frm, "Forms.Label.1", "lblSearch", 28, 88, 240, 18, "Live search", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.Label.1", "lblFilterField", 294, 88, 180, 18, "Filter by real column", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.Label.1", "lblFilterValue", 490, 88, 180, 18, "Filter value", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.Label.1", "lblEditorTitle", 28, 474, 400, 20, "Record fields - exact worksheet headers", RGB(255, 255, 255), RGB(15, 23, 42), True, 10
    AddCtl frm, "Forms.Label.1", "lblStatus", 16, 738, 1168, 26, "Ready.", RGB(255, 255, 255), RGB(22, 163, 74), True, 9
    AddCtl frm, "Forms.ComboBox.1", "cboRole", 986, 120, 180, 26, "", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.ComboBox.1", "cboDataset", 986, 194, 180, 26, "", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.ComboBox.1", "cboFilterField", 294, 108, 180, 26, "", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.ComboBox.1", "cboFilterValue", 490, 108, 180, 26, "", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.TextBox.1", "txtSearch", 28, 108, 250, 26, "", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.ListBox.1", "lstHeaders", 16, 152, 938, 24, "", RGB(16, 42, 67), RGB(255, 255, 255), False, 9
    AddCtl frm, "Forms.ListBox.1", "lstData", 16, 176, 938, 282, "", RGB(255, 255, 255), RGB(15, 23, 42), False, 9
    AddCtl frm, "Forms.CommandButton.1", "btnSwitchMode", 986, 152, 180, 30, "Switch Mode", RGB(37, 99, 235), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnRefresh", 986, 298, 180, 30, "Refresh", RGB(37, 99, 235), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnSheetView", 986, 336, 180, 30, "Switch to Excel Sheet", RGB(37, 99, 235), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnUnlockAll", 986, 380, 180, 30, "Unlock Sheets", RGB(37, 99, 235), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnLockForAgent", 986, 418, 180, 30, "Lock For Agent", RGB(100, 116, 139), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnClearFilter", 686, 108, 120, 26, "Clear Filter", RGB(100, 116, 139), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnAdd", 28, 694, 140, 30, "Add", RGB(37, 99, 235), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnUpdate", 180, 694, 140, 30, "Update", RGB(245, 158, 11), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnDelete", 332, 694, 140, 30, "Delete", RGB(220, 38, 38), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnClear", 484, 694, 140, 30, "Clear", RGB(100, 116, 139), RGB(255, 255, 255), True, 9
    AddCtl frm, "Forms.CommandButton.1", "btnClose", 800, 694, 140, 30, "Close", RGB(100, 116, 139), RGB(255, 255, 255), True, 9
    Exit Sub
ErrorHandler:
    BuilderHandleError "AddAllControls"
End Sub

Private Function ModuleSource() As String
    On Error GoTo ErrorHandler
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' Raad Inventory App - shared configuration, startup, role and protection logic." & vbCrLf
    s = s & "' All worksheet names and field captions are read from the real workbook at run" & vbCrLf
    s = s & "' time. This source deliberately contains no hard-coded Arabic field names." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Type AppDataset" & vbCrLf
    s = s & "    Key As String" & vbCrLf
    s = s & "    DisplayName As String" & vbCrLf
    s = s & "    SheetName As String" & vbCrLf
    s = s & "    HeaderRow As Long" & vbCrLf
    s = s & "    FirstCol As Long" & vbCrLf
    s = s & "    LastCol As Long" & vbCrLf
    s = s & "    FirstDataRow As Long" & vbCrLf
    s = s & "    LastDataRow As Long" & vbCrLf
    s = s & "    AllowAdd As Boolean" & vbCrLf
    s = s & "    AllowUpdate As Boolean" & vbCrLf
    s = s & "    AllowDelete As Boolean" & vbCrLf
    s = s & "    RequiredCols As String" & vbCrLf
    s = s & "    Description As String" & vbCrLf
    s = s & "End Type" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Type AppField" & vbCrLf
    s = s & "    Index As Long" & vbCrLf
    s = s & "    ColNumber As Long" & vbCrLf
    s = s & "    ColLetter As String" & vbCrLf
    s = s & "    Header As String" & vbCrLf
    s = s & "    LabelText As String" & vbCrLf
    s = s & "    Kind As String" & vbCrLf
    s = s & "    InputName As String" & vbCrLf
    s = s & "    Editable As Boolean" & vbCrLf
    s = s & "    ValidationList As String" & vbCrLf
    s = s & "End Type" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Const ROLE_AGENT As String = ""AGENT""" & vbCrLf
    s = s & "Public Const ROLE_ADMIN As String = ""ADMIN""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public gDatasets() As AppDataset" & vbCrLf
    s = s & "Public gDatasetCount As Long" & vbCrLf
    s = s & "Public gRole As String" & vbCrLf
    s = s & "Public gAppWriting As Boolean" & vbCrLf
    s = s & "Public gWorkbookLocked As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const SETTINGS_APP As String = ""RaadInventoryApp""" & vbCrLf
    s = s & "Private Const SETTINGS_SECTION As String = ""Security""" & vbCrLf
    s = s & "Private Const SETTINGS_KEY As String = ""AdminPIN""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub StartApplication()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    EnsureDatasets" & vbCrLf
    s = s & "    gRole = ROLE_AGENT" & vbCrLf
    s = s & "    AddLauncherButton" & vbCrLf
    s = s & "    LockForAgent" & vbCrLf
    s = s & "    ShowRaadApp" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""StartApplication""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ShowRaadApp()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    EnsureDatasets" & vbCrLf
    s = s & "    frmRaadApp.Show vbModeless" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""ShowRaadApp""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub RefreshOpenForm(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If gAppWriting Then Exit Sub" & vbCrLf
    s = s & "    If Not IsFormLoaded(""frmRaadApp"") Then Exit Sub" & vbCrLf
    s = s & "    frmRaadApp.ExternalSheetRefresh CStr(Sh.Name), Target" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""RefreshOpenForm""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub CleanupBeforeClose()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    gAppWriting = False" & vbCrLf
    s = s & "    Application.EnableEvents = True" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""CleanupBeforeClose""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsFormLoaded(ByVal formName As String) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim f As Object" & vbCrLf
    s = s & "    For Each f In UserForms" & vbCrLf
    s = s & "        If StrComp(f.Name, formName, vbTextCompare) = 0 Then" & vbCrLf
    s = s & "            IsFormLoaded = True" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next f" & vbCrLf
    s = s & "    IsFormLoaded = False" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""IsFormLoaded""" & vbCrLf
    s = s & "    IsFormLoaded = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub HandleGlobalError(ByVal procName As String)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim n As Long" & vbCrLf
    s = s & "    Dim desc As String" & vbCrLf
    s = s & "    Dim msg As String" & vbCrLf
    s = s & "    n = Err.Number" & vbCrLf
    s = s & "    desc = Err.Description" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    msg = ""A recoverable error occurred in "" & procName & ""."" & vbCrLf & _" & vbCrLf
    s = s & "          ""Error "" & CStr(n) & "": "" & desc & vbCrLf & _" & vbCrLf
    s = s & "          ""No Debug/End window was shown. Please review the last action.""" & vbCrLf
    s = s & "    MsgBox msg, vbExclamation, ""Raad App""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub EnsureDatasets()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If gDatasetCount = 0 Then InitDatasets" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""EnsureDatasets""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function FindMainDatasetIndex() As Long" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    EnsureDatasets" & vbCrLf
    s = s & "    For i = 1 To gDatasetCount" & vbCrLf
    s = s & "        If StrComp(gDatasets(i).Key, ""MAIN_ARCHIVE"", vbBinaryCompare) = 0 Then" & vbCrLf
    s = s & "            FindMainDatasetIndex = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    FindMainDatasetIndex = 1" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""FindMainDatasetIndex""" & vbCrLf
    s = s & "    FindMainDatasetIndex = 1" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub InitDatasets()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim blockRows As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim hdr As Long" & vbCrLf
    s = s & "    Dim lastRow As Long" & vbCrLf
    s = s & "    gDatasetCount = 0" & vbCrLf
    s = s & "    Erase gDatasets" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    AddDataset ""DASH_KPI_LABELS"", 1, ""Dashboard KPI labels"", 13, 2, 10, 14, 14, False, False, False, """", ""Read-only KPI labels row.""" & vbCrLf
    s = s & "    AddDataset ""DASH_KPI_TABLE"", 1, ""Dashboard KPI table"", 17, 2, 11, 18, 28, False, True, False, """", ""Formula-backed KPI table; calculated cells are read-only.""" & vbCrLf
    s = s & "    AddDataset ""DAILY_INPUT"", 2, ""Daily input station"", 5, 1, 7, 6, 16, False, True, False, """", ""A:C editable; linked formula columns stay read-only.""" & vbCrLf
    s = s & "    AddDataset ""BANANA_MASTER"", 3, ""Unified bun master"", 5, 1, 10, 6, 6, False, True, False, """", ""Single master record with live formula columns.""" & vbCrLf
    s = s & "    AddDataset ""BANANA_DETAIL"", 3, ""Unified bun distribution"", 10, 1, 6, 11, 14, False, True, False, """", ""Contains the real Yes/No validation column.""" & vbCrLf
    s = s & "    AddDataset ""TIME_ANALYTICS"", 4, ""Time analytics"", 6, 1, 9, 7, 17, False, False, False, """", ""Read-only linked analytics formulas.""" & vbCrLf
    s = s & "    AddDataset ""MAIN_ARCHIVE"", 5, ""Daily archive - FULL CRUD"", 5, 1, 9, 6, 200, True, True, True, ""1,2,3,4,5,6,7,8,9"", ""Flat database table; Add/Update/Delete write directly here.""" & vbCrLf
    s = s & "    AddDataset ""BEV_PEPSI"", 6, ""Beverages - block 1"", 3, 1, 9, 4, 4, False, True, False, """", ""Single-record beverage block.""" & vbCrLf
    s = s & "    AddDataset ""BEV_WATER"", 6, ""Beverages - block 2"", 8, 1, 9, 9, 9, False, True, False, """", ""Single-record beverage block.""" & vbCrLf
    s = s & "    AddDataset ""FAKHER_FULL"", 7, ""Full section grid"", 2, 1, 13, 3, 21, False, True, False, """", ""Full-width grid; formula columns are preserved.""" & vbCrLf
    s = s & "    AddDataset ""MASHAWI_FULL"", 8, ""Full section grid"", 2, 1, 20, 3, 23, False, True, False, """", ""Full-width grid including helper columns.""" & vbCrLf
    s = s & "    AddDataset ""BURGER_SALES"", 9, ""Sales grid"", 1, 1, 7, 2, 12, False, True, False, """", ""Sales grid; formula columns are preserved.""" & vbCrLf
    s = s & "    AddDataset ""BURGER_INV"", 9, ""Inventory grid"", 14, 1, 8, 15, 19, False, True, False, """", ""Inventory grid; formula columns are preserved.""" & vbCrLf
    s = s & "    AddDataset ""MADGHOUT_FULL"", 10, ""Full section grid"", 3, 1, 15, 4, 15, False, True, False, """", ""Full-width grid; formula columns are preserved.""" & vbCrLf
    s = s & "    AddDataset ""ZURBIAN_FULL"", 11, ""Full section grid"", 3, 1, 16, 4, 12, False, True, False, """", ""Full-width grid; formula columns are preserved.""" & vbCrLf
    s = s & "    AddDataset ""KUNAFA_GRID"", 12, ""Formula grid"", 2, 1, 17, 3, 11, False, True, False, """", ""Formula-heavy grid; calculated cells are read-only.""" & vbCrLf
    s = s & "    AddDataset ""KUNAFA_INV"", 12, ""Inventory grid"", 14, 1, 8, 15, 24, False, True, False, """", ""Inventory grid; formula columns are preserved.""" & vbCrLf
    s = s & "    AddDataset ""FRIDGE_FULL"", 13, ""Fridge grid"", 2, 1, 15, 3, 19, False, True, False, """", ""Fridge grid; status/formula columns are preserved.""" & vbCrLf
    s = s & "    AddDataset ""BROAST_SALES"", 14, ""Broast sales"", 5, 2, 8, 6, 12, False, True, False, """", ""Sales table; formula column E is read-only.""" & vbCrLf
    s = s & "    AddDataset ""BROAST_INV"", 14, ""Broast inventory"", 16, 2, 9, 17, 21, False, True, False, """", ""Inventory table; formula columns are preserved.""" & vbCrLf
    s = s & "    AddDataset ""PIZZA_SALES"", 16, ""Pizza sales"", 1, 1, 4, 2, 9, False, True, False, """", ""Sales grid; formula column D is read-only.""" & vbCrLf
    s = s & "    AddDataset ""PIZZA_INV"", 16, ""Pizza inventory"", 13, 1, 8, 14, 15, False, True, False, """", ""Inventory table; formula columns are preserved.""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    blockRows = Array(5, 33, 61, 90, 118, 145, 171, 199, 229, 256, 284, 311, 338, 366)" & vbCrLf
    s = s & "    For i = LBound(blockRows) To UBound(blockRows)" & vbCrLf
    s = s & "        hdr = CLng(blockRows(i))" & vbCrLf
    s = s & "        If i < UBound(blockRows) Then" & vbCrLf
    s = s & "            lastRow = CLng(blockRows(i + 1)) - 1" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            lastRow = 390" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        AddDataset ""BROAST_HIST_"" & Format(i + 1, ""00""), 15, _" & vbCrLf
    s = s & "                   ""Historical daily block "" & Format(i + 1, ""00""), _" & vbCrLf
    s = s & "                   hdr, 1, 17, hdr + 1, lastRow, False, False, False, """", _" & vbCrLf
    s = s & "                   ""Read-only historical block; use Excel sheet view for raw edits.""" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""InitDatasets""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub AddDataset(ByVal key As String, ByVal sheetIndex As Long, ByVal label As String, _" & vbCrLf
    s = s & "                       ByVal headerRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, _" & vbCrLf
    s = s & "                       ByVal firstDataRow As Long, ByVal lastDataRow As Long, _" & vbCrLf
    s = s & "                       ByVal allowAdd As Boolean, ByVal allowUpdate As Boolean, ByVal allowDelete As Boolean, _" & vbCrLf
    s = s & "                       ByVal requiredCols As String, ByVal description As String)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim sheetName As String" & vbCrLf
    s = s & "    sheetName = ThisWorkbook.Worksheets(sheetIndex).Name" & vbCrLf
    s = s & "    gDatasetCount = gDatasetCount + 1" & vbCrLf
    s = s & "    If gDatasetCount = 1 Then" & vbCrLf
    s = s & "        ReDim gDatasets(1 To 1)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ReDim Preserve gDatasets(1 To gDatasetCount)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    With gDatasets(gDatasetCount)" & vbCrLf
    s = s & "        .Key = key" & vbCrLf
    s = s & "        .DisplayName = sheetName & "" | "" & label" & vbCrLf
    s = s & "        .SheetName = sheetName" & vbCrLf
    s = s & "        .HeaderRow = headerRow" & vbCrLf
    s = s & "        .FirstCol = firstCol" & vbCrLf
    s = s & "        .LastCol = lastCol" & vbCrLf
    s = s & "        .FirstDataRow = firstDataRow" & vbCrLf
    s = s & "        .LastDataRow = lastDataRow" & vbCrLf
    s = s & "        .AllowAdd = allowAdd" & vbCrLf
    s = s & "        .AllowUpdate = allowUpdate" & vbCrLf
    s = s & "        .AllowDelete = allowDelete" & vbCrLf
    s = s & "        .RequiredCols = requiredCols" & vbCrLf
    s = s & "        .Description = description" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""AddDataset""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function ValidateAdminLogin() As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim pin As String" & vbCrLf
    s = s & "    Dim entered As String" & vbCrLf
    s = s & "    pin = GetSetting(SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, """")" & vbCrLf
    s = s & "    If Len(pin) = 0 Then" & vbCrLf
    s = s & "        entered = InputBox(""Create an Admin PIN (at least 4 characters). It will be stored in your user settings."", ""First Admin Setup"")" & vbCrLf
    s = s & "        If Len(Trim$(entered)) < 4 Then" & vbCrLf
    s = s & "            MsgBox ""Admin setup was not completed. Agent mode remains active."", vbInformation, ""Raad App""" & vbCrLf
    s = s & "            ValidateAdminLogin = False" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            SaveSetting SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, Trim$(entered)" & vbCrLf
    s = s & "            ValidateAdminLogin = True" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        entered = InputBox(""Enter Admin PIN:"", ""Admin Login"")" & vbCrLf
    s = s & "        If StrComp(Trim$(entered), pin, vbBinaryCompare) = 0 Then" & vbCrLf
    s = s & "            ValidateAdminLogin = True" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            MsgBox ""Incorrect Admin PIN. Agent mode remains active."", vbExclamation, ""Raad App""" & vbCrLf
    s = s & "            ValidateAdminLogin = False" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""ValidateAdminLogin""" & vbCrLf
    s = s & "    ValidateAdminLogin = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub LockForAgent()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim oldEvents As Boolean" & vbCrLf
    s = s & "    Dim oldAlerts As Boolean" & vbCrLf
    s = s & "    Dim oldScreen As Boolean" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    EnsureDatasets" & vbCrLf
    s = s & "    oldEvents = Application.EnableEvents" & vbCrLf
    s = s & "    oldAlerts = Application.DisplayAlerts" & vbCrLf
    s = s & "    oldScreen = Application.ScreenUpdating" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "    Application.DisplayAlerts = False" & vbCrLf
    s = s & "    Application.ScreenUpdating = False" & vbCrLf
    s = s & "    ThisWorkbook.Unprotect" & vbCrLf
    s = s & "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    s = s & "        ws.Unprotect" & vbCrLf
    s = s & "        LockSheetCells ws" & vbCrLf
    s = s & "        LockFormulaCells ws" & vbCrLf
    s = s & "        UnlockEditableCells ws" & vbCrLf
    s = s & "        ws.Protect Password:="""", DrawingObjects:=False, Contents:=True, Scenarios:=False, _" & vbCrLf
    s = s & "                   UserInterfaceOnly:=True, AllowFormattingCells:=False, _" & vbCrLf
    s = s & "                   AllowFormattingColumns:=False, AllowFormattingRows:=False, _" & vbCrLf
    s = s & "                   AllowInsertingColumns:=False, AllowInsertingRows:=False, _" & vbCrLf
    s = s & "                   AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, _" & vbCrLf
    s = s & "                   AllowDeletingRows:=False, AllowSorting:=True, AllowFiltering:=True, _" & vbCrLf
    s = s & "                   AllowUsingPivotTables:=True" & vbCrLf
    s = s & "    Next ws" & vbCrLf
    s = s & "    ThisWorkbook.Protect Password:="""", Structure:=True, Windows:=False" & vbCrLf
    s = s & "    gWorkbookLocked = True" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    Application.EnableEvents = oldEvents" & vbCrLf
    s = s & "    Application.DisplayAlerts = oldAlerts" & vbCrLf
    s = s & "    Application.ScreenUpdating = oldScreen" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""LockForAgent""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LockSheetCells(ByVal ws As Worksheet)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ws.UsedRange.Locked = True" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UnlockEditableCells(ByVal ws As Worksheet)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    For i = 1 To gDatasetCount" & vbCrLf
    s = s & "        d = gDatasets(i)" & vbCrLf
    s = s & "        If StrComp(d.SheetName, ws.Name, vbTextCompare) = 0 And d.AllowUpdate Then" & vbCrLf
    s = s & "            For r = d.FirstDataRow To d.LastDataRow" & vbCrLf
    s = s & "                For c = d.FirstCol To d.LastCol" & vbCrLf
    s = s & "                    If Not ws.Cells(r, c).HasFormula Then ws.Cells(r, c).Locked = False" & vbCrLf
    s = s & "                Next c" & vbCrLf
    s = s & "            Next r" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub UnlockAllSheets()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim oldEvents As Boolean" & vbCrLf
    s = s & "    Dim oldAlerts As Boolean" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    oldEvents = Application.EnableEvents" & vbCrLf
    s = s & "    oldAlerts = Application.DisplayAlerts" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "    Application.DisplayAlerts = False" & vbCrLf
    s = s & "    ThisWorkbook.Unprotect" & vbCrLf
    s = s & "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    s = s & "        ws.Unprotect" & vbCrLf
    s = s & "    Next ws" & vbCrLf
    s = s & "    gWorkbookLocked = False" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    Application.EnableEvents = oldEvents" & vbCrLf
    s = s & "    Application.DisplayAlerts = oldAlerts" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""UnlockAllSheets""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LockFormulaCells(ByVal ws As Worksheet)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim rng As Range" & vbCrLf
    s = s & "    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)" & vbCrLf
    s = s & "    rng.Locked = True" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    If Err.Number <> 1004 Then HandleGlobalError ""LockFormulaCells""" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub AddLauncherButton()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Dim shp As Shape" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(1)" & vbCrLf
    s = s & "    If LauncherShapeExists(ws) Then ws.Shapes(""shpRaadLauncher"").Delete" & vbCrLf
    s = s & "    Set shp = ws.Shapes.AddShape(5, 860, 8, 150, 32)" & vbCrLf
    s = s & "    shp.Name = ""shpRaadLauncher""" & vbCrLf
    s = s & "    shp.OnAction = ""ShowRaadApp""" & vbCrLf
    s = s & "    shp.TextFrame.Characters.Text = ""Open Raad App""" & vbCrLf
    s = s & "    shp.TextFrame.Characters.Font.Name = ""Segoe UI""" & vbCrLf
    s = s & "    shp.TextFrame.Characters.Font.Size = 10" & vbCrLf
    s = s & "    shp.TextFrame.Characters.Font.Bold = True" & vbCrLf
    s = s & "    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)" & vbCrLf
    s = s & "    shp.Fill.ForeColor.RGB = RGB(37, 99, 235)" & vbCrLf
    s = s & "    shp.Line.Visible = 0" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleGlobalError ""AddLauncherButton""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function LauncherShapeExists(ByVal ws As Worksheet) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim shp As Shape" & vbCrLf
    s = s & "    Set shp = ws.Shapes(""shpRaadLauncher"")" & vbCrLf
    s = s & "    LauncherShapeExists = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    LauncherShapeExists = False" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Function" & vbCrLf
    ModuleSource = s
    Exit Function
ErrorHandler:
    BuilderHandleError "ModuleSource"
    ModuleSource = ""
End Function

Private Function FormSource() As String
    On Error GoTo ErrorHandler
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "' frmRaadApp - modern, dynamic UserForm code." & vbCrLf
    s = s & "' Static shell controls are listed in UI_Layout_and_Controls.md." & vbCrLf
    s = s & "' Field controls are created at run time from the exact worksheet header cells." & vbCrLf
    s = s & "'===============================================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mDatasets() As AppDataset" & vbCrLf
    s = s & "Private mDatasetCount As Long" & vbCrLf
    s = s & "Private mDatasetIndex As Long" & vbCrLf
    s = s & "Private mFields() As AppField" & vbCrLf
    s = s & "Private mFieldCount As Long" & vbCrLf
    s = s & "Private mSheet As Worksheet" & vbCrLf
    s = s & "Private mSelectedRow As Long" & vbCrLf
    s = s & "Private mRowCount As Long" & vbCrLf
    s = s & "Private mRowMap() As Long" & vbCrLf
    s = s & "Private mInitializing As Boolean" & vbCrLf
    s = s & "Private mSuppressEvents As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mBg As Long" & vbCrLf
    s = s & "Private mHeader As Long" & vbCrLf
    s = s & "Private mHeaderDark As Long" & vbCrLf
    s = s & "Private mCard As Long" & vbCrLf
    s = s & "Private mWhite As Long" & vbCrLf
    s = s & "Private mText As Long" & vbCrLf
    s = s & "Private mMuted As Long" & vbCrLf
    s = s & "Private mBorder As Long" & vbCrLf
    s = s & "Private mReadOnly As Long" & vbCrLf
    s = s & "Private mPrimary As Long" & vbCrLf
    s = s & "Private mPrimaryHover As Long" & vbCrLf
    s = s & "Private mOrange As Long" & vbCrLf
    s = s & "Private mOrangeHover As Long" & vbCrLf
    s = s & "Private mDanger As Long" & vbCrLf
    s = s & "Private mDangerHover As Long" & vbCrLf
    s = s & "Private mGray As Long" & vbCrLf
    s = s & "Private mGrayHover As Long" & vbCrLf
    s = s & "Private mSuccess As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' UserForm lifecycle" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    InitializeApp" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""UserForm_Initialize""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Activate()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Me.lblClock.Caption = Format(Now, ""yyyy-mm-dd hh:nn"")" & vbCrLf
    s = s & "    ApplyRoleVisual" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""UserForm_Activate""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    gAppWriting = False" & vbCrLf
    s = s & "    mSuppressEvents = False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""UserForm_QueryClose""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""UserForm_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Static control events" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboDataset_Change()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If mInitializing Or mSuppressEvents Then Exit Sub" & vbCrLf
    s = s & "    ApplyDataset" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""cboDataset_Change""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If mInitializing Or mSuppressEvents Then Exit Sub" & vbCrLf
    s = s & "    RefreshGrid" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""txtSearch_Change""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboFilterField_Change()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If mInitializing Or mSuppressEvents Then Exit Sub" & vbCrLf
    s = s & "    PopulateFilterValues" & vbCrLf
    s = s & "    RefreshGrid" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""cboFilterField_Change""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboFilterValue_Change()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If mInitializing Or mSuppressEvents Then Exit Sub" & vbCrLf
    s = s & "    RefreshGrid" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""cboFilterValue_Change""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstData_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim idx As Long" & vbCrLf
    s = s & "    If Me.lstData.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "    idx = Me.lstData.ListIndex + 1" & vbCrLf
    s = s & "    If idx > mRowCount Then Exit Sub" & vbCrLf
    s = s & "    LoadRecordIntoForm mRowMap(idx)" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""lstData_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    lstData_Click" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""lstData_DblClick""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnAdd_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    AddRecord" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnAdd_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnUpdate_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    UpdateRecord" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnUpdate_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    DeleteRecord" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnDelete_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnClear_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ClearInputs" & vbCrLf
    s = s & "    ShowStatus ""Form cleared. Select a grid row or enter a new record."", False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnClear_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnClearFilter_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim oldFlag As Boolean" & vbCrLf
    s = s & "    oldFlag = mSuppressEvents" & vbCrLf
    s = s & "    mSuppressEvents = True" & vbCrLf
    s = s & "    Me.txtSearch.Text = """"" & vbCrLf
    s = s & "    Me.cboFilterField.ListIndex = 0" & vbCrLf
    s = s & "    PopulateFilterValues" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    mSuppressEvents = oldFlag" & vbCrLf
    s = s & "    RefreshGrid" & vbCrLf
    s = s & "    ShowStatus ""Search and filters cleared."", False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnClearFilter_Click""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnRefresh_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If Not mSheet Is Nothing Then mSheet.Calculate" & vbCrLf
    s = s & "    RefreshGrid" & vbCrLf
    s = s & "    If mSelectedRow > 0 Then LoadRecordIntoForm mSelectedRow" & vbCrLf
    s = s & "    ShowStatus ""Refreshed directly from the real worksheet."", False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnRefresh_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSheetView_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    If mSheet Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    TrySetApplication ""Visible"", True" & vbCrLf
    s = s & "    mSheet.Activate" & vbCrLf
    s = s & "    Application.Goto mSheet.Range(mSheet.Cells(d.HeaderRow, d.FirstCol), _" & vbCrLf
    s = s & "                                  mSheet.Cells(d.LastDataRow, d.LastCol)), True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    AppActivate Application.Caption" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnSheetView_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSwitchMode_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim wantsAdmin As Boolean" & vbCrLf
    s = s & "    wantsAdmin = (StrComp(CStr(Me.cboRole.Value), ""Admin"", vbTextCompare) = 0)" & vbCrLf
    s = s & "    If wantsAdmin Then" & vbCrLf
    s = s & "        If modRaadApp.ValidateAdminLogin() Then" & vbCrLf
    s = s & "            gRole = ROLE_ADMIN" & vbCrLf
    s = s & "            modRaadApp.UnlockAllSheets" & vbCrLf
    s = s & "            ShowStatus ""Admin mode enabled. Sheets and workbook structure are unlocked."", False" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            gRole = ROLE_AGENT" & vbCrLf
    s = s & "            modRaadApp.LockForAgent" & vbCrLf
    s = s & "            ShowStatus ""Admin login failed. Agent mode remains active and locked."", True" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        gRole = ROLE_AGENT" & vbCrLf
    s = s & "        modRaadApp.LockForAgent" & vbCrLf
    s = s & "        ShowStatus ""Agent mode active. Structure and formula cells are protected."", False" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    ApplyRoleVisual" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnSwitchMode_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnUnlockAll_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If gRole <> ROLE_ADMIN Then" & vbCrLf
    s = s & "        ShowStatus ""Only Admin mode can unlock all sheets."", True" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    modRaadApp.UnlockAllSheets" & vbCrLf
    s = s & "    ShowStatus ""All sheets and workbook structure are unlocked for Admin editing."", False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnUnlockAll_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnLockForAgent_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If gRole <> ROLE_ADMIN Then" & vbCrLf
    s = s & "        ShowStatus ""Only Admin mode can change the lock state."", True" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    gRole = ROLE_AGENT" & vbCrLf
    s = s & "    modRaadApp.LockForAgent" & vbCrLf
    s = s & "    ApplyRoleVisual" & vbCrLf
    s = s & "    ShowStatus ""Agent mode restored. Structural editing is locked again."", False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnLockForAgent_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnClose_Click""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Hover events" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnAdd_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnAdd" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnAdd_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnUpdate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnUpdate" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnUpdate_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnDelete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnDelete" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnDelete_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnClear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnClear" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnClear_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnClose" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnClose_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnClearFilter_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnClearFilter" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnClearFilter_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnRefresh_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnRefresh" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnRefresh_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSheetView_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnSheetView" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnSheetView_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSwitchMode_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnSwitchMode" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnSwitchMode_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnUnlockAll_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnUnlockAll" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnUnlockAll_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnLockForAgent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    HighlightButton Me.btnLockForAgent" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""btnLockForAgent_MouseMove""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Public entry point called by ThisWorkbook" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ExternalSheetRefresh(ByVal sheetName As String, ByVal Target As Range)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    Dim watchRange As Range" & vbCrLf
    s = s & "    If mInitializing Or gAppWriting Then Exit Sub" & vbCrLf
    s = s & "    If mDatasetIndex < 1 Or mSheet Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    If StrComp(d.SheetName, sheetName, vbTextCompare) <> 0 Then Exit Sub" & vbCrLf
    s = s & "    Set watchRange = mSheet.Range(mSheet.Cells(d.HeaderRow, d.FirstCol), _" & vbCrLf
    s = s & "                                  mSheet.Cells(d.LastDataRow, d.LastCol))" & vbCrLf
    s = s & "    If Intersect(Target, watchRange) Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    RefreshGrid" & vbCrLf
    s = s & "    If mSelectedRow > 0 Then LoadRecordIntoForm mSelectedRow" & vbCrLf
    s = s & "    ShowStatus ""Worksheet edit detected. Grid and inputs refreshed."", False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ExternalSheetRefresh""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Initialization and visual design" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub InitializeApp()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    mInitializing = True" & vbCrLf
    s = s & "    modRaadApp.EnsureDatasets" & vbCrLf
    s = s & "    InitPalette" & vbCrLf
    s = s & "    ApplyLayout" & vbCrLf
    s = s & "    StyleAllControls" & vbCrLf
    s = s & "    ApplyCaptions" & vbCrLf
    s = s & "    ApplyPalette" & vbCrLf
    s = s & "    PopulateRoles" & vbCrLf
    s = s & "    PopulateDatasets" & vbCrLf
    s = s & "    mDatasetIndex = modRaadApp.FindMainDatasetIndex()" & vbCrLf
    s = s & "    If mDatasetIndex < 1 Then mDatasetIndex = 1" & vbCrLf
    s = s & "    Me.cboDataset.ListIndex = mDatasetIndex - 1" & vbCrLf
    s = s & "    gRole = ROLE_AGENT" & vbCrLf
    s = s & "    Me.cboRole.ListIndex = 0" & vbCrLf
    s = s & "    ApplyDataset" & vbCrLf
    s = s & "    ShowStatus ""Ready. Agent mode is active and every field is connected to the real sheet."", False" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    mInitializing = False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""InitializeApp""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub InitPalette()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    mBg = RGB(244, 247, 251)" & vbCrLf
    s = s & "    mHeader = RGB(16, 42, 67)" & vbCrLf
    s = s & "    mHeaderDark = RGB(12, 31, 50)" & vbCrLf
    s = s & "    mCard = RGB(255, 255, 255)" & vbCrLf
    s = s & "    mWhite = RGB(255, 255, 255)" & vbCrLf
    s = s & "    mText = RGB(15, 23, 42)" & vbCrLf
    s = s & "    mMuted = RGB(100, 116, 139)" & vbCrLf
    s = s & "    mBorder = RGB(203, 213, 225)" & vbCrLf
    s = s & "    mReadOnly = RGB(241, 245, 249)" & vbCrLf
    s = s & "    mPrimary = RGB(37, 99, 235)" & vbCrLf
    s = s & "    mPrimaryHover = RGB(29, 78, 216)" & vbCrLf
    s = s & "    mOrange = RGB(245, 158, 11)" & vbCrLf
    s = s & "    mOrangeHover = RGB(217, 119, 6)" & vbCrLf
    s = s & "    mDanger = RGB(220, 38, 38)" & vbCrLf
    s = s & "    mDangerHover = RGB(185, 28, 28)" & vbCrLf
    s = s & "    mGray = RGB(100, 116, 139)" & vbCrLf
    s = s & "    mGrayHover = RGB(71, 85, 105)" & vbCrLf
    s = s & "    mSuccess = RGB(22, 163, 74)" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""InitPalette""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ApplyLayout()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Me.Width = 1200" & vbCrLf
    s = s & "    Me.Height = 780" & vbCrLf
    s = s & "    Me.StartUpPosition = 2" & vbCrLf
    s = s & "    PlaceControl Me.fraHeader, 0, 0, 1200, 64" & vbCrLf
    s = s & "    PlaceControl Me.lblTitle, 24, 8, 620, 26" & vbCrLf
    s = s & "    PlaceControl Me.lblSubTitle, 24, 34, 760, 20" & vbCrLf
    s = s & "    PlaceControl Me.lblModeBadge, 970, 16, 90, 28" & vbCrLf
    s = s & "    PlaceControl Me.lblClock, 1070, 16, 110, 28" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PlaceControl Me.fraSidebar, 970, 76, 214, 622" & vbCrLf
    s = s & "    PlaceControl Me.lblNav, 986, 92, 180, 22" & vbCrLf
    s = s & "    PlaceControl Me.cboRole, 986, 120, 180, 26" & vbCrLf
    s = s & "    PlaceControl Me.btnSwitchMode, 986, 152, 180, 30" & vbCrLf
    s = s & "    PlaceControl Me.cboDataset, 986, 194, 180, 26" & vbCrLf
    s = s & "    PlaceControl Me.lblDatasetInfo, 986, 226, 180, 62" & vbCrLf
    s = s & "    PlaceControl Me.btnRefresh, 986, 298, 180, 30" & vbCrLf
    s = s & "    PlaceControl Me.btnSheetView, 986, 336, 180, 30" & vbCrLf
    s = s & "    PlaceControl Me.btnUnlockAll, 986, 380, 180, 30" & vbCrLf
    s = s & "    PlaceControl Me.btnLockForAgent, 986, 418, 180, 30" & vbCrLf
    s = s & "    PlaceControl Me.lblRecordCount, 986, 466, 180, 22" & vbCrLf
    s = s & "    PlaceControl Me.lblSelectedRow, 986, 494, 180, 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PlaceControl Me.fraToolbar, 16, 76, 938, 68" & vbCrLf
    s = s & "    PlaceControl Me.lblSearch, 28, 88, 240, 18" & vbCrLf
    s = s & "    PlaceControl Me.txtSearch, 28, 108, 250, 26" & vbCrLf
    s = s & "    PlaceControl Me.lblFilterField, 294, 88, 180, 18" & vbCrLf
    s = s & "    PlaceControl Me.cboFilterField, 294, 108, 180, 26" & vbCrLf
    s = s & "    PlaceControl Me.lblFilterValue, 490, 88, 180, 18" & vbCrLf
    s = s & "    PlaceControl Me.cboFilterValue, 490, 108, 180, 26" & vbCrLf
    s = s & "    PlaceControl Me.btnClearFilter, 686, 108, 120, 26" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PlaceControl Me.lstHeaders, 16, 152, 938, 24" & vbCrLf
    s = s & "    PlaceControl Me.lstData, 16, 176, 938, 282" & vbCrLf
    s = s & "    PlaceControl Me.fraEditor, 16, 466, 938, 210" & vbCrLf
    s = s & "    PlaceControl Me.lblEditorTitle, 28, 474, 400, 20" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    PlaceControl Me.fraButtons, 16, 684, 938, 48" & vbCrLf
    s = s & "    PlaceControl Me.btnAdd, 28, 694, 140, 30" & vbCrLf
    s = s & "    PlaceControl Me.btnUpdate, 180, 694, 140, 30" & vbCrLf
    s = s & "    PlaceControl Me.btnDelete, 332, 694, 140, 30" & vbCrLf
    s = s & "    PlaceControl Me.btnClear, 484, 694, 140, 30" & vbCrLf
    s = s & "    PlaceControl Me.btnClose, 800, 694, 140, 30" & vbCrLf
    s = s & "    PlaceControl Me.lblStatus, 16, 738, 1168, 26" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ApplyLayout""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ApplyCaptions()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Me.Caption = ""Raad | Live Inventory & Variance Console""" & vbCrLf
    s = s & "    SetCaptionSafe Me.fraHeader, """"" & vbCrLf
    s = s & "    SetCaptionSafe Me.fraSidebar, """"" & vbCrLf
    s = s & "    SetCaptionSafe Me.fraToolbar, """"" & vbCrLf
    s = s & "    SetCaptionSafe Me.fraEditor, """"" & vbCrLf
    s = s & "    SetCaptionSafe Me.fraButtons, """"" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblTitle, ""Raad Inventory & Daily Variance""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblSubTitle, ""Live two-way connection to your real worksheets, fields, formulas and dropdowns""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblModeBadge, ""AGENT""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblClock, Format(Now, ""yyyy-mm-dd hh:nn"")" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblNav, ""Workspace / Role""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblDatasetInfo, ""Select a real sheet dataset""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblRecordCount, ""Records: 0""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblSelectedRow, ""Selected row: none""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblSearch, ""Live search""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblFilterField, ""Filter by real column""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblFilterValue, ""Filter value""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblEditorTitle, ""Record fields - exact worksheet headers""" & vbCrLf
    s = s & "    SetCaptionSafe Me.lblStatus, ""Ready.""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnAdd, ""Add""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnUpdate, ""Update""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnDelete, ""Delete""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnClear, ""Clear""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnClose, ""Close""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnClearFilter, ""Clear Filter""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnRefresh, ""Refresh""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnSheetView, ""Switch to Excel Sheet""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnSwitchMode, ""Switch Mode""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnUnlockAll, ""Unlock Sheets""" & vbCrLf
    s = s & "    SetCaptionSafe Me.btnLockForAgent, ""Lock For Agent""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ApplyCaptions""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ApplyPalette()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Me.BackColor = mBg" & vbCrLf
    s = s & "    SetBack Me.fraHeader, mHeader" & vbCrLf
    s = s & "    SetBack Me.fraSidebar, mHeaderDark" & vbCrLf
    s = s & "    SetBack Me.fraToolbar, mCard" & vbCrLf
    s = s & "    SetBack Me.fraEditor, mCard" & vbCrLf
    s = s & "    SetBack Me.fraButtons, mCard" & vbCrLf
    s = s & "    SetBorderColor Me.fraToolbar, mBorder" & vbCrLf
    s = s & "    SetBorderColor Me.fraEditor, mBorder" & vbCrLf
    s = s & "    SetBorderColor Me.fraButtons, mBorder" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SetBack Me.lblTitle, mHeader" & vbCrLf
    s = s & "    SetFore Me.lblTitle, mWhite" & vbCrLf
    s = s & "    SetBack Me.lblSubTitle, mHeader" & vbCrLf
    s = s & "    SetFore Me.lblSubTitle, RGB(203, 213, 225)" & vbCrLf
    s = s & "    SetBack Me.lblModeBadge, mPrimary" & vbCrLf
    s = s & "    SetFore Me.lblModeBadge, mWhite" & vbCrLf
    s = s & "    SetBack Me.lblClock, mHeader" & vbCrLf
    s = s & "    SetFore Me.lblClock, mWhite" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SetBack Me.lblNav, mHeaderDark" & vbCrLf
    s = s & "    SetFore Me.lblNav, mWhite" & vbCrLf
    s = s & "    SetBack Me.lblDatasetInfo, mHeaderDark" & vbCrLf
    s = s & "    SetFore Me.lblDatasetInfo, mWhite" & vbCrLf
    s = s & "    SetBack Me.lblRecordCount, mHeaderDark" & vbCrLf
    s = s & "    SetFore Me.lblRecordCount, mWhite" & vbCrLf
    s = s & "    SetBack Me.lblSelectedRow, mHeaderDark" & vbCrLf
    s = s & "    SetFore Me.lblSelectedRow, mWhite" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SetBack Me.lblSearch, mCard" & vbCrLf
    s = s & "    SetFore Me.lblSearch, mText" & vbCrLf
    s = s & "    SetBack Me.lblFilterField, mCard" & vbCrLf
    s = s & "    SetFore Me.lblFilterField, mText" & vbCrLf
    s = s & "    SetBack Me.lblFilterValue, mCard" & vbCrLf
    s = s & "    SetFore Me.lblFilterValue, mText" & vbCrLf
    s = s & "    SetBack Me.lblEditorTitle, mCard" & vbCrLf
    s = s & "    SetFore Me.lblEditorTitle, mText" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SetBack Me.txtSearch, mWhite" & vbCrLf
    s = s & "    SetFore Me.txtSearch, mText" & vbCrLf
    s = s & "    SetBorderColor Me.txtSearch, mBorder" & vbCrLf
    s = s & "    SetBack Me.cboRole, mWhite" & vbCrLf
    s = s & "    SetFore Me.cboRole, mText" & vbCrLf
    s = s & "    SetBorderColor Me.cboRole, mBorder" & vbCrLf
    s = s & "    SetBack Me.cboDataset, mWhite" & vbCrLf
    s = s & "    SetFore Me.cboDataset, mText" & vbCrLf
    s = s & "    SetBorderColor Me.cboDataset, mBorder" & vbCrLf
    s = s & "    SetBack Me.cboFilterField, mWhite" & vbCrLf
    s = s & "    SetFore Me.cboFilterField, mText" & vbCrLf
    s = s & "    SetBorderColor Me.cboFilterField, mBorder" & vbCrLf
    s = s & "    SetBack Me.cboFilterValue, mWhite" & vbCrLf
    s = s & "    SetFore Me.cboFilterValue, mText" & vbCrLf
    s = s & "    SetBorderColor Me.cboFilterValue, mBorder" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SetBack Me.lstHeaders, mHeader" & vbCrLf
    s = s & "    SetFore Me.lstHeaders, mWhite" & vbCrLf
    s = s & "    SetBorderColor Me.lstHeaders, mBorder" & vbCrLf
    s = s & "    SetBack Me.lstData, mWhite" & vbCrLf
    s = s & "    SetFore Me.lstData, mText" & vbCrLf
    s = s & "    SetBorderColor Me.lstData, mBorder" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SetBack Me.lblStatus, mCard" & vbCrLf
    s = s & "    SetFore Me.lblStatus, mSuccess" & vbCrLf
    s = s & "    SetBorderColor Me.lblStatus, mBorder" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ResetButtonColors" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ApplyPalette""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub StyleAllControls()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim ctl As MSForms.Control" & vbCrLf
    s = s & "    Dim child As Object" & vbCrLf
    s = s & "    Dim container As Object" & vbCrLf
    s = s & "    For Each ctl In Me.Controls" & vbCrLf
    s = s & "        StyleControl ctl" & vbCrLf
    s = s & "        If TypeName(ctl) = ""Frame"" Then" & vbCrLf
    s = s & "            Set container = ctl" & vbCrLf
    s = s & "            For Each child In container.Controls" & vbCrLf
    s = s & "                StyleControl child" & vbCrLf
    s = s & "            Next child" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next ctl" & vbCrLf
    s = s & "    ApplyTitleFonts" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""StyleAllControls""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub StyleControl(ByVal ctl As Object)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    TrySetFont ctl, 9, False" & vbCrLf
    s = s & "    TrySet ctl, ""SpecialEffect"", 0" & vbCrLf
    s = s & "    TrySet ctl, ""BorderStyle"", 1" & vbCrLf
    s = s & "    TrySet ctl, ""BorderColor"", mBorder" & vbCrLf
    s = s & "    TrySet ctl, ""BackStyle"", 1" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""StyleControl""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ApplyTitleFonts()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    TrySetFont Me.lblTitle, 16, True" & vbCrLf
    s = s & "    TrySetFont Me.lblSubTitle, 9, False" & vbCrLf
    s = s & "    TrySetFont Me.lblModeBadge, 10, True" & vbCrLf
    s = s & "    TrySetFont Me.lblClock, 9, False" & vbCrLf
    s = s & "    TrySetFont Me.lblNav, 10, True" & vbCrLf
    s = s & "    TrySetFont Me.lblEditorTitle, 10, True" & vbCrLf
    s = s & "    TrySetFont Me.lblStatus, 9, True" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ApplyTitleFonts""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ResetButtonColors()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    SetBack Me.btnAdd, mPrimary" & vbCrLf
    s = s & "    SetBack Me.btnUpdate, mOrange" & vbCrLf
    s = s & "    SetBack Me.btnDelete, mDanger" & vbCrLf
    s = s & "    SetBack Me.btnClear, mGray" & vbCrLf
    s = s & "    SetBack Me.btnClose, mGray" & vbCrLf
    s = s & "    SetBack Me.btnClearFilter, mGray" & vbCrLf
    s = s & "    SetBack Me.btnRefresh, mPrimary" & vbCrLf
    s = s & "    SetBack Me.btnSheetView, mPrimary" & vbCrLf
    s = s & "    SetBack Me.btnSwitchMode, mPrimary" & vbCrLf
    s = s & "    SetBack Me.btnUnlockAll, mPrimary" & vbCrLf
    s = s & "    SetBack Me.btnLockForAgent, mGray" & vbCrLf
    s = s & "    SetFore Me.btnAdd, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnUpdate, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnDelete, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnClear, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnClose, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnClearFilter, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnRefresh, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnSheetView, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnSwitchMode, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnUnlockAll, mWhite" & vbCrLf
    s = s & "    SetFore Me.btnLockForAgent, mWhite" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ResetButtonColors""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub HighlightButton(ByVal btn As Object)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Select Case btn.Name" & vbCrLf
    s = s & "        Case ""btnAdd"", ""btnRefresh"", ""btnSheetView"", ""btnSwitchMode"", ""btnUnlockAll""" & vbCrLf
    s = s & "            c = mPrimaryHover" & vbCrLf
    s = s & "        Case ""btnUpdate""" & vbCrLf
    s = s & "            c = mOrangeHover" & vbCrLf
    s = s & "        Case ""btnDelete""" & vbCrLf
    s = s & "            c = mDangerHover" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            c = mGrayHover" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    SetBack btn, c" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""HighlightButton""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ShowStatus(ByVal message As String, Optional ByVal isError As Boolean = False)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Me.lblStatus.Caption = message" & vbCrLf
    s = s & "    If isError Then" & vbCrLf
    s = s & "        SetFore Me.lblStatus, mDanger" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        SetFore Me.lblStatus, mSuccess" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SetBack Me.lblStatus, mCard" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub HandleError(ByVal procName As String)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim n As Long" & vbCrLf
    s = s & "    Dim desc As String" & vbCrLf
    s = s & "    Dim msg As String" & vbCrLf
    s = s & "    n = Err.Number" & vbCrLf
    s = s & "    desc = Err.Description" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    msg = ""A recoverable error occurred in "" & procName & ""."" & vbCrLf & _" & vbCrLf
    s = s & "          ""Error "" & CStr(n) & "": "" & desc & vbCrLf & _" & vbCrLf
    s = s & "          ""No Debug/End window was shown. Please review the last action.""" & vbCrLf
    s = s & "    ShowStatus msg, True" & vbCrLf
    s = s & "    MsgBox msg, vbExclamation, ""Raad App""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub TrySet(ByVal ctl As Object, ByVal propName As String, ByVal value As Variant)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    CallByName ctl, propName, VbLet, value" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub TrySetFont(ByVal ctl As Object, ByVal size As Single, Optional ByVal bold As Boolean = False)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim f As Object" & vbCrLf
    s = s & "    Set f = CallByName(ctl, ""Font"", VbGet)" & vbCrLf
    s = s & "    CallByName f, ""Name"", VbLet, ""Segoe UI""" & vbCrLf
    s = s & "    CallByName f, ""Size"", VbLet, size" & vbCrLf
    s = s & "    CallByName f, ""Bold"", VbLet, bold" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub TrySetApplication(ByVal propName As String, ByVal value As Variant)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    CallByName Application, propName, VbLet, value" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetBack(ByVal ctl As Object, ByVal colorValue As Long)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    TrySet ctl, ""BackColor"", colorValue" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""SetBack""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetFore(ByVal ctl As Object, ByVal colorValue As Long)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    TrySet ctl, ""ForeColor"", colorValue" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""SetFore""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetBorderColor(ByVal ctl As Object, ByVal colorValue As Long)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    TrySet ctl, ""BorderColor"", colorValue" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""SetBorderColor""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetCaptionSafe(ByVal ctl As Object, ByVal caption As String)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    TrySet ctl, ""Caption"", caption" & vbCrLf
    s = s & "    TrySet ctl, ""AutoSize"", False" & vbCrLf
    s = s & "    TrySet ctl, ""WordWrap"", True" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""SetCaptionSafe""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PlaceControl(ByVal ctl As Object, ByVal leftValue As Single, ByVal topValue As Single, _" & vbCrLf
    s = s & "                         ByVal widthValue As Single, ByVal heightValue As Single)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    TrySet ctl, ""Left"", leftValue" & vbCrLf
    s = s & "    TrySet ctl, ""Top"", topValue" & vbCrLf
    s = s & "    TrySet ctl, ""Width"", widthValue" & vbCrLf
    s = s & "    TrySet ctl, ""Height"", heightValue" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""PlaceControl""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Dataset, role and dynamic field construction" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PopulateRoles()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim oldFlag As Boolean" & vbCrLf
    s = s & "    oldFlag = mSuppressEvents" & vbCrLf
    s = s & "    mSuppressEvents = True" & vbCrLf
    s = s & "    Me.cboRole.Clear" & vbCrLf
    s = s & "    Me.cboRole.AddItem ""Agent""" & vbCrLf
    s = s & "    Me.cboRole.AddItem ""Admin""" & vbCrLf
    s = s & "    Me.cboRole.ListIndex = 0" & vbCrLf
    s = s & "    TrySet Me.cboRole, ""Style"", 2" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    mSuppressEvents = oldFlag" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""PopulateRoles""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PopulateDatasets()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim oldFlag As Boolean" & vbCrLf
    s = s & "    oldFlag = mSuppressEvents" & vbCrLf
    s = s & "    mSuppressEvents = True" & vbCrLf
    s = s & "    ReDim mDatasets(1 To gDatasetCount)" & vbCrLf
    s = s & "    For i = 1 To gDatasetCount" & vbCrLf
    s = s & "        mDatasets(i) = gDatasets(i)" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    mDatasetCount = gDatasetCount" & vbCrLf
    s = s & "    Me.cboDataset.Clear" & vbCrLf
    s = s & "    For i = 1 To mDatasetCount" & vbCrLf
    s = s & "        Me.cboDataset.AddItem mDatasets(i).DisplayName" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    TrySet Me.cboDataset, ""Style"", 2" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    mSuppressEvents = oldFlag" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""PopulateDatasets""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ApplyDataset()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    If Me.cboDataset.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "    mDatasetIndex = Me.cboDataset.ListIndex + 1" & vbCrLf
    s = s & "    If mDatasetIndex < 1 Or mDatasetIndex > mDatasetCount Then Exit Sub" & vbCrLf
    s = s & "    d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    Set mSheet = ThisWorkbook.Worksheets(d.SheetName)" & vbCrLf
    s = s & "    BuildFieldMetadata" & vbCrLf
    s = s & "    BuildFieldControls" & vbCrLf
    s = s & "    PopulateFilterFields" & vbCrLf
    s = s & "    PopulateFilterValues" & vbCrLf
    s = s & "    ClearInputs" & vbCrLf
    s = s & "    mSelectedRow = 0" & vbCrLf
    s = s & "    Me.lblSelectedRow.Caption = ""Selected row: none""" & vbCrLf
    s = s & "    Me.lblDatasetInfo.Caption = d.DisplayName & vbCrLf & d.Description" & vbCrLf
    s = s & "    RefreshGrid" & vbCrLf
    s = s & "    ApplyRoleVisual" & vbCrLf
    s = s & "    StyleAllControls" & vbCrLf
    s = s & "    ShowStatus ""Dataset loaded: "" & d.DisplayName & "". Exact headers were read from the sheet."", False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ApplyDataset""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ApplyRoleVisual()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim isAdmin As Boolean" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    isAdmin = (gRole = ROLE_ADMIN)" & vbCrLf
    s = s & "    If mDatasetIndex > 0 Then d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    Me.btnAdd.Enabled = d.AllowAdd" & vbCrLf
    s = s & "    Me.btnUpdate.Enabled = d.AllowUpdate" & vbCrLf
    s = s & "    Me.btnDelete.Enabled = isAdmin And d.AllowDelete" & vbCrLf
    s = s & "    Me.btnUnlockAll.Enabled = isAdmin" & vbCrLf
    s = s & "    Me.btnLockForAgent.Enabled = isAdmin" & vbCrLf
    s = s & "    Me.btnSheetView.Enabled = True" & vbCrLf
    s = s & "    Me.btnRefresh.Enabled = True" & vbCrLf
    s = s & "    Me.btnSwitchMode.Enabled = True" & vbCrLf
    s = s & "    If isAdmin Then" & vbCrLf
    s = s & "        Me.lblModeBadge.Caption = ""ADMIN""" & vbCrLf
    s = s & "        SetBack Me.lblModeBadge, mDanger" & vbCrLf
    s = s & "        Me.cboRole.ListIndex = 1" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Me.lblModeBadge.Caption = ""AGENT""" & vbCrLf
    s = s & "        SetBack Me.lblModeBadge, mPrimary" & vbCrLf
    s = s & "        Me.cboRole.ListIndex = 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SetFore Me.lblModeBadge, mWhite" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ApplyRoleVisual""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildFieldMetadata()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim colCount As Long" & vbCrLf
    s = s & "    Dim vlist As String" & vbCrLf
    s = s & "    d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    colCount = d.LastCol - d.FirstCol + 1" & vbCrLf
    s = s & "    ReDim mFields(1 To colCount)" & vbCrLf
    s = s & "    mFieldCount = colCount" & vbCrLf
    s = s & "    For i = 1 To colCount" & vbCrLf
    s = s & "        mFields(i).Index = i" & vbCrLf
    s = s & "        mFields(i).ColNumber = d.FirstCol + i - 1" & vbCrLf
    s = s & "        mFields(i).ColLetter = ColumnLetter(mFields(i).ColNumber)" & vbCrLf
    s = s & "        mFields(i).Header = CStr(mSheet.Cells(d.HeaderRow, mFields(i).ColNumber).Value)" & vbCrLf
    s = s & "        If Len(Trim$(mFields(i).Header)) = 0 Then" & vbCrLf
    s = s & "            mFields(i).LabelText = ""["" & mFields(i).ColLetter & ""] (no header)""" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mFields(i).LabelText = mFields(i).Header" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        vlist = GetValidationList(d, mFields(i).ColNumber)" & vbCrLf
    s = s & "        If Len(vlist) > 0 Then" & vbCrLf
    s = s & "            mFields(i).Kind = ""LIST""" & vbCrLf
    s = s & "            mFields(i).Editable = True" & vbCrLf
    s = s & "            mFields(i).InputName = ""cboF"" & CStr(i)" & vbCrLf
    s = s & "            mFields(i).ValidationList = vlist" & vbCrLf
    s = s & "        ElseIf IsFormulaColumn(d, mFields(i).ColNumber) Then" & vbCrLf
    s = s & "            mFields(i).Kind = ""FORMULA""" & vbCrLf
    s = s & "            mFields(i).Editable = False" & vbCrLf
    s = s & "            mFields(i).InputName = ""txtF"" & CStr(i)" & vbCrLf
    s = s & "            mFields(i).ValidationList = """"" & vbCrLf
    s = s & "        ElseIf IsDateColumn(d, mFields(i).ColNumber) Then" & vbCrLf
    s = s & "            mFields(i).Kind = ""DATE""" & vbCrLf
    s = s & "            mFields(i).Editable = True" & vbCrLf
    s = s & "            mFields(i).InputName = ""txtF"" & CStr(i)" & vbCrLf
    s = s & "            mFields(i).ValidationList = """"" & vbCrLf
    s = s & "        ElseIf IsNumberColumn(d, mFields(i).ColNumber) Then" & vbCrLf
    s = s & "            mFields(i).Kind = ""NUMBER""" & vbCrLf
    s = s & "            mFields(i).Editable = True" & vbCrLf
    s = s & "            mFields(i).InputName = ""txtF"" & CStr(i)" & vbCrLf
    s = s & "            mFields(i).ValidationList = """"" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mFields(i).Kind = ""TEXT""" & vbCrLf
    s = s & "            mFields(i).Editable = True" & vbCrLf
    s = s & "            mFields(i).InputName = ""txtF"" & CStr(i)" & vbCrLf
    s = s & "            mFields(i).ValidationList = """"" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""BuildFieldMetadata""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildFieldControls()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim k As Long" & vbCrLf
    s = s & "    Dim nm As String" & vbCrLf
    s = s & "    Dim nCols As Long" & vbCrLf
    s = s & "    Dim rowPos As Long" & vbCrLf
    s = s & "    Dim colPos As Long" & vbCrLf
    s = s & "    Dim rows As Long" & vbCrLf
    s = s & "    Dim leftPos As Single" & vbCrLf
    s = s & "    Dim topPos As Single" & vbCrLf
    s = s & "    Dim lbl As Object" & vbCrLf
    s = s & "    Dim inp As Object" & vbCrLf
    s = s & "    Dim items As Variant" & vbCrLf
    s = s & "    Dim itemText As String" & vbCrLf
    s = s & "    For i = Me.fraEditor.Controls.Count - 1 To 0 Step -1" & vbCrLf
    s = s & "        nm = Me.fraEditor.Controls(i).Name" & vbCrLf
    s = s & "        If Left$(nm, 4) = ""lblF"" Or Left$(nm, 4) = ""txtF"" Or Left$(nm, 4) = ""cboF"" Then" & vbCrLf
    s = s & "            Me.fraEditor.Controls.Remove nm" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    nCols = 3" & vbCrLf
    s = s & "    rows = Int((mFieldCount + nCols - 1) / nCols)" & vbCrLf
    s = s & "    TrySet Me.fraEditor, ""ScrollBars"", 2" & vbCrLf
    s = s & "    TrySet Me.fraEditor, ""ScrollHeight"", rows * 62 + 60" & vbCrLf
    s = s & "    TrySet Me.fraEditor, ""ScrollTop"", 0" & vbCrLf
    s = s & "    TrySet Me.fraEditor, ""KeepScrollBarsVisible"", 2" & vbCrLf
    s = s & "    For i = 1 To mFieldCount" & vbCrLf
    s = s & "        colPos = (i - 1) Mod nCols" & vbCrLf
    s = s & "        rowPos = Int((i - 1) / nCols)" & vbCrLf
    s = s & "        leftPos = 12 + colPos * 302" & vbCrLf
    s = s & "        topPos = 34 + rowPos * 62" & vbCrLf
    s = s & "        Set lbl = Me.fraEditor.Controls.Add(""Forms.Label.1"", ""lblF"" & CStr(i), True)" & vbCrLf
    s = s & "        PlaceControl lbl, leftPos, topPos, 280, 18" & vbCrLf
    s = s & "        lbl.Caption = mFields(i).LabelText" & vbCrLf
    s = s & "        SetBack lbl, mCard" & vbCrLf
    s = s & "        SetFore lbl, mText" & vbCrLf
    s = s & "        TrySetFont lbl, 9, True" & vbCrLf
    s = s & "        TrySet lbl, ""TextAlign"", 3" & vbCrLf
    s = s & "        TrySet lbl, ""RightToLeft"", True" & vbCrLf
    s = s & "        TrySet lbl, ""SpecialEffect"", 0" & vbCrLf
    s = s & "        TrySet lbl, ""BorderStyle"", 1" & vbCrLf
    s = s & "        If mFields(i).Kind = ""LIST"" Then" & vbCrLf
    s = s & "            Set inp = Me.fraEditor.Controls.Add(""Forms.ComboBox.1"", mFields(i).InputName, True)" & vbCrLf
    s = s & "            TrySet inp, ""Style"", 0" & vbCrLf
    s = s & "            inp.Clear" & vbCrLf
    s = s & "            items = Split(mFields(i).ValidationList, ""|"")" & vbCrLf
    s = s & "            For k = LBound(items) To UBound(items)" & vbCrLf
    s = s & "                itemText = Trim$(CStr(items(k)))" & vbCrLf
    s = s & "                If Len(itemText) > 0 Then inp.AddItem itemText" & vbCrLf
    s = s & "            Next k" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            Set inp = Me.fraEditor.Controls.Add(""Forms.TextBox.1"", mFields(i).InputName, True)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        PlaceControl inp, leftPos, topPos + 20, 280, 26" & vbCrLf
    s = s & "        SetBack inp, mWhite" & vbCrLf
    s = s & "        SetFore inp, mText" & vbCrLf
    s = s & "        SetBorderColor inp, mBorder" & vbCrLf
    s = s & "        TrySetFont inp, 9, False" & vbCrLf
    s = s & "        TrySet inp, ""SpecialEffect"", 0" & vbCrLf
    s = s & "        TrySet inp, ""BorderStyle"", 1" & vbCrLf
    s = s & "        TrySet inp, ""TextAlign"", 3" & vbCrLf
    s = s & "        TrySet inp, ""RightToLeft"", True" & vbCrLf
    s = s & "        If Not mFields(i).Editable Then" & vbCrLf
    s = s & "            TrySet inp, ""Locked"", True" & vbCrLf
    s = s & "            TrySet inp, ""TabStop"", False" & vbCrLf
    s = s & "            SetBack inp, mReadOnly" & vbCrLf
    s = s & "            SetFore inp, mMuted" & vbCrLf
    s = s & "            TrySet inp, ""ControlTipText"", ""Read-only: calculated by Excel. The formula is never overwritten.""" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    StyleAllControls" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""BuildFieldControls""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ColumnLetter(ByVal colNumber As Long) As String" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ColumnLetter = Split(mSheet.Cells(1, colNumber).Address(True, False), ""$"")(0)" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ColumnLetter""" & vbCrLf
    s = s & "    ColumnLetter = ""A""" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsFormulaColumn(ByVal d As AppDataset, ByVal col As Long) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim stopRow As Long" & vbCrLf
    s = s & "    stopRow = d.FirstDataRow + 20" & vbCrLf
    s = s & "    If stopRow > d.LastDataRow Then stopRow = d.LastDataRow" & vbCrLf
    s = s & "    For r = d.FirstDataRow To stopRow" & vbCrLf
    s = s & "        If mSheet.Cells(r, col).HasFormula Then" & vbCrLf
    s = s & "            IsFormulaColumn = True" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    IsFormulaColumn = False" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""IsFormulaColumn""" & vbCrLf
    s = s & "    IsFormulaColumn = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsDateColumn(ByVal d As AppDataset, ByVal col As Long) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim stopRow As Long" & vbCrLf
    s = s & "    Dim fmt As String" & vbCrLf
    s = s & "    stopRow = d.FirstDataRow + 20" & vbCrLf
    s = s & "    If stopRow > d.LastDataRow Then stopRow = d.LastDataRow" & vbCrLf
    s = s & "    For r = d.FirstDataRow To stopRow" & vbCrLf
    s = s & "        If Not mSheet.Cells(r, col).HasFormula Then" & vbCrLf
    s = s & "            If IsDate(mSheet.Cells(r, col).Value) Then" & vbCrLf
    s = s & "                IsDateColumn = True" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            fmt = LCase$(mSheet.Cells(r, col).NumberFormat)" & vbCrLf
    s = s & "            If InStr(1, fmt, ""yy"", vbTextCompare) > 0 Or InStr(1, fmt, ""dd"", vbTextCompare) > 0 Then" & vbCrLf
    s = s & "                IsDateColumn = True" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    IsDateColumn = False" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""IsDateColumn""" & vbCrLf
    s = s & "    IsDateColumn = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsNumberColumn(ByVal d As AppDataset, ByVal col As Long) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim stopRow As Long" & vbCrLf
    s = s & "    Dim fmt As String" & vbCrLf
    s = s & "    stopRow = d.FirstDataRow + 20" & vbCrLf
    s = s & "    If stopRow > d.LastDataRow Then stopRow = d.LastDataRow" & vbCrLf
    s = s & "    For r = d.FirstDataRow To stopRow" & vbCrLf
    s = s & "        If Not mSheet.Cells(r, col).HasFormula Then" & vbCrLf
    s = s & "            If IsNumeric(mSheet.Cells(r, col).Value) And Not IsDate(mSheet.Cells(r, col).Value) Then" & vbCrLf
    s = s & "                IsNumberColumn = True" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            fmt = LCase$(mSheet.Cells(r, col).NumberFormat)" & vbCrLf
    s = s & "            If InStr(1, fmt, ""0"", vbTextCompare) > 0 Or InStr(1, fmt, ""#"", vbTextCompare) > 0 Then" & vbCrLf
    s = s & "                IsNumberColumn = True" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    IsNumberColumn = False" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""IsNumberColumn""" & vbCrLf
    s = s & "    IsNumberColumn = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function GetValidationList(ByVal d As AppDataset, ByVal col As Long) As String" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim f As String" & vbCrLf
    s = s & "    Dim arr As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim part As String" & vbCrLf
    s = s & "    Dim result As String" & vbCrLf
    s = s & "    If mSheet.Cells(d.FirstDataRow, col).Validation.Type <> xlValidateList Then Exit Function" & vbCrLf
    s = s & "    f = mSheet.Cells(d.FirstDataRow, col).Validation.Formula1" & vbCrLf
    s = s & "    If Left$(f, 1) = """""""" Then f = Mid$(f, 2, Len(f) - 2)" & vbCrLf
    s = s & "    f = Replace(f, "";"", ""|"")" & vbCrLf
    s = s & "    f = Replace(f, ""،"", ""|"")" & vbCrLf
    s = s & "    f = Replace(f, "","", ""|"")" & vbCrLf
    s = s & "    f = Replace(f, vbCr, ""|"")" & vbCrLf
    s = s & "    f = Replace(f, vbLf, ""|"")" & vbCrLf
    s = s & "    arr = Split(f, ""|"")" & vbCrLf
    s = s & "    For i = LBound(arr) To UBound(arr)" & vbCrLf
    s = s & "        part = Trim$(Replace$(CStr(arr(i)), """""""", """"))" & vbCrLf
    s = s & "        If Len(part) > 0 Then" & vbCrLf
    s = s & "            If Len(result) > 0 Then result = result & ""|""" & vbCrLf
    s = s & "            result = result & part" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    GetValidationList = result" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Filters, grid and record loading" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PopulateFilterFields()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim oldFlag As Boolean" & vbCrLf
    s = s & "    oldFlag = mSuppressEvents" & vbCrLf
    s = s & "    mSuppressEvents = True" & vbCrLf
    s = s & "    Me.cboFilterField.Clear" & vbCrLf
    s = s & "    Me.cboFilterField.AddItem ""<All fields>""" & vbCrLf
    s = s & "    For i = 1 To mFieldCount" & vbCrLf
    s = s & "        Me.cboFilterField.AddItem mFields(i).LabelText" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Me.cboFilterField.ListIndex = 0" & vbCrLf
    s = s & "    TrySet Me.cboFilterField, ""Style"", 2" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    mSuppressEvents = oldFlag" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""PopulateFilterFields""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PopulateFilterValues()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    Dim fld As Long" & vbCrLf
    s = s & "    Dim col As Long" & vbCrLf
    s = s & "    Dim values As String" & vbCrLf
    s = s & "    Dim arr As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim t As String" & vbCrLf
    s = s & "    Dim oldFlag As Boolean" & vbCrLf
    s = s & "    oldFlag = mSuppressEvents" & vbCrLf
    s = s & "    mSuppressEvents = True" & vbCrLf
    s = s & "    Me.cboFilterValue.Clear" & vbCrLf
    s = s & "    Me.cboFilterValue.AddItem ""<All values>""" & vbCrLf
    s = s & "    fld = Me.cboFilterField.ListIndex" & vbCrLf
    s = s & "    If fld > 0 And mDatasetIndex > 0 Then" & vbCrLf
    s = s & "        d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "        col = d.FirstCol + fld - 1" & vbCrLf
    s = s & "        values = GetDistinctColumnValues(d, col)" & vbCrLf
    s = s & "        If Len(values) > 0 Then" & vbCrLf
    s = s & "            arr = Split(values, ""|"")" & vbCrLf
    s = s & "            For i = LBound(arr) To UBound(arr)" & vbCrLf
    s = s & "                t = Trim$(CStr(arr(i)))" & vbCrLf
    s = s & "                If Len(t) > 0 Then Me.cboFilterValue.AddItem t" & vbCrLf
    s = s & "            Next i" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Me.cboFilterValue.ListIndex = 0" & vbCrLf
    s = s & "    TrySet Me.cboFilterValue, ""Style"", 0" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    mSuppressEvents = oldFlag" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""PopulateFilterValues""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function GetDistinctColumnValues(ByVal d As AppDataset, ByVal col As Long) As String" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim coll As Collection" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim t As String" & vbCrLf
    s = s & "    Dim result As String" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    Set coll = New Collection" & vbCrLf
    s = s & "    For r = d.FirstDataRow To d.LastDataRow" & vbCrLf
    s = s & "        t = Trim$(CellDisplayText(mSheet.Cells(r, col)))" & vbCrLf
    s = s & "        If Len(t) > 240 Then t = Left$(t, 240)" & vbCrLf
    s = s & "        If Len(t) > 0 Then" & vbCrLf
    s = s & "            If Not CollectionContains(coll, t) Then coll.Add t, t" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    For Each v In coll" & vbCrLf
    s = s & "        If Len(result) > 0 Then result = result & ""|""" & vbCrLf
    s = s & "        result = result & CStr(v)" & vbCrLf
    s = s & "    Next v" & vbCrLf
    s = s & "    GetDistinctColumnValues = result" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""GetDistinctColumnValues""" & vbCrLf
    s = s & "    GetDistinctColumnValues = """"" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function CollectionContains(ByVal coll As Collection, ByVal key As String) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim tmp As Variant" & vbCrLf
    s = s & "    tmp = coll.Item(key)" & vbCrLf
    s = s & "    CollectionContains = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    CollectionContains = False" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub RefreshGrid()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    Dim colCount As Long" & vbCrLf
    s = s & "    Dim maxRows As Long" & vbCrLf
    s = s & "    Dim widths As String" & vbCrLf
    s = s & "    Dim headerArr() As Variant" & vbCrLf
    s = s & "    Dim tempArr() As Variant" & vbCrLf
    s = s & "    Dim finalArr() As Variant" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim j As Long" & vbCrLf
    s = s & "    Dim rowCount As Long" & vbCrLf
    s = s & "    If mDatasetIndex < 1 Or mSheet Is Nothing Then Exit Sub" & vbCrLf
    s = s & "    d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    colCount = d.LastCol - d.FirstCol + 1" & vbCrLf
    s = s & "    widths = BuildColumnWidths(colCount)" & vbCrLf
    s = s & "    SetGridColumns colCount, widths" & vbCrLf
    s = s & "    ReDim headerArr(1 To 1, 1 To colCount)" & vbCrLf
    s = s & "    For c = 1 To colCount" & vbCrLf
    s = s & "        headerArr(1, c) = mFields(c).LabelText" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "    Me.lstHeaders.List = headerArr" & vbCrLf
    s = s & "    maxRows = d.LastDataRow - d.FirstDataRow + 1" & vbCrLf
    s = s & "    If maxRows < 1 Then" & vbCrLf
    s = s & "        Me.lstData.Clear" & vbCrLf
    s = s & "        mRowCount = 0" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    ReDim tempArr(1 To maxRows, 1 To colCount)" & vbCrLf
    s = s & "    ReDim mRowMap(1 To maxRows)" & vbCrLf
    s = s & "    rowCount = 0" & vbCrLf
    s = s & "    For r = d.FirstDataRow To d.LastDataRow" & vbCrLf
    s = s & "        If IsDataRow(r, d) Then" & vbCrLf
    s = s & "            If RowMatchesFilters(r, d) Then" & vbCrLf
    s = s & "                rowCount = rowCount + 1" & vbCrLf
    s = s & "                mRowMap(rowCount) = r" & vbCrLf
    s = s & "                For c = 1 To colCount" & vbCrLf
    s = s & "                    tempArr(rowCount, c) = CellDisplayText(mSheet.Cells(r, d.FirstCol + c - 1))" & vbCrLf
    s = s & "                Next c" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    mRowCount = rowCount" & vbCrLf
    s = s & "    If rowCount > 0 Then" & vbCrLf
    s = s & "        ReDim finalArr(1 To rowCount, 1 To colCount)" & vbCrLf
    s = s & "        For i = 1 To rowCount" & vbCrLf
    s = s & "            For j = 1 To colCount" & vbCrLf
    s = s & "                finalArr(i, j) = tempArr(i, j)" & vbCrLf
    s = s & "            Next j" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "        Me.lstData.List = finalArr" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Me.lstData.Clear" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Me.lblRecordCount.Caption = ""Records: "" & CStr(rowCount)" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""RefreshGrid""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetGridColumns(ByVal colCount As Long, ByVal widths As String)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    TrySet Me.lstHeaders, ""ColumnCount"", colCount" & vbCrLf
    s = s & "    TrySet Me.lstHeaders, ""ColumnWidths"", widths" & vbCrLf
    s = s & "    TrySet Me.lstHeaders, ""Enabled"", False" & vbCrLf
    s = s & "    TrySet Me.lstHeaders, ""SpecialEffect"", 0" & vbCrLf
    s = s & "    TrySet Me.lstHeaders, ""BorderStyle"", 1" & vbCrLf
    s = s & "    TrySet Me.lstData, ""ColumnCount"", colCount" & vbCrLf
    s = s & "    TrySet Me.lstData, ""ColumnWidths"", widths" & vbCrLf
    s = s & "    TrySet Me.lstData, ""BoundColumn"", 1" & vbCrLf
    s = s & "    TrySet Me.lstData, ""MultiSelect"", 0" & vbCrLf
    s = s & "    TrySet Me.lstData, ""ListStyle"", 0" & vbCrLf
    s = s & "    TrySet Me.lstData, ""IntegralHeight"", False" & vbCrLf
    s = s & "    TrySet Me.lstData, ""ScrollBars"", 3" & vbCrLf
    s = s & "    TrySet Me.lstData, ""SpecialEffect"", 0" & vbCrLf
    s = s & "    TrySet Me.lstData, ""BorderStyle"", 1" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""SetGridColumns""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function BuildColumnWidths(ByVal colCount As Long) As String" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim w As Long" & vbCrLf
    s = s & "    Dim result As String" & vbCrLf
    s = s & "    For i = 1 To colCount" & vbCrLf
    s = s & "        w = Len(mFields(i).LabelText) * 7 + 30" & vbCrLf
    s = s & "        If w < 85 Then w = 85" & vbCrLf
    s = s & "        If w > 190 Then w = 190" & vbCrLf
    s = s & "        If i > 1 Then result = result & "";""" & vbCrLf
    s = s & "        result = result & CStr(w)" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    BuildColumnWidths = result" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""BuildColumnWidths""" & vbCrLf
    s = s & "    BuildColumnWidths = ""100""" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsDataRow(ByVal row As Long, ByVal d As AppDataset) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    For c = d.FirstCol To d.LastCol" & vbCrLf
    s = s & "        If Len(Trim$(CellDisplayText(mSheet.Cells(row, c)))) > 0 Then" & vbCrLf
    s = s & "            IsDataRow = True" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "    IsDataRow = False" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""IsDataRow""" & vbCrLf
    s = s & "    IsDataRow = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function RowMatchesFilters(ByVal row As Long, ByVal d As AppDataset) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim q As String" & vbCrLf
    s = s & "    Dim fld As Long" & vbCrLf
    s = s & "    Dim fv As String" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim t As String" & vbCrLf
    s = s & "    Dim found As Boolean" & vbCrLf
    s = s & "    q = Trim$(Me.txtSearch.Text)" & vbCrLf
    s = s & "    fld = Me.cboFilterField.ListIndex" & vbCrLf
    s = s & "    fv = Trim$(CStr(Me.cboFilterValue.Value))" & vbCrLf
    s = s & "    If StrComp(fv, ""<All values>"", vbTextCompare) = 0 Then fv = """"" & vbCrLf
    s = s & "    If Len(q) > 0 Then" & vbCrLf
    s = s & "        found = False" & vbCrLf
    s = s & "        For c = d.FirstCol To d.LastCol" & vbCrLf
    s = s & "            t = CellDisplayText(mSheet.Cells(row, c))" & vbCrLf
    s = s & "            If InStr(1, t, q, vbTextCompare) > 0 Then" & vbCrLf
    s = s & "                found = True" & vbCrLf
    s = s & "                Exit For" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next c" & vbCrLf
    s = s & "        If Not found Then" & vbCrLf
    s = s & "            RowMatchesFilters = False" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If fld > 0 And Len(fv) > 0 Then" & vbCrLf
    s = s & "        t = CellDisplayText(mSheet.Cells(row, d.FirstCol + fld - 1))" & vbCrLf
    s = s & "        If StrComp(Trim$(t), fv, vbTextCompare) <> 0 Then" & vbCrLf
    s = s & "            RowMatchesFilters = False" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    RowMatchesFilters = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""RowMatchesFilters""" & vbCrLf
    s = s & "    RowMatchesFilters = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function CellDisplayText(ByVal cell As Range) As String" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    If IsError(cell.Value) Then" & vbCrLf
    s = s & "        CellDisplayText = ""#ERROR""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cell.HasFormula Then" & vbCrLf
    s = s & "        CellDisplayText = CStr(cell.Text)" & vbCrLf
    s = s & "    ElseIf IsDate(cell.Value) Then" & vbCrLf
    s = s & "        CellDisplayText = Format(cell.Value, ""yyyy-mm-dd"")" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        CellDisplayText = CStr(cell.Value)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""CellDisplayText""" & vbCrLf
    s = s & "    CellDisplayText = """"" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadRecordIntoForm(ByVal row As Long)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    mSelectedRow = row" & vbCrLf
    s = s & "    For i = 1 To mFieldCount" & vbCrLf
    s = s & "        SetInputValue i, CellDisplayText(mSheet.Cells(row, mFields(i).ColNumber))" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Me.lblSelectedRow.Caption = ""Selected row: "" & CStr(row)" & vbCrLf
    s = s & "    ShowStatus ""Loaded worksheet row "" & CStr(row) & "" into every field."", False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""LoadRecordIntoForm""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FormControlByName(ByVal controlName As String) As Object" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim ctl As MSForms.Control" & vbCrLf
    s = s & "    Set FormControlByName = Nothing" & vbCrLf
    s = s & "    For Each ctl In Me.fraEditor.Controls" & vbCrLf
    s = s & "        If StrComp(ctl.Name, controlName, vbTextCompare) = 0 Then" & vbCrLf
    s = s & "            Set FormControlByName = ctl" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next ctl" & vbCrLf
    s = s & "    For Each ctl In Me.Controls" & vbCrLf
    s = s & "        If StrComp(ctl.Name, controlName, vbTextCompare) = 0 Then" & vbCrLf
    s = s & "            Set FormControlByName = ctl" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next ctl" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""FormControlByName""" & vbCrLf
    s = s & "    Set FormControlByName = Nothing" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function GetInputValue(ByVal fieldIndex As Long) As String" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim ctl As Object" & vbCrLf
    s = s & "    Set ctl = FormControlByName(mFields(fieldIndex).InputName)" & vbCrLf
    s = s & "    If ctl Is Nothing Then" & vbCrLf
    s = s & "        GetInputValue = """"" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        GetInputValue = CStr(ctl.Value)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""GetInputValue""" & vbCrLf
    s = s & "    GetInputValue = """"" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetInputValue(ByVal fieldIndex As Long, ByVal value As String)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim ctl As Object" & vbCrLf
    s = s & "    Set ctl = FormControlByName(mFields(fieldIndex).InputName)" & vbCrLf
    s = s & "    If Not ctl Is Nothing Then TrySet ctl, ""Value"", value" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""SetInputValue""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ClearInputs()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim oldFlag As Boolean" & vbCrLf
    s = s & "    oldFlag = mSuppressEvents" & vbCrLf
    s = s & "    mSuppressEvents = True" & vbCrLf
    s = s & "    For i = 1 To mFieldCount" & vbCrLf
    s = s & "        SetInputValue i, """"" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    mSelectedRow = 0" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    mSuppressEvents = oldFlag" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ClearInputs""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "' Validation, CRUD and auto-save" & vbCrLf
    s = s & "'-------------------------------------------------------------------------------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsRequiredColumn(ByVal d As AppDataset, ByVal fieldIndex As Long) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim arr As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    If Len(Trim$(d.RequiredCols)) = 0 Then" & vbCrLf
    s = s & "        IsRequiredColumn = False" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    arr = Split(d.RequiredCols, "","")" & vbCrLf
    s = s & "    For i = LBound(arr) To UBound(arr)" & vbCrLf
    s = s & "        If StrComp(Trim$(CStr(arr(i))), CStr(fieldIndex), vbTextCompare) = 0 Then" & vbCrLf
    s = s & "            IsRequiredColumn = True" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    IsRequiredColumn = False" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""IsRequiredColumn""" & vbCrLf
    s = s & "    IsRequiredColumn = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ValidateInputs(ByVal isAdd As Boolean) As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim v As String" & vbCrLf
    s = s & "    d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    For i = 1 To mFieldCount" & vbCrLf
    s = s & "        If mFields(i).Editable Then" & vbCrLf
    s = s & "            v = Trim$(GetInputValue(i))" & vbCrLf
    s = s & "            If IsRequiredColumn(d, i) And Len(v) = 0 Then" & vbCrLf
    s = s & "                ShowStatus ""Required field is empty: "" & mFields(i).LabelText, True" & vbCrLf
    s = s & "                ValidateInputs = False" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            If Len(v) > 0 Then" & vbCrLf
    s = s & "                Select Case mFields(i).Kind" & vbCrLf
    s = s & "                    Case ""DATE""" & vbCrLf
    s = s & "                        If Not IsDate(v) Then" & vbCrLf
    s = s & "                            ShowStatus ""Invalid date in "" & mFields(i).LabelText & "". Use a real date."", True" & vbCrLf
    s = s & "                            ValidateInputs = False" & vbCrLf
    s = s & "                            Exit Function" & vbCrLf
    s = s & "                        End If" & vbCrLf
    s = s & "                    Case ""NUMBER""" & vbCrLf
    s = s & "                        If Not IsNumeric(v) Then" & vbCrLf
    s = s & "                            ShowStatus ""Invalid number in "" & mFields(i).LabelText & ""."", True" & vbCrLf
    s = s & "                            ValidateInputs = False" & vbCrLf
    s = s & "                            Exit Function" & vbCrLf
    s = s & "                        End If" & vbCrLf
    s = s & "                End Select" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    ValidateInputs = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ValidateInputs""" & vbCrLf
    s = s & "    ValidateInputs = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub AddRecord()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    Dim newRow As Long" & vbCrLf
    s = s & "    Dim oldEvents As Boolean" & vbCrLf
    s = s & "    Dim oldAlerts As Boolean" & vbCrLf
    s = s & "    Dim oldScreen As Boolean" & vbCrLf
    s = s & "    Dim oldCalc As Long" & vbCrLf
    s = s & "    Dim saved As Boolean" & vbCrLf
    s = s & "    d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    If Not d.AllowAdd Then" & vbCrLf
    s = s & "        ShowStatus ""Add is disabled for this fixed formula layout. The daily archive dataset supports new records."", True" & vbCrLf
    s = s & "        GoTo SafeExit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not ValidateInputs(True) Then GoTo SafeExit" & vbCrLf
    s = s & "    newRow = FindNextBlankRow(d)" & vbCrLf
    s = s & "    If newRow < 1 Then" & vbCrLf
    s = s & "        ShowStatus ""No empty row remains inside the configured real data range."", True" & vbCrLf
    s = s & "        GoTo SafeExit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    oldEvents = Application.EnableEvents" & vbCrLf
    s = s & "    oldAlerts = Application.DisplayAlerts" & vbCrLf
    s = s & "    oldScreen = Application.ScreenUpdating" & vbCrLf
    s = s & "    oldCalc = Application.Calculation" & vbCrLf
    s = s & "    gAppWriting = True" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "    Application.DisplayAlerts = False" & vbCrLf
    s = s & "    Application.ScreenUpdating = False" & vbCrLf
    s = s & "    Application.Calculation = xlCalculationManual" & vbCrLf
    s = s & "    CopyRowFormat mSheet, d, d.FirstDataRow, newRow" & vbCrLf
    s = s & "    WriteInputsToRow newRow" & vbCrLf
    s = s & "    ExtendFormulaColumns d, newRow" & vbCrLf
    s = s & "    mSheet.Calculate" & vbCrLf
    s = s & "    saved = SaveWorkbookSafely()" & vbCrLf
    s = s & "    If saved Then" & vbCrLf
    s = s & "        ShowStatus ""Added row "" & CStr(newRow) & "" and saved the workbook."", False" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ShowStatus ""Row was added in memory but the workbook was not saved."", True" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    TrySetApplication ""Calculation"", oldCalc" & vbCrLf
    s = s & "    TrySetApplication ""EnableEvents"", oldEvents" & vbCrLf
    s = s & "    TrySetApplication ""DisplayAlerts"", oldAlerts" & vbCrLf
    s = s & "    TrySetApplication ""ScreenUpdating"", oldScreen" & vbCrLf
    s = s & "    gAppWriting = False" & vbCrLf
    s = s & "    If saved Then" & vbCrLf
    s = s & "        LoadRecordIntoForm newRow" & vbCrLf
    s = s & "        RefreshGrid" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""AddRecord""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateRecord()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    Dim oldEvents As Boolean" & vbCrLf
    s = s & "    Dim oldAlerts As Boolean" & vbCrLf
    s = s & "    Dim oldScreen As Boolean" & vbCrLf
    s = s & "    Dim oldCalc As Long" & vbCrLf
    s = s & "    Dim saved As Boolean" & vbCrLf
    s = s & "    d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    If Not d.AllowUpdate Then" & vbCrLf
    s = s & "        ShowStatus ""Update is read-only for this dataset. Admin can use Excel sheet view."", True" & vbCrLf
    s = s & "        GoTo SafeExit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If mSelectedRow < 1 Then" & vbCrLf
    s = s & "        ShowStatus ""Select a record in the grid first."", True" & vbCrLf
    s = s & "        GoTo SafeExit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not ValidateInputs(False) Then GoTo SafeExit" & vbCrLf
    s = s & "    oldEvents = Application.EnableEvents" & vbCrLf
    s = s & "    oldAlerts = Application.DisplayAlerts" & vbCrLf
    s = s & "    oldScreen = Application.ScreenUpdating" & vbCrLf
    s = s & "    oldCalc = Application.Calculation" & vbCrLf
    s = s & "    gAppWriting = True" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "    Application.DisplayAlerts = False" & vbCrLf
    s = s & "    Application.ScreenUpdating = False" & vbCrLf
    s = s & "    Application.Calculation = xlCalculationManual" & vbCrLf
    s = s & "    WriteInputsToRow mSelectedRow" & vbCrLf
    s = s & "    mSheet.Calculate" & vbCrLf
    s = s & "    saved = SaveWorkbookSafely()" & vbCrLf
    s = s & "    If saved Then" & vbCrLf
    s = s & "        ShowStatus ""Updated exact worksheet row "" & CStr(mSelectedRow) & "" and saved the workbook."", False" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ShowStatus ""Row was updated in memory but the workbook was not saved."", True" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    TrySetApplication ""Calculation"", oldCalc" & vbCrLf
    s = s & "    TrySetApplication ""EnableEvents"", oldEvents" & vbCrLf
    s = s & "    TrySetApplication ""DisplayAlerts"", oldAlerts" & vbCrLf
    s = s & "    TrySetApplication ""ScreenUpdating"", oldScreen" & vbCrLf
    s = s & "    gAppWriting = False" & vbCrLf
    s = s & "    If saved Then" & vbCrLf
    s = s & "        LoadRecordIntoForm mSelectedRow" & vbCrLf
    s = s & "        RefreshGrid" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""UpdateRecord""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub DeleteRecord()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim d As AppDataset" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    Dim lastRow As Long" & vbCrLf
    s = s & "    Dim answer As VbMsgBoxResult" & vbCrLf
    s = s & "    Dim oldEvents As Boolean" & vbCrLf
    s = s & "    Dim oldAlerts As Boolean" & vbCrLf
    s = s & "    Dim oldScreen As Boolean" & vbCrLf
    s = s & "    Dim oldCalc As Long" & vbCrLf
    s = s & "    Dim saved As Boolean" & vbCrLf
    s = s & "    d = mDatasets(mDatasetIndex)" & vbCrLf
    s = s & "    If gRole <> ROLE_ADMIN Then" & vbCrLf
    s = s & "        ShowStatus ""Delete is disabled in Agent mode. Switch to Admin mode first."", True" & vbCrLf
    s = s & "        GoTo SafeExit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not d.AllowDelete Then" & vbCrLf
    s = s & "        ShowStatus ""Delete is disabled for formula-backed layouts. Use Admin Excel sheet view for raw structural edits."", True" & vbCrLf
    s = s & "        GoTo SafeExit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    For i = 1 To mFieldCount" & vbCrLf
    s = s & "        If mFields(i).Kind = ""FORMULA"" Then" & vbCrLf
    s = s & "            ShowStatus ""Delete is blocked because this dataset contains live formulas."", True" & vbCrLf
    s = s & "            GoTo SafeExit" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    If mSelectedRow < 1 Then" & vbCrLf
    s = s & "        ShowStatus ""Select a record in the grid first."", True" & vbCrLf
    s = s & "        GoTo SafeExit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    answer = MsgBox(""Are you sure?"" & vbCrLf & _" & vbCrLf
    s = s & "                    ""This deletes only worksheet row "" & CStr(mSelectedRow) & _" & vbCrLf
    s = s & "                    "" from the selected dataset range."", _" & vbCrLf
    s = s & "                    vbYesNo + vbQuestion + vbDefaultButton2, ""Confirm Delete"")" & vbCrLf
    s = s & "    If answer <> vbYes Then" & vbCrLf
    s = s & "        ShowStatus ""Delete cancelled; no worksheet data was changed."", False" & vbCrLf
    s = s & "        GoTo SafeExit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lastRow = FindLastUsedRow(d)" & vbCrLf
    s = s & "    If lastRow < mSelectedRow Then" & vbCrLf
    s = s & "        ShowStatus ""The selected row is no longer present in the sheet."", True" & vbCrLf
    s = s & "        GoTo SafeExit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    oldEvents = Application.EnableEvents" & vbCrLf
    s = s & "    oldAlerts = Application.DisplayAlerts" & vbCrLf
    s = s & "    oldScreen = Application.ScreenUpdating" & vbCrLf
    s = s & "    oldCalc = Application.Calculation" & vbCrLf
    s = s & "    gAppWriting = True" & vbCrLf
    s = s & "    Application.EnableEvents = False" & vbCrLf
    s = s & "    Application.DisplayAlerts = False" & vbCrLf
    s = s & "    Application.ScreenUpdating = False" & vbCrLf
    s = s & "    Application.Calculation = xlCalculationManual" & vbCrLf
    s = s & "    If lastRow = mSelectedRow Then" & vbCrLf
    s = s & "        mSheet.Range(mSheet.Cells(mSelectedRow, d.FirstCol), _" & vbCrLf
    s = s & "                     mSheet.Cells(mSelectedRow, d.LastCol)).ClearContents" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        For c = d.FirstCol To d.LastCol" & vbCrLf
    s = s & "            mSheet.Range(mSheet.Cells(mSelectedRow, c), mSheet.Cells(lastRow, c)).Delete Shift:=xlUp" & vbCrLf
    s = s & "        Next c" & vbCrLf
    s = s & "        CopyRowFormat mSheet, d, d.FirstDataRow, lastRow" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mSheet.Calculate" & vbCrLf
    s = s & "    saved = SaveWorkbookSafely()" & vbCrLf
    s = s & "    If saved Then" & vbCrLf
    s = s & "        ShowStatus ""Deleted worksheet row "" & CStr(mSelectedRow) & "" and saved the workbook."", False" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ShowStatus ""Row was deleted in memory but the workbook was not saved."", True" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    TrySetApplication ""Calculation"", oldCalc" & vbCrLf
    s = s & "    TrySetApplication ""EnableEvents"", oldEvents" & vbCrLf
    s = s & "    TrySetApplication ""DisplayAlerts"", oldAlerts" & vbCrLf
    s = s & "    TrySetApplication ""ScreenUpdating"", oldScreen" & vbCrLf
    s = s & "    gAppWriting = False" & vbCrLf
    s = s & "    If saved Then" & vbCrLf
    s = s & "        ClearInputs" & vbCrLf
    s = s & "        RefreshGrid" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""DeleteRecord""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub WriteInputsToRow(ByVal row As Long)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim v As String" & vbCrLf
    s = s & "    Dim cell As Range" & vbCrLf
    s = s & "    For i = 1 To mFieldCount" & vbCrLf
    s = s & "        If mFields(i).Editable Then" & vbCrLf
    s = s & "            Set cell = mSheet.Cells(row, mFields(i).ColNumber)" & vbCrLf
    s = s & "            v = Trim$(GetInputValue(i))" & vbCrLf
    s = s & "            If Len(v) = 0 Then" & vbCrLf
    s = s & "                cell.ClearContents" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                Select Case mFields(i).Kind" & vbCrLf
    s = s & "                    Case ""DATE""" & vbCrLf
    s = s & "                        cell.Value = CDate(v)" & vbCrLf
    s = s & "                    Case ""NUMBER""" & vbCrLf
    s = s & "                        cell.Value = CDbl(v)" & vbCrLf
    s = s & "                    Case Else" & vbCrLf
    s = s & "                        cell.Value = v" & vbCrLf
    s = s & "                End Select" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""WriteInputsToRow""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ExtendFormulaColumns(ByVal d As AppDataset, ByVal row As Long)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim col As Long" & vbCrLf
    s = s & "    For i = 1 To mFieldCount" & vbCrLf
    s = s & "        If mFields(i).Kind = ""FORMULA"" Then" & vbCrLf
    s = s & "            col = mFields(i).ColNumber" & vbCrLf
    s = s & "            If mSheet.Cells(d.FirstDataRow, col).HasFormula Then" & vbCrLf
    s = s & "                mSheet.Cells(row, col).FormulaR1C1 = mSheet.Cells(d.FirstDataRow, col).FormulaR1C1" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""ExtendFormulaColumns""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FindNextBlankRow(ByVal d As AppDataset) As Long" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    For r = d.FirstDataRow To d.LastDataRow" & vbCrLf
    s = s & "        If Not IsDataRow(r, d) Then" & vbCrLf
    s = s & "            FindNextBlankRow = r" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    FindNextBlankRow = 0" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""FindNextBlankRow""" & vbCrLf
    s = s & "    FindNextBlankRow = 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FindLastUsedRow(ByVal d As AppDataset) As Long" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    For r = d.LastDataRow To d.FirstDataRow Step -1" & vbCrLf
    s = s & "        If IsDataRow(r, d) Then" & vbCrLf
    s = s & "            FindLastUsedRow = r" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    FindLastUsedRow = 0" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""FindLastUsedRow""" & vbCrLf
    s = s & "    FindLastUsedRow = 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub CopyRowFormat(ByVal ws As Worksheet, ByVal d As AppDataset, ByVal sourceRow As Long, ByVal targetRow As Long)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    Dim src As Range" & vbCrLf
    s = s & "    Dim dst As Range" & vbCrLf
    s = s & "    Set src = ws.Range(ws.Cells(sourceRow, d.FirstCol), ws.Cells(sourceRow, d.LastCol))" & vbCrLf
    s = s & "    Set dst = ws.Range(ws.Cells(targetRow, d.FirstCol), ws.Cells(targetRow, d.LastCol))" & vbCrLf
    s = s & "    src.Copy" & vbCrLf
    s = s & "    PasteSpecialSafe dst, xlPasteFormats" & vbCrLf
    s = s & "    PasteSpecialSafe dst, xlPasteValidation" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "    TrySetApplication ""CutCopyMode"", False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""CopyRowFormat""" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub PasteSpecialSafe(ByVal rng As Range, ByVal pasteType As Long)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    rng.PasteSpecial Paste:=pasteType, Operation:=xlNone, SkipBlanks:=False, Transpose:=False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    Resume SafeExit" & vbCrLf
    s = s & "SafeExit:" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function SaveWorkbookSafely() As Boolean" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    ThisWorkbook.Save" & vbCrLf
    s = s & "    SaveWorkbookSafely = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    HandleError ""SaveWorkbookSafely""" & vbCrLf
    s = s & "    SaveWorkbookSafely = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    FormSource = s
    Exit Function
ErrorHandler:
    BuilderHandleError "FormSource"
    FormSource = ""
End Function

Private Function WorkbookSource() As String
    On Error GoTo ErrorHandler
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Workbook_Open()" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    modRaadApp.StartApplication" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    modRaadApp.HandleGlobalError ""Workbook_Open""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    modRaadApp.RefreshOpenForm Sh, Target" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    modRaadApp.HandleGlobalError ""Workbook_SheetChange""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf
    s = s & "    On Error GoTo ErrorHandler" & vbCrLf
    s = s & "    modRaadApp.CleanupBeforeClose" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrorHandler:" & vbCrLf
    s = s & "    modRaadApp.HandleGlobalError ""Workbook_BeforeClose""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    WorkbookSource = s
    Exit Function
ErrorHandler:
    BuilderHandleError "WorkbookSource"
    WorkbookSource = ""
End Function

