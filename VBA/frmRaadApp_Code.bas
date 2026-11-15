Option Explicit

'===============================================================================
' frmRaadApp - modern, dynamic UserForm code.
' Static shell controls are listed in UI_Layout_and_Controls.md.
' Field controls are created at run time from the exact worksheet header cells.
'===============================================================================

Private mDatasets() As AppDataset
Private mDatasetCount As Long
Private mDatasetIndex As Long
Private mFields() As AppField
Private mFieldCount As Long
Private mSheet As Worksheet
Private mSelectedRow As Long
Private mRowCount As Long
Private mRowMap() As Long
Private mInitializing As Boolean
Private mSuppressEvents As Boolean

Private mBg As Long
Private mHeader As Long
Private mHeaderDark As Long
Private mCard As Long
Private mWhite As Long
Private mText As Long
Private mMuted As Long
Private mBorder As Long
Private mReadOnly As Long
Private mPrimary As Long
Private mPrimaryHover As Long
Private mOrange As Long
Private mOrangeHover As Long
Private mDanger As Long
Private mDangerHover As Long
Private mGray As Long
Private mGrayHover As Long
Private mSuccess As Long

'-------------------------------------------------------------------------------
' UserForm lifecycle
'-------------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    InitializeApp
    Exit Sub
ErrorHandler:
    HandleError "UserForm_Initialize"
End Sub

Private Sub UserForm_Activate()
    On Error GoTo ErrorHandler
    Me.lblClock.Caption = Format(Now, "yyyy-mm-dd hh:nn")
    ApplyRoleVisual
    Exit Sub
ErrorHandler:
    HandleError "UserForm_Activate"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler
    gAppWriting = False
    mSuppressEvents = False
    Exit Sub
ErrorHandler:
    HandleError "UserForm_QueryClose"
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    Exit Sub
ErrorHandler:
    HandleError "UserForm_MouseMove"
End Sub

'-------------------------------------------------------------------------------
' Static control events
'-------------------------------------------------------------------------------

Private Sub cboDataset_Change()
    On Error GoTo ErrorHandler
    If mInitializing Or mSuppressEvents Then Exit Sub
    ApplyDataset
    Exit Sub
ErrorHandler:
    HandleError "cboDataset_Change"
End Sub

Private Sub txtSearch_Change()
    On Error GoTo ErrorHandler
    If mInitializing Or mSuppressEvents Then Exit Sub
    RefreshGrid
    Exit Sub
ErrorHandler:
    HandleError "txtSearch_Change"
End Sub

Private Sub cboFilterField_Change()
    On Error GoTo ErrorHandler
    If mInitializing Or mSuppressEvents Then Exit Sub
    PopulateFilterValues
    RefreshGrid
    Exit Sub
ErrorHandler:
    HandleError "cboFilterField_Change"
End Sub

Private Sub cboFilterValue_Change()
    On Error GoTo ErrorHandler
    If mInitializing Or mSuppressEvents Then Exit Sub
    RefreshGrid
    Exit Sub
ErrorHandler:
    HandleError "cboFilterValue_Change"
End Sub

Private Sub lstData_Click()
    On Error GoTo ErrorHandler
    Dim idx As Long
    If Me.lstData.ListIndex < 0 Then Exit Sub
    idx = Me.lstData.ListIndex + 1
    If idx > mRowCount Then Exit Sub
    LoadRecordIntoForm mRowMap(idx)
    Exit Sub
ErrorHandler:
    HandleError "lstData_Click"
End Sub

Private Sub lstData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    lstData_Click
    Exit Sub
ErrorHandler:
    HandleError "lstData_DblClick"
End Sub

Private Sub btnAdd_Click()
    On Error GoTo ErrorHandler
    AddRecord
    Exit Sub
ErrorHandler:
    HandleError "btnAdd_Click"
End Sub

Private Sub btnUpdate_Click()
    On Error GoTo ErrorHandler
    UpdateRecord
    Exit Sub
ErrorHandler:
    HandleError "btnUpdate_Click"
End Sub

Private Sub btnDelete_Click()
    On Error GoTo ErrorHandler
    DeleteRecord
    Exit Sub
ErrorHandler:
    HandleError "btnDelete_Click"
End Sub

Private Sub btnClear_Click()
    On Error GoTo ErrorHandler
    ClearInputs
    ShowStatus "Form cleared. Select a grid row or enter a new record.", False
    Exit Sub
ErrorHandler:
    HandleError "btnClear_Click"
End Sub

Private Sub btnClearFilter_Click()
    On Error GoTo ErrorHandler
    Dim oldFlag As Boolean
    oldFlag = mSuppressEvents
    mSuppressEvents = True
    Me.txtSearch.Text = ""
    Me.cboFilterField.ListIndex = 0
    PopulateFilterValues
SafeExit:
    mSuppressEvents = oldFlag
    RefreshGrid
    ShowStatus "Search and filters cleared.", False
    Exit Sub
ErrorHandler:
    HandleError "btnClearFilter_Click"
    Resume SafeExit
End Sub

Private Sub btnRefresh_Click()
    On Error GoTo ErrorHandler
    If Not mSheet Is Nothing Then mSheet.Calculate
    RefreshGrid
    If mSelectedRow > 0 Then LoadRecordIntoForm mSelectedRow
    ShowStatus "Refreshed directly from the real worksheet.", False
    Exit Sub
ErrorHandler:
    HandleError "btnRefresh_Click"
End Sub

Private Sub btnSheetView_Click()
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    If mSheet Is Nothing Then Exit Sub
    d = mDatasets(mDatasetIndex)
    TrySetApplication "Visible", True
    mSheet.Activate
    Application.Goto mSheet.Range(mSheet.Cells(d.HeaderRow, d.FirstCol), _
                                  mSheet.Cells(d.LastDataRow, d.LastCol)), True
    Me.Hide
    AppActivate Application.Caption
    Exit Sub
ErrorHandler:
    HandleError "btnSheetView_Click"
End Sub

Private Sub btnSwitchMode_Click()
    On Error GoTo ErrorHandler
    Dim wantsAdmin As Boolean
    wantsAdmin = (StrComp(CStr(Me.cboRole.Value), "Admin", vbTextCompare) = 0)
    If wantsAdmin Then
        If modRaadApp.ValidateAdminLogin() Then
            gRole = ROLE_ADMIN
            modRaadApp.UnlockAllSheets
            ShowStatus "Admin mode enabled. Sheets and workbook structure are unlocked.", False
        Else
            gRole = ROLE_AGENT
            modRaadApp.LockForAgent
            ShowStatus "Admin login failed. Agent mode remains active and locked.", True
        End If
    Else
        gRole = ROLE_AGENT
        modRaadApp.LockForAgent
        ShowStatus "Agent mode active. Structure and formula cells are protected.", False
    End If
    ApplyRoleVisual
    Exit Sub
ErrorHandler:
    HandleError "btnSwitchMode_Click"
End Sub

Private Sub btnUnlockAll_Click()
    On Error GoTo ErrorHandler
    If gRole <> ROLE_ADMIN Then
        ShowStatus "Only Admin mode can unlock all sheets.", True
        Exit Sub
    End If
    modRaadApp.UnlockAllSheets
    ShowStatus "All sheets and workbook structure are unlocked for Admin editing.", False
    Exit Sub
ErrorHandler:
    HandleError "btnUnlockAll_Click"
End Sub

Private Sub btnLockForAgent_Click()
    On Error GoTo ErrorHandler
    If gRole <> ROLE_ADMIN Then
        ShowStatus "Only Admin mode can change the lock state.", True
        Exit Sub
    End If
    gRole = ROLE_AGENT
    modRaadApp.LockForAgent
    ApplyRoleVisual
    ShowStatus "Agent mode restored. Structural editing is locked again.", False
    Exit Sub
ErrorHandler:
    HandleError "btnLockForAgent_Click"
End Sub

Private Sub btnClose_Click()
    On Error GoTo ErrorHandler
    Me.Hide
    Exit Sub
ErrorHandler:
    HandleError "btnClose_Click"
End Sub

'-------------------------------------------------------------------------------
' Hover events
'-------------------------------------------------------------------------------

Private Sub btnAdd_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnAdd
    Exit Sub
ErrorHandler:
    HandleError "btnAdd_MouseMove"
End Sub

Private Sub btnUpdate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnUpdate
    Exit Sub
ErrorHandler:
    HandleError "btnUpdate_MouseMove"
End Sub

Private Sub btnDelete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnDelete
    Exit Sub
ErrorHandler:
    HandleError "btnDelete_MouseMove"
End Sub

Private Sub btnClear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnClear
    Exit Sub
ErrorHandler:
    HandleError "btnClear_MouseMove"
End Sub

Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnClose
    Exit Sub
ErrorHandler:
    HandleError "btnClose_MouseMove"
End Sub

Private Sub btnClearFilter_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnClearFilter
    Exit Sub
ErrorHandler:
    HandleError "btnClearFilter_MouseMove"
End Sub

Private Sub btnRefresh_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnRefresh
    Exit Sub
ErrorHandler:
    HandleError "btnRefresh_MouseMove"
End Sub

Private Sub btnSheetView_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnSheetView
    Exit Sub
ErrorHandler:
    HandleError "btnSheetView_MouseMove"
End Sub

Private Sub btnSwitchMode_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnSwitchMode
    Exit Sub
ErrorHandler:
    HandleError "btnSwitchMode_MouseMove"
End Sub

Private Sub btnUnlockAll_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnUnlockAll
    Exit Sub
ErrorHandler:
    HandleError "btnUnlockAll_MouseMove"
End Sub

Private Sub btnLockForAgent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler
    ResetButtonColors
    HighlightButton Me.btnLockForAgent
    Exit Sub
ErrorHandler:
    HandleError "btnLockForAgent_MouseMove"
End Sub

'-------------------------------------------------------------------------------
' Public entry point called by ThisWorkbook
'-------------------------------------------------------------------------------

Public Sub ExternalSheetRefresh(ByVal sheetName As String, ByVal Target As Range)
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    Dim watchRange As Range
    If mInitializing Or gAppWriting Then Exit Sub
    If mDatasetIndex < 1 Or mSheet Is Nothing Then Exit Sub
    d = mDatasets(mDatasetIndex)
    If StrComp(d.SheetName, sheetName, vbTextCompare) <> 0 Then Exit Sub
    Set watchRange = mSheet.Range(mSheet.Cells(d.HeaderRow, d.FirstCol), _
                                  mSheet.Cells(d.LastDataRow, d.LastCol))
    If Intersect(Target, watchRange) Is Nothing Then Exit Sub
    RefreshGrid
    If mSelectedRow > 0 Then LoadRecordIntoForm mSelectedRow
    ShowStatus "Worksheet edit detected. Grid and inputs refreshed.", False
    Exit Sub
ErrorHandler:
    HandleError "ExternalSheetRefresh"
End Sub

'-------------------------------------------------------------------------------
' Initialization and visual design
'-------------------------------------------------------------------------------

Private Sub InitializeApp()
    On Error GoTo ErrorHandler
    mInitializing = True
    modRaadApp.EnsureDatasets
    InitPalette
    ApplyLayout
    StyleAllControls
    ApplyCaptions
    ApplyPalette
    PopulateRoles
    PopulateDatasets
    mDatasetIndex = modRaadApp.FindMainDatasetIndex()
    If mDatasetIndex < 1 Then mDatasetIndex = 1
    Me.cboDataset.ListIndex = mDatasetIndex - 1
    gRole = ROLE_AGENT
    Me.cboRole.ListIndex = 0
    ApplyDataset
    ShowStatus "Ready. Agent mode is active and every field is connected to the real sheet.", False
SafeExit:
    mInitializing = False
    Exit Sub
ErrorHandler:
    HandleError "InitializeApp"
    Resume SafeExit
End Sub

Private Sub InitPalette()
    On Error GoTo ErrorHandler
    mBg = RGB(244, 247, 251)
    mHeader = RGB(16, 42, 67)
    mHeaderDark = RGB(12, 31, 50)
    mCard = RGB(255, 255, 255)
    mWhite = RGB(255, 255, 255)
    mText = RGB(15, 23, 42)
    mMuted = RGB(100, 116, 139)
    mBorder = RGB(203, 213, 225)
    mReadOnly = RGB(241, 245, 249)
    mPrimary = RGB(37, 99, 235)
    mPrimaryHover = RGB(29, 78, 216)
    mOrange = RGB(245, 158, 11)
    mOrangeHover = RGB(217, 119, 6)
    mDanger = RGB(220, 38, 38)
    mDangerHover = RGB(185, 28, 28)
    mGray = RGB(100, 116, 139)
    mGrayHover = RGB(71, 85, 105)
    mSuccess = RGB(22, 163, 74)
    Exit Sub
ErrorHandler:
    HandleError "InitPalette"
End Sub

Private Sub ApplyLayout()
    On Error GoTo ErrorHandler
    Me.Width = 1200
    Me.Height = 780
    Me.StartUpPosition = 2
    PlaceControl Me.fraHeader, 0, 0, 1200, 64
    PlaceControl Me.lblTitle, 24, 8, 620, 26
    PlaceControl Me.lblSubTitle, 24, 34, 760, 20
    PlaceControl Me.lblModeBadge, 970, 16, 90, 28
    PlaceControl Me.lblClock, 1070, 16, 110, 28

    PlaceControl Me.fraSidebar, 970, 76, 214, 622
    PlaceControl Me.lblNav, 986, 92, 180, 22
    PlaceControl Me.cboRole, 986, 120, 180, 26
    PlaceControl Me.btnSwitchMode, 986, 152, 180, 30
    PlaceControl Me.cboDataset, 986, 194, 180, 26
    PlaceControl Me.lblDatasetInfo, 986, 226, 180, 62
    PlaceControl Me.btnRefresh, 986, 298, 180, 30
    PlaceControl Me.btnSheetView, 986, 336, 180, 30
    PlaceControl Me.btnUnlockAll, 986, 380, 180, 30
    PlaceControl Me.btnLockForAgent, 986, 418, 180, 30
    PlaceControl Me.lblRecordCount, 986, 466, 180, 22
    PlaceControl Me.lblSelectedRow, 986, 494, 180, 22

    PlaceControl Me.fraToolbar, 16, 76, 938, 68
    PlaceControl Me.lblSearch, 28, 88, 240, 18
    PlaceControl Me.txtSearch, 28, 108, 250, 26
    PlaceControl Me.lblFilterField, 294, 88, 180, 18
    PlaceControl Me.cboFilterField, 294, 108, 180, 26
    PlaceControl Me.lblFilterValue, 490, 88, 180, 18
    PlaceControl Me.cboFilterValue, 490, 108, 180, 26
    PlaceControl Me.btnClearFilter, 686, 108, 120, 26

    PlaceControl Me.lstHeaders, 16, 152, 938, 24
    PlaceControl Me.lstData, 16, 176, 938, 282
    PlaceControl Me.fraEditor, 16, 466, 938, 210
    PlaceControl Me.lblEditorTitle, 28, 474, 400, 20

    PlaceControl Me.fraButtons, 16, 684, 938, 48
    PlaceControl Me.btnAdd, 28, 694, 140, 30
    PlaceControl Me.btnUpdate, 180, 694, 140, 30
    PlaceControl Me.btnDelete, 332, 694, 140, 30
    PlaceControl Me.btnClear, 484, 694, 140, 30
    PlaceControl Me.btnClose, 800, 694, 140, 30
    PlaceControl Me.lblStatus, 16, 738, 1168, 26
    Exit Sub
ErrorHandler:
    HandleError "ApplyLayout"
End Sub

Private Sub ApplyCaptions()
    On Error GoTo ErrorHandler
    Me.Caption = "Raad | Live Inventory & Variance Console"
    SetCaptionSafe Me.fraHeader, ""
    SetCaptionSafe Me.fraSidebar, ""
    SetCaptionSafe Me.fraToolbar, ""
    SetCaptionSafe Me.fraEditor, ""
    SetCaptionSafe Me.fraButtons, ""
    SetCaptionSafe Me.lblTitle, "Raad Inventory & Daily Variance"
    SetCaptionSafe Me.lblSubTitle, "Live two-way connection to your real worksheets, fields, formulas and dropdowns"
    SetCaptionSafe Me.lblModeBadge, "AGENT"
    SetCaptionSafe Me.lblClock, Format(Now, "yyyy-mm-dd hh:nn")
    SetCaptionSafe Me.lblNav, "Workspace / Role"
    SetCaptionSafe Me.lblDatasetInfo, "Select a real sheet dataset"
    SetCaptionSafe Me.lblRecordCount, "Records: 0"
    SetCaptionSafe Me.lblSelectedRow, "Selected row: none"
    SetCaptionSafe Me.lblSearch, "Live search"
    SetCaptionSafe Me.lblFilterField, "Filter by real column"
    SetCaptionSafe Me.lblFilterValue, "Filter value"
    SetCaptionSafe Me.lblEditorTitle, "Record fields - exact worksheet headers"
    SetCaptionSafe Me.lblStatus, "Ready."
    SetCaptionSafe Me.btnAdd, "Add"
    SetCaptionSafe Me.btnUpdate, "Update"
    SetCaptionSafe Me.btnDelete, "Delete"
    SetCaptionSafe Me.btnClear, "Clear"
    SetCaptionSafe Me.btnClose, "Close"
    SetCaptionSafe Me.btnClearFilter, "Clear Filter"
    SetCaptionSafe Me.btnRefresh, "Refresh"
    SetCaptionSafe Me.btnSheetView, "Switch to Excel Sheet"
    SetCaptionSafe Me.btnSwitchMode, "Switch Mode"
    SetCaptionSafe Me.btnUnlockAll, "Unlock Sheets"
    SetCaptionSafe Me.btnLockForAgent, "Lock For Agent"
    Exit Sub
ErrorHandler:
    HandleError "ApplyCaptions"
End Sub

Private Sub ApplyPalette()
    On Error GoTo ErrorHandler
    Me.BackColor = mBg
    SetBack Me.fraHeader, mHeader
    SetBack Me.fraSidebar, mHeaderDark
    SetBack Me.fraToolbar, mCard
    SetBack Me.fraEditor, mCard
    SetBack Me.fraButtons, mCard
    SetBorderColor Me.fraToolbar, mBorder
    SetBorderColor Me.fraEditor, mBorder
    SetBorderColor Me.fraButtons, mBorder

    SetBack Me.lblTitle, mHeader
    SetFore Me.lblTitle, mWhite
    SetBack Me.lblSubTitle, mHeader
    SetFore Me.lblSubTitle, RGB(203, 213, 225)
    SetBack Me.lblModeBadge, mPrimary
    SetFore Me.lblModeBadge, mWhite
    SetBack Me.lblClock, mHeader
    SetFore Me.lblClock, mWhite

    SetBack Me.lblNav, mHeaderDark
    SetFore Me.lblNav, mWhite
    SetBack Me.lblDatasetInfo, mHeaderDark
    SetFore Me.lblDatasetInfo, mWhite
    SetBack Me.lblRecordCount, mHeaderDark
    SetFore Me.lblRecordCount, mWhite
    SetBack Me.lblSelectedRow, mHeaderDark
    SetFore Me.lblSelectedRow, mWhite

    SetBack Me.lblSearch, mCard
    SetFore Me.lblSearch, mText
    SetBack Me.lblFilterField, mCard
    SetFore Me.lblFilterField, mText
    SetBack Me.lblFilterValue, mCard
    SetFore Me.lblFilterValue, mText
    SetBack Me.lblEditorTitle, mCard
    SetFore Me.lblEditorTitle, mText

    SetBack Me.txtSearch, mWhite
    SetFore Me.txtSearch, mText
    SetBorderColor Me.txtSearch, mBorder
    SetBack Me.cboRole, mWhite
    SetFore Me.cboRole, mText
    SetBorderColor Me.cboRole, mBorder
    SetBack Me.cboDataset, mWhite
    SetFore Me.cboDataset, mText
    SetBorderColor Me.cboDataset, mBorder
    SetBack Me.cboFilterField, mWhite
    SetFore Me.cboFilterField, mText
    SetBorderColor Me.cboFilterField, mBorder
    SetBack Me.cboFilterValue, mWhite
    SetFore Me.cboFilterValue, mText
    SetBorderColor Me.cboFilterValue, mBorder

    SetBack Me.lstHeaders, mHeader
    SetFore Me.lstHeaders, mWhite
    SetBorderColor Me.lstHeaders, mBorder
    SetBack Me.lstData, mWhite
    SetFore Me.lstData, mText
    SetBorderColor Me.lstData, mBorder

    SetBack Me.lblStatus, mCard
    SetFore Me.lblStatus, mSuccess
    SetBorderColor Me.lblStatus, mBorder

    ResetButtonColors
    Exit Sub
ErrorHandler:
    HandleError "ApplyPalette"
End Sub

Private Sub StyleAllControls()
    On Error GoTo ErrorHandler
    Dim ctl As MSForms.Control
    Dim child As Object
    Dim container As Object
    For Each ctl In Me.Controls
        StyleControl ctl
        If TypeName(ctl) = "Frame" Then
            Set container = ctl
            For Each child In container.Controls
                StyleControl child
            Next child
        End If
    Next ctl
    ApplyTitleFonts
    Exit Sub
ErrorHandler:
    HandleError "StyleAllControls"
End Sub

Private Sub StyleControl(ByVal ctl As Object)
    On Error GoTo ErrorHandler
    TrySetFont ctl, 9, False
    TrySet ctl, "SpecialEffect", 0
    TrySet ctl, "BorderStyle", 1
    TrySet ctl, "BorderColor", mBorder
    TrySet ctl, "BackStyle", 1
    Exit Sub
ErrorHandler:
    HandleError "StyleControl"
End Sub

Private Sub ApplyTitleFonts()
    On Error GoTo ErrorHandler
    TrySetFont Me.lblTitle, 16, True
    TrySetFont Me.lblSubTitle, 9, False
    TrySetFont Me.lblModeBadge, 10, True
    TrySetFont Me.lblClock, 9, False
    TrySetFont Me.lblNav, 10, True
    TrySetFont Me.lblEditorTitle, 10, True
    TrySetFont Me.lblStatus, 9, True
    Exit Sub
ErrorHandler:
    HandleError "ApplyTitleFonts"
End Sub

Private Sub ResetButtonColors()
    On Error GoTo ErrorHandler
    SetBack Me.btnAdd, mPrimary
    SetBack Me.btnUpdate, mOrange
    SetBack Me.btnDelete, mDanger
    SetBack Me.btnClear, mGray
    SetBack Me.btnClose, mGray
    SetBack Me.btnClearFilter, mGray
    SetBack Me.btnRefresh, mPrimary
    SetBack Me.btnSheetView, mPrimary
    SetBack Me.btnSwitchMode, mPrimary
    SetBack Me.btnUnlockAll, mPrimary
    SetBack Me.btnLockForAgent, mGray
    SetFore Me.btnAdd, mWhite
    SetFore Me.btnUpdate, mWhite
    SetFore Me.btnDelete, mWhite
    SetFore Me.btnClear, mWhite
    SetFore Me.btnClose, mWhite
    SetFore Me.btnClearFilter, mWhite
    SetFore Me.btnRefresh, mWhite
    SetFore Me.btnSheetView, mWhite
    SetFore Me.btnSwitchMode, mWhite
    SetFore Me.btnUnlockAll, mWhite
    SetFore Me.btnLockForAgent, mWhite
    Exit Sub
ErrorHandler:
    HandleError "ResetButtonColors"
End Sub

Private Sub HighlightButton(ByVal btn As Object)
    On Error GoTo ErrorHandler
    Dim c As Long
    Select Case btn.Name
        Case "btnAdd", "btnRefresh", "btnSheetView", "btnSwitchMode", "btnUnlockAll"
            c = mPrimaryHover
        Case "btnUpdate"
            c = mOrangeHover
        Case "btnDelete"
            c = mDangerHover
        Case Else
            c = mGrayHover
    End Select
    SetBack btn, c
    Exit Sub
ErrorHandler:
    HandleError "HighlightButton"
End Sub

Private Sub ShowStatus(ByVal message As String, Optional ByVal isError As Boolean = False)
    On Error GoTo ErrorHandler
    Me.lblStatus.Caption = message
    If isError Then
        SetFore Me.lblStatus, mDanger
    Else
        SetFore Me.lblStatus, mSuccess
    End If
    SetBack Me.lblStatus, mCard
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub HandleError(ByVal procName As String)
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
    ShowStatus msg, True
    MsgBox msg, vbExclamation, "Raad App"
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub TrySet(ByVal ctl As Object, ByVal propName As String, ByVal value As Variant)
    On Error GoTo ErrorHandler
    CallByName ctl, propName, VbLet, value
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub TrySetFont(ByVal ctl As Object, ByVal size As Single, Optional ByVal bold As Boolean = False)
    On Error GoTo ErrorHandler
    Dim f As Object
    Set f = CallByName(ctl, "Font", VbGet)
    CallByName f, "Name", VbLet, "Segoe UI"
    CallByName f, "Size", VbLet, size
    CallByName f, "Bold", VbLet, bold
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub TrySetApplication(ByVal propName As String, ByVal value As Variant)
    On Error GoTo ErrorHandler
    CallByName Application, propName, VbLet, value
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Sub SetBack(ByVal ctl As Object, ByVal colorValue As Long)
    On Error GoTo ErrorHandler
    TrySet ctl, "BackColor", colorValue
    Exit Sub
ErrorHandler:
    HandleError "SetBack"
End Sub

Private Sub SetFore(ByVal ctl As Object, ByVal colorValue As Long)
    On Error GoTo ErrorHandler
    TrySet ctl, "ForeColor", colorValue
    Exit Sub
ErrorHandler:
    HandleError "SetFore"
End Sub

Private Sub SetBorderColor(ByVal ctl As Object, ByVal colorValue As Long)
    On Error GoTo ErrorHandler
    TrySet ctl, "BorderColor", colorValue
    Exit Sub
ErrorHandler:
    HandleError "SetBorderColor"
End Sub

Private Sub SetCaptionSafe(ByVal ctl As Object, ByVal caption As String)
    On Error GoTo ErrorHandler
    TrySet ctl, "Caption", caption
    TrySet ctl, "AutoSize", False
    TrySet ctl, "WordWrap", True
    Exit Sub
ErrorHandler:
    HandleError "SetCaptionSafe"
End Sub

Private Sub PlaceControl(ByVal ctl As Object, ByVal leftValue As Single, ByVal topValue As Single, _
                         ByVal widthValue As Single, ByVal heightValue As Single)
    On Error GoTo ErrorHandler
    TrySet ctl, "Left", leftValue
    TrySet ctl, "Top", topValue
    TrySet ctl, "Width", widthValue
    TrySet ctl, "Height", heightValue
    Exit Sub
ErrorHandler:
    HandleError "PlaceControl"
End Sub

'-------------------------------------------------------------------------------
' Dataset, role and dynamic field construction
'-------------------------------------------------------------------------------

Private Sub PopulateRoles()
    On Error GoTo ErrorHandler
    Dim oldFlag As Boolean
    oldFlag = mSuppressEvents
    mSuppressEvents = True
    Me.cboRole.Clear
    Me.cboRole.AddItem "Agent"
    Me.cboRole.AddItem "Admin"
    Me.cboRole.ListIndex = 0
    TrySet Me.cboRole, "Style", 2
SafeExit:
    mSuppressEvents = oldFlag
    Exit Sub
ErrorHandler:
    HandleError "PopulateRoles"
    Resume SafeExit
End Sub

Private Sub PopulateDatasets()
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim oldFlag As Boolean
    oldFlag = mSuppressEvents
    mSuppressEvents = True
    ReDim mDatasets(1 To gDatasetCount)
    For i = 1 To gDatasetCount
        mDatasets(i) = gDatasets(i)
    Next i
    mDatasetCount = gDatasetCount
    Me.cboDataset.Clear
    For i = 1 To mDatasetCount
        Me.cboDataset.AddItem mDatasets(i).DisplayName
    Next i
    TrySet Me.cboDataset, "Style", 2
SafeExit:
    mSuppressEvents = oldFlag
    Exit Sub
ErrorHandler:
    HandleError "PopulateDatasets"
    Resume SafeExit
End Sub

Private Sub ApplyDataset()
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    If Me.cboDataset.ListIndex < 0 Then Exit Sub
    mDatasetIndex = Me.cboDataset.ListIndex + 1
    If mDatasetIndex < 1 Or mDatasetIndex > mDatasetCount Then Exit Sub
    d = mDatasets(mDatasetIndex)
    Set mSheet = ThisWorkbook.Worksheets(d.SheetName)
    BuildFieldMetadata
    BuildFieldControls
    PopulateFilterFields
    PopulateFilterValues
    ClearInputs
    mSelectedRow = 0
    Me.lblSelectedRow.Caption = "Selected row: none"
    Me.lblDatasetInfo.Caption = d.DisplayName & vbCrLf & d.Description
    RefreshGrid
    ApplyRoleVisual
    StyleAllControls
    ShowStatus "Dataset loaded: " & d.DisplayName & ". Exact headers were read from the sheet.", False
    Exit Sub
ErrorHandler:
    HandleError "ApplyDataset"
End Sub

Private Sub ApplyRoleVisual()
    On Error GoTo ErrorHandler
    Dim isAdmin As Boolean
    Dim d As AppDataset
    isAdmin = (gRole = ROLE_ADMIN)
    If mDatasetIndex > 0 Then d = mDatasets(mDatasetIndex)
    Me.btnAdd.Enabled = d.AllowAdd
    Me.btnUpdate.Enabled = d.AllowUpdate
    Me.btnDelete.Enabled = isAdmin And d.AllowDelete
    Me.btnUnlockAll.Enabled = isAdmin
    Me.btnLockForAgent.Enabled = isAdmin
    Me.btnSheetView.Enabled = True
    Me.btnRefresh.Enabled = True
    Me.btnSwitchMode.Enabled = True
    If isAdmin Then
        Me.lblModeBadge.Caption = "ADMIN"
        SetBack Me.lblModeBadge, mDanger
        Me.cboRole.ListIndex = 1
    Else
        Me.lblModeBadge.Caption = "AGENT"
        SetBack Me.lblModeBadge, mPrimary
        Me.cboRole.ListIndex = 0
    End If
    SetFore Me.lblModeBadge, mWhite
    Exit Sub
ErrorHandler:
    HandleError "ApplyRoleVisual"
End Sub

Private Sub BuildFieldMetadata()
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    Dim i As Long
    Dim colCount As Long
    Dim vlist As String
    d = mDatasets(mDatasetIndex)
    colCount = d.LastCol - d.FirstCol + 1
    ReDim mFields(1 To colCount)
    mFieldCount = colCount
    For i = 1 To colCount
        mFields(i).Index = i
        mFields(i).ColNumber = d.FirstCol + i - 1
        mFields(i).ColLetter = ColumnLetter(mFields(i).ColNumber)
        mFields(i).Header = CStr(mSheet.Cells(d.HeaderRow, mFields(i).ColNumber).Value)
        If Len(Trim$(mFields(i).Header)) = 0 Then
            mFields(i).LabelText = "[" & mFields(i).ColLetter & "] (no header)"
        Else
            mFields(i).LabelText = mFields(i).Header
        End If
        vlist = GetValidationList(d, mFields(i).ColNumber)
        If Len(vlist) > 0 Then
            mFields(i).Kind = "LIST"
            mFields(i).Editable = True
            mFields(i).InputName = "cboF" & CStr(i)
            mFields(i).ValidationList = vlist
        ElseIf IsFormulaColumn(d, mFields(i).ColNumber) Then
            mFields(i).Kind = "FORMULA"
            mFields(i).Editable = False
            mFields(i).InputName = "txtF" & CStr(i)
            mFields(i).ValidationList = ""
        ElseIf IsDateColumn(d, mFields(i).ColNumber) Then
            mFields(i).Kind = "DATE"
            mFields(i).Editable = True
            mFields(i).InputName = "txtF" & CStr(i)
            mFields(i).ValidationList = ""
        ElseIf IsNumberColumn(d, mFields(i).ColNumber) Then
            mFields(i).Kind = "NUMBER"
            mFields(i).Editable = True
            mFields(i).InputName = "txtF" & CStr(i)
            mFields(i).ValidationList = ""
        Else
            mFields(i).Kind = "TEXT"
            mFields(i).Editable = True
            mFields(i).InputName = "txtF" & CStr(i)
            mFields(i).ValidationList = ""
        End If
    Next i
    Exit Sub
ErrorHandler:
    HandleError "BuildFieldMetadata"
End Sub

Private Sub BuildFieldControls()
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim k As Long
    Dim nm As String
    Dim nCols As Long
    Dim rowPos As Long
    Dim colPos As Long
    Dim rows As Long
    Dim leftPos As Single
    Dim topPos As Single
    Dim lbl As Object
    Dim inp As Object
    Dim items As Variant
    Dim itemText As String
    For i = Me.fraEditor.Controls.Count - 1 To 0 Step -1
        nm = Me.fraEditor.Controls(i).Name
        If Left$(nm, 4) = "lblF" Or Left$(nm, 4) = "txtF" Or Left$(nm, 4) = "cboF" Then
            Me.fraEditor.Controls.Remove nm
        End If
    Next i
    nCols = 3
    rows = Int((mFieldCount + nCols - 1) / nCols)
    TrySet Me.fraEditor, "ScrollBars", 2
    TrySet Me.fraEditor, "ScrollHeight", rows * 62 + 60
    TrySet Me.fraEditor, "ScrollTop", 0
    TrySet Me.fraEditor, "KeepScrollBarsVisible", 2
    For i = 1 To mFieldCount
        colPos = (i - 1) Mod nCols
        rowPos = Int((i - 1) / nCols)
        leftPos = 12 + colPos * 302
        topPos = 34 + rowPos * 62
        Set lbl = Me.fraEditor.Controls.Add("Forms.Label.1", "lblF" & CStr(i), True)
        PlaceControl lbl, leftPos, topPos, 280, 18
        lbl.Caption = mFields(i).LabelText
        SetBack lbl, mCard
        SetFore lbl, mText
        TrySetFont lbl, 9, True
        TrySet lbl, "TextAlign", 3
        TrySet lbl, "RightToLeft", True
        TrySet lbl, "SpecialEffect", 0
        TrySet lbl, "BorderStyle", 1
        If mFields(i).Kind = "LIST" Then
            Set inp = Me.fraEditor.Controls.Add("Forms.ComboBox.1", mFields(i).InputName, True)
            TrySet inp, "Style", 0
            inp.Clear
            items = Split(mFields(i).ValidationList, "|")
            For k = LBound(items) To UBound(items)
                itemText = Trim$(CStr(items(k)))
                If Len(itemText) > 0 Then inp.AddItem itemText
            Next k
        Else
            Set inp = Me.fraEditor.Controls.Add("Forms.TextBox.1", mFields(i).InputName, True)
        End If
        PlaceControl inp, leftPos, topPos + 20, 280, 26
        SetBack inp, mWhite
        SetFore inp, mText
        SetBorderColor inp, mBorder
        TrySetFont inp, 9, False
        TrySet inp, "SpecialEffect", 0
        TrySet inp, "BorderStyle", 1
        TrySet inp, "TextAlign", 3
        TrySet inp, "RightToLeft", True
        If Not mFields(i).Editable Then
            TrySet inp, "Locked", True
            TrySet inp, "TabStop", False
            SetBack inp, mReadOnly
            SetFore inp, mMuted
            TrySet inp, "ControlTipText", "Read-only: calculated by Excel. The formula is never overwritten."
        End If
    Next i
    StyleAllControls
    Exit Sub
ErrorHandler:
    HandleError "BuildFieldControls"
End Sub

Private Function ColumnLetter(ByVal colNumber As Long) As String
    On Error GoTo ErrorHandler
    ColumnLetter = Split(mSheet.Cells(1, colNumber).Address(True, False), "$")(0)
    Exit Function
ErrorHandler:
    HandleError "ColumnLetter"
    ColumnLetter = "A"
End Function

Private Function IsFormulaColumn(ByVal d As AppDataset, ByVal col As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim r As Long
    Dim stopRow As Long
    stopRow = d.FirstDataRow + 20
    If stopRow > d.LastDataRow Then stopRow = d.LastDataRow
    For r = d.FirstDataRow To stopRow
        If mSheet.Cells(r, col).HasFormula Then
            IsFormulaColumn = True
            Exit Function
        End If
    Next r
    IsFormulaColumn = False
    Exit Function
ErrorHandler:
    HandleError "IsFormulaColumn"
    IsFormulaColumn = False
End Function

Private Function IsDateColumn(ByVal d As AppDataset, ByVal col As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim r As Long
    Dim stopRow As Long
    Dim fmt As String
    stopRow = d.FirstDataRow + 20
    If stopRow > d.LastDataRow Then stopRow = d.LastDataRow
    For r = d.FirstDataRow To stopRow
        If Not mSheet.Cells(r, col).HasFormula Then
            If IsDate(mSheet.Cells(r, col).Value) Then
                IsDateColumn = True
                Exit Function
            End If
            fmt = LCase$(mSheet.Cells(r, col).NumberFormat)
            If InStr(1, fmt, "yy", vbTextCompare) > 0 Or InStr(1, fmt, "dd", vbTextCompare) > 0 Then
                IsDateColumn = True
                Exit Function
            End If
        End If
    Next r
    IsDateColumn = False
    Exit Function
ErrorHandler:
    HandleError "IsDateColumn"
    IsDateColumn = False
End Function

Private Function IsNumberColumn(ByVal d As AppDataset, ByVal col As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim r As Long
    Dim stopRow As Long
    Dim fmt As String
    stopRow = d.FirstDataRow + 20
    If stopRow > d.LastDataRow Then stopRow = d.LastDataRow
    For r = d.FirstDataRow To stopRow
        If Not mSheet.Cells(r, col).HasFormula Then
            If IsNumeric(mSheet.Cells(r, col).Value) And Not IsDate(mSheet.Cells(r, col).Value) Then
                IsNumberColumn = True
                Exit Function
            End If
            fmt = LCase$(mSheet.Cells(r, col).NumberFormat)
            If InStr(1, fmt, "0", vbTextCompare) > 0 Or InStr(1, fmt, "#", vbTextCompare) > 0 Then
                IsNumberColumn = True
                Exit Function
            End If
        End If
    Next r
    IsNumberColumn = False
    Exit Function
ErrorHandler:
    HandleError "IsNumberColumn"
    IsNumberColumn = False
End Function

Private Function GetValidationList(ByVal d As AppDataset, ByVal col As Long) As String
    On Error GoTo ErrorHandler
    Dim f As String
    Dim arr As Variant
    Dim i As Long
    Dim part As String
    Dim result As String
    If mSheet.Cells(d.FirstDataRow, col).Validation.Type <> xlValidateList Then Exit Function
    f = mSheet.Cells(d.FirstDataRow, col).Validation.Formula1
    If Left$(f, 1) = """" Then f = Mid$(f, 2, Len(f) - 2)
    f = Replace(f, ";", "|")
    f = Replace(f, "،", "|")
    f = Replace(f, ",", "|")
    f = Replace(f, vbCr, "|")
    f = Replace(f, vbLf, "|")
    arr = Split(f, "|")
    For i = LBound(arr) To UBound(arr)
        part = Trim$(Replace$(CStr(arr(i)), """", ""))
        If Len(part) > 0 Then
            If Len(result) > 0 Then result = result & "|"
            result = result & part
        End If
    Next i
    GetValidationList = result
    Exit Function
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Function

'-------------------------------------------------------------------------------
' Filters, grid and record loading
'-------------------------------------------------------------------------------

Private Sub PopulateFilterFields()
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim oldFlag As Boolean
    oldFlag = mSuppressEvents
    mSuppressEvents = True
    Me.cboFilterField.Clear
    Me.cboFilterField.AddItem "<All fields>"
    For i = 1 To mFieldCount
        Me.cboFilterField.AddItem mFields(i).LabelText
    Next i
    Me.cboFilterField.ListIndex = 0
    TrySet Me.cboFilterField, "Style", 2
SafeExit:
    mSuppressEvents = oldFlag
    Exit Sub
ErrorHandler:
    HandleError "PopulateFilterFields"
    Resume SafeExit
End Sub

Private Sub PopulateFilterValues()
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    Dim fld As Long
    Dim col As Long
    Dim values As String
    Dim arr As Variant
    Dim i As Long
    Dim t As String
    Dim oldFlag As Boolean
    oldFlag = mSuppressEvents
    mSuppressEvents = True
    Me.cboFilterValue.Clear
    Me.cboFilterValue.AddItem "<All values>"
    fld = Me.cboFilterField.ListIndex
    If fld > 0 And mDatasetIndex > 0 Then
        d = mDatasets(mDatasetIndex)
        col = d.FirstCol + fld - 1
        values = GetDistinctColumnValues(d, col)
        If Len(values) > 0 Then
            arr = Split(values, "|")
            For i = LBound(arr) To UBound(arr)
                t = Trim$(CStr(arr(i)))
                If Len(t) > 0 Then Me.cboFilterValue.AddItem t
            Next i
        End If
    End If
    Me.cboFilterValue.ListIndex = 0
    TrySet Me.cboFilterValue, "Style", 0
SafeExit:
    mSuppressEvents = oldFlag
    Exit Sub
ErrorHandler:
    HandleError "PopulateFilterValues"
    Resume SafeExit
End Sub

Private Function GetDistinctColumnValues(ByVal d As AppDataset, ByVal col As Long) As String
    On Error GoTo ErrorHandler
    Dim coll As Collection
    Dim r As Long
    Dim t As String
    Dim result As String
    Dim v As Variant
    Set coll = New Collection
    For r = d.FirstDataRow To d.LastDataRow
        t = Trim$(CellDisplayText(mSheet.Cells(r, col)))
        If Len(t) > 240 Then t = Left$(t, 240)
        If Len(t) > 0 Then
            If Not CollectionContains(coll, t) Then coll.Add t, t
        End If
    Next r
    For Each v In coll
        If Len(result) > 0 Then result = result & "|"
        result = result & CStr(v)
    Next v
    GetDistinctColumnValues = result
    Exit Function
ErrorHandler:
    HandleError "GetDistinctColumnValues"
    GetDistinctColumnValues = ""
End Function

Private Function CollectionContains(ByVal coll As Collection, ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    Dim tmp As Variant
    tmp = coll.Item(key)
    CollectionContains = True
    Exit Function
ErrorHandler:
    Err.Clear
    CollectionContains = False
    Resume SafeExit
SafeExit:
End Function

Private Sub RefreshGrid()
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    Dim colCount As Long
    Dim maxRows As Long
    Dim widths As String
    Dim headerArr() As Variant
    Dim tempArr() As Variant
    Dim finalArr() As Variant
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim rowCount As Long
    If mDatasetIndex < 1 Or mSheet Is Nothing Then Exit Sub
    d = mDatasets(mDatasetIndex)
    colCount = d.LastCol - d.FirstCol + 1
    widths = BuildColumnWidths(colCount)
    SetGridColumns colCount, widths
    ReDim headerArr(1 To 1, 1 To colCount)
    For c = 1 To colCount
        headerArr(1, c) = mFields(c).LabelText
    Next c
    Me.lstHeaders.List = headerArr
    maxRows = d.LastDataRow - d.FirstDataRow + 1
    If maxRows < 1 Then
        Me.lstData.Clear
        mRowCount = 0
        Exit Sub
    End If
    ReDim tempArr(1 To maxRows, 1 To colCount)
    ReDim mRowMap(1 To maxRows)
    rowCount = 0
    For r = d.FirstDataRow To d.LastDataRow
        If IsDataRow(r, d) Then
            If RowMatchesFilters(r, d) Then
                rowCount = rowCount + 1
                mRowMap(rowCount) = r
                For c = 1 To colCount
                    tempArr(rowCount, c) = CellDisplayText(mSheet.Cells(r, d.FirstCol + c - 1))
                Next c
            End If
        End If
    Next r
    mRowCount = rowCount
    If rowCount > 0 Then
        ReDim finalArr(1 To rowCount, 1 To colCount)
        For i = 1 To rowCount
            For j = 1 To colCount
                finalArr(i, j) = tempArr(i, j)
            Next j
        Next i
        Me.lstData.List = finalArr
    Else
        Me.lstData.Clear
    End If
    Me.lblRecordCount.Caption = "Records: " & CStr(rowCount)
    Exit Sub
ErrorHandler:
    HandleError "RefreshGrid"
End Sub

Private Sub SetGridColumns(ByVal colCount As Long, ByVal widths As String)
    On Error GoTo ErrorHandler
    TrySet Me.lstHeaders, "ColumnCount", colCount
    TrySet Me.lstHeaders, "ColumnWidths", widths
    TrySet Me.lstHeaders, "Enabled", False
    TrySet Me.lstHeaders, "SpecialEffect", 0
    TrySet Me.lstHeaders, "BorderStyle", 1
    TrySet Me.lstData, "ColumnCount", colCount
    TrySet Me.lstData, "ColumnWidths", widths
    TrySet Me.lstData, "BoundColumn", 1
    TrySet Me.lstData, "MultiSelect", 0
    TrySet Me.lstData, "ListStyle", 0
    TrySet Me.lstData, "IntegralHeight", False
    TrySet Me.lstData, "ScrollBars", 3
    TrySet Me.lstData, "SpecialEffect", 0
    TrySet Me.lstData, "BorderStyle", 1
    Exit Sub
ErrorHandler:
    HandleError "SetGridColumns"
End Sub

Private Function BuildColumnWidths(ByVal colCount As Long) As String
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim w As Long
    Dim result As String
    For i = 1 To colCount
        w = Len(mFields(i).LabelText) * 7 + 30
        If w < 85 Then w = 85
        If w > 190 Then w = 190
        If i > 1 Then result = result & ";"
        result = result & CStr(w)
    Next i
    BuildColumnWidths = result
    Exit Function
ErrorHandler:
    HandleError "BuildColumnWidths"
    BuildColumnWidths = "100"
End Function

Private Function IsDataRow(ByVal row As Long, ByVal d As AppDataset) As Boolean
    On Error GoTo ErrorHandler
    Dim c As Long
    For c = d.FirstCol To d.LastCol
        If Len(Trim$(CellDisplayText(mSheet.Cells(row, c)))) > 0 Then
            IsDataRow = True
            Exit Function
        End If
    Next c
    IsDataRow = False
    Exit Function
ErrorHandler:
    HandleError "IsDataRow"
    IsDataRow = False
End Function

Private Function RowMatchesFilters(ByVal row As Long, ByVal d As AppDataset) As Boolean
    On Error GoTo ErrorHandler
    Dim q As String
    Dim fld As Long
    Dim fv As String
    Dim c As Long
    Dim t As String
    Dim found As Boolean
    q = Trim$(Me.txtSearch.Text)
    fld = Me.cboFilterField.ListIndex
    fv = Trim$(CStr(Me.cboFilterValue.Value))
    If StrComp(fv, "<All values>", vbTextCompare) = 0 Then fv = ""
    If Len(q) > 0 Then
        found = False
        For c = d.FirstCol To d.LastCol
            t = CellDisplayText(mSheet.Cells(row, c))
            If InStr(1, t, q, vbTextCompare) > 0 Then
                found = True
                Exit For
            End If
        Next c
        If Not found Then
            RowMatchesFilters = False
            Exit Function
        End If
    End If
    If fld > 0 And Len(fv) > 0 Then
        t = CellDisplayText(mSheet.Cells(row, d.FirstCol + fld - 1))
        If StrComp(Trim$(t), fv, vbTextCompare) <> 0 Then
            RowMatchesFilters = False
            Exit Function
        End If
    End If
    RowMatchesFilters = True
    Exit Function
ErrorHandler:
    HandleError "RowMatchesFilters"
    RowMatchesFilters = False
End Function

Private Function CellDisplayText(ByVal cell As Range) As String
    On Error GoTo ErrorHandler
    If IsError(cell.Value) Then
        CellDisplayText = "#ERROR"
        Exit Function
    End If
    If cell.HasFormula Then
        CellDisplayText = CStr(cell.Text)
    ElseIf IsDate(cell.Value) Then
        CellDisplayText = Format(cell.Value, "yyyy-mm-dd")
    Else
        CellDisplayText = CStr(cell.Value)
    End If
    Exit Function
ErrorHandler:
    HandleError "CellDisplayText"
    CellDisplayText = ""
End Function

Private Sub LoadRecordIntoForm(ByVal row As Long)
    On Error GoTo ErrorHandler
    Dim i As Long
    mSelectedRow = row
    For i = 1 To mFieldCount
        SetInputValue i, CellDisplayText(mSheet.Cells(row, mFields(i).ColNumber))
    Next i
    Me.lblSelectedRow.Caption = "Selected row: " & CStr(row)
    ShowStatus "Loaded worksheet row " & CStr(row) & " into every field.", False
    Exit Sub
ErrorHandler:
    HandleError "LoadRecordIntoForm"
End Sub

Private Function FormControlByName(ByVal controlName As String) As Object
    On Error GoTo ErrorHandler
    Dim ctl As MSForms.Control
    Set FormControlByName = Nothing
    For Each ctl In Me.fraEditor.Controls
        If StrComp(ctl.Name, controlName, vbTextCompare) = 0 Then
            Set FormControlByName = ctl
            Exit Function
        End If
    Next ctl
    For Each ctl In Me.Controls
        If StrComp(ctl.Name, controlName, vbTextCompare) = 0 Then
            Set FormControlByName = ctl
            Exit Function
        End If
    Next ctl
    Exit Function
ErrorHandler:
    HandleError "FormControlByName"
    Set FormControlByName = Nothing
End Function

Private Function GetInputValue(ByVal fieldIndex As Long) As String
    On Error GoTo ErrorHandler
    Dim ctl As Object
    Set ctl = FormControlByName(mFields(fieldIndex).InputName)
    If ctl Is Nothing Then
        GetInputValue = ""
    Else
        GetInputValue = CStr(ctl.Value)
    End If
    Exit Function
ErrorHandler:
    HandleError "GetInputValue"
    GetInputValue = ""
End Function

Private Sub SetInputValue(ByVal fieldIndex As Long, ByVal value As String)
    On Error GoTo ErrorHandler
    Dim ctl As Object
    Set ctl = FormControlByName(mFields(fieldIndex).InputName)
    If Not ctl Is Nothing Then TrySet ctl, "Value", value
    Exit Sub
ErrorHandler:
    HandleError "SetInputValue"
End Sub

Private Sub ClearInputs()
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim oldFlag As Boolean
    oldFlag = mSuppressEvents
    mSuppressEvents = True
    For i = 1 To mFieldCount
        SetInputValue i, ""
    Next i
    mSelectedRow = 0
SafeExit:
    mSuppressEvents = oldFlag
    Exit Sub
ErrorHandler:
    HandleError "ClearInputs"
    Resume SafeExit
End Sub

'-------------------------------------------------------------------------------
' Validation, CRUD and auto-save
'-------------------------------------------------------------------------------

Private Function IsRequiredColumn(ByVal d As AppDataset, ByVal fieldIndex As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim arr As Variant
    Dim i As Long
    If Len(Trim$(d.RequiredCols)) = 0 Then
        IsRequiredColumn = False
        Exit Function
    End If
    arr = Split(d.RequiredCols, ",")
    For i = LBound(arr) To UBound(arr)
        If StrComp(Trim$(CStr(arr(i))), CStr(fieldIndex), vbTextCompare) = 0 Then
            IsRequiredColumn = True
            Exit Function
        End If
    Next i
    IsRequiredColumn = False
    Exit Function
ErrorHandler:
    HandleError "IsRequiredColumn"
    IsRequiredColumn = False
End Function

Private Function ValidateInputs(ByVal isAdd As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    Dim i As Long
    Dim v As String
    d = mDatasets(mDatasetIndex)
    For i = 1 To mFieldCount
        If mFields(i).Editable Then
            v = Trim$(GetInputValue(i))
            If IsRequiredColumn(d, i) And Len(v) = 0 Then
                ShowStatus "Required field is empty: " & mFields(i).LabelText, True
                ValidateInputs = False
                Exit Function
            End If
            If Len(v) > 0 Then
                Select Case mFields(i).Kind
                    Case "DATE"
                        If Not IsDate(v) Then
                            ShowStatus "Invalid date in " & mFields(i).LabelText & ". Use a real date.", True
                            ValidateInputs = False
                            Exit Function
                        End If
                    Case "NUMBER"
                        If Not IsNumeric(v) Then
                            ShowStatus "Invalid number in " & mFields(i).LabelText & ".", True
                            ValidateInputs = False
                            Exit Function
                        End If
                End Select
            End If
        End If
    Next i
    ValidateInputs = True
    Exit Function
ErrorHandler:
    HandleError "ValidateInputs"
    ValidateInputs = False
End Function

Private Sub AddRecord()
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    Dim newRow As Long
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim saved As Boolean
    d = mDatasets(mDatasetIndex)
    If Not d.AllowAdd Then
        ShowStatus "Add is disabled for this fixed formula layout. The daily archive dataset supports new records.", True
        GoTo SafeExit
    End If
    If Not ValidateInputs(True) Then GoTo SafeExit
    newRow = FindNextBlankRow(d)
    If newRow < 1 Then
        ShowStatus "No empty row remains inside the configured real data range.", True
        GoTo SafeExit
    End If
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    gAppWriting = True
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    CopyRowFormat mSheet, d, d.FirstDataRow, newRow
    WriteInputsToRow newRow
    ExtendFormulaColumns d, newRow
    mSheet.Calculate
    saved = SaveWorkbookSafely()
    If saved Then
        ShowStatus "Added row " & CStr(newRow) & " and saved the workbook.", False
    Else
        ShowStatus "Row was added in memory but the workbook was not saved.", True
    End If
SafeExit:
    TrySetApplication "Calculation", oldCalc
    TrySetApplication "EnableEvents", oldEvents
    TrySetApplication "DisplayAlerts", oldAlerts
    TrySetApplication "ScreenUpdating", oldScreen
    gAppWriting = False
    If saved Then
        LoadRecordIntoForm newRow
        RefreshGrid
    End If
    Exit Sub
ErrorHandler:
    HandleError "AddRecord"
    Resume SafeExit
End Sub

Private Sub UpdateRecord()
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim saved As Boolean
    d = mDatasets(mDatasetIndex)
    If Not d.AllowUpdate Then
        ShowStatus "Update is read-only for this dataset. Admin can use Excel sheet view.", True
        GoTo SafeExit
    End If
    If mSelectedRow < 1 Then
        ShowStatus "Select a record in the grid first.", True
        GoTo SafeExit
    End If
    If Not ValidateInputs(False) Then GoTo SafeExit
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    gAppWriting = True
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    WriteInputsToRow mSelectedRow
    mSheet.Calculate
    saved = SaveWorkbookSafely()
    If saved Then
        ShowStatus "Updated exact worksheet row " & CStr(mSelectedRow) & " and saved the workbook.", False
    Else
        ShowStatus "Row was updated in memory but the workbook was not saved.", True
    End If
SafeExit:
    TrySetApplication "Calculation", oldCalc
    TrySetApplication "EnableEvents", oldEvents
    TrySetApplication "DisplayAlerts", oldAlerts
    TrySetApplication "ScreenUpdating", oldScreen
    gAppWriting = False
    If saved Then
        LoadRecordIntoForm mSelectedRow
        RefreshGrid
    End If
    Exit Sub
ErrorHandler:
    HandleError "UpdateRecord"
    Resume SafeExit
End Sub

Private Sub DeleteRecord()
    On Error GoTo ErrorHandler
    Dim d As AppDataset
    Dim i As Long
    Dim c As Long
    Dim lastRow As Long
    Dim answer As VbMsgBoxResult
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim saved As Boolean
    d = mDatasets(mDatasetIndex)
    If gRole <> ROLE_ADMIN Then
        ShowStatus "Delete is disabled in Agent mode. Switch to Admin mode first.", True
        GoTo SafeExit
    End If
    If Not d.AllowDelete Then
        ShowStatus "Delete is disabled for formula-backed layouts. Use Admin Excel sheet view for raw structural edits.", True
        GoTo SafeExit
    End If
    For i = 1 To mFieldCount
        If mFields(i).Kind = "FORMULA" Then
            ShowStatus "Delete is blocked because this dataset contains live formulas.", True
            GoTo SafeExit
        End If
    Next i
    If mSelectedRow < 1 Then
        ShowStatus "Select a record in the grid first.", True
        GoTo SafeExit
    End If
    answer = MsgBox("Are you sure?" & vbCrLf & _
                    "This deletes only worksheet row " & CStr(mSelectedRow) & _
                    " from the selected dataset range.", _
                    vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete")
    If answer <> vbYes Then
        ShowStatus "Delete cancelled; no worksheet data was changed.", False
        GoTo SafeExit
    End If
    lastRow = FindLastUsedRow(d)
    If lastRow < mSelectedRow Then
        ShowStatus "The selected row is no longer present in the sheet.", True
        GoTo SafeExit
    End If
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    gAppWriting = True
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    If lastRow = mSelectedRow Then
        mSheet.Range(mSheet.Cells(mSelectedRow, d.FirstCol), _
                     mSheet.Cells(mSelectedRow, d.LastCol)).ClearContents
    Else
        For c = d.FirstCol To d.LastCol
            mSheet.Range(mSheet.Cells(mSelectedRow, c), mSheet.Cells(lastRow, c)).Delete Shift:=xlUp
        Next c
        CopyRowFormat mSheet, d, d.FirstDataRow, lastRow
    End If
    mSheet.Calculate
    saved = SaveWorkbookSafely()
    If saved Then
        ShowStatus "Deleted worksheet row " & CStr(mSelectedRow) & " and saved the workbook.", False
    Else
        ShowStatus "Row was deleted in memory but the workbook was not saved.", True
    End If
SafeExit:
    TrySetApplication "Calculation", oldCalc
    TrySetApplication "EnableEvents", oldEvents
    TrySetApplication "DisplayAlerts", oldAlerts
    TrySetApplication "ScreenUpdating", oldScreen
    gAppWriting = False
    If saved Then
        ClearInputs
        RefreshGrid
    End If
    Exit Sub
ErrorHandler:
    HandleError "DeleteRecord"
    Resume SafeExit
End Sub

Private Sub WriteInputsToRow(ByVal row As Long)
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim v As String
    Dim cell As Range
    For i = 1 To mFieldCount
        If mFields(i).Editable Then
            Set cell = mSheet.Cells(row, mFields(i).ColNumber)
            v = Trim$(GetInputValue(i))
            If Len(v) = 0 Then
                cell.ClearContents
            Else
                Select Case mFields(i).Kind
                    Case "DATE"
                        cell.Value = CDate(v)
                    Case "NUMBER"
                        cell.Value = CDbl(v)
                    Case Else
                        cell.Value = v
                End Select
            End If
        End If
    Next i
    Exit Sub
ErrorHandler:
    HandleError "WriteInputsToRow"
End Sub

Private Sub ExtendFormulaColumns(ByVal d As AppDataset, ByVal row As Long)
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim col As Long
    For i = 1 To mFieldCount
        If mFields(i).Kind = "FORMULA" Then
            col = mFields(i).ColNumber
            If mSheet.Cells(d.FirstDataRow, col).HasFormula Then
                mSheet.Cells(row, col).FormulaR1C1 = mSheet.Cells(d.FirstDataRow, col).FormulaR1C1
            End If
        End If
    Next i
    Exit Sub
ErrorHandler:
    HandleError "ExtendFormulaColumns"
End Sub

Private Function FindNextBlankRow(ByVal d As AppDataset) As Long
    On Error GoTo ErrorHandler
    Dim r As Long
    For r = d.FirstDataRow To d.LastDataRow
        If Not IsDataRow(r, d) Then
            FindNextBlankRow = r
            Exit Function
        End If
    Next r
    FindNextBlankRow = 0
    Exit Function
ErrorHandler:
    HandleError "FindNextBlankRow"
    FindNextBlankRow = 0
End Function

Private Function FindLastUsedRow(ByVal d As AppDataset) As Long
    On Error GoTo ErrorHandler
    Dim r As Long
    For r = d.LastDataRow To d.FirstDataRow Step -1
        If IsDataRow(r, d) Then
            FindLastUsedRow = r
            Exit Function
        End If
    Next r
    FindLastUsedRow = 0
    Exit Function
ErrorHandler:
    HandleError "FindLastUsedRow"
    FindLastUsedRow = 0
End Function

Private Sub CopyRowFormat(ByVal ws As Worksheet, ByVal d As AppDataset, ByVal sourceRow As Long, ByVal targetRow As Long)
    On Error GoTo ErrorHandler
    Dim src As Range
    Dim dst As Range
    Set src = ws.Range(ws.Cells(sourceRow, d.FirstCol), ws.Cells(sourceRow, d.LastCol))
    Set dst = ws.Range(ws.Cells(targetRow, d.FirstCol), ws.Cells(targetRow, d.LastCol))
    src.Copy
    PasteSpecialSafe dst, xlPasteFormats
    PasteSpecialSafe dst, xlPasteValidation
SafeExit:
    TrySetApplication "CutCopyMode", False
    Exit Sub
ErrorHandler:
    HandleError "CopyRowFormat"
    Resume SafeExit
End Sub

Private Sub PasteSpecialSafe(ByVal rng As Range, ByVal pasteType As Long)
    On Error GoTo ErrorHandler
    rng.PasteSpecial Paste:=pasteType, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Exit Sub
ErrorHandler:
    Err.Clear
    Resume SafeExit
SafeExit:
End Sub

Private Function SaveWorkbookSafely() As Boolean
    On Error GoTo ErrorHandler
    ThisWorkbook.Save
    SaveWorkbookSafely = True
    Exit Function
ErrorHandler:
    HandleError "SaveWorkbookSafely"
    SaveWorkbookSafely = False
End Function
