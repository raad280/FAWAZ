VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    modRaadApp.StartApplication
    Exit Sub
ErrorHandler:
    modRaadApp.HandleGlobalError "Workbook_Open"
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler
    modRaadApp.RefreshOpenForm Sh, Target
    Exit Sub
ErrorHandler:
    modRaadApp.HandleGlobalError "Workbook_SheetChange"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrorHandler
    modRaadApp.CleanupBeforeClose
    Exit Sub
ErrorHandler:
    modRaadApp.HandleGlobalError "Workbook_BeforeClose"
End Sub
