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
