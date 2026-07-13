' Paste into the ThisWorkbook module in the VBA editor.

Option Explicit

Private Sub Workbook_Open()
    modSheetRenameMonitor.InitializeSheetNameTracking
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modSheetRenameMonitor.TearDownSheetNameTracking
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    modSheetRenameMonitor.RegisterSheetName Sh
End Sub

Private Sub Workbook_BeforeSheetDelete(ByVal Sh As Object)
    modSheetRenameMonitor.UnregisterSheetName Sh
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    modSheetRenameMonitor.CheckForSheetRenames
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    modSheetRenameMonitor.CheckForSheetRenames
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    modSheetRenameMonitor.CheckForSheetRenames
End Sub
