' Paste into the ThisWorkbook module in the VBA editor.

Option Explicit

Private Sub Workbook_Open()
    modC1Bootstrap.InitializeC1Handlers
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modC1Bootstrap.TearDownC1Handlers
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    modC1Bootstrap.RegisterSheetForC1Monitoring Sh
End Sub
