Attribute VB_Name = "modC1Bootstrap"
Option Explicit

'==============================================================================
' Registers one clsSheetC1Handler per monitored worksheet at workbook open.
' The module-level collection keeps handler instances alive so WithEvents work.
'==============================================================================

Private SheetHandlers As Collection

Public Sub InitializeC1Handlers()
    Dim ws As Worksheet
    Dim handler As clsSheetC1Handler

    TearDownC1Handlers
    Set SheetHandlers = New Collection

    For Each ws In ThisWorkbook.Worksheets
        If modC1RecalcCore.ShouldMonitorSheet(ws) Then
            Set handler = New clsSheetC1Handler
            handler.Initialize ws
            SheetHandlers.Add handler, ws.CodeName
        End If
    Next ws
End Sub

Public Sub TearDownC1Handlers()
    Set SheetHandlers = Nothing
End Sub

Public Sub RefreshAllC1Snapshots()
    Dim handler As clsSheetC1Handler

    If SheetHandlers Is Nothing Then Exit Sub

    For Each handler In SheetHandlers
        handler.RefreshSnapshot
    Next handler
End Sub

Public Sub RegisterSheetForC1Monitoring(ByVal ws As Worksheet)
    Dim handler As clsSheetC1Handler

    If SheetHandlers Is Nothing Then
        Set SheetHandlers = New Collection
    End If

    On Error Resume Next
    SheetHandlers.Remove ws.CodeName
    On Error GoTo 0

    If modC1RecalcCore.ShouldMonitorSheet(ws) Then
        Set handler = New clsSheetC1Handler
        handler.Initialize ws
        SheetHandlers.Add handler, ws.CodeName
    End If
End Sub

Public Sub UnregisterSheetFromC1Monitoring(ByVal ws As Worksheet)
    If SheetHandlers Is Nothing Then Exit Sub

    On Error Resume Next
    SheetHandlers.Remove ws.CodeName
    On Error GoTo 0
End Sub
