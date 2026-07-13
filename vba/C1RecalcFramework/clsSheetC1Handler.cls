VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetC1Handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Per-worksheet event handler. One instance per monitored sheet.
' Keeps a private snapshot of C1 and reacts only when that sheet's C1 changes.
'==============================================================================

Public WithEvents MonitoredSheet As Worksheet

Private PrevC1Value As Variant
Private IsInitialized As Boolean
Private IsHandling As Boolean

Public Sub Initialize(ByVal ws As Worksheet)
    Set MonitoredSheet = ws
    PrevC1Value = modC1RecalcCore.ReadC1Value(ws)
    IsInitialized = True
    IsHandling = False
End Sub

Public Sub RefreshSnapshot()
    If MonitoredSheet Is Nothing Then Exit Sub
    PrevC1Value = modC1RecalcCore.ReadC1Value(MonitoredSheet)
End Sub

Private Sub MonitoredSheet_Calculate()
    Dim currentValue As Variant

    If MonitoredSheet Is Nothing Then Exit Sub
    If IsHandling Then Exit Sub

    currentValue = modC1RecalcCore.ReadC1Value(MonitoredSheet)

    If Not IsInitialized Then
        PrevC1Value = currentValue
        IsInitialized = True
        Exit Sub
    End If

    If modC1RecalcCore.ValuesEqual(currentValue, PrevC1Value) Then Exit Sub

    IsHandling = True
    Application.EnableEvents = False

    On Error GoTo CleanUp
    modC1RecalcCore.ProcessC1Change MonitoredSheet, currentValue, PrevC1Value
    PrevC1Value = currentValue

CleanUp:
    Application.EnableEvents = True
    IsHandling = False
End Sub
