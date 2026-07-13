'==============================================================================
' ALTERNATIVE: Per-sheet module approach (no class module required)
'
' Copy this entire block into EACH worksheet module that contains a C1 formula.
' Replace "Sheet1" in the module name context - each sheet module is already
' scoped to its own sheet, so Me refers to that sheet only.
'==============================================================================

Option Explicit

Private PrevC1Value As Variant
Private C1IsInitialized As Boolean
Private C1IsHandling As Boolean

Private Sub Worksheet_Calculate()
    Dim currentValue As Variant

    If C1IsHandling Then Exit Sub

    currentValue = modC1RecalcCore.ReadC1Value(Me)

    If Not C1IsInitialized Then
        PrevC1Value = currentValue
        C1IsInitialized = True
        Exit Sub
    End If

    If modC1RecalcCore.ValuesEqual(currentValue, PrevC1Value) Then Exit Sub

    C1IsHandling = True
    Application.EnableEvents = False

    On Error GoTo CleanUp
    modC1RecalcCore.ProcessC1Change Me, currentValue, PrevC1Value
    PrevC1Value = currentValue

CleanUp:
    Application.EnableEvents = True
    C1IsHandling = False
End Sub

Private Sub Worksheet_Activate()
    ' Optional: resync snapshot when user returns to the sheet.
    PrevC1Value = modC1RecalcCore.ReadC1Value(Me)
    C1IsInitialized = True
End Sub
