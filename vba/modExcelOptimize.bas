Attribute VB_Name = "modExcelOptimize"
Option Explicit

'==============================================================================
' Excel Performance Optimizer
'
' One-call helper for long-running macros:
'
'   OptimizeExcel True   ' turn optimizations ON at the start of your code
'   OptimizeExcel False  ' restore previous settings at the end
'
' Always pair them in an error-safe Finally-style pattern so Excel is never
' left with calculation or screen updating permanently disabled.
'==============================================================================

Private SavedScreenUpdating As Boolean
Private SavedEnableEvents As Boolean
Private SavedDisplayAlerts As Boolean
Private SavedCalculation As XlCalculation
Private SavedStatusBar As Variant
Private OptimizationDepth As Long
Private StateIsSaved As Boolean

'------------------------------------------------------------------------------
' OptimizeExcel True  - disables screen updates, events, alerts, and calculation
' OptimizeExcel False - restores the settings that were active beforehand
'
' Nested calls are supported: only the outermost False restores Excel.
'------------------------------------------------------------------------------
Public Sub OptimizeExcel(ByVal OptimizeOn As Boolean)
    If OptimizeOn Then
        ApplyOptimizations
    Else
        RestoreDefaults
    End If
End Sub

'------------------------------------------------------------------------------
Private Sub ApplyOptimizations()
    On Error Resume Next

    OptimizationDepth = OptimizationDepth + 1

    ' Only capture the live Excel state on the outermost enable.
    If OptimizationDepth = 1 Or Not StateIsSaved Then
        SavedScreenUpdating = Application.ScreenUpdating
        SavedEnableEvents = Application.EnableEvents
        SavedDisplayAlerts = Application.DisplayAlerts
        SavedCalculation = Application.Calculation
        SavedStatusBar = Application.StatusBar
        StateIsSaved = True
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = False
End Sub

'------------------------------------------------------------------------------
Private Sub RestoreDefaults()
    On Error Resume Next

    If OptimizationDepth > 0 Then
        OptimizationDepth = OptimizationDepth - 1
    End If

    ' Nested callers: wait until the outermost False to restore.
    If OptimizationDepth > 0 Then Exit Sub
    If Not StateIsSaved Then Exit Sub

    Application.ScreenUpdating = SavedScreenUpdating
    Application.EnableEvents = SavedEnableEvents
    Application.DisplayAlerts = SavedDisplayAlerts
    Application.Calculation = SavedCalculation

    If IsEmpty(SavedStatusBar) Or SavedStatusBar = False Then
        Application.StatusBar = False
    Else
        Application.StatusBar = SavedStatusBar
    End If

    StateIsSaved = False
End Sub
