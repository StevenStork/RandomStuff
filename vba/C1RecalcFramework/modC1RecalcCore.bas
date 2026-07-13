Attribute VB_Name = "modC1RecalcCore"
Option Explicit

'==============================================================================
' C1 Recalculation Framework - Core Module
'
' Detects when cell C1's calculated value changes on a worksheet and runs
' sheet-specific logic. Only the sheet where C1 actually changed is affected.
'
' Primary integration: use clsSheetC1Handler (WithEvents) via modC1Bootstrap.
' Alternative: copy WorksheetModuleTemplate.bas into each sheet module.
'==============================================================================

Public Const C1_MONITOR_CELL As String = "C1"

'------------------------------------------------------------------------------
' Entry point: called only for the worksheet whose C1 value changed.
' Replace the body of ProcessC1Change with your business logic.
'------------------------------------------------------------------------------
Public Sub ProcessC1Change(ByVal ws As Worksheet, ByVal newValue As Variant, ByVal oldValue As Variant)
    On Error GoTo ErrHandler

    ' --- YOUR CODE STARTS HERE ---
    ' Example: write a timestamp on the same sheet only
    ' ws.Range("D1").Value = Now
    ' ws.Range("E1").Value = "C1 changed from " & CStr(oldValue) & " to " & CStr(newValue)

    Debug.Print "C1 changed on '" & ws.Name & "': " & ValueToDebugString(oldValue) & " -> " & ValueToDebugString(newValue)
    ' --- YOUR CODE ENDS HERE ---

    Exit Sub

ErrHandler:
    Debug.Print "ProcessC1Change error on '" & ws.Name & "': " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Returns True if this worksheet should be monitored for C1 recalculation.
' Customize the filter logic here (by name, tab color, a list range, etc.).
'------------------------------------------------------------------------------
Public Function ShouldMonitorSheet(ByVal ws As Worksheet) As Boolean
    On Error GoTo ErrHandler

    ' Default: monitor every worksheet except this optional exclusions list.
    Select Case ws.Name
        Case "Config", "ReadMe"
            ShouldMonitorSheet = False
        Case Else
            ShouldMonitorSheet = True
    End Select

    Exit Function

ErrHandler:
    ShouldMonitorSheet = False
End Function

'------------------------------------------------------------------------------
' Compares two cell values, including errors, blanks, and numeric edge cases.
' Uses Value2 so formatting differences do not trigger false positives.
'------------------------------------------------------------------------------
Public Function ValuesEqual(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean
    Dim err1 As Long
    Dim err2 As Long

    If IsEmpty(v1) And IsEmpty(v2) Then
        ValuesEqual = True
        Exit Function
    End If

    If IsError(v1) And Not IsError(v2) Then
        ValuesEqual = False
        Exit Function
    End If

    If IsError(v2) And Not IsError(v1) Then
        ValuesEqual = False
        Exit Function
    End If

    If IsError(v1) And IsError(v2) Then
        On Error Resume Next
        err1 = CVErr(v1)
        err2 = CVErr(v2)
        On Error GoTo 0
        ValuesEqual = (err1 = err2)
        Exit Function
    End If

    On Error Resume Next
    ValuesEqual = (v1 = v2)
    If Err.Number <> 0 Then
        Err.Clear
        ValuesEqual = False
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
Public Function ReadC1Value(ByVal ws As Worksheet) As Variant
    ReadC1Value = ws.Range(C1_MONITOR_CELL).Value2
End Function

'------------------------------------------------------------------------------
Public Function ValueToDebugString(ByVal v As Variant) As String
    If IsEmpty(v) Then
        ValueToDebugString = "<Empty>"
    ElseIf IsError(v) Then
        ValueToDebugString = "#" & CStr(v)
    ElseIf IsNull(v) Then
        ValueToDebugString = "<Null>"
    Else
        ValueToDebugString = CStr(v)
    End If
End Function
