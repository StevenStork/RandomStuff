Attribute VB_Name = "modSheetRenameMonitor"
Option Explicit

'==============================================================================
' Sheet Rename Monitor
'
' Excel has no native "sheet renamed" event. This module tracks each sheet's
' CodeName -> Name mapping and detects renames during common workbook events.
'
' When a rename is detected, it calls the predefined macro named below, passing
' the sheet's new name as a String argument.
'==============================================================================

' Name of the Public Sub to call (must exist in a standard module in this workbook).
Public Const SHEET_RENAME_MACRO As String = "OnSheetNameChanged"

Private SheetNames As Object   ' Scripting.Dictionary: CodeName -> Name
Private IsChecking As Boolean

'------------------------------------------------------------------------------
' Initialize or rebuild the CodeName -> Name snapshot for all sheets.
'------------------------------------------------------------------------------
Public Sub InitializeSheetNameTracking()
    Dim sh As Object

    Set SheetNames = CreateObject("Scripting.Dictionary")

    For Each sh In ThisWorkbook.Sheets
        If SupportsRenameTracking(sh) Then
            SheetNames(sh.CodeName) = sh.Name
        End If
    Next sh
End Sub

'------------------------------------------------------------------------------
' Tear down tracking state (optional, e.g. on workbook close).
'------------------------------------------------------------------------------
Public Sub TearDownSheetNameTracking()
    Set SheetNames = Nothing
End Sub

'------------------------------------------------------------------------------
' Compare live sheet names to the snapshot and invoke the callback on changes.
' Safe to call from multiple workbook events; only the changed sheet is reported.
'------------------------------------------------------------------------------
Public Sub CheckForSheetRenames()
    Dim sh As Object
    Dim codeName As String
    Dim oldName As String
    Dim newName As String

    If IsChecking Then Exit Sub
    If SheetNames Is Nothing Then InitializeSheetNameTracking

    IsChecking = True
    On Error GoTo CleanUp

    For Each sh In ThisWorkbook.Sheets
        If Not SupportsRenameTracking(sh) Then GoTo NextSheet

        codeName = sh.CodeName
        newName = sh.Name

        If SheetNames.Exists(codeName) Then
            oldName = CStr(SheetNames(codeName))

            If StrComp(oldName, newName, vbBinaryCompare) <> 0 Then
                SheetNames(codeName) = newName
                CallSheetRenameMacro newName, oldName
            End If
        Else
            ' New sheet that was not registered via Workbook_NewSheet yet.
            SheetNames(codeName) = newName
        End If

NextSheet:
    Next sh

CleanUp:
    IsChecking = False
End Sub

'------------------------------------------------------------------------------
' Register a newly added sheet without treating it as a rename.
'------------------------------------------------------------------------------
Public Sub RegisterSheetName(ByVal sh As Object)
    If SheetNames Is Nothing Then InitializeSheetNameTracking
    If Not SupportsRenameTracking(sh) Then Exit Sub

    SheetNames(sh.CodeName) = sh.Name
End Sub

'------------------------------------------------------------------------------
' Remove a sheet from tracking before it is deleted.
'------------------------------------------------------------------------------
Public Sub UnregisterSheetName(ByVal sh As Object)
    If SheetNames Is Nothing Then Exit Sub

    On Error Resume Next
    SheetNames.Remove sh.CodeName
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Invokes the predefined macro, passing the new sheet name.
' Uses Application.Run so you can point SHEET_RENAME_MACRO at any Public Sub
' in this workbook that accepts at least one String argument.
'------------------------------------------------------------------------------
Private Sub CallSheetRenameMacro(ByVal sheetName As String, ByVal oldName As String)
    Dim qualifiedMacro As String

    On Error GoTo ErrHandler

    qualifiedMacro = "'" & ThisWorkbook.Name & "'!" & SHEET_RENAME_MACRO
    Application.Run qualifiedMacro, sheetName

    Exit Sub

ErrHandler:
    Debug.Print "Sheet rename macro '" & SHEET_RENAME_MACRO & "' failed: " & Err.Description
    Debug.Print "  Old name: " & oldName & "  New name: " & sheetName
End Sub

'------------------------------------------------------------------------------
Private Function SupportsRenameTracking(ByVal sh As Object) As Boolean
    On Error GoTo ErrHandler

    SupportsRenameTracking = (Len(sh.CodeName) > 0) And (Len(sh.Name) > 0)
    Exit Function

ErrHandler:
    SupportsRenameTracking = False
End Function

'==============================================================================
' PREDEFINED MACRO STUB
' Rename this procedure only if you also update SHEET_RENAME_MACRO above.
'==============================================================================
Public Sub OnSheetNameChanged(ByVal sheetName As String)
    ' sheetName = the sheet's name after the rename
    '
    ' Example:
    '   Debug.Print "Renamed to: " & sheetName
    '   ThisWorkbook.Worksheets(sheetName).Range("A1").Value = "Renamed at " & Now

    Debug.Print "Sheet renamed to: " & sheetName
End Sub
