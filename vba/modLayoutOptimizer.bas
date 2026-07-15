Attribute VB_Name = "modLayoutOptimizer"
Option Explicit

'==============================================================================
' Station Layout Optimizer
'
' Moves and rotates stations on the Stations sheet, then calls
' RecalculateStationDistances from modFactoryAStar to score the layout.
' Repeated trials search for a layout that minimizes total operator travel.
'
' Manual helpers:
'   MoveStation "Press1", 2, -1
'   RotateStation "Press1", 90
'   RecalculateLayoutDistances
'
' Auto search:
'   OptimizeStationLayout
'
' Requires: modFactoryAStar.bas and modExcelOptimize.bas
'==============================================================================

' --- Search settings (edit these) -------------------------------------------
Private Const OPT_ITERATIONS As Long = 150
Private Const OPT_RESTARTS As Long = 3
Private Const OPT_MAX_MOVE As Double = 5#          ' max translation per try
Private Const OPT_ROTATE_CHANCE As Double = 0.35   ' probability of rotating vs moving
Private Const OPT_CLEARANCE As Double = 0.5        ' min gap between machines
Private Const FACTORY_MIN_X As Double = 0#
Private Const FACTORY_MIN_Y As Double = 0#
Private Const FACTORY_MAX_X As Double = 100#       ' set to your floor size
Private Const FACTORY_MAX_Y As Double = 100#
Private Const START_TEMPERATURE As Double = 25#    ' simulated annealing
Private Const COOLING_RATE As Double = 0.97

Private Type LayoutStation
    Name As String
    BottomX As Double
    BottomY As Double
    TopX As Double
    TopY As Double
    OpX As Double
    OpY As Double
    SheetRow As Long
End Type

'------------------------------------------------------------------------------
' Move a station by (dx, dy). Updates footprint and operator point together.
'------------------------------------------------------------------------------
Public Sub MoveStation(ByVal stationName As String, ByVal dx As Double, ByVal dy As Double)
    Dim stations() As LayoutStation
    Dim n As Long
    Dim idx As Long

    n = LoadStations(stations)
    idx = FindStationIndex(stations, n, stationName)
    If idx = 0 Then
        MsgBox "Station not found: " & stationName, vbExclamation
        Exit Sub
    End If

    TranslateStation stations(idx), dx, dy
    If Not StationInsideFactory(stations(idx)) Then
        MsgBox "Move would place '" & stationName & "' outside the factory bounds.", vbExclamation
        Exit Sub
    End If

    WriteStationRow stations(idx)
End Sub

'------------------------------------------------------------------------------
' Rotate a station by 90 / 180 / 270 degrees around its rectangle center.
' Operator stand point rotates with the machine.
'------------------------------------------------------------------------------
Public Sub RotateStation(ByVal stationName As String, ByVal degrees As Long)
    Dim stations() As LayoutStation
    Dim n As Long
    Dim idx As Long
    Dim steps As Long

    n = LoadStations(stations)
    idx = FindStationIndex(stations, n, stationName)
    If idx = 0 Then
        MsgBox "Station not found: " & stationName, vbExclamation
        Exit Sub
    End If

    steps = NormalizeRotationSteps(degrees)
    If steps = 0 Then Exit Sub

    RotateStationInPlace stations(idx), steps
    If Not StationInsideFactory(stations(idx)) Then
        MsgBox "Rotation would place '" & stationName & "' outside the factory bounds.", vbExclamation
        Exit Sub
    End If

    WriteStationRow stations(idx)
End Sub

'------------------------------------------------------------------------------
' Convenience wrapper: rewrite A* distance matrix for the current layout.
'------------------------------------------------------------------------------
Public Sub RecalculateLayoutDistances()
    FloorBoundMaxX = FACTORY_MAX_X
    FloorBoundMaxY = FACTORY_MAX_Y
    CalculateStationDistances
End Sub

'------------------------------------------------------------------------------
' Search for a lower-travel layout by repeatedly moving/rotating stations
' and scoring each candidate with A*.
' Writes the best layout back to Stations and regenerates Distances.
'------------------------------------------------------------------------------
Public Sub OptimizeStationLayout()
    Dim current() As LayoutStation
    Dim candidate() As LayoutStation
    Dim best() As LayoutStation
    Dim n As Long
    Dim restart As Long
    Dim iter As Long
    Dim currentScore As Double
    Dim candidateScore As Double
    Dim bestScore As Double
    Dim temperature As Double
    Dim accepted As Long
    Dim evaluated As Long
    Dim startScore As Double
    Dim errNum As Long
    Dim errDesc As String

    OptimizeExcel True
    On Error GoTo CleanUp

    FloorBoundMaxX = FACTORY_MAX_X
    FloorBoundMaxY = FACTORY_MAX_Y
    Randomize

    n = LoadStations(current)
    If n < 2 Then
        Err.Raise vbObjectError + 200, "OptimizeStationLayout", _
            "Need at least 2 stations on the Stations sheet."
    End If

    If StationsOverlap(current, n) Then
        Err.Raise vbObjectError + 201, "OptimizeStationLayout", _
            "Starting layout has overlapping machines. Separate them before optimizing."
    End If

    CopyLayout current, n, best
    WriteAllStations best, n
    startScore = RecalculateStationDistances(False, False)
    If startScore < 0 Then
        Err.Raise vbObjectError + 202, "OptimizeStationLayout", _
            "Starting layout has unreachable operator pairs. Fix access paths first."
    End If
    bestScore = startScore

    For restart = 1 To OPT_RESTARTS
        CopyLayout best, n, current
        WriteAllStations current, n
        currentScore = RecalculateStationDistances(False, False)
        If currentScore < 0 Then currentScore = bestScore

        temperature = START_TEMPERATURE

        For iter = 1 To OPT_ITERATIONS
            CopyLayout current, n, candidate
            MutateLayout candidate, n

            If Not LayoutFeasible(candidate, n) Then GoTo NextIter

            WriteAllStations candidate, n
            candidateScore = RecalculateStationDistances(False, False)
            evaluated = evaluated + 1

            If candidateScore < 0 Then GoTo NextIter

            If ShouldAccept(currentScore, candidateScore, temperature) Then
                CopyLayout candidate, n, current
                currentScore = candidateScore
                accepted = accepted + 1

                If candidateScore + 0.0000001 < bestScore Then
                    CopyLayout candidate, n, best
                    bestScore = candidateScore
                End If
            End If

NextIter:
            temperature = temperature * COOLING_RATE
            If iter Mod 10 = 0 Then
                Application.StatusBar = "Layout optimize restart " & restart & "/" & OPT_RESTARTS & _
                    "  iter " & iter & "/" & OPT_ITERATIONS & _
                    "  best=" & Format$(bestScore, "0.00")
            End If
        Next iter
    Next restart

    WriteAllStations best, n
    RecalculateStationDistances True, False

CleanUp:
    errNum = Err.Number
    errDesc = Err.Description
    Application.StatusBar = False
    OptimizeExcel False

    If errNum <> 0 Then
        MsgBox "OptimizeStationLayout failed:" & vbCrLf & errDesc, vbExclamation
    Else
        MsgBox "Layout optimization complete." & vbCrLf & vbCrLf & _
               "Starting total travel: " & Format$(startScore, "0.00") & vbCrLf & _
               "Best total travel:     " & Format$(bestScore, "0.00") & vbCrLf & _
               "Improvement:           " & Format$(startScore - bestScore, "0.00") & vbCrLf & _
               "Candidates evaluated:  " & evaluated & vbCrLf & _
               "Accepted moves:        " & accepted, vbInformation
    End If
End Sub

'==============================================================================
' Mutations
'==============================================================================
Private Sub MutateLayout(ByRef stations() As LayoutStation, ByVal n As Long)
    Dim idx As Long
    Dim dx As Double
    Dim dy As Double
    Dim steps As Long

    idx = 1 + Int(Rnd() * n)

    If Rnd() < OPT_ROTATE_CHANCE Then
        steps = 1 + Int(Rnd() * 3)   ' 90, 180, or 270
        RotateStationInPlace stations(idx), steps
    Else
        dx = (Rnd() * 2# - 1#) * OPT_MAX_MOVE
        dy = (Rnd() * 2# - 1#) * OPT_MAX_MOVE
        TranslateStation stations(idx), dx, dy
    End If
End Sub

Private Sub TranslateStation(ByRef st As LayoutStation, ByVal dx As Double, ByVal dy As Double)
    st.BottomX = st.BottomX + dx
    st.BottomY = st.BottomY + dy
    st.TopX = st.TopX + dx
    st.TopY = st.TopY + dy
    st.OpX = st.OpX + dx
    st.OpY = st.OpY + dy
End Sub

' steps = 1 => 90° CCW, 2 => 180°, 3 => 270° CCW about rectangle center
Private Sub RotateStationInPlace(ByRef st As LayoutStation, ByVal steps As Long)
    Dim s As Long
    Dim cx As Double
    Dim cy As Double
    Dim halfW As Double
    Dim halfH As Double
    Dim relX As Double
    Dim relY As Double
    Dim newRelX As Double
    Dim newRelY As Double
    Dim w As Double
    Dim h As Double

    steps = ((steps Mod 4) + 4) Mod 4
    If steps = 0 Then Exit Sub

    For s = 1 To steps
        cx = (st.BottomX + st.TopX) / 2#
        cy = (st.BottomY + st.TopY) / 2#
        w = st.TopX - st.BottomX
        h = st.TopY - st.BottomY

        ' Operator: 90° CCW around center => (x,y) -> (-y, x)
        relX = st.OpX - cx
        relY = st.OpY - cy
        newRelX = -relY
        newRelY = relX
        st.OpX = cx + newRelX
        st.OpY = cy + newRelY

        ' AABB swaps width/height for a 90° turn.
        halfW = h / 2#
        halfH = w / 2#
        st.BottomX = cx - halfW
        st.BottomY = cy - halfH
        st.TopX = cx + halfW
        st.TopY = cy + halfH
    Next s
End Sub

Private Function NormalizeRotationSteps(ByVal degrees As Long) As Long
    Dim d As Long
    d = ((degrees Mod 360) + 360) Mod 360
    Select Case d
        Case 0
            NormalizeRotationSteps = 0
        Case 90
            NormalizeRotationSteps = 1
        Case 180
            NormalizeRotationSteps = 2
        Case 270
            NormalizeRotationSteps = 3
        Case Else
            Err.Raise vbObjectError + 203, "RotateStation", _
                "Rotation must be 0, 90, 180, or 270 degrees."
    End Select
End Function

Private Function ShouldAccept(ByVal currentScore As Double, ByVal candidateScore As Double, ByVal temperature As Double) As Boolean
    Dim delta As Double
    Dim p As Double

    If candidateScore <= currentScore Then
        ShouldAccept = True
        Exit Function
    End If

    If temperature <= 0.0000001 Then
        ShouldAccept = False
        Exit Function
    End If

    delta = candidateScore - currentScore
    p = Exp(-delta / temperature)
    ShouldAccept = (Rnd() < p)
End Function

'==============================================================================
' Feasibility
'==============================================================================
Private Function LayoutFeasible(ByRef stations() As LayoutStation, ByVal n As Long) As Boolean
    Dim i As Long

    For i = 1 To n
        If Not StationInsideFactory(stations(i)) Then
            LayoutFeasible = False
            Exit Function
        End If
    Next i

    LayoutFeasible = Not StationsOverlap(stations, n)
End Function

Private Function StationInsideFactory(ByRef st As LayoutStation) As Boolean
    StationInsideFactory = _
        st.BottomX >= FACTORY_MIN_X And st.BottomY >= FACTORY_MIN_Y And _
        st.TopX <= FACTORY_MAX_X And st.TopY <= FACTORY_MAX_Y And _
        st.OpX >= FACTORY_MIN_X And st.OpY >= FACTORY_MIN_Y And _
        st.OpX <= FACTORY_MAX_X And st.OpY <= FACTORY_MAX_Y
End Function

Private Function StationsOverlap(ByRef stations() As LayoutStation, ByVal n As Long) As Boolean
    Dim i As Long
    Dim j As Long

    For i = 1 To n - 1
        For j = i + 1 To n
            If AabbOverlap(stations(i), stations(j), OPT_CLEARANCE) Then
                StationsOverlap = True
                Exit Function
            End If
        Next j
    Next i
    StationsOverlap = False
End Function

Private Function AabbOverlap(ByRef a As LayoutStation, ByRef b As LayoutStation, ByVal clearance As Double) As Boolean
    AabbOverlap = Not ( _
        a.TopX + clearance <= b.BottomX Or _
        b.TopX + clearance <= a.BottomX Or _
        a.TopY + clearance <= b.BottomY Or _
        b.TopY + clearance <= a.BottomY)
End Function

'==============================================================================
' Sheet I/O
'==============================================================================
Private Function LoadStations(ByRef stations() As LayoutStation) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim n As Long
    Dim bx As Double
    Dim by As Double
    Dim tx As Double
    Dim ty As Double

    Set ws = ThisWorkbook.Worksheets(STATIONS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, STATIONS_COL_NAME).End(xlUp).Row
    If lastRow < STATIONS_DATA_START_ROW Then
        LoadStations = 0
        Exit Function
    End If

    ReDim stations(1 To lastRow - STATIONS_DATA_START_ROW + 1)
    n = 0

    For r = STATIONS_DATA_START_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(r, STATIONS_COL_NAME).Value2))) = 0 Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_BOTTOM_X).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_BOTTOM_Y).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_TOP_X).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_TOP_Y).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_OP_X).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_OP_Y).Value2) Then GoTo NextRow

        n = n + 1
        bx = CDbl(ws.Cells(r, STATIONS_COL_BOTTOM_X).Value2)
        by = CDbl(ws.Cells(r, STATIONS_COL_BOTTOM_Y).Value2)
        tx = CDbl(ws.Cells(r, STATIONS_COL_TOP_X).Value2)
        ty = CDbl(ws.Cells(r, STATIONS_COL_TOP_Y).Value2)

        stations(n).Name = CStr(ws.Cells(r, STATIONS_COL_NAME).Value2)
        stations(n).BottomX = MinD(bx, tx)
        stations(n).BottomY = MinD(by, ty)
        stations(n).TopX = MaxD(bx, tx)
        stations(n).TopY = MaxD(by, ty)
        stations(n).OpX = CDbl(ws.Cells(r, STATIONS_COL_OP_X).Value2)
        stations(n).OpY = CDbl(ws.Cells(r, STATIONS_COL_OP_Y).Value2)
        stations(n).SheetRow = r
NextRow:
    Next r

    If n > 0 Then ReDim Preserve stations(1 To n)
    LoadStations = n
End Function

Private Sub WriteStationRow(ByRef st As LayoutStation)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(STATIONS_SHEET)

    ws.Cells(st.SheetRow, STATIONS_COL_NAME).Value = st.Name
    ws.Cells(st.SheetRow, STATIONS_COL_BOTTOM_X).Value = st.BottomX
    ws.Cells(st.SheetRow, STATIONS_COL_BOTTOM_Y).Value = st.BottomY
    ws.Cells(st.SheetRow, STATIONS_COL_TOP_X).Value = st.TopX
    ws.Cells(st.SheetRow, STATIONS_COL_TOP_Y).Value = st.TopY
    ws.Cells(st.SheetRow, STATIONS_COL_OP_X).Value = st.OpX
    ws.Cells(st.SheetRow, STATIONS_COL_OP_Y).Value = st.OpY
End Sub

Private Sub WriteAllStations(ByRef stations() As LayoutStation, ByVal n As Long)
    Dim i As Long
    For i = 1 To n
        WriteStationRow stations(i)
    Next i
End Sub

Private Function FindStationIndex(ByRef stations() As LayoutStation, ByVal n As Long, ByVal stationName As String) As Long
    Dim i As Long
    For i = 1 To n
        If StrComp(stations(i).Name, stationName, vbTextCompare) = 0 Then
            FindStationIndex = i
            Exit Function
        End If
    Next i
    FindStationIndex = 0
End Function

Private Sub CopyLayout(ByRef source() As LayoutStation, ByVal n As Long, ByRef dest() As LayoutStation)
    Dim i As Long
    ReDim dest(1 To n)
    For i = 1 To n
        dest(i) = source(i)
    Next i
End Sub

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function
