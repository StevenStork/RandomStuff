Attribute VB_Name = "modFactoryAStar"
Option Explicit

'==============================================================================
' Factory Station Distance Matrix (A* pathfinding)
'
' Reads machine rectangles and operator stand points from a Stations sheet,
' treats machines as obstacles on a grid (factory origin at 0,0), then writes
' the walking distance between every pair of operator points.
'
' Expected input sheet: Stations (headers in row 1)
'   A: StationName
'   B: BottomX
'   C: BottomY
'   D: TopX
'   E: TopY
'   F: OpX
'   G: OpY
'
' Output sheet: Distances
'   Symmetric matrix of A* path lengths (same units as your coordinates).
'   Empty / "Unreachable" if no walkable path exists.
'
' Run:
'   CalculateStationDistances
'==============================================================================

' --- Sheet / column layout (edit if your workbook differs) -------------------
Private Const STATIONS_SHEET As String = "Stations"
Private Const DISTANCES_SHEET As String = "Distances"
Private Const HEADER_ROW As Long = 1
Private Const DATA_START_ROW As Long = 2
Private Const COL_NAME As Long = 1
Private Const COL_BOTTOM_X As Long = 2
Private Const COL_BOTTOM_Y As Long = 3
Private Const COL_TOP_X As Long = 4
Private Const COL_TOP_Y As Long = 5
Private Const COL_OP_X As Long = 6
Private Const COL_OP_Y As Long = 7

' --- Grid / movement settings -----------------------------------------------
' One grid cell = this many coordinate units. Smaller = more accurate, slower.
Private Const CELL_SIZE As Double = 1#
' Padding past the furthest machine / op point so paths can skirt edges.
Private Const BOUND_PADDING As Double = 2#
' True = 8-way movement (diagonals cost Sqrt(2)); False = 4-way only.
Private Const ALLOW_DIAGONALS As Boolean = True

Private Type StationRec
    Name As String
    BottomX As Double
    BottomY As Double
    TopX As Double
    TopY As Double
    OpX As Double
    OpY As Double
End Type

Private Type GridPoint
    X As Long
    Y As Long
End Type

'------------------------------------------------------------------------------
' Public entry point
'------------------------------------------------------------------------------
Public Sub CalculateStationDistances()
    Dim stations() As StationRec
    Dim stationCount As Long
    Dim blocked() As Boolean
    Dim cols As Long
    Dim rows As Long
    Dim originX As Double
    Dim originY As Double
    Dim opCells() As GridPoint
    Dim distances() As Double
    Dim i As Long
    Dim j As Long
    Dim pathLen As Double
    Dim wsOut As Worksheet

    OptimizeExcel True
    On Error GoTo CleanUp

    stationCount = ReadStations(stations)
    If stationCount < 2 Then
        Err.Raise vbObjectError + 1, "CalculateStationDistances", _
            "Need at least 2 stations with complete Bottom/Top/Op coordinates on '" & STATIONS_SHEET & "'."
    End If

    BuildOccupancyGrid stations, stationCount, blocked, cols, rows, originX, originY
    MapOperatorCells stations, stationCount, blocked, cols, rows, originX, originY, opCells

    ReDim distances(1 To stationCount, 1 To stationCount)

    For i = 1 To stationCount
        distances(i, i) = 0
        For j = i + 1 To stationCount
            pathLen = AStarDistance(blocked, cols, rows, opCells(i), opCells(j))
            distances(i, j) = pathLen
            distances(j, i) = pathLen
        Next j
    Next i

    Set wsOut = GetOrCreateSheet(DISTANCES_SHEET)
    WriteDistanceMatrix wsOut, stations, stationCount, distances

CleanUp:
    OptimizeExcel False
    If Err.Number <> 0 Then
        MsgBox "CalculateStationDistances failed:" & vbCrLf & Err.Description, vbExclamation
    End If
End Sub

'==============================================================================
' Input
'==============================================================================
Private Function ReadStations(ByRef stations() As StationRec) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim n As Long
    Dim bx As Double
    Dim by As Double
    Dim tx As Double
    Dim ty As Double

    Set ws = ThisWorkbook.Worksheets(STATIONS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, COL_NAME).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        ReadStations = 0
        Exit Function
    End If

    ReDim stations(1 To lastRow - DATA_START_ROW + 1)
    n = 0

    For r = DATA_START_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(r, COL_NAME).Value2))) = 0 Then GoTo NextRow
        If IsBlankNumeric(ws.Cells(r, COL_BOTTOM_X)) Then GoTo NextRow
        If IsBlankNumeric(ws.Cells(r, COL_BOTTOM_Y)) Then GoTo NextRow
        If IsBlankNumeric(ws.Cells(r, COL_TOP_X)) Then GoTo NextRow
        If IsBlankNumeric(ws.Cells(r, COL_TOP_Y)) Then GoTo NextRow
        If IsBlankNumeric(ws.Cells(r, COL_OP_X)) Then GoTo NextRow
        If IsBlankNumeric(ws.Cells(r, COL_OP_Y)) Then GoTo NextRow

        n = n + 1
        stations(n).Name = CStr(ws.Cells(r, COL_NAME).Value2)
        bx = CDbl(ws.Cells(r, COL_BOTTOM_X).Value2)
        by = CDbl(ws.Cells(r, COL_BOTTOM_Y).Value2)
        tx = CDbl(ws.Cells(r, COL_TOP_X).Value2)
        ty = CDbl(ws.Cells(r, COL_TOP_Y).Value2)

        ' Normalize so Bottom is min corner and Top is max corner.
        stations(n).BottomX = MinD(bx, tx)
        stations(n).BottomY = MinD(by, ty)
        stations(n).TopX = MaxD(bx, tx)
        stations(n).TopY = MaxD(by, ty)
        stations(n).OpX = CDbl(ws.Cells(r, COL_OP_X).Value2)
        stations(n).OpY = CDbl(ws.Cells(r, COL_OP_Y).Value2)
NextRow:
    Next r

    If n > 0 Then
        ReDim Preserve stations(1 To n)
    End If
    ReadStations = n
End Function

Private Function IsBlankNumeric(ByVal cell As Range) As Boolean
    IsBlankNumeric = IsEmpty(cell.Value2) Or Len(Trim$(CStr(cell.Value2))) = 0 Or Not IsNumeric(cell.Value2)
End Function

'==============================================================================
' Grid construction
'==============================================================================
Private Sub BuildOccupancyGrid( _
    ByRef stations() As StationRec, _
    ByVal stationCount As Long, _
    ByRef blocked() As Boolean, _
    ByRef cols As Long, _
    ByRef rows As Long, _
    ByRef originX As Double, _
    ByRef originY As Double)

    Dim maxX As Double
    Dim maxY As Double
    Dim i As Long
    Dim gx As Long
    Dim gy As Long
    Dim worldX As Double
    Dim worldY As Double

    originX = 0#
    originY = 0#
    maxX = 0#
    maxY = 0#

    For i = 1 To stationCount
        maxX = MaxD(maxX, stations(i).TopX)
        maxX = MaxD(maxX, stations(i).OpX)
        maxY = MaxD(maxY, stations(i).TopY)
        maxY = MaxD(maxY, stations(i).OpY)
    Next i

    maxX = maxX + BOUND_PADDING
    maxY = maxY + BOUND_PADDING

    cols = MaxL(1, CLng(Application.WorksheetFunction.Ceiling(maxX / CELL_SIZE, 1)))
    rows = MaxL(1, CLng(Application.WorksheetFunction.Ceiling(maxY / CELL_SIZE, 1)))

    ReDim blocked(0 To cols - 1, 0 To rows - 1)

    For gx = 0 To cols - 1
        For gy = 0 To rows - 1
            worldX = originX + (CDbl(gx) + 0.5) * CELL_SIZE
            worldY = originY + (CDbl(gy) + 0.5) * CELL_SIZE
            blocked(gx, gy) = PointInsideAnyMachine(worldX, worldY, stations, stationCount)
        Next gy
    Next gx
End Sub

Private Function PointInsideAnyMachine( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByRef stations() As StationRec, _
    ByVal stationCount As Long) As Boolean

    Dim i As Long
    For i = 1 To stationCount
        If x >= stations(i).BottomX And x <= stations(i).TopX _
           And y >= stations(i).BottomY And y <= stations(i).TopY Then
            PointInsideAnyMachine = True
            Exit Function
        End If
    Next i
    PointInsideAnyMachine = False
End Function

Private Sub MapOperatorCells( _
    ByRef stations() As StationRec, _
    ByVal stationCount As Long, _
    ByRef blocked() As Boolean, _
    ByVal cols As Long, _
    ByVal rows As Long, _
    ByVal originX As Double, _
    ByVal originY As Double, _
    ByRef opCells() As GridPoint)

    Dim i As Long
    Dim gp As GridPoint

    ReDim opCells(1 To stationCount)

    For i = 1 To stationCount
        gp = WorldToGrid(stations(i).OpX, stations(i).OpY, originX, originY, cols, rows)
        ' Operator stand points must remain walkable even if they sit on a machine edge.
        blocked(gp.X, gp.Y) = False
        opCells(i) = gp
    Next i
End Sub

Private Function WorldToGrid( _
    ByVal worldX As Double, _
    ByVal worldY As Double, _
    ByVal originX As Double, _
    ByVal originY As Double, _
    ByVal cols As Long, _
    ByVal rows As Long) As GridPoint

    Dim gx As Long
    Dim gy As Long

    gx = CLng(Int((worldX - originX) / CELL_SIZE))
    gy = CLng(Int((worldY - originY) / CELL_SIZE))

    If gx < 0 Then gx = 0
    If gy < 0 Then gy = 0
    If gx > cols - 1 Then gx = cols - 1
    If gy > rows - 1 Then gy = rows - 1

    WorldToGrid.X = gx
    WorldToGrid.Y = gy
End Function

'==============================================================================
' A* pathfinding
' Returns path length in world units, or -1 if unreachable.
'==============================================================================
Private Function AStarDistance( _
    ByRef blocked() As Boolean, _
    ByVal cols As Long, _
    ByVal rows As Long, _
    ByRef startPt As GridPoint, _
    ByRef goalPt As GridPoint) As Double

    Dim gScore() As Double
    Dim fScore() As Double
    Dim closed() As Boolean
    Dim inOpen() As Boolean
    Dim cameFromX() As Long
    Dim cameFromY() As Long

    Dim openX() As Long
    Dim openY() As Long
    Dim openCount As Long

    Dim curX As Long
    Dim curY As Long
    Dim bestIdx As Long
    Dim i As Long
    Dim nx As Long
    Dim ny As Long
    Dim stepCost As Double
    Dim tentG As Double
    Dim dx() As Long
    Dim dy() As Long
    Dim moveCount As Long
    Dim m As Long
    Dim pathLen As Double
    Dim px As Long
    Dim py As Long
    Dim cx As Long
    Dim cy As Long

    Const INF As Double = 1E+308

    If startPt.X = goalPt.X And startPt.Y = goalPt.Y Then
        AStarDistance = 0
        Exit Function
    End If

    If blocked(startPt.X, startPt.Y) Or blocked(goalPt.X, goalPt.Y) Then
        AStarDistance = -1
        Exit Function
    End If

    ReDim gScore(0 To cols - 1, 0 To rows - 1)
    ReDim fScore(0 To cols - 1, 0 To rows - 1)
    ReDim closed(0 To cols - 1, 0 To rows - 1)
    ReDim inOpen(0 To cols - 1, 0 To rows - 1)
    ReDim cameFromX(0 To cols - 1, 0 To rows - 1)
    ReDim cameFromY(0 To cols - 1, 0 To rows - 1)
    ReDim openX(0 To cols * rows)
    ReDim openY(0 To cols * rows)

    For nx = 0 To cols - 1
        For ny = 0 To rows - 1
            gScore(nx, ny) = INF
            fScore(nx, ny) = INF
            cameFromX(nx, ny) = -1
            cameFromY(nx, ny) = -1
        Next ny
    Next nx

    BuildMoveOffsets dx, dy, moveCount

    gScore(startPt.X, startPt.Y) = 0
    fScore(startPt.X, startPt.Y) = Heuristic(startPt.X, startPt.Y, goalPt.X, goalPt.Y)
    openCount = 1
    openX(1) = startPt.X
    openY(1) = startPt.Y
    inOpen(startPt.X, startPt.Y) = True

    Do While openCount > 0
        bestIdx = 1
        For i = 2 To openCount
            If fScore(openX(i), openY(i)) < fScore(openX(bestIdx), openY(bestIdx)) Then
                bestIdx = i
            End If
        Next i

        curX = openX(bestIdx)
        curY = openY(bestIdx)

        openX(bestIdx) = openX(openCount)
        openY(bestIdx) = openY(openCount)
        openCount = openCount - 1
        inOpen(curX, curY) = False
        closed(curX, curY) = True

        If curX = goalPt.X And curY = goalPt.Y Then
            ' Reconstruct length from parent chain (avoids float drift in gScore scale).
            pathLen = 0
            cx = curX
            cy = curY
            Do While Not (cx = startPt.X And cy = startPt.Y)
                px = cameFromX(cx, cy)
                py = cameFromY(cx, cy)
                If px < 0 Or py < 0 Then
                    AStarDistance = -1
                    Exit Function
                End If
                If Abs(cx - px) + Abs(cy - py) = 2 Then
                    pathLen = pathLen + CELL_SIZE * Sqr(2)
                Else
                    pathLen = pathLen + CELL_SIZE
                End If
                cx = px
                cy = py
            Loop
            AStarDistance = pathLen
            Exit Function
        End If

        For m = 1 To moveCount
            nx = curX + dx(m)
            ny = curY + dy(m)

            If nx < 0 Or ny < 0 Or nx >= cols Or ny >= rows Then GoTo NextMove
            If blocked(nx, ny) Then GoTo NextMove
            If closed(nx, ny) Then GoTo NextMove

            ' Prevent diagonal corner-cutting through blocked cells.
            If Abs(dx(m)) + Abs(dy(m)) = 2 Then
                If blocked(curX + dx(m), curY) Or blocked(curX, curY + dy(m)) Then GoTo NextMove
            End If

            If Abs(dx(m)) + Abs(dy(m)) = 2 Then
                stepCost = CELL_SIZE * Sqr(2)
            Else
                stepCost = CELL_SIZE
            End If

            tentG = gScore(curX, curY) + stepCost
            If tentG + 0.0000001 < gScore(nx, ny) Then
                cameFromX(nx, ny) = curX
                cameFromY(nx, ny) = curY
                gScore(nx, ny) = tentG
                fScore(nx, ny) = tentG + Heuristic(nx, ny, goalPt.X, goalPt.Y)

                If Not inOpen(nx, ny) Then
                    openCount = openCount + 1
                    openX(openCount) = nx
                    openY(openCount) = ny
                    inOpen(nx, ny) = True
                End If
            End If
NextMove:
        Next m
    Loop

    AStarDistance = -1
End Function

Private Sub BuildMoveOffsets(ByRef dx() As Long, ByRef dy() As Long, ByRef moveCount As Long)
    If ALLOW_DIAGONALS Then
        moveCount = 8
        ReDim dx(1 To 8)
        ReDim dy(1 To 8)
        dx(1) = 1: dy(1) = 0
        dx(2) = -1: dy(2) = 0
        dx(3) = 0: dy(3) = 1
        dx(4) = 0: dy(4) = -1
        dx(5) = 1: dy(5) = 1
        dx(6) = 1: dy(6) = -1
        dx(7) = -1: dy(7) = 1
        dx(8) = -1: dy(8) = -1
    Else
        moveCount = 4
        ReDim dx(1 To 4)
        ReDim dy(1 To 4)
        dx(1) = 1: dy(1) = 0
        dx(2) = -1: dy(2) = 0
        dx(3) = 0: dy(3) = 1
        dx(4) = 0: dy(4) = -1
    End If
End Sub

Private Function Heuristic(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Double
    Dim dx As Double
    Dim dy As Double
    dx = Abs(CDbl(x1 - x2))
    dy = Abs(CDbl(y1 - y2))

    If ALLOW_DIAGONALS Then
        ' Octile distance matches 8-way step costs.
        Heuristic = CELL_SIZE * (dx + dy + (Sqr(2) - 2) * MinD(dx, dy))
    Else
        Heuristic = CELL_SIZE * (dx + dy)
    End If
End Function

'==============================================================================
' Output
'==============================================================================
Private Sub WriteDistanceMatrix( _
    ByVal ws As Worksheet, _
    ByRef stations() As StationRec, _
    ByVal stationCount As Long, _
    ByRef distances() As Double)

    Dim i As Long
    Dim j As Long

    ws.Cells.Clear

    ws.Cells(1, 1).Value = "From \ To"
    For i = 1 To stationCount
        ws.Cells(1, i + 1).Value = stations(i).Name
        ws.Cells(i + 1, 1).Value = stations(i).Name
    Next i

    For i = 1 To stationCount
        For j = 1 To stationCount
            If distances(i, j) < 0 Then
                ws.Cells(i + 1, j + 1).Value = "Unreachable"
            Else
                ws.Cells(i + 1, j + 1).Value = Round(distances(i, j), 4)
            End If
        Next j
    Next i

    ws.Columns.AutoFit
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws
End Function

'==============================================================================
' Small helpers
'==============================================================================
Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

'==============================================================================
' Optional: create a blank Stations sheet with headers
'==============================================================================
Public Sub CreateStationsTemplate()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(STATIONS_SHEET)
    ws.Cells.Clear

    ws.Cells(1, COL_NAME).Value = "StationName"
    ws.Cells(1, COL_BOTTOM_X).Value = "BottomX"
    ws.Cells(1, COL_BOTTOM_Y).Value = "BottomY"
    ws.Cells(1, COL_TOP_X).Value = "TopX"
    ws.Cells(1, COL_TOP_Y).Value = "TopY"
    ws.Cells(1, COL_OP_X).Value = "OpX"
    ws.Cells(1, COL_OP_Y).Value = "OpY"
    ws.Columns.AutoFit
End Sub
