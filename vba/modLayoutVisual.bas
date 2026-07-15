Attribute VB_Name = "modLayoutVisual"
Option Explicit

'==============================================================================
' Layout Visual
'
' Draws the factory on sheet "Layout Visual":
'   - Light grey rectangle  = factory floor
'   - Blue rectangles       = equipment (Stations footprints)
'   - Red rectangles        = Obstacles
'   - Small black dots      = station operator points
'
' Every run fully deletes existing shapes on that sheet, then rebuilds them
' so repeated draws do not accumulate file bloat.
'
' Run:
'   DrawFactoryLayout
'
' Requires Stations (and optionally Obstacles) sheets used by modFactoryAStar.
'==============================================================================

Public Const LAYOUT_VISUAL_SHEET As String = "Layout Visual"

' Factory bounds in the same units as Stations / Obstacles coordinates.
Private Const FACTORY_MIN_X As Double = 0#
Private Const FACTORY_MIN_Y As Double = 0#
Private Const FACTORY_MAX_X As Double = 100#
Private Const FACTORY_MAX_Y As Double = 100#

' Drawing scale: Excel points per factory coordinate unit.
Private Const SCALE_PT As Double = 6#
' Top-left corner of the drawn factory on the sheet (points).
Private Const ORIGIN_LEFT_PT As Double = 36#
Private Const ORIGIN_TOP_PT As Double = 36#
' Operator-point marker size (points).
Private Const DOT_SIZE_PT As Double = 8#

Private Const COLOR_OUTLINE_R As Long = 102
Private Const COLOR_OUTLINE_G As Long = 102
Private Const COLOR_OUTLINE_B As Long = 102

Private Type RectRec
    Name As String
    BottomX As Double
    BottomY As Double
    TopX As Double
    TopY As Double
    OpX As Double
    OpY As Double
    HasOp As Boolean
End Type

'------------------------------------------------------------------------------
' Public entry point — wipe all shapes, then redraw the layout.
'------------------------------------------------------------------------------
Public Sub DrawFactoryLayout()
    Dim ws As Worksheet
    Dim equipment() As RectRec
    Dim obstacles() As RectRec
    Dim equipCount As Long
    Dim obstacleCount As Long
    Dim i As Long
    Dim errNum As Long
    Dim errDesc As String

    OptimizeExcel True
    On Error GoTo CleanUp

    Set ws = GetOrCreateLayoutSheet()
    ClearAllShapes ws

    equipCount = ReadEquipment(equipment)
    obstacleCount = ReadObstacleRects(obstacles)

    DrawFactoryFloor ws

    For i = 1 To equipCount
        DrawWorldRect ws, equipment(i), RGB(70, 130, 200), "Equip_" & SanitizeName(equipment(i).Name, i)
    Next i

    For i = 1 To obstacleCount
        DrawWorldRect ws, obstacles(i), RGB(200, 60, 60), "Obs_" & SanitizeName(obstacles(i).Name, i)
    Next i

    For i = 1 To equipCount
        If equipment(i).HasOp Then
            DrawStationDot ws, equipment(i).OpX, equipment(i).OpY, _
                "Station_" & SanitizeName(equipment(i).Name, i)
        End If
    Next i

    ws.Activate

CleanUp:
    errNum = Err.Number
    errDesc = Err.Description
    OptimizeExcel False
    If errNum <> 0 Then
        MsgBox "DrawFactoryLayout failed:" & vbCrLf & errDesc, vbExclamation
    End If
End Sub

'==============================================================================
' Clear / sheet helpers
'==============================================================================
Private Sub ClearAllShapes(ByVal ws As Worksheet)
    Dim i As Long

    ' Delete every shape, chart, button, etc. so nothing accumulates.
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Private Function GetOrCreateLayoutSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LAYOUT_VISUAL_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        On Error Resume Next
        ws.Name = LAYOUT_VISUAL_SHEET
        On Error GoTo 0
    End If

    Set GetOrCreateLayoutSheet = ws
End Function

'==============================================================================
' Data readers (Stations = equipment + operator dots; Obstacles = red blocks)
'==============================================================================
Private Function ReadEquipment(ByRef equipment() As RectRec) As Long
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
        ReadEquipment = 0
        Exit Function
    End If

    ReDim equipment(1 To lastRow - STATIONS_DATA_START_ROW + 1)
    n = 0

    For r = STATIONS_DATA_START_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(r, STATIONS_COL_NAME).Value2))) = 0 Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_BOTTOM_X).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_BOTTOM_Y).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_TOP_X).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, STATIONS_COL_TOP_Y).Value2) Then GoTo NextRow

        n = n + 1
        bx = CDbl(ws.Cells(r, STATIONS_COL_BOTTOM_X).Value2)
        by = CDbl(ws.Cells(r, STATIONS_COL_BOTTOM_Y).Value2)
        tx = CDbl(ws.Cells(r, STATIONS_COL_TOP_X).Value2)
        ty = CDbl(ws.Cells(r, STATIONS_COL_TOP_Y).Value2)

        equipment(n).Name = CStr(ws.Cells(r, STATIONS_COL_NAME).Value2)
        equipment(n).BottomX = MinD(bx, tx)
        equipment(n).BottomY = MinD(by, ty)
        equipment(n).TopX = MaxD(bx, tx)
        equipment(n).TopY = MaxD(by, ty)

        If IsNumeric(ws.Cells(r, STATIONS_COL_OP_X).Value2) And IsNumeric(ws.Cells(r, STATIONS_COL_OP_Y).Value2) Then
            equipment(n).OpX = CDbl(ws.Cells(r, STATIONS_COL_OP_X).Value2)
            equipment(n).OpY = CDbl(ws.Cells(r, STATIONS_COL_OP_Y).Value2)
            equipment(n).HasOp = True
        Else
            equipment(n).HasOp = False
        End If
NextRow:
    Next r

    If n > 0 Then ReDim Preserve equipment(1 To n)
    ReadEquipment = n
End Function

Private Function ReadObstacleRects(ByRef obstacles() As RectRec) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim n As Long
    Dim bx As Double
    Dim by As Double
    Dim tx As Double
    Dim ty As Double

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(OBSTACLES_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        ReadObstacleRects = 0
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, OBSTACLES_COL_NAME).End(xlUp).Row
    If lastRow < OBSTACLES_DATA_START_ROW Then
        ReadObstacleRects = 0
        Exit Function
    End If

    ReDim obstacles(1 To lastRow - OBSTACLES_DATA_START_ROW + 1)
    n = 0

    For r = OBSTACLES_DATA_START_ROW To lastRow
        If Not IsNumeric(ws.Cells(r, OBSTACLES_COL_BOTTOM_X).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, OBSTACLES_COL_BOTTOM_Y).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, OBSTACLES_COL_TOP_X).Value2) Then GoTo NextRow
        If Not IsNumeric(ws.Cells(r, OBSTACLES_COL_TOP_Y).Value2) Then GoTo NextRow

        n = n + 1
        bx = CDbl(ws.Cells(r, OBSTACLES_COL_BOTTOM_X).Value2)
        by = CDbl(ws.Cells(r, OBSTACLES_COL_BOTTOM_Y).Value2)
        tx = CDbl(ws.Cells(r, OBSTACLES_COL_TOP_X).Value2)
        ty = CDbl(ws.Cells(r, OBSTACLES_COL_TOP_Y).Value2)

        If Len(Trim$(CStr(ws.Cells(r, OBSTACLES_COL_NAME).Value2))) = 0 Then
            obstacles(n).Name = "Obstacle" & CStr(n)
        Else
            obstacles(n).Name = CStr(ws.Cells(r, OBSTACLES_COL_NAME).Value2)
        End If

        obstacles(n).BottomX = MinD(bx, tx)
        obstacles(n).BottomY = MinD(by, ty)
        obstacles(n).TopX = MaxD(bx, tx)
        obstacles(n).TopY = MaxD(by, ty)
        obstacles(n).HasOp = False
NextRow:
    Next r

    If n > 0 Then ReDim Preserve obstacles(1 To n)
    ReadObstacleRects = n
End Function

'==============================================================================
' Drawing
'==============================================================================
Private Sub DrawFactoryFloor(ByVal ws As Worksheet)
    Dim floorRect As RectRec
    Dim shp As Shape

    floorRect.Name = "Factory"
    floorRect.BottomX = FACTORY_MIN_X
    floorRect.BottomY = FACTORY_MIN_Y
    floorRect.TopX = FACTORY_MAX_X
    floorRect.TopY = FACTORY_MAX_Y

    Set shp = AddWorldRect(ws, floorRect, "FactoryFloor")
    StyleFilledRect shp, RGB(220, 220, 220), RGB(COLOR_OUTLINE_R, COLOR_OUTLINE_G, COLOR_OUTLINE_B)
    shp.ZOrder msoSendToBack
End Sub

Private Sub DrawWorldRect(ByVal ws As Worksheet, ByRef rec As RectRec, ByVal fillColor As Long, ByVal shapeName As String)
    Dim shp As Shape
    Set shp = AddWorldRect(ws, rec, shapeName)
    StyleFilledRect shp, fillColor, RGB(COLOR_OUTLINE_R, COLOR_OUTLINE_G, COLOR_OUTLINE_B)
End Sub

Private Function AddWorldRect(ByVal ws As Worksheet, ByRef rec As RectRec, ByVal shapeName As String) As Shape
    Dim leftPt As Double
    Dim topPt As Double
    Dim widthPt As Double
    Dim heightPt As Double
    Dim shp As Shape

    WorldRectToSheet rec, leftPt, topPt, widthPt, heightPt
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPt, topPt, _
        MaxD(widthPt, 1#), MaxD(heightPt, 1#))
    shp.Name = shapeName
    Set AddWorldRect = shp
End Function

Private Sub DrawStationDot(ByVal ws As Worksheet, ByVal opX As Double, ByVal opY As Double, ByVal shapeName As String)
    Dim leftPt As Double
    Dim topPt As Double
    Dim shp As Shape

    leftPt = WorldXToLeft(opX) - (DOT_SIZE_PT / 2#)
    topPt = WorldYToTop(opY) - (DOT_SIZE_PT / 2#)

    Set shp = ws.Shapes.AddShape(msoShapeOval, leftPt, topPt, DOT_SIZE_PT, DOT_SIZE_PT)
    shp.Name = shapeName
    shp.Fill.Visible = msoTrue
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = RGB(0, 0, 0)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(0, 0, 0)
    shp.Line.Weight = 0.5
    shp.ZOrder msoBringToFront
End Sub

Private Sub StyleFilledRect(ByVal shp As Shape, ByVal fillColor As Long, ByVal lineColor As Long)
    shp.Fill.Visible = msoTrue
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = lineColor
    shp.Line.Weight = 0.75
End Sub

' World coords: origin at factory bottom-left, Y up.
' Sheet shapes: origin at top-left, Y down.
Private Sub WorldRectToSheet( _
    ByRef rec As RectRec, _
    ByRef leftPt As Double, _
    ByRef topPt As Double, _
    ByRef widthPt As Double, _
    ByRef heightPt As Double)

    leftPt = WorldXToLeft(rec.BottomX)
    topPt = WorldYToTop(rec.TopY)
    widthPt = (rec.TopX - rec.BottomX) * SCALE_PT
    heightPt = (rec.TopY - rec.BottomY) * SCALE_PT
End Sub

Private Function WorldXToLeft(ByVal worldX As Double) As Double
    WorldXToLeft = ORIGIN_LEFT_PT + (worldX - FACTORY_MIN_X) * SCALE_PT
End Function

Private Function WorldYToTop(ByVal worldY As Double) As Double
    WorldYToTop = ORIGIN_TOP_PT + (FACTORY_MAX_Y - worldY) * SCALE_PT
End Function

Private Function SanitizeName(ByVal rawName As String, ByVal fallbackIndex As Long) As String
    Dim s As String
    Dim i As Long
    Dim ch As String

    s = Trim$(rawName)
    If Len(s) = 0 Then
        SanitizeName = "Item" & CStr(fallbackIndex)
        Exit Function
    End If

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case Asc(ch)
            Case 48 To 57, 65 To 90, 97 To 122, 95
                ' keep
            Case Else
                Mid$(s, i, 1) = "_"
        End Select
    Next i

    If Len(s) > 20 Then s = Left$(s, 20)
    SanitizeName = s
End Function

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function
