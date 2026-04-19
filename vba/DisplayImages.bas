Option Explicit

' ============================================================================
' Module:      PlotLinkModule
' Purpose:     Refreshable plot display for =PlotLink(...) cells using
'              pixel-based sizing and offsets.
'
' Design:
'              - PlotLink(...) returns a plot path only.
'              - Optional arguments let the user request display width/height
'                and offsets in pixels.
'              - On recalculation, event code refreshes the corresponding image.
'              - Refresh scans the worksheet UsedRange for formulas beginning
'                with =PlotLink(
'
' PlotLink signature:
'   PlotLink(plotPath, [plotWidthPx], [plotHeightPx], [topOffsetPx], [leftOffsetPx])
'
' Example:
'   =PlotLink(RPlot("plot(1:10)","BasicPlot",800,600,A1:A10), 900, 500, 25, 0)
'
' Default display:
'   Width       = 800 px
'   Height      = 500 px
'   Top offset  = 20 px
'   Left offset = 0 px
' ============================================================================

Private Const DEFAULT_PLOT_WIDTH_PX As Double = 800
Private Const DEFAULT_PLOT_HEIGHT_PX As Double = 500
Private Const DEFAULT_TOP_OFFSET_PX As Double = 20
Private Const DEFAULT_LEFT_OFFSET_PX As Double = 0

Private Const MIN_PLOT_WIDTH_PX As Double = 120
Private Const MIN_PLOT_HEIGHT_PX As Double = 120

Private Type PlotLinkDisplaySpec
    WidthPx As Double
    HeightPx As Double
    TopOffsetPx As Double
    LeftOffsetPx As Double
End Type

Public Function PlotLink( _
    ByVal plotPath As Variant, _
    Optional ByVal plotWidthPx As Variant, _
    Optional ByVal plotHeightPx As Variant, _
    Optional ByVal TopOffsetPx As Variant, _
    Optional ByVal LeftOffsetPx As Variant) As Variant

    Dim p As String

    Application.Volatile False

    If IsError(plotPath) Then
        PlotLink = plotPath
        Exit Function
    End If

    p = Trim$(CStr(plotPath))

    If Len(p) = 0 Then
        PlotLink = ""
    Else
        PlotLink = p
    End If

    ' Optional layout arguments are intentionally not used here.
    ' They are read later from the formula text by the refresh code.
End Function

Public Sub RefreshPlotLinksInSheet(ByVal ws As Worksheet)

    Dim candidates As Collection
    Dim item As Variant

    On Error GoTo CleanFail

    If ws Is Nothing Then Exit Sub

    Set candidates = GetPlotLinkCandidateCells(ws)
    If candidates Is Nothing Then Exit Sub
    If candidates.Count = 0 Then Exit Sub

    For Each item In candidates
        RefreshSinglePlotLinkCell ws, item
    Next item

    Exit Sub

CleanFail:
    Debug.Print "RefreshPlotLinksInSheet error: " & Err.Number & " - " & Err.Description
End Sub

Public Sub RefreshSinglePlotLinkCell(ByVal ws As Worksheet, ByVal cell As Range)

    Dim formulaText As String
    Dim plotPath As String
    Dim shapeName As String
    Dim existingTag As String
    Dim layoutTag As String
    Dim desiredTag As String
    Dim spec As PlotLinkDisplaySpec

    On Error GoTo CleanFail

    If ws Is Nothing Then Exit Sub
    If cell Is Nothing Then Exit Sub

    shapeName = PlotLinkShapeName(ws, cell)

    If Not cell.HasFormula Then
        DeleteShapeIfExists ws, shapeName
        Exit Sub
    End If

    formulaText = LCase$(Trim$(cell.Formula))

    If Left$(formulaText, 10) <> "=plotlink(" Then
        DeleteShapeIfExists ws, shapeName
        Exit Sub
    End If

    plotPath = Trim$(CStr(cell.Value))
    plotPath = Replace(plotPath, "/", "\")

    If Len(plotPath) = 0 Or Dir$(plotPath) = "" Then
        DeleteShapeIfExists ws, shapeName
        Exit Sub
    End If

    spec = GetDisplaySpecFromFormula(cell)
    desiredTag = BuildLayoutTag(plotPath, spec)

    existingTag = GetShapeAltText(ws, shapeName)
    layoutTag = GetShapeTitle(ws, shapeName)

    If Not ShapeExists(ws, shapeName) Then
        InsertOrReplacePlotPicture ws, cell, plotPath, shapeName, spec
    ElseIf existingTag <> plotPath Or layoutTag <> desiredTag Then
        InsertOrReplacePlotPicture ws, cell, plotPath, shapeName, spec
    End If

    Exit Sub

CleanFail:
    Debug.Print "RefreshSinglePlotLinkCell error: " & Err.Number & " - " & Err.Description
End Sub

Private Function GetDisplaySpecFromFormula(ByVal formulaCell As Range) As PlotLinkDisplaySpec

    Dim spec As PlotLinkDisplaySpec
    Dim args As Collection

    spec.WidthPx = DEFAULT_PLOT_WIDTH_PX
    spec.HeightPx = DEFAULT_PLOT_HEIGHT_PX
    spec.TopOffsetPx = DEFAULT_TOP_OFFSET_PX
    spec.LeftOffsetPx = DEFAULT_LEFT_OFFSET_PX

    On Error GoTo CleanExit

    Set args = GetTopLevelFormulaArguments(formulaCell.Formula)

    ' PlotLink arguments:
    ' 1 plotPath
    ' 2 plotWidthPx
    ' 3 plotHeightPx
    ' 4 topOffsetPx
    ' 5 leftOffsetPx

    If args.Count >= 2 Then spec.WidthPx = GetNumericArgumentValue(formulaCell, CStr(args(2)), DEFAULT_PLOT_WIDTH_PX)
    If args.Count >= 3 Then spec.HeightPx = GetNumericArgumentValue(formulaCell, CStr(args(3)), DEFAULT_PLOT_HEIGHT_PX)
    If args.Count >= 4 Then spec.TopOffsetPx = GetNumericArgumentValue(formulaCell, CStr(args(4)), DEFAULT_TOP_OFFSET_PX)
    If args.Count >= 5 Then spec.LeftOffsetPx = GetNumericArgumentValue(formulaCell, CStr(args(5)), DEFAULT_LEFT_OFFSET_PX)

    If spec.WidthPx < MIN_PLOT_WIDTH_PX Then spec.WidthPx = MIN_PLOT_WIDTH_PX
    If spec.HeightPx < MIN_PLOT_HEIGHT_PX Then spec.HeightPx = MIN_PLOT_HEIGHT_PX

CleanExit:
    GetDisplaySpecFromFormula = spec
End Function

Private Function GetTopLevelFormulaArguments(ByVal formulaText As String) As Collection

    Dim results As New Collection
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim currentArg As String
    Dim depth As Long
    Dim inQuotes As Boolean
    Dim openPos As Long

    s = Trim$(formulaText)
    openPos = InStr(1, s, "(", vbTextCompare)

    If openPos = 0 Then
        Set GetTopLevelFormulaArguments = results
        Exit Function
    End If

    s = Mid$(s, openPos + 1)

    depth = 0
    inQuotes = False
    currentArg = ""

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch = """" Then
            currentArg = currentArg & ch

            If i < Len(s) And Mid$(s, i + 1, 1) = """" Then
                currentArg = currentArg & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If

        ElseIf Not inQuotes Then

            Select Case ch
                Case "("
                    depth = depth + 1
                    currentArg = currentArg & ch

                Case ")"
                    If depth = 0 Then
                        results.Add Trim$(currentArg)
                        Exit For
                    Else
                        depth = depth - 1
                        currentArg = currentArg & ch
                    End If

                Case ","
                    If depth = 0 Then
                        results.Add Trim$(currentArg)
                        currentArg = ""
                    Else
                        currentArg = currentArg & ch
                    End If

                Case Else
                    currentArg = currentArg & ch
            End Select

        Else
            currentArg = currentArg & ch
        End If
    Next i

    Set GetTopLevelFormulaArguments = results
End Function

Private Function GetNumericArgumentValue( _
    ByVal formulaCell As Range, _
    ByVal argText As String, _
    ByVal defaultValue As Double) As Double

    Dim v As Variant
    Dim s As String

    On Error GoTo UseDefault

    s = Trim$(argText)
    If Len(s) = 0 Then GoTo UseDefault

    If IsNumeric(s) Then
        GetNumericArgumentValue = CDbl(s)
        Exit Function
    End If

    v = formulaCell.Worksheet.Evaluate(s)

    If IsError(v) Then GoTo UseDefault
    If Not IsNumeric(v) Then GoTo UseDefault

    GetNumericArgumentValue = CDbl(v)
    Exit Function

UseDefault:
    GetNumericArgumentValue = defaultValue
End Function

Public Sub InsertOrReplacePlotPicture( _
    ByVal ws As Worksheet, _
    ByVal anchorCell As Range, _
    ByVal plotPath As String, _
    ByVal shapeName As String, _
    ByRef spec As PlotLinkDisplaySpec)

    Dim shp As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim layoutTag As String

    On Error GoTo CleanFail

    plotPath = Trim$(plotPath)
    plotPath = Replace(plotPath, "/", "\")

    If ws Is Nothing Then Exit Sub
    If anchorCell Is Nothing Then Exit Sub
    If Len(plotPath) = 0 Then Exit Sub
    If Dir$(plotPath) = "" Then Exit Sub

    DeleteShapeIfExists ws, shapeName

    leftPos = anchorCell.Left + spec.LeftOffsetPx
    topPos = anchorCell.Top + spec.TopOffsetPx

    Set shp = ws.Shapes.AddPicture( _
        Filename:=plotPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=spec.WidthPx, _
        Height:=spec.HeightPx)

    shp.Name = shapeName
    shp.Placement = xlMoveAndSize
    shp.AlternativeText = plotPath

    layoutTag = BuildLayoutTag(plotPath, spec)
    On Error Resume Next
    shp.Title = layoutTag
    On Error GoTo CleanFail

    Exit Sub

CleanFail:
    Debug.Print "InsertOrReplacePlotPicture error: " & Err.Number & " - " & Err.Description
End Sub

Private Function BuildLayoutTag(ByVal plotPath As String, ByRef spec As PlotLinkDisplaySpec) As String
    Dim fileStamp As String

    fileStamp = GetFileStamp(plotPath)

    BuildLayoutTag = plotPath & "|" & _
                     CStr(spec.WidthPx) & "|" & _
                     CStr(spec.HeightPx) & "|" & _
                     CStr(spec.TopOffsetPx) & "|" & _
                     CStr(spec.LeftOffsetPx) & "|" & _
                     fileStamp
End Function

Private Function GetFileStamp(ByVal plotPath As String) As String

    On Error GoTo Fail

    If Len(plotPath) = 0 Then
        GetFileStamp = ""
        Exit Function
    End If

    If Dir$(plotPath) = "" Then
        GetFileStamp = ""
        Exit Function
    End If

    GetFileStamp = Format$(FileDateTime(plotPath), "yyyy-mm-dd hh:nn:ss")
    Exit Function

Fail:
    GetFileStamp = ""
End Function

Public Function GetPlotLinkCandidateCells(ByVal ws As Worksheet) As Collection

    Dim results As New Collection
    Dim formulaCells As Range
    Dim area As Range
    Dim cell As Range
    Dim f As String

    On Error GoTo CleanExit

    If ws Is Nothing Then
        Set GetPlotLinkCandidateCells = results
        Exit Function
    End If

    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo CleanExit

    If formulaCells Is Nothing Then
        Set GetPlotLinkCandidateCells = results
        Exit Function
    End If

    For Each area In formulaCells.Areas
        For Each cell In area.Cells
            f = Trim$(cell.Formula)
            If InStr(1, f, "=PlotLink(", vbTextCompare) = 1 Then
                results.Add cell
            End If
        Next cell
    Next area

CleanExit:
    Set GetPlotLinkCandidateCells = results
End Function

Public Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Public Function ShapeExists(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    Dim shp As Shape

    On Error GoTo NotFound
    Set shp = ws.Shapes(shapeName)
    ShapeExists = True
    Exit Function

NotFound:
    ShapeExists = False
End Function

Public Function GetShapeAltText(ByVal ws As Worksheet, ByVal shapeName As String) As String
    On Error GoTo NotFound
    GetShapeAltText = ws.Shapes(shapeName).AlternativeText
    Exit Function

NotFound:
    GetShapeAltText = ""
End Function

Public Function GetShapeTitle(ByVal ws As Worksheet, ByVal shapeName As String) As String
    On Error GoTo NotFound
    GetShapeTitle = ws.Shapes(shapeName).Title
    Exit Function

NotFound:
    GetShapeTitle = ""
End Function

Public Function PlotLinkShapeName(ByVal ws As Worksheet, ByVal c As Range) As String
    PlotLinkShapeName = "PlotLink_" & CleanShapeToken(ws.Name) & "_" & c.Row & "_" & c.Column
End Function

Public Function CleanShapeToken(ByVal s As String) As String

    Dim t As String

    t = s
    t = Replace(t, " ", "_")
    t = Replace(t, ".", "_")
    t = Replace(t, ":", "_")
    t = Replace(t, "\", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, "[", "_")
    t = Replace(t, "]", "_")

    CleanShapeToken = t
End Function

Public Sub RefreshPlotLinksActiveSheet()
    If Not ActiveSheet Is Nothing Then
        RefreshPlotLinksInSheet ActiveSheet
    End If
End Sub

Public Sub RefreshPlotLinksAllWorksheets()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        RefreshPlotLinksInSheet ws
    Next ws
End Sub

