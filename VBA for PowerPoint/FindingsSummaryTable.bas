Attribute VB_Name = "FindingsSummaryTable"
Function ParseAllSlidesForFindings()
    Dim fnds As Object
    Dim fnd As finding
    Dim numberOfFindings As Integer
    Set fnds = CreateObject("System.Collections.ArrayList")
    With Application.ActivePresentation
        For I = 1 To .Slides.Count
             Set fnd = ParseSlideForFindings((I), numberOfFindings)
             If fnd.ready Then
                numberOfFindings = numberOfFindings + 1
                fnd.number = numberOfFindings
                fnds.Add fnd
                'FixIndexInFindingTitle ((I))
             End If
        Next I
    End With
    Set ParseAllSlidesForFindings = fnds
End Function
'Sub FixIndexInFindingTitleInSlide(slideNumber As I)
'
'End Sub
'Function FormatFindingsIndex(ByVal index As Integer) As String
'    If inputNumber < 1 Or inputNumber > 99 Then
'        Err.Raise number:=vbObjectError + 1000, description:="Input must be between 1 and 99 (inclusive)."
'    Else
'        FormatFindingsIndex = Format(inputNumber, "00")
'    End If
'End Function
Function ParseSlideForFindings(slideNum As Long, numberOfFindings As Integer)
    Dim result As finding
    Dim t As String
    Dim pptShape As shape
    Set result = New finding
    Set ParseSlideForFindings = result
    With ActivePresentation.Slides(slideNum)
        For Each pptShape In .Shapes
            If pptShape.HasTextFrame Then
                t = pptShape.TextFrame.TextRange.text
                Dim foundTitle As Boolean
                If Not result.titleReady Then
                    result.parseTitleForFinding pptShape, numberOfFindings
                    foundTitle = result.titleReady
                End If
                If Not result.bodyReady And Not foundTitle Then
                    result.parseBodyForFinding (t)
                End If
                If result.titleReady And result.bodyReady Then
                    result.ready = True
                    Exit Function
                End If
            End If
        Next
    End With
End Function

Public Function ParseFindingsTitle(text As String)
    Set fndng = FindingText.parseTitleForFinding((text))
    Set ParseFindingsTitle = fndng
    Exit Function
End Function

Public Sub BuildFindingsSummaryTable()
    Dim fnds As Object
    Set fnds = ParseAllSlidesForFindings
    MakeSlideWithFindingsTable 1, fnds
End Sub
Function cm2pt(cm As Double)
    cm2pt = cm / 0.035278
End Function
Sub MakeSlideWithFindingsTable(slideNumber As Integer, fnds As Object)
    Set pptPres = Application.ActivePresentation
    Set pptSlide = pptPres.Slides.Add(slideNumber, ppLayoutTitleOnly)
    ActiveWindow.View.GotoSlide slideNumber
    
    With pptSlide.Shapes.Title
        With .TextFrame.TextRange
            .text = "Summary of Findings and Recommendations"
            .Font.Size = 20
        End With
    End With
    
    numRows = fnds.Count + 2
    numColumns = 8
    headingColour = "#E76500"
    
    'Set titleShape = pptSlid.Shapes.AddTextbox()
    
    Set pptShape = pptSlide.Shapes.AddTable(numRows, numColumns, _
        cm2pt(1.4), _
        cm2pt(4), _
        cm2pt(31.076), 100)
        
    Set subtitleShape = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, cm2pt(1.4), cm2pt(2.24), cm2pt(30), cm2pt(1))
    With subtitleShape.TextFrame.TextRange
        .text = "Note: the findings & recommendations presented in this section will be synchronised with the Issues & Actions manager on SAP Cloud ALM."
        .Font.Size = 11
    End With
    
        
    
    columnWidths = Array( _
        4.25, _
        1, _
        9.2, _
        9.2, _
        0.58, _
        2, _
        2, _
        2 _
    )
    For I = LBound(columnWidths) To UBound(columnWidths)
        columnWidths(I) = cm2pt((columnWidths(I)))
    Next I
    
    For I = LBound(columnWidths) To UBound(columnWidths)
        pptShape.Table.Columns(I + 1).Width = columnWidths(I)
    Next I
    
    For I = 1 To 5
        Set cell1 = pptShape.Table.Cell(1, I)
        Set cell2 = pptShape.Table.Cell(2, I)
        cell1.Merge (cell2)
    Next I
    
    With pptShape.Table
        .Cell(1, 6).Merge (.Cell(1, 7))
        .Cell(1, 6).Merge (.Cell(1, 8))
    End With
    
    headers = Array( _
        "Category", _
        "#", _
        "Finding", _
        "Description", _
        "Priority", _
        "Effort" _
    )
    
    effortHeaders = Array( _
        "Short" & vbNewLine & "(< 2 mo.)", _
        "Medium" & vbNewLine & "(2-6 mo.)", _
        "Long" & vbNewLine & "(> 6 mo.)" _
    )
    effortColors = Array( _
        Ratings.ColorRGB(Low), _
        Ratings.ColorRGB(Medium), _
        Ratings.ColorRGB(High) _
    )
    
    With pptShape.Table
        For row = 1 To .Rows.Count
            For col = 1 To .Columns.Count
                Set Cell = .Cell(row, col)
                Cell.shape.TextFrame.VerticalAnchor = msoAnchorMiddle
                Cell.shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                With Cell.shape.TextFrame.TextRange.Font
                    .Size = 12
                End With
            Next col
        Next row
    End With
    
    For I = LBound(headers) To UBound(headers)
        With pptShape.Table.Cell(1, I + 1)
            With .shape.TextFrame
                .VerticalAnchor = msoAnchorBottom
                With .TextRange
                    .text = headers(I)
                    .ParagraphFormat.Alignment = ppAlignLeft
                    With .Font
                        .Color.rgb = rgb(203, 101, 0)
                        .Size = 12
                    End With
                End With
            End With
        End With
    Next I
    For I = 6 To 8
        With pptShape.Table.Cell(2, I)
            With .Borders(ppBorderTop)
                .Weight = 0.75
                .Style = msoLineSolid
                .ForeColor.rgb = rgb(0, 0, 0)
                .Visible = msoTrue
            End With
            With .shape.TextFrame.TextRange.ParagraphFormat
                .Alignment = msoAlignCenter
            End With
            With .shape.TextFrame.TextRange
                efforti = I - 6
                .text = effortHeaders(efforti)
                With .Font
                    .Size = 9
                    .Color.rgb = effortColors(efforti)
                End With
            End With
        End With
    Next I
    
    With pptShape.Table.Cell(1, 2)
        .shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    End With
    With pptShape.Table.Cell(1, 5)
        .shape.TextFrame.Orientation = msoTextOrientationUpward
        .shape.TextFrame.VerticalAnchor = msoAnchorMiddle
        .shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    End With
    With pptShape.Table.Cell(1, 6)
        .shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    End With
    
    With pptShape.Table
        For row = 0 To fnds.Count - 1
            Dim fnd As finding
            Set fnd = fnds(row)
            With .Cell(row + 3, 1).shape
                rgbc = SOMCategories.ColorRGB(fnd.categoryText)
                .Fill.ForeColor.rgb = rgb(rgbc(0), rgbc(1), rgbc(2))
                With .TextFrame.TextRange
                    .text = fnd.categoryText
                    .Font.Size = 10
                    .Font.Color.rgb = rgb(255, 255, 255)
                End With
            End With
            With .Cell(row + 3, 2).shape.TextFrame.TextRange
                .text = fnds(row).numberText
                .Font.Size = 10
            End With
            With .Cell(row + 3, 3).shape.TextFrame.TextRange
                .text = fnd.issueTitle
                .Font.Size = 10
                .ParagraphFormat.Alignment = ppAlignLeft
            End With
            With .Cell(row + 3, 5).shape
                .Fill.ForeColor.rgb = Ratings.ColorRGB(fnd.priority)
            End With
            
            effortColumn = 5 + fnd.effort
            With .Cell(row + 3, effortColumn).shape.TextFrame
                .VerticalAnchor = msoAnchorMiddle
                With .TextRange
                    .text = ChrW(&H2713)
                    .Font.Name = "Arial"
                    .Font.Size = 12
                    .Font.Color.rgb = Ratings.ColorRGB(Low)
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = ppAlignCenter
                End With
            End With
            With .Cell(row + 3, 4).shape.TextFrame.TextRange
                .text = fnd.description
                .Font.Size = 10
                .ParagraphFormat.Alignment = ppAlignLeft
            End With
        Next row
        
        For row = 0 To fnds.Count - 1
            For col = 1 To .Columns.Count
                With .Cell(row + 3, col).Borders(ppBorderTop)
                    .Weight = 0.75
                    .Style = msoLineSolid
                    .ForeColor.rgb = rgb(0, 0, 0)
                    .Visible = msoTrue
                End With
            Next col
        Next row
    End With
    
    If fnds.Count > 1 Then
        For I = 3 To 3 + fnds.Count - 2
            With pptShape.Table
                Set cell1 = .Cell(I, 1)
                Set cell2 = .Cell(I + 1, 1)
                If fnds(I - 3).categoryText = fnds(I - 2).categoryText Then
                    cell1.Merge (cell2)
                    cell1.shape.TextFrame.TextRange.text = fnds(I - 3).categoryText
                End If
            End With
        Next I
    End If
    
End Sub
Public Function HexToRgb(hexColor As String) As String

    Dim red As String
    Dim green As String
    Dim blue As String

    red = Left(hexColor, 2)
    green = Mid(hexColor, 3, 2)
    blue = Right(hexColor, 2)

    HexToRgb = blue & green & red

End Function
