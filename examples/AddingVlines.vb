Sub Macro2()
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.AddChart2(297, xlBarStacked).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$2:$A$7,Sheet1!$B$2:$C$7")
    Range("P23").Select
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveSheet.ChartObjects("Chart 2").Activate
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "=""Average"""
    ActiveChart.FullSeriesCollection(3).Values = "=Sheet1!$D$2:$D$3"
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.FullSeriesCollection(3).ChartType = xlXYScatterLinesNoMarkers
    ActiveSheet.ChartObjects("Chart 2").Activate
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(3).XValues = "=Sheet1!$D$2:$D$3"
    ActiveChart.FullSeriesCollection(3).Values = "=Sheet1!$E$2:$E$3"
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.Axes(xlValue, xlSecondary).Select
    ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = 1
    Selection.TickLabelPosition = xlNone
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.FullSeriesCollection(3).Points(2).Select
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 1.5
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .DashStyle = msoLineSysDash
    End With
    Range("M10").Select
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.FullSeriesCollection(3).Points(2).Select
    ActiveChart.FullSeriesCollection(3).Points(2).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.Position = xlLabelPositionAbove
    ActiveChart.ChartTitle.Select
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    ActiveChart.FullSeriesCollection(3).Points(2).DataLabel.Select
    ActiveChart.SeriesCollection(3).DataLabels(2).Format.TextFrame2.TextRange. _
        Characters.Text = ""
    ActiveChart.SeriesCollection(3).DataLabels(2).Format.TextFrame2.TextRange. _
        InsertAfter "Something" & Chr(13) & ""
    With ActiveChart.SeriesCollection(3).DataLabels(2).Format.TextFrame2.TextRange. _
        Characters(1, 10).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(64, 64, 64)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 9
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    With ActiveChart.SeriesCollection(3).DataLabels(2).Format.TextFrame2.TextRange. _
        Characters(1, 10).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    ActiveChart.FullSeriesCollection(3).Points(2).DataLabel.Select
    Selection.Format.TextFrame2.TextRange.Font.Italic = msoTrue
    ActiveChart.SeriesCollection(3).DataLabels(2).Format.TextFrame2.TextRange. _
        Characters.Text = ""
    ActiveChart.SeriesCollection(3).DataLabels(2).Format.TextFrame2.TextRange. _
        InsertAfter "Something" & Chr(13) & ""
    With ActiveChart.SeriesCollection(3).DataLabels(2).Format.TextFrame2.TextRange. _
        Characters(1, 10).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(64, 64, 64)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 9
        .Italic = msoTrue
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    With ActiveChart.SeriesCollection(3).DataLabels(2).Format.TextFrame2.TextRange. _
        Characters(1, 10).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Application.CommandBars("Format Object").Visible = False
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    Range("M11").Select
End Sub
