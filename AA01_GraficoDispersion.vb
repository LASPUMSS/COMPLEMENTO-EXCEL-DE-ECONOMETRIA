Module AA01_GraficoDispersion

    Public Sub graficoDispersionMetodo(ByVal sigY As String, ByVal sigX As String, ByVal rangoY As String, ByVal rangoX As String, ByVal nombreHoja As String)

        Dim tituloGrafico As String

        tituloGrafico = "Grafíco de Dispersión, Explicar: " & sigY & " Mediante: " & sigX

        Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddChart2(240, Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter).Select()
        Globals.ThisAddIn.Application.ActiveChart.SeriesCollection.NewSeries
        Globals.ThisAddIn.Application.ActiveChart.FullSeriesCollection(1).Name = "=""" & tituloGrafico & """"
        Globals.ThisAddIn.Application.ActiveChart.FullSeriesCollection(1).XValues = "=" & nombreHoja & "!" & rangoX
        Globals.ThisAddIn.Application.ActiveChart.FullSeriesCollection(1).Values = "=" & nombreHoja & "!" & rangoY


        Globals.ThisAddIn.Application.ActiveChart.FullSeriesCollection(1).Trendlines.Add
        Globals.ThisAddIn.Application.ActiveChart.FullSeriesCollection(1).Trendlines(1).Select
        Globals.ThisAddIn.Application.Selection.DisplayEquation = True
        Globals.ThisAddIn.Application.Selection.DisplayRSquared = True

        Globals.ThisAddIn.Application.ActiveChart.ClearToMatchStyle()
        Globals.ThisAddIn.Application.ActiveChart.ChartStyle = 248

        Globals.ThisAddIn.Application.ActiveChart.Location(Where:=Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsNewSheet)

        Globals.ThisAddIn.Application.ActiveChart.ChartArea.Select()
        Globals.ThisAddIn.Application.ActiveChart.FullSeriesCollection(1).Trendlines(1).DataLabel.Select
        Globals.ThisAddIn.Application.Selection.Format.TextFrame2.TextRange.Font.Size = 18
        Globals.ThisAddIn.Application.Selection.Left = 257.99
        Globals.ThisAddIn.Application.Selection.Top = 47.697

        Globals.ThisAddIn.Application.ActiveChart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 32
    End Sub

End Module
