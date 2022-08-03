Imports Microsoft.Office.Interop
Module AA01_GraficoDispersion

    Public Sub graficoDispersionMetodo(ByVal sigY As String, ByVal sigX As String, ByVal rangoY As String, ByVal rangoX As String, ByVal nombreHoja As String)
        Dim tituloGrafico As String
        tituloGrafico = $"Grafíco de Dispersión, Explicar: {sigY} Mediante: {sigX}"

        With Globals.ThisAddIn.Application
            .ActiveSheet.Shapes.AddChart2(240, Excel.XlChartType.xlXYScatter).Select() 'Type of chart
            .ActiveChart.SeriesCollection.NewSeries
            .ActiveChart.FullSeriesCollection(1).Name = $"=""{tituloGrafico}"""
            .ActiveChart.FullSeriesCollection(1).XValues = $"={nombreHoja}!{rangoX}"
            .ActiveChart.FullSeriesCollection(1).Values = $"={nombreHoja}!{rangoY}"

            .ActiveChart.FullSeriesCollection(1).Trendlines.Add
            .ActiveChart.FullSeriesCollection(1).Trendlines(1).Select
            .Selection.DisplayEquation = True
            .Selection.DisplayRSquared = True

            .ActiveChart.ClearToMatchStyle()
            .ActiveChart.ChartStyle = 248

            .ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsNewSheet)

            .ActiveChart.ChartArea.Select()
            .ActiveChart.FullSeriesCollection(1).Trendlines(1).DataLabel.Select
            .Selection.Format.TextFrame2.TextRange.Font.Size = 18
            .Selection.Left = 257.99
            .Selection.Top = 47.697

            .ActiveChart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 32
        End With
    End Sub

End Module
