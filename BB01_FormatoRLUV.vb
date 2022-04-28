Module BB01_FormatoRLUV
    Public Sub formatoTitulos()
        Dim nCol As Integer = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim rangoTitulos As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1), Globals.ThisAddIn.Application.ActiveSheet.Cells(1, nCol))

        With rangoTitulos

            .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
            .MergeCells = False

            .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
            .Font.TintAndShade = -0.249977111117893

            .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
            .Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
            .Interior.TintAndShade = -0.249977111117893
            .Interior.PatternTintAndShade = 0

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

        End With

        With Globals.ThisAddIn.Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        Globals.ThisAddIn.Application.ActiveWindow.FreezePanes = True

    End Sub

    Public Sub formatoObservaciones()
        Dim nCol As Integer = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
        Dim nFil As Long = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim rangoObs As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(2, 1), Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil, nCol))
        Dim rangoNumObs As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(2, 1), Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil, 1))


        With rangoObs
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .NumberFormat = "#,##0.00"
        End With

        rangoNumObs.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
        rangoNumObs.NumberFormat = "#,##0"
    End Sub

    Public Sub formatoPromTls()

        Dim nCol As Integer = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
        Dim nFil As Long = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row - 1
        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim rangoTls As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil, 1),
                                                       Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil, nCol))
        Dim rangoProm As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil + 1, 1),
                                                          Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil + 1, 3))

        With rangoTls
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
            .MergeCells = False

            .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
            .Font.TintAndShade = -0.249977111117893
            .Font.Bold = True

            .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
            .Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
            .Interior.TintAndShade = -0.249977111117893
            .Interior.PatternTintAndShade = 0

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            .NumberFormat = "#,##0.00"
        End With

        With rangoProm
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
            .MergeCells = False

            .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
            .Font.TintAndShade = -0.249977111117893
            .Font.Bold = True

            .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
            .Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
            .Interior.TintAndShade = -0.249977111117893
            .Interior.PatternTintAndShade = 0

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium

            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = 0
            .Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            .NumberFormat = "#,##0.00"
        End With

    End Sub

    Public Sub ajustarAnchoCol()
        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        hojaActiva.Columns("A:S").ColumnWidth = 16.71
    End Sub
End Module
