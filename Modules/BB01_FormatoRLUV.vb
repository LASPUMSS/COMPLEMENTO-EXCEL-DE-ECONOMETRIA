Imports Microsoft.Office.Interop
Module BB01_FormatoRLUV
    Public Sub formatoTitulos()
        Dim nCol As Integer = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight).Column
        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim rangoTitulos As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1), Globals.ThisAddIn.Application.ActiveSheet.Cells(1, nCol))

        With rangoTitulos

            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False

            .Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .Font.TintAndShade = -0.249977111117893

            .Interior.Pattern = Excel.Constants.xlSolid
            .Interior.PatternColorIndex = Excel.Constants.xlAutomatic
            .Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
            .Interior.TintAndShade = -0.249977111117893
            .Interior.PatternTintAndShade = 0

            .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeTop).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeRight).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlInsideVertical).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlInsideVertical).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlMedium

        End With

        With Globals.ThisAddIn.Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        Globals.ThisAddIn.Application.ActiveWindow.FreezePanes = True

    End Sub

    Public Sub formatoObservaciones()
        Dim nCol As Integer = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight).Column
        Dim nFil As Long = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlDown).Row
        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim rangoObs As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(2, 1), Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil, nCol))
        Dim rangoNumObs As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(2, 1), Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil, 1))


        With rangoObs
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeTop).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeRight).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlInsideVertical).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlInsideVertical).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlMedium

            .NumberFormat = "#,##0.00"
        End With

        rangoNumObs.HorizontalAlignment = Excel.Constants.xlCenter
        rangoNumObs.NumberFormat = "#,##0"
    End Sub

    Public Sub formatoPromTls()

        Dim nCol As Integer = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight).Column
        Dim nFil As Long = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlDown).Row - 1
        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim rangoTls As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil, 1),
                                                       Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil, nCol))
        Dim rangoProm As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil + 1, 1),
                                                          Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil + 1, 3))

        With rangoTls
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False

            .Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .Font.TintAndShade = -0.249977111117893
            .Font.Bold = True

            .Interior.Pattern = Excel.Constants.xlSolid
            .Interior.PatternColorIndex = Excel.Constants.xlAutomatic
            .Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
            .Interior.TintAndShade = -0.249977111117893
            .Interior.PatternTintAndShade = 0

            .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeTop).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeRight).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlInsideVertical).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlInsideVertical).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlMedium
            .NumberFormat = "#,##0.00"
        End With

        With rangoProm
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False

            .Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            .Font.TintAndShade = -0.249977111117893
            .Font.Bold = True

            .Interior.Pattern = Excel.Constants.xlSolid
            .Interior.PatternColorIndex = Excel.Constants.xlAutomatic
            .Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
            .Interior.TintAndShade = -0.249977111117893
            .Interior.PatternTintAndShade = 0

            .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeTop).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeRight).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlInsideVertical).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlInsideVertical).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium

            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = Excel.Constants.xlAutomatic
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = 0
            .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlMedium
            .NumberFormat = "#,##0.00"
        End With

    End Sub

    Public Sub ajustarAnchoCol()
        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        hojaActiva.Columns("A:S").ColumnWidth = 16.71
    End Sub
End Module
