Module BB02_FormatoRLVV
    Public Sub AJUSTE_01()
        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim n2 As Long
            Dim n3 As Long

            n = (.Cells(1, 2).Value) / 2
            n2 = (.Cells(2, 6).Value) / 2
            n3 = .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Column

            With .Cells(1, 1)
                .Value = "N° FILAS Y"
                .Font.Bold = True
                .Font.Size = 8
            End With
            With .Cells(1, 2)
                .Font.Bold = True
                .Font.Size = 11
            End With
            With .Cells(2, 1)
                .Value = "N° COLUMNAS Y"
                .Font.Bold = True
                .Font.Size = 8
            End With
            With .Cells(2, 2)
                .Font.Bold = True
                .Font.Size = 11
            End With

            With .Cells(1, 5)
                .Value = "N° FILAS X"
                .Font.Bold = True
                .Font.Size = 8
            End With
            With .Cells(1, 6)
                .Font.Bold = True
                .Font.Size = 11
            End With
            With .Cells(2, 5)
                .Value = "N° COLUMNAS X"
                .Font.Bold = True
                .Font.Size = 8
            End With
            With .Cells(2, 6)
                .Font.Bold = True
                .Font.Size = 11
            End With

            .Cells(4, 2).Select
            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 44
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .Cells(n, 1)
                .Value = " Y ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With

            .Cells(4, 5).Select
            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 43
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .Cells(n, 4)
                .Value = " X ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With

            .Cells(4, n3).Select
            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .Cells(n2 + 4, n3 - 1)
                .Value = "XT ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With
            With .Cells(n2 + 4, n3 - 1).Characters(Start:=2, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 11
                .Strikethrough = False
                .Superscript = True
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With

            .Cells(4, n3).Select
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n2, -1)
                .Value = "XT*X ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With
            With .ActiveCell.Offset(n2, -1).Characters(Start:=2, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 11
                .Strikethrough = False
                .Superscript = True
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With

            .Cells(4, n3).Select
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n2, -1)
                .Value = "(XT*X)-1 ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With
            With .ActiveCell.Offset(n2, -1).Characters(Start:=3, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 11
                .Strikethrough = False
                .Superscript = True
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            With .ActiveCell.Offset(n2, -1).Characters(Start:=7, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 11
                .Strikethrough = False
                .Superscript = True
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            With .ActiveCell.Offset(n2, -1).Characters(Start:=8, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 11
                .Strikethrough = False
                .Superscript = True
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With

            .Cells(4, n3).Select
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n2, -1)
                .Value = "XT*Y ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With
            With .ActiveCell.Offset(n2, -1).Characters(Start:=2, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 11
                .Strikethrough = False
                .Superscript = True
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With

            .Cells(4, n3).Select
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n2, -1)
                .Value = "B ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With

        End With
    End Sub

    Public Sub AJUSTE_02()
        With Globals.ThisAddIn.Application


            .Range("E1:F2").Select()
            .Selection.Cut
            .Range("A1").Select()
            .ActiveSheet.Paste
            .Range("A1").Select()
            .ActiveCell.FormulaR1C1 = "n"
            .Range("A2").Select()
            .ActiveCell.FormulaR1C1 = "k"
            .Range("C1").Select()
            .ActiveCell.FormulaR1C1 = "n-k"
            .Range("D1").Select()
            .Application.CutCopyMode = False
            .ActiveCell.FormulaR1C1 = "=RC[-2]-R[1]C[-2]"
            .Range("C2").Select()
            .ActiveCell.FormulaR1C1 = "k-1"
            .Range("D2").Select()
            .ActiveCell.FormulaR1C1 = "=RC[-2]-1"
            .Range("E1").Select()
            .ActiveCell.FormulaR1C1 = "Y promedio"
            .Range("E2").Select()
            .ActiveCell.FormulaR1C1 = "SRC"
            .Range("G1").Select()
            .ActiveCell.FormulaR1C1 = "STC"
            .Range("G2").Select()
            .ActiveCell.FormulaR1C1 = "SEC"
            .Range("I1").Select()
            .ActiveCell.FormulaR1C1 = "R2"
            .Range("I2").Select()
            .ActiveCell.FormulaR1C1 = "R2 AJUST"
            .Range("K1").Select()
            .ActiveCell.FormulaR1C1 = "F"
            .Range("K2").Select()
            .ActiveCell.FormulaR1C1 = "Prob F"
            .Range("M1").Select()
            .ActiveCell.FormulaR1C1 = "r"
            .Range("M2").Select()
            .ActiveCell.FormulaR1C1 = "o2"
            .Range("A1:N2").Select()

            With .Selection.Font
                .Name = "Calibri"
                .Size = 14
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            With .Selection.Font
                .Name = "Calibri"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            .Selection.Font.Bold = False
            .Selection.Font.Bold = True
            With .Selection
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                .MergeCells = False
            End With

        End With
    End Sub

    Public Sub AJUSTE_03()
        With Globals.ThisAddIn.Application
            .Range("H2").Select()
            Globals.ThisAddIn.Application.CutCopyMode = False
            .ActiveCell.FormulaR1C1 = "=R[-1]C-RC[-2]"

            .Range("J1").Select()
            .ActiveCell.FormulaR1C1 = "=1-(R[1]C[-4]/RC[-2])"

            .Range("J2").Select()
            .ActiveCell.FormulaR1C1 =
    "=1-(((RC[-4]/R[-1]C[-6]))/((R[-1]C[-2])/(R[-1]C[-8]-1)))"

            .Range("L1").Select()
            .ActiveCell.FormulaR1C1 = "=(R[1]C[-4]/R[1]C[-8])/(R[1]C[-6]/RC[-8])"

            .Range("L2").Select()
            .ActiveCell.FormulaR1C1 = "=F.DIST.RT(R[-1]C,RC[-8],R[-1]C[-8])"

            .Range("N1").Select()
            .ActiveCell.FormulaR1C1 = "=SQRT(RC[-4])"

            .Range("N2").Select()
            Globals.ThisAddIn.Application.CutCopyMode = False
            .ActiveCell.FormulaR1C1 = "=RC[-8]/R[-1]C[-10]"
            .Range("N3").Select()

            .Range("I1").Select()
            .ActiveCell.FormulaR1C1 = "R2"
            With .ActiveCell.Characters(Start:=1, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Negrita"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            With .ActiveCell.Characters(Start:=2, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Negrita"
                .Size = 11
                .Strikethrough = False
                .Superscript = True
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            .Range("I2").Select()
            .ActiveCell.FormulaR1C1 = "R2 AJUST"
            With .ActiveCell.Characters(Start:=1, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Negrita"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            With .ActiveCell.Characters(Start:=2, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Negrita"
                .Size = 11
                .Strikethrough = False
                .Superscript = True
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            With .ActiveCell.Characters(Start:=3, Length:=6).Font
                .Name = "Calibri"
                .FontStyle = "Negrita"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            .Range("M2").Select()
            .ActiveCell.FormulaR1C1 = "o2"
            With .ActiveCell.Characters(Start:=1, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Negrita"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            With .ActiveCell.Characters(Start:=2, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Negrita"
                .Size = 11
                .Strikethrough = False
                .Superscript = True
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With

        End With
    End Sub

    Public Sub AJUSTE_04()
        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim n2 As Long
            Dim n3 As Long


            n = (.Cells(2, 2).Value) / 2
            n2 = .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Column
            n3 = (.Cells(1, 2).Value) / 2
            .Cells(4, n2).Select()
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown) _
                .End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n, -1)
                .Value = " VarCov B ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 10
            End With

            .Cells(4, n2).Select
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown) _
                    .End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown) _
                    .End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n, -1)
                .Value = " Var B ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With
            '##############
            .Cells(4, n2).Select
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown) _
.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown) _
.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n, -1)
                .Value = " ee B ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With

            '##############
            .Cells(4, n2).Select()
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown) _
.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown) _
.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n, -1)
                .Value = "B ="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With
            With .ActiveCell.Offset(-1, 0)
                .Value = "t"
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With
            With .ActiveCell.Offset(-1, 1)
                .Value = "Prob t"
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With

            '#################################################################
            .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()
            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n3, -1)
                .Value = "Y' = "
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 14
            End With

            .Cells(.Rows.Count, 4).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n3, -1)
                .Value = "(Y-Y')^2="
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 9
            End With

            .Cells(.Rows.Count, 6).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select

            .ActiveCell.CurrentRegion.Select()
            .Selection.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Selection.Interior.ColorIndex = 2
            .Selection.Font.Bold = True

            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            End With
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone

            With .ActiveCell.Offset(n3, -1)
                .Value = "(Y-Yprom)^2= "
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.Size = 9
            End With

            .Range("A1:N2").Select()
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With

            Dim rangoTitulos(6) As String
            Dim i As Integer
            rangoTitulos(0) = "A1:A2"
            rangoTitulos(1) = "C1:C2"
            rangoTitulos(2) = "E1:E2"
            rangoTitulos(3) = "G1:G2"
            rangoTitulos(4) = "I1:I2"
            rangoTitulos(5) = "K1:K2"
            rangoTitulos(6) = "M1:M2"


            For i = 0 To rangoTitulos.Length - 1

                .Range(rangoTitulos(i)).Select()

                With .Selection.Interior
                    .Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                    .PatternColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                    .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With .Selection.Font
                    .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                    .TintAndShade = 0
                End With

            Next

            .Range("A1").Select()


        End With
    End Sub

End Module
