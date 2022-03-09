Module HipotesisProcedimientos
    Public Sub llamadaHipotesis(ByVal txt_Hipotesis As String, ByVal txt_Hp As String, ByVal txt_NvSf As String, ByVal ComB_Betas As String)

        Dim filaInicial As String
        Dim filaFinal As String
        Dim titulo As String

        filaInicial = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).Value
        filaFinal = Globals.ThisAddIn.Application.ActiveSheet.Cells(29, 1).Value
        titulo = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 5).Value

        'Debug.Print filaInicial, filaFinal, titulo
        If filaInicial = "DETERMINACION DE BETAS" And filaFinal = "SEC" _
            And titulo = "PRUEBAS DE HIPOTESIS, TABLAS RESUMEN, CUADROS ANOVA, P VALOR Y INTERVALOS DE CONFIANZA" Then

            Globals.ThisAddIn.Application.ActiveSheet.Cells(Globals.ThisAddIn.Application.ActiveSheet.Rows.Count, 5).Select()
            Globals.ThisAddIn.Application.ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()
            Globals.ThisAddIn.Application.ActiveCell.Offset(3, 0).Select()

            hipotesisSimples(txt_Hipotesis, txt_Hp, txt_NvSf, ComB_Betas)

        Else
            MsgBox("LA HOJA NO ESTA LISTA PARA EJECUTAR ESTE PROCEDIMIENTO")
        End If
    End Sub

    Public Sub hipotesisSimples(ByVal txt_Hipotesis As String, ByVal txt_Hp As String, ByVal txt_NvSf As String, ByVal ComB_Betas As String)

        With Globals.ThisAddIn.Application
            .ActiveCell.Value = "Probar la hipótesis:"
            .ActiveCell.Font.Bold = True
            .ActiveCell.Offset(1, 0).Value = txt_Hipotesis

            'PASO1:PLATEAMIENTO DE LA HIPOTESIS
            .ActiveCell.Offset(3, 0).Value = "Paso 1. Planteamiento de la hipótesis:"
            .ActiveCell.Offset(3, 0).Font.Bold = True

            'PASO2:DETERMINACIÓN DE LA SIGNIFICANCIA
            .ActiveCell.Offset(7, 0).Value = "Paso 2. Determinación de significancia:"
            .ActiveCell.Offset(7, 0).Font.Bold = True

            'PASO3:DETERMINACION DE T CRITICO
            .ActiveCell.Offset(11, 0).Value = "Paso 3. Determinar t critico:"
            .ActiveCell.Offset(11, 0).Font.Bold = True

            'Paso4: DETERMINACION DE T CALCULADO
            .ActiveCell.Offset(15, 0).Value = "Paso 4. Determinar t calculado:"
            .ActiveCell.Offset(15, 0).Font.Bold = True

            'PASO5: SE ACEPTA O SE RECHAZA LA HIPOTESIS
            .ActiveCell.Offset(19, 0).Value = "Paso 5. Se acepta o se rechaza la hipotesis:"
            .ActiveCell.Offset(19, 0).Font.Bold = True

            'PASO6: P valor
            .ActiveCell.Offset(23, 0).Value = "Paso 6. Probabilidad de t calculado:"
            .ActiveCell.Offset(23, 0).Font.Bold = True

            'PASO1
            .ActiveCell.Offset(4, 1).Value = "H0"
            .ActiveCell.Offset(4, 1).Font.Bold = False
            .ActiveCell.Offset(4, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            With .ActiveCell.Offset(4, 1).Characters(Start:=2, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = True
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With

            .ActiveCell.Offset(5, 1).Value = "H1"
            .ActiveCell.Offset(5, 1).Font.Bold = False
            .ActiveCell.Offset(5, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            With .ActiveCell.Offset(5, 1).Characters(Start:=2, Length:=1).Font
                .Name = "Calibri"
                .FontStyle = "Normal"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = True
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = Microsoft.Office.Interop.Excel.XlThemeFont.xlThemeFontMinor
            End With
            .ActiveCell.Offset(4, 2).Value = "'="
            .ActiveCell.Offset(4, 2).Font.Bold = False
            .ActiveCell.Offset(4, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(5, 2).Value = "'=/="
            .ActiveCell.Offset(5, 2).Font.Bold = False
            .ActiveCell.Offset(5, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(4, 3).Value = txt_Hp
            .ActiveCell.Offset(4, 3).Font.Bold = False
            .ActiveCell.Offset(4, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(5, 3).Value = txt_Hp
            .ActiveCell.Offset(5, 3).Font.Bold = False
            .ActiveCell.Offset(5, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            'PASO 2
            .ActiveCell.Offset(8, 1).Value = "a"
            .ActiveCell.Offset(8, 1).Font.Bold = False
            .ActiveCell.Offset(8, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(9, 1).Value = "1-a"
            .ActiveCell.Offset(9, 1).Font.Bold = False
            .ActiveCell.Offset(9, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(8, 2).Value = "'="
            .ActiveCell.Offset(8, 2).Font.Bold = False
            .ActiveCell.Offset(8, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(9, 2).Value = "'="
            .ActiveCell.Offset(9, 2).Font.Bold = False
            .ActiveCell.Offset(9, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(8, 3).Value = txt_NvSf
            .ActiveCell.Offset(8, 3).Font.Bold = False
            .ActiveCell.Offset(8, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(9, 3).FormulaR1C1 = "=1-R[-1]C"
            .ActiveCell.Offset(9, 3).Font.Bold = False
            .ActiveCell.Offset(9, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            'PASO 3
            .ActiveCell.Offset(13, 1).Value = "t="
            .ActiveCell.Offset(13, 1).Font.Bold = False
            .ActiveCell.Offset(13, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(13, 2).FormulaR1C1 = "=T.INV.2T(R[-5]C[1],R10C2)"
            .ActiveCell.Offset(13, 2).Font.Bold = False
            .ActiveCell.Offset(13, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            'PASO 4
            .ActiveCell.Offset(17, 1).Value = "t="
            .ActiveCell.Offset(17, 1).Font.Bold = False
            .ActiveCell.Offset(17, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            If ComB_Betas = "BETA 1" Then
                .ActiveCell.Offset(17, 2).FormulaR1C1 = "=(R6C2-R[-13]C[1])/(R18C2)"
                .ActiveCell.Offset(17, 2).Font.Bold = False
                .ActiveCell.Offset(17, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            ElseIf ComB_Betas = "BETA 2" Then
                .ActiveCell.Offset(17, 2).FormulaR1C1 = "=(R5C2-R[-13]C[1])/(R15C2)"
                .ActiveCell.Offset(17, 2).Font.Bold = False
                .ActiveCell.Offset(17, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End If

            'PASO 5
            .ActiveCell.Offset(21, 1).FormulaR1C1 =
                    "=IF(ABS(R[-4]C[1])>R[-8]C[1],""Se rechaza la hipotesis nula."",IF(ABS(R[-4]C[1])<=R[-8]C[1],""Se acepta la hipotesis nula."",""""))"
            .ActiveCell.Offset(21, 1).Font.Bold = False
            .ActiveCell.Offset(21, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            'PASO 6
            .ActiveCell.Offset(24, 1).Value = "p valor="
            .ActiveCell.Offset(24, 1).Font.Bold = False
            .ActiveCell.Offset(24, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(24, 2).FormulaR1C1 = "=T.DIST.2T(ABS(R[-7]C),R10C2)"
            .ActiveCell.Offset(24, 2).Font.Bold = False
            .ActiveCell.Offset(24, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .ActiveCell.Offset(25, 1).FormulaR1C1 =
                    "=IF(R[-1]C[1]<R[-17]C[2],""Se rechaza la hipotesis nula."",IF(R[-1]C[1]>=R[-17]C[2],""Se acepta la hipotesis nula."",""""))"
            .ActiveCell.Offset(25, 1).Font.Bold = False
            .ActiveCell.Offset(25, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

        End With

        With Globals.ThisAddIn.Application.Selection.Font
            .Color = -16777024
            .TintAndShade = 0
        End With
        Globals.ThisAddIn.Application.Selection.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle
    End Sub
End Module
