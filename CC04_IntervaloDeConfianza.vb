Module CC04_IntervaloDeConfianza
    Public Sub intervaloDeConfianzaMet(ByVal txt_NvSf As String, ByVal ComB_Betas As String)
        On Error Resume Next

        Dim filaInicial As String
        Dim filaFinal As String
        Dim titulo As String

        filaInicial = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).Value
        filaFinal = Globals.ThisAddIn.Application.ActiveSheet.Cells(29, 1).Value
        titulo = Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 5).Value


        With Globals.ThisAddIn.Application
            If filaInicial = "DETERMINACION DE BETAS" And filaFinal = "SEC" And titulo = "PRUEBAS DE HIPOTESIS, TABLAS RESUMEN, CUADROS ANOVA, P VALOR Y INTERVALOS DE CONFIANZA" Then

                .Cells(Globals.ThisAddIn.Application.ActiveSheet.Rows.Count, 5).Select()
                .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()
                .ActiveCell.Offset(4, 0).Select()

                .ActiveCell.Value = "Intervalo de confianza:"
                With .Selection.Font
                    .Color = -16777024
                    .TintAndShade = 0
                    .Bold = True
                End With
                .Selection.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                .ActiveCell.Offset(1, 0).Value = "De " & ComB_Betas & ", a un nivel de significancia de " _
                & txt_NvSf
                .ActiveCell.Offset(1, 0).Font.Bold = True

                .ActiveCell.Offset(1, 5).Value = "tc="
                With .ActiveCell.Offset(1, 5).Characters(Start:=2, Length:=1).Font
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


                .ActiveCell.Offset(1, 5).Font.Bold = True
                .ActiveCell.Offset(1, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .ActiveCell.Offset(1, 6).FormulaR1C1 = "=T.INV.2T(" & txt_NvSf & ",R10C2)"

                If ComB_Betas = "BETA 1" Then
                    .ActiveCell.Offset(2, 2).FormulaR1C1 = "=R6C2-R18C2*R[-1]C[4]"
                    .ActiveCell.Offset(2, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .ActiveCell.Offset(2, 3).Value = "B"
                    .ActiveCell.Offset(2, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .ActiveCell.Offset(2, 3).Font.Bold = True
                    .ActiveCell.Offset(2, 4).FormulaR1C1 = "=R6C2+R18C2*R[-1]C[2]"
                    .ActiveCell.Offset(2, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                ElseIf ComB_Betas = "BETA 2" Then
                    .ActiveCell.Offset(2, 2).FormulaR1C1 = "=R5C2-R15C2*R[-1]C[4]"
                    .ActiveCell.Offset(2, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .ActiveCell.Offset(2, 3).Value = "B"
                    .ActiveCell.Offset(2, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                    .ActiveCell.Offset(2, 3).Font.Bold = True
                    .ActiveCell.Offset(2, 4).FormulaR1C1 = "=R5C2+R15C2*R[-1]C[2]"
                    .ActiveCell.Offset(2, 4).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                End If
            Else

                MsgBox("LA HOJA NO ESTA LISTA PARA EJECUTAR ESTE PROCEDIMIENTO")

            End If
        End With
    End Sub
End Module
