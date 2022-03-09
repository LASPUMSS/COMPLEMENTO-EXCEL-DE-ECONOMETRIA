Module CrearHojaHipotesis
    Public Sub hojaHipotesis()

        'On Error Resume Next
        Dim Verificar1 As String
        Dim Verificar2 As String
        Dim Verificar3 As String
        Dim Verificar4 As String

        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim celdaIncial As Excel.Range = hojaActiva.Cells(1, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)


        Verificar1 = celdaIncial.Value
        Verificar2 = celdaIncial.Offset(2, 0).Value
        Verificar3 = celdaIncial.Offset(2, 9).Value
        Verificar4 = celdaIncial.Offset(9, 0).Value


        If Verificar1 = "PROMEDIO" And Verificar2 = "DETERMINACION DE BETAS" And Verificar3 = "DETERMINACIÓN DE VARIANZA Y DESVIACIÓN ESTANDAR." And Verificar4 = "PROPIEDADES:" Then
            'PROCEDIMIENTO QUE PREPARA UNA HOJA PARA HACER HIPOTESIS SIMPLES SOBRE LOS RESULATADOS DE UNA REGRESION LINEAL A UNA VARIABLE
            Dim HOJA_DATOS As Excel.Worksheet
            Dim HOJA_HIPOTESIS As Excel.Worksheet
            Dim n As Integer

            HOJA_DATOS = Globals.ThisAddIn.Application.ActiveSheet
            HOJA_DATOS.Activate()
            Globals.ThisAddIn.Application.ActiveSheet.Range("A1").Select()
            n = Globals.ThisAddIn.Application.ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row + 2

            Globals.ThisAddIn.Application.Sheets.Add()
            HOJA_HIPOTESIS = Globals.ThisAddIn.Application.ActiveSheet

            HOJA_DATOS.Activate()
            Globals.ThisAddIn.Application.ActiveSheet.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(n, 1), Globals.ThisAddIn.Application.ActiveSheet.Cells(n + 5, 2)).Select()
            Globals.ThisAddIn.Application.Selection.Copy()

            HOJA_HIPOTESIS.Activate()
            Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 1).Select()
            Globals.ThisAddIn.Application.Selection.PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues,
                                                                 Operation:=Microsoft.Office.Interop.Excel.Constants.xlNone,
                                                                 SkipBlanks:=False, Transpose:=False)

            HOJA_DATOS.Activate()
            Globals.ThisAddIn.Application.ActiveSheet.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(n, 10), Globals.ThisAddIn.Application.ActiveSheet.Cells(n + 18, 12)).Select()
            Globals.ThisAddIn.Application.Selection.Copy()

            HOJA_HIPOTESIS.Activate()
            Globals.ThisAddIn.Application.ActiveSheet.Cells(8, 1).Select()
            Globals.ThisAddIn.Application.Selection.PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues,
                                                                 Operation:=Microsoft.Office.Interop.Excel.Constants.xlNone,
                                                                 SkipBlanks:=False, Transpose:=False)

            HOJA_DATOS.Activate()
            Globals.ThisAddIn.Application.ActiveSheet.Cells(n - 3, 15).Select()
            Globals.ThisAddIn.Application.Selection.Copy()

            HOJA_HIPOTESIS.Activate()
            Globals.ThisAddIn.Application.ActiveSheet.Cells(28, 2).Select()
            Globals.ThisAddIn.Application.Selection.PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues,
                                                                 Operation:=Microsoft.Office.Interop.Excel.Constants.xlNone,
                                                                 SkipBlanks:=False, Transpose:=False)

            Globals.ThisAddIn.Application.ActiveSheet.Cells(28, 1).Select()
            Globals.ThisAddIn.Application.ActiveCell.Value = "SRC"

            HOJA_DATOS.Activate()
            Globals.ThisAddIn.Application.ActiveSheet.Cells(n - 3, 17).Select()
            Globals.ThisAddIn.Application.Selection.Copy()

            HOJA_HIPOTESIS.Activate()
            Globals.ThisAddIn.Application.ActiveSheet.Cells(29, 2).Select()
            Globals.ThisAddIn.Application.Selection.PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues,
                                                                 Operation:=Microsoft.Office.Interop.Excel.Constants.xlNone,
                                                                 SkipBlanks:=False, Transpose:=False)

            Globals.ThisAddIn.Application.Cells(29, 1).Select()
            Globals.ThisAddIn.Application.ActiveCell.Value = "SEC"

            Globals.ThisAddIn.Application.Range("E1").Select()
            Globals.ThisAddIn.Application.Selection.Value = "PRUEBAS DE HIPOTESIS, TABLAS RESUMEN, CUADROS ANOVA, P VALOR Y INTERVALOS DE CONFIANZA"

            formatoTitulosHipotesis()


        Else
            MsgBox("La hoja no es la apropiada para ejecutar este procedimeinto.")
        End If

    End Sub




End Module
