Module BB03_RLUV

    Public Sub preparaHoja(ByVal HojaDatos As String, ByVal RangoY As String, ByVal RangoX As String, ByVal explicarY As String, ByVal medianteX As String)


        Dim HOJA_DATOS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(HojaDatos)
        Dim HOJA_RESULTADOS As Excel.Worksheet

        Dim filasX As Long
        Dim ColumnasX As Integer
        Dim filasY As Long
        Dim ColumnasY As Integer


        Globals.ThisAddIn.Application.Sheets.Add()
        HOJA_RESULTADOS = Globals.ThisAddIn.Application.ActiveSheet

        HOJA_DATOS.Activate()
        Globals.ThisAddIn.Application.Range(RangoX).Select()
        filasX = Globals.ThisAddIn.Application.Selection.Rows.Count
        ColumnasX = Globals.ThisAddIn.Application.Selection.Columns.Count


        HOJA_DATOS.Activate()
        Globals.ThisAddIn.Application.Range(RangoY).Select()
        filasY = Globals.ThisAddIn.Application.Selection.Rows.Count
        ColumnasY = Globals.ThisAddIn.Application.Selection.Columns.Count


        If filasX = filasY And ColumnasX = 1 And ColumnasY = 1 Then

            HOJA_DATOS.Activate()
            Globals.ThisAddIn.Application.Range(RangoX).Select()
            Globals.ThisAddIn.Application.Selection.Copy

            HOJA_RESULTADOS.Activate()
            Globals.ThisAddIn.Application.Range("B2").Select()
            Globals.ThisAddIn.Application.ActiveSheet.Paste(Link:=True)
            Globals.ThisAddIn.Application.Application.CutCopyMode = False

            HOJA_DATOS.Activate()
            Globals.ThisAddIn.Application.Range(RangoY).Select()
            Globals.ThisAddIn.Application.Selection.Copy

            HOJA_RESULTADOS.Activate()
            Globals.ThisAddIn.Application.Range("C2").Select()
            Globals.ThisAddIn.Application.ActiveSheet.Paste(Link:=True)
            Globals.ThisAddIn.Application.Application.CutCopyMode = False

            Dim i As Long
            For i = 2 To filasX + 1
                Globals.ThisAddIn.Application.Cells(i, 1).Select
                Globals.ThisAddIn.Application.Selection.Value = i - 1
            Next i

            colocarTitulos()
            calculosDeColumnas()
            calculosPromediosTotales()
            calculoBetas(explicarY, medianteX)
            propiedadesMCO()


        Else
            Globals.ThisAddIn.Application.Application.DisplayAlerts = False
            MsgBox("LAS LONGITUDES DE ""X"" Y ""Y"" NO SON IGUALES", vbExclamation)
            HOJA_RESULTADOS.Delete()
        End If

    End Sub

    Public Sub colocarTitulos()

        Globals.ThisAddIn.Application.Cells(1, 1).Select 'Esta linea es para controlar un error que coloca en negrita la celda celecionada

        With Globals.ThisAddIn.Application.ActiveSheet

            .Cells(1, 1).Value = "N°"
            .Cells(1, 1).Font.Bold = True

            .Cells(1, 2).Value = "X"
            .Cells(1, 2).Font.Bold = True

            .Cells(1, 3).Value = "Y"
            .Cells(1, 3).Font.Bold = True

            .Cells(1, 4).Value = "XY"
            .Cells(1, 4).Font.Bold = True

            .Cells(1, 5).Value = "X^2"
            .Cells(1, 5).Font.Bold = True

            .Cells(1, 6).Value = "xi"
            .Cells(1, 6).Font.Bold = True

            .Cells(1, 7).Value = "yi"
            .Cells(1, 7).Font.Bold = True

            .Cells(1, 8).Value = "xi*yi"
            .Cells(1, 8).Font.Bold = True

            .Cells(1, 9).Value = "xi^2"
            .Cells(1, 9).Font.Bold = True

            .Cells(1, 10).Value = "Y estimada"
            .Cells(1, 10).Font.Bold = True

            .Cells(1, 11).Value = "e"
            .Cells(1, 11).Font.Bold = True

            .Cells(1, 12).Value = "(Y est.)*e"
            .Cells(1, 12).Font.Bold = True

            .Cells(1, 13).Value = "X*e"
            .Cells(1, 13).Font.Bold = True

            '#################################################################

            .Cells(1, 14).Value = "yi estimada"
            .Cells(1, 14).Font.Bold = True

            .Cells(1, 15).Value = "e^2"
            .Cells(1, 15).Font.Bold = True

            .Cells(1, 16).Value = "yi^2"
            .Cells(1, 16).Font.Bold = True

            .Cells(1, 17).Value = "yi est^2"
            .Cells(1, 17).Font.Bold = True

            .Cells(1, 18).Value = "Y^2"
            .Cells(1, 18).Font.Bold = True

            .Cells(1, 19).Value = "yi*yi est"
            .Cells(1, 19).Font.Bold = True

        End With

        formatoTitulos()
        ajustarAnchoCol()

    End Sub

    Public Sub calculosDeColumnas()
        Dim i As Integer
        Dim n As Long
        Dim c As Integer
        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim celdaActiva As Excel.Range = hojaActiva.Cells(1, 1)

        n = celdaActiva.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
        c = celdaActiva.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column

        i = 2

        With Globals.ThisAddIn.Application

            'Bucles para realizar las operaciones fila por fila
            'Columna de del producto XY
            .ActiveSheet.Cells(i, 4).FormulaR1C1 = "=(RC[-2])*(RC[-1])"

            'Columna del producto X^2
            .ActiveSheet.Cells(i, 5).FormulaR1C1 = "=(RC[-3])^2"

            'Columna del producto Xi-X(prom)
            .ActiveSheet.Cells(i, 6).FormulaR1C1 = "=RC[-4]-R" & n + 2 & "C2"

            'Columna del producto Yi-Y(prom)
            .ActiveSheet.Cells(i, 7).FormulaR1C1 = "=RC[-4]-R" & n + 2 & "C3"

            'Columna del xi^2
            .ActiveSheet.Cells(i, 8).FormulaR1C1 = "=RC[-2]*RC[-1]"

            'Columna del xi^2
            .ActiveSheet.Cells(i, 9).FormulaR1C1 = "=RC[-3]^2"


            'Columna de Y estimada
            .ActiveSheet.Cells(i, 10).FormulaR1C1 = "=R" & n + 9 & "C2+R" & n + 8 & "C2*RC[-8]"

            'Columna de e
            .ActiveSheet.Cells(i, 11).FormulaR1C1 = "=RC[-8]-RC[-1]"

            'Columna de Y estimada * e
            .ActiveSheet.Cells(i, 12).FormulaR1C1 = "=RC[-1]*RC[-2]"

            'Columna de X*e
            .ActiveSheet.Cells(i, 13).FormulaR1C1 = "=RC[-11]*RC[-2]"

            '###########################################################

            'Columna de yi estimada
            .ActiveSheet.Cells(i, 14).FormulaR1C1 = "=RC[-4]-R" & n + 2 & "C3"

            'Columna e^2
            .ActiveSheet.Cells(i, 15).FormulaR1C1 = "=RC[-4]^2"

            'Columna yi^2
            .ActiveSheet.Cells(i, 16).FormulaR1C1 = "=RC[-9]^2"

            'Columna yi estimada ^2
            .ActiveSheet.Cells(i, 17).FormulaR1C1 = "=RC[-3]^2"

            'Columna Y^2
            .ActiveSheet.Cells(i, 18).FormulaR1C1 = "=RC[-15]^2"

            'Colmna yi * yi estimada
            .ActiveSheet.Cells(i, 19).FormulaR1C1 = "=RC[-12]*RC[-5]"


        End With


        Dim rangoExp As Excel.Range

        rangoExp = Globals.ThisAddIn.Application.ActiveSheet.Range(
            Globals.ThisAddIn.Application.ActiveSheet.Cells(2, 4),
            Globals.ThisAddIn.Application.ActiveSheet.Cells(2, c))

        rangoExp.AutoFill(Destination:=Globals.ThisAddIn.Application.ActiveSheet.Range(
                          Globals.ThisAddIn.Application.ActiveSheet.Cells(2, 4),
                          Globals.ThisAddIn.Application.ActiveSheet.Cells(n, c)))

        formatoObservaciones()


    End Sub

    Public Sub calculosPromediosTotales()

        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim celdaInicial As Excel.Range = hojaActiva.Cells(1, 1)
        Dim num_filas_sum As Long = celdaInicial.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row + 1
        Dim num_columnas_sum As Integer = celdaInicial.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column - 2
        Dim n1 As Long = celdaInicial.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
        Dim c1 As Integer


        With hojaActiva
            .Cells(n1 + 1, 1).Value = "SUMA"
            .Cells(n1 + 1, 1).Font.Bold = True

            .Cells(n1 + 2, 1).Value = "PROMEDIO"
            .Cells(n1 + 2, 1).Font.Bold = True

            .Cells(n1 + 2, 15).Value = "SRC"
            .Cells(n1 + 2, 15).Font.Bold = True
            .Cells(n1 + 2, 15).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .Cells(n1 + 2, 16).Value = "STC"
            .Cells(n1 + 2, 16).Font.Bold = True
            .Cells(n1 + 2, 16).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .Cells(n1 + 2, 17).Value = "SEC"
            .Cells(n1 + 2, 17).Font.Bold = True
            .Cells(n1 + 2, 17).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

        End With


        For c1 = 0 To num_columnas_sum Step 1

            hojaActiva.Cells(num_filas_sum, 2 + c1).FormulaR1C1 = "=SUM(R[-" & num_filas_sum - 2 & "]C:R[-1]C)"

        Next

        For c1 = 0 To 1 Step 1

            hojaActiva.Cells(num_filas_sum + 1, 2 + c1).FormulaR1C1 = "=AVERAGE(R[-" & num_filas_sum - 1 & "]C:R[-2]C)"

        Next

        formatoPromTls()


    End Sub

    Public Sub calculoBetas(ByRef explicarY As String, ByVal medianteX As String)

        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim n As Long = hojaActiva.Cells(1, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row


        With Globals.ThisAddIn.Application
            '############################################################
            '############################################################

            .ActiveSheet.Cells(n + 2, 1).FormulaR1C1 = "DETERMINACION DE BETAS"
            .ActiveSheet.Cells(n + 2, 1).Font.Bold = True

            .ActiveSheet.Cells(n + 3, 1).FormulaR1C1 = "Forma Clasica:"
            .ActiveSheet.Cells(n + 3, 1).Font.Bold = True

            .ActiveSheet.Cells(n + 4, 1).FormulaR1C1 = "Numerador"
            .ActiveSheet.Cells(n + 4, 1).Font.Bold = True

            .ActiveSheet.Cells(n + 5, 1).FormulaR1C1 = "Denominador"
            .ActiveSheet.Cells(n + 5, 1).Font.Bold = True

            .ActiveSheet.Cells(n + 6, 1).FormulaR1C1 = "Beta 2"
            .ActiveSheet.Cells(n + 6, 1).Font.Bold = True

            .ActiveSheet.Cells(n + 7, 1).FormulaR1C1 = "Beta 1"
            .ActiveSheet.Cells(n + 7, 1).Font.Bold = True

            .ActiveSheet.Cells(n + 3, 4).FormulaR1C1 = "Por desvios:"
            .ActiveSheet.Cells(n + 3, 4).Font.Bold = True

            .ActiveSheet.Cells(n + 4, 4).FormulaR1C1 = "Beta 2"
            .ActiveSheet.Cells(n + 4, 4).Font.Bold = True

            .ActiveSheet.Cells(n + 5, 4).FormulaR1C1 = "Beta 1"
            .ActiveSheet.Cells(n + 5, 4).Font.Bold = True

            .ActiveSheet.Cells(n + 6, 4).FormulaR1C1 = "OBJETIVO"
            .ActiveSheet.Cells(n + 6, 4).Font.Bold = True

            .ActiveSheet.Cells(n + 7, 4).FormulaR1C1 = "Explicar:"
            .ActiveSheet.Cells(n + 7, 4).Font.Bold = True

            .ActiveSheet.Cells(n + 8, 4).FormulaR1C1 = "Mediante:"
            .ActiveSheet.Cells(n + 8, 4).Font.Bold = True

            '//////////////////////////////////////////////////////////////////////////////////////////////////////
            'FORMA CLASICA
            '############################################################
            '#######        DETERMINACION DEL NUMERADOR     #############
            '############################################################

            .ActiveSheet.Cells(n + 4, 2).FormulaR1C1 = "=R[-5]C[2]-R[-6]C[-1]*R[-4]C*R[-4]C[1]"


            '############################################################
            '#######        DETERMINACION DEL DENOMINADOR     ###########
            '############################################################

            .ActiveSheet.Cells(n + 5, 2).FormulaR1C1 = "=R[-6]C[3]-R[-7]C[-1]*R[-5]C^2"


            '############################################################
            '#######        DETERMINACION DEL BETA 2          ###########
            '############################################################

            .ActiveSheet.Cells(n + 6, 2).FormulaR1C1 = "=R[-2]C/R[-1]C"


            '############################################################
            '#######        DETERMINACION DEL BETA 1          ###########
            '############################################################

            .ActiveSheet.Cells(n + 7, 2).FormulaR1C1 = "=R[-7]C[1]-R[-1]C*R[-7]C"

            '///////////////////////////////////////////////////////////////////////////////////
            'POR DEVIOS
            '############################################################
            '#######        DETERMINACION DEL BETA 2          ###########
            '############################################################

            .ActiveSheet.Cells(n + 4, 5).FormulaR1C1 = "=R[-5]C[3]/R[-5]C[4]"


            '############################################################
            '#######        DETERMINACION DEL BETA 1          ###########
            '############################################################

            .ActiveSheet.Cells(n + 5, 5).FormulaR1C1 = "=R[-5]C[-2]-R[-1]C*R[-5]C[-3]"

            .ActiveSheet.Cells(n + 7, 5).Value = explicarY
            .ActiveSheet.Cells(n + 8, 5).Value = medianteX

            '###########################################################
            '###########################################################
            '##### TITULOS DE VARIACIONES Y DESVIACIONES ESTANDAR  #####
            '###########################################################
            '###########################################################

            .ActiveSheet.Cells(n + 2, 10).FormulaR1C1 = "DETERMINACIÓN DE VARIANZA Y DESVIACIÓN ESTANDAR."
            .ActiveSheet.Cells(n + 2, 10).Font.Bold = True

            .ActiveSheet.Cells(n + 3, 10).FormulaR1C1 = "PARA LAS DESVIACIONES:"
            .ActiveSheet.Cells(n + 3, 10).Font.Bold = True

            .ActiveSheet.Cells(n + 4, 10).FormulaR1C1 = "n-k"
            .ActiveSheet.Cells(n + 4, 10).Font.Bold = True

            .ActiveSheet.Cells(n + 5, 10).FormulaR1C1 = "Varianza u"

            .ActiveSheet.Cells(n + 6, 10).FormulaR1C1 = "Err. Estandar"

            .ActiveSheet.Cells(n + 7, 10).FormulaR1C1 = "PARA BETA 2:"
            .ActiveSheet.Cells(n + 7, 10).Font.Bold = True

            .ActiveSheet.Cells(n + 8, 10).FormulaR1C1 = "Varianza"

            .ActiveSheet.Cells(n + 9, 10).FormulaR1C1 = "Err. Estandar"

            .ActiveSheet.Cells(n + 10, 10).FormulaR1C1 = "PARA BETA 1:"
            .ActiveSheet.Cells(n + 10, 10).Font.Bold = True

            .ActiveSheet.Cells(n + 11, 10).FormulaR1C1 = "Varianza"

            .ActiveSheet.Cells(n + 12, 10).FormulaR1C1 = "Err. Estandar"

            '###########################################################

            .ActiveSheet.Cells(n + 14, 10).FormulaR1C1 = "DETERMINACIÓN DEL COEFICIENTE DE DETERMINACIÓN:"
            .ActiveSheet.Cells(n + 14, 10).Font.Bold = True

            .ActiveSheet.Cells(n + 15, 10).FormulaR1C1 = "R^2"

            .ActiveSheet.Cells(n + 17, 10).FormulaR1C1 = "DETERMINACIÓN DEL COEFICIENTE DE CORRELACIÓN:"
            .ActiveSheet.Cells(n + 17, 10).Font.Bold = True

            .ActiveSheet.Cells(n + 18, 10).FormulaR1C1 = "r"
            .ActiveSheet.Cells(n + 18, 12).FormulaR1C1 = "Formula clasica."

            .ActiveSheet.Cells(n + 19, 10).FormulaR1C1 = "r"
            .ActiveSheet.Cells(n + 19, 12).FormulaR1C1 = "Formula por desvios."

            .ActiveSheet.Cells(n + 20, 10).FormulaR1C1 = "r"
            .ActiveSheet.Cells(n + 20, 12).FormulaR1C1 = "Atajo desde R^2."

            '###########################################################
            '###########################################################
            '##### FORMULAS DE VARIACIONES Y DESVIACIONES ESTANDAR  #####
            '###########################################################
            '###########################################################


            'Determinacion de grados de libertad

            .ActiveSheet.Cells(n + 4, 11).FormulaR1C1 = "=R[-6]C[-10]-2"

            'Varianza y error estandar de las desviaciones

            .ActiveSheet.Cells(n + 5, 11).FormulaR1C1 = "=R[-6]C[4]/R[-1]C"
            .ActiveSheet.Cells(n + 6, 11).FormulaR1C1 = "=SQRT(R[-1]C)"

            'Varianza y error estandar de beta 2

            .ActiveSheet.Cells(n + 8, 11).FormulaR1C1 = "=R[-3]C/R[-9]C[-2]"
            .ActiveSheet.Cells(n + 9, 11).FormulaR1C1 = "=SQRT(R[-1]C)"

            'Varianza y error estandar de beta 1

            .ActiveSheet.Cells(n + 11, 11).FormulaR1C1 = "=(R[-6]C*R[-12]C[-6])/(R[-13]C[-10]*R[-12]C[-2])"
            .ActiveSheet.Cells(n + 12, 11).FormulaR1C1 = "=SQRT(R[-1]C)"

            'Determinacion del coeficiente de determinacion.

            .ActiveSheet.Cells(n + 15, 11).FormulaR1C1 = "=R[-16]C[6]/R[-16]C[5]"
            .ActiveSheet.Cells(n + 15, 11).Style = "Percent"

            'Determinacion de los coeficientes de correlacion
            .ActiveSheet.Cells(n + 18, 11).FormulaR1C1 = "=(R[-20]C[-10]*R[-19]C[-7]-R[-19]C[-9]*R[-19]C[-8])/" &
            "(SQRT((R[-20]C[-10]*R[-19]C[-6]-(R[-19]C[-9])^2)*(R[-20]C[-10]*R[-19]C[7]-(R[-19]C[-8])^2)))"

            '    .ActiveCell.FormulaR1C1 = "Formula clasica."
            .ActiveSheet.Cells(n + 19, 11).FormulaR1C1 = "=(R[-20]C[8])/(SQRT(R[-20]C[5]*R[-20]C[6]))"

            '    .ActiveCell.FormulaR1C1 = "Formula por desvios."
            .ActiveSheet.Cells(n + 20, 11).FormulaR1C1 = "=SQRT(R[-5]C)"
            '.ActiveCell.FormulaR1C1 = "Atajo desde R^2."

        End With


    End Sub

    Public Sub propiedadesMCO()

        Dim FI As Integer = Globals.ThisAddIn.Application.Cells(Globals.ThisAddIn.Application.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row + 2

        With Globals.ThisAddIn.Application

            'TITULOS CULUMNA UNO DE LAS PROPIEDADES

            .Cells(FI, 1).Value = "PROPIEDADES:"
            .Cells(FI, 1).Font.Bold = True

            .Cells(FI + 1, 1).Value = "1ERA PROPIEDAD: Que el promedio de Y observada es igual a Y estimada."
            .Cells(FI + 1, 1).Font.Bold = True

            .Cells(FI + 6, 1).Value = "2DA PROPIEDAD: Que el promedio de Y observada es igual al promedio de Y estimada."
            .Cells(FI + 6, 1).Font.Bold = True

            .Cells(FI + 7, 1).Value = "O su equivalente que la suma de Y observada es igual  la suma de Y estmada."
            .ActiveCell.Font.Bold = True

            .Cells(FI + 12, 1).Value = "3RA PROPIEDAD: Que el promedio de los residuos es igual a cero."
            .Cells(FI + 12, 1).Font.Bold = True

            .Cells(FI + 13).Value = "O lo que es lo mismo decir, que la suma de los residuos es cero."
            .Cells(FI + 13).Font.Bold = True

            .Cells(FI + 18, 1).Value = "4TA PROPIEDAD: Que la suma de los productos resultantes de Y estimada por sus desviaciones correspondientes, es igual a cero."
            .Cells(FI + 18, 1).Font.Bold = True

            .Cells(FI + 19, 1).Value = "Los residuos no tienen niguna relación con el valor estimado."
            .Cells(FI + 19, 1).Font.Bold = True

            .Cells(FI + 24, 1).Value = "5TA PROPIEDAD: Que la suma de los productos resultantes de X observada por sus desviaciones correspondientes, es igual a cero."
            .Cells(FI + 24, 1).Font.Bold = True

            'TITULOS COLUMNA 2

            .Cells(FI + 3, 2).Value = "Prom. Y."
            .Cells(FI + 3, 2).Font.Bold = True

            .Cells(FI + 9, 2).Value = "Suma Y"
            .Cells(FI + 9, 2).Font.Bold = True

            .Cells(FI + 15, 2).Value = "Suma de e."
            .Cells(FI + 15, 2).Font.Bold = True

            .Cells(FI + 21, 2).Value = "Suma de Y estimada por e."
            .Cells(FI + 21, 2).Font.Bold = True

            .Cells(FI + 26, 2).Value = "Suma de X estimada por e."
            .Cells(FI + 26, 2).Font.Bold = True

            'TITULOS COLUMNA 3
            .Cells(FI + 3, 3).Value = "Y estimada."
            .Cells(FI + 3, 3).Font.Bold = True

            .Cells(FI + 9, 3).Value = "Suma Y estimada"
            .Cells(FI + 9, 3).Font.Bold = True

            'AHORA VAMOS HACER LAS FORMULAR PARA LAS PROPIEDADES
            'EMPEZANDO DESDE LA TERCERA COLUMNA
            .Cells(FI + 4, 3).FormulaR1C1 = "=ROUND((R[-6]C[-1]+R[-7]C[-1]*R[-13]C[-1]),2)"
            .Cells(FI + 4, 3).NumberFormat = "#,##0.00"

            .Cells(FI + 10, 3).FormulaR1C1 = "=ROUND(R[-20]C[7],2)"
            .Cells(FI + 10, 3).NumberFormat = "#,##0.00"

            'CONTINUAMOS CON LA SEGUNDA COLUMNA
            .Cells(FI + 4, 2).FormulaR1C1 = "=ROUND(R[-13]C[1],2)"
            .Cells(FI + 4, 2).NumberFormat = "#,##0.00"

            .Cells(FI + 10, 2).FormulaR1C1 = "=ROUND(R[-20]C[1],2)"
            .Cells(FI + 10, 2).NumberFormat = "#,##0.00"

            .Cells(FI + 16, 2).FormulaR1C1 = "=ROUND(R[-26]C[9],2)"

            .Cells(FI + 22, 2).FormulaR1C1 = "=ROUND(R[-32]C[10],2)"

            .Cells(FI + 27, 2).FormulaR1C1 = "=ROUND(R[-37]C[11],2)"

            'CONTINUAMOS CON LA PRIMERA COLUMNA
            .Cells(FI + 4, 1).FormulaR1C1 = "=IF(RC[1]=RC[2],""Se Cumple."",""No Cumple."")"

            .Cells(FI + 10, 1).FormulaR1C1 = "=IF(RC[1]=RC[2],""Se Cumple."",""No Cumple."")"

            .Cells(FI + 16, 1).FormulaR1C1 = "=IF(RC[1]=0,""Se Cumple."",""No Cumple."")"

            .Cells(FI + 22, 1).FormulaR1C1 = "=IF(RC[1]=0,""Se Cumple."",""No Cumple."")"

            .Cells(FI + 27, 1).FormulaR1C1 = "=IF(RC[1]=0,""Se Cumple."",""No Cumple."")"

        End With



    End Sub
End Module
