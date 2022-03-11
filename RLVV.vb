Module RLVV
    Public Sub metodoPrincipalRLVV(ByVal txt_RgX As String, ByVal txt_RgY As String)
        copiarDatosRLVV(txt_RgX, txt_RgY)

        M_TRANSPUESTA_X()
        MULT_MATRICES()
        Y_ESTIMADA_SRC_ECT()
        M_FALTANTES()

    End Sub

    Public Sub copiarDatosRLVV(ByVal txt_RgX As String, ByVal txt_RgY As String)
        Dim hojaDatos As Excel.Worksheet
        Dim hojaEjercicio As Excel.Worksheet

        Dim filasMX As Long
        Dim filasMY As Long
        Dim colMX As Integer
        Dim colMY As Integer

        Dim i As Integer
        Dim n As Long

        hojaDatos = Globals.ThisAddIn.Application.ActiveSheet
        Globals.ThisAddIn.Application.Sheets.Add()
        hojaEjercicio = Globals.ThisAddIn.Application.ActiveSheet

        hojaDatos.Activate()
        filasMX = hojaDatos.Range(txt_RgX).Rows.Count
        filasMY = hojaDatos.Range(txt_RgY).Rows.Count

        colMX = hojaDatos.Range(txt_RgX).Columns.Count
        colMY = hojaDatos.Range(txt_RgY).Columns.Count

        With Globals.ThisAddIn.Application
            If filasMX = filasMY And colMY = 1 Then

                .Range(txt_RgY).Select()
                .Selection.Copy
                hojaEjercicio.Activate()
                .Range("B4").Select()
                .ActiveSheet.Paste(Link:=True)

                hojaDatos.Activate()
                .Range(txt_RgX).Select()
                .Selection.Copy
                hojaEjercicio.Activate()
                .Range("F4").Select()
                .ActiveSheet.Paste(Link:=True)

                i = 4
                Do While CStr(.Cells(i, 6).Value) <> ""
                    .Cells(i, 5).Value = 1
                    i = i + 1
                Loop

                .Cells(1, 2).Value = filasMY
                .Cells(2, 2).Value = colMY
                .Cells(1, 6).Value = filasMX
                .Cells(2, 6).Value = colMX + 1

            Else
                MsgBox("El numero de filas de la Matrix Y, no coinciden con el numero de filas de la Matriz X.")
            End If
        End With
    End Sub

    Public Sub M_TRANSPUESTA_X()
        'PRIMERO SELECIONAMOS LA CANTIDAD EXANTA DE CELDAS QUE OCUPAREMOS PARA
        'PODER APLICAR LA FORMULA TIPO MATRIZ
        Dim filasX As Long
        Dim colX As Integer

        With Globals.ThisAddIn.Application

            filasX = .Range("F1").Value
            colX = .Range("F2").Value

            Dim PI_MTX As String
            Dim PF_MTX As String

            .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Offset(0, 3).Select()
            PI_MTX = .ActiveCell.Address
            .ActiveCell.Offset(colX - 1, filasX - 1).Select()
            PF_MTX = .ActiveCell.Address

            .Range(PI_MTX & ":" & PF_MTX).Select()

            'Selection.FormulaArray = "=TRANSPOSE(RC[-8]:R[22]C[-3])"
            .Selection.FormulaArray = "=TRANSPOSE(RC[-" & 2 + colX & "]:R[" & filasX - 1 & "]C[-3])"

        End With

    End Sub

    Public Sub MULT_MATRICES()

        On Error Resume Next
        With Globals.ThisAddIn.Application

            'UBICACION DE LA MATRIZ Y
            Dim PIF_MY As Long
            Dim PFF_MY As Long
            Dim PIC_MY As Long

            .Cells(4, 2).Select
            PIF_MY = .ActiveCell.Row
            PFF_MY = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            PIC_MY = .ActiveCell.Column

            'UBICACION DE LA MATRIZ X
            Dim PIF_MX As Long
            Dim PFF_MX As Long
            Dim PIC_MX As Long
            Dim PFC_MX As Long

            .Cells(4, 5).Select
            PIF_MX = .ActiveCell.Row
            PFF_MX = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            PIC_MX = .ActiveCell.Column
            PFC_MX = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column

            'UBICACION DE LA MATRIZ TRANPUESTA X
            Dim PIF_MTX As Long
            Dim PFF_MTX As Long
            Dim PIC_MTX As Long
            Dim PFC_MTX As Long

            .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Select()
            PIF_MTX = .ActiveCell.Row
            PFF_MTX = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            PIC_MTX = .ActiveCell.Column
            PFC_MTX = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column


            '#################################################################

            Dim filasX As Long
            Dim colX As Integer
            Dim filasY As Long
            Dim colY As Integer

            Dim PI_M As String
            Dim PF_M As String

            Dim n As Long

            filasX = .Range("F1").Value
            colX = .Range("F2").Value
            filasY = .Range("B1").Value
            colY = .Range("B2").Value

            .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Select()
            n = .ActiveCell.Column

            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(3, 0).Select()
            PI_M = .ActiveCell.Address

            .ActiveCell.Offset(colX - 1, colX - 1).Select()
            PF_M = .ActiveCell.Address

            .Range(PI_M & ":" & PF_M).Select()

            .Selection.FormulaArray = "=MMULT(R" & PIF_MTX & "C" & PIC_MTX & ":R" & PFF_MTX & "C" & PFC_MTX & ",R" &
            PIF_MX & "C" & PIC_MX & ":R" & PFF_MX & "C" & PFC_MX & ")"


            'UBICACION DE LA MATRIZ PRODUCTO DE LA TRANPUESTA X * LA MATRIZ X
            Dim PIF_MTXX As Long
            Dim PFF_MTXX As Long
            Dim PIC_MTXX As Long
            Dim PFC_MTXX As Long

            PIF_MTXX = .ActiveCell.Row
            PFF_MTXX = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            PIC_MTXX = .ActiveCell.Column
            PFC_MTXX = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column


            '#################################################################
            .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Select()
            n = .ActiveCell.Column

            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(3, 0).Select()
            PI_M = .ActiveCell.Address

            .ActiveCell.Offset(colX - 1, colX - 1).Select()
            PF_M = .ActiveCell.Address

            .Range(PI_M & ":" & PF_M).Select()
            .Selection.FormulaArray = "=MINVERSE(R" & PIF_MTXX & "C" & PIC_MTXX & ":R" & PFF_MTXX & "C" & PFC_MTXX & ")"


            'UBICACION DE LA MATRIZ INVERSA DEL PRODUCTO DE LA TRANPUESTA X * LA MATRIZ X
            Dim PIF_MTXIN As Long
            Dim PFF_MTXIN As Long
            Dim PIC_MTXIN As Long
            Dim PFC_MTXIN As Long

            PIF_MTXIN = .ActiveCell.Row
            PFF_MTXIN = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            PIC_MTXIN = .ActiveCell.Column
            PFC_MTXIN = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column


            '#################################################################

            .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Select()
            n = .ActiveCell.Column

            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(3, 0).Select()
            PI_M = .ActiveCell.Address

            .ActiveCell.Offset(colX - 1, 0).Select()
            PF_M = .ActiveCell.Address

            .Range(PI_M & ":" & PF_M).Select()
            .Selection.FormulaArray = "=MMULT(R" & PIF_MTX & "C" & PIC_MTX & ":R" & PFF_MTX & "C" & PFC_MTX & ",R" _
            & PIF_MY & "C" & PIC_MY & ":R" & PFF_MY & "C" & PIC_MY & ")"


            'UBICACION DE LA MATRIZ DEL PRODUCTO DE LA TRANPUESTA X * LA MATRIZ Y
            Dim PIF_MTXY As Long
            Dim PFF_MTXY As Long
            Dim PIC_MTXY As Long

            PIF_MTXY = .ActiveCell.Row
            PFF_MTXY = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            PIC_MTXY = .ActiveCell.Column


            '#################################################################

            .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Select()
            n = .ActiveCell.Column

            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(3, 0).Select()
            PI_M = .ActiveCell.Address

            .ActiveCell.Offset(colX - 1, 0).Select()
            PF_M = .ActiveCell.Address

            .Range(PI_M & ":" & PF_M).Select()
            .Selection.FormulaArray = "=MMULT(R" & PIF_MTXIN & "C" & PIC_MTXIN & ":R" & PFF_MTXIN & "C" & PFC_MTXIN & ",R" _
            & PIF_MTXY & "C" & PIC_MTXY & ":R" & PFF_MTXY & "C" & PIC_MTXY & ")"

        End With
        AJUSTE_01()
        AJUSTE_02()

    End Sub

    Public Sub Y_ESTIMADA_SRC_ECT()
        On Error Resume Next

        With Globals.ThisAddIn.Application

            Dim PIF_MY As Long
            Dim PFF_MY As Long
            Dim PC_MY As Integer


            PIF_MY = .Cells(4, 2).Row
            PFF_MY = .Cells(4, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            PC_MY = .Cells(4, 2).Column


            'UBICACION DE DE LA MATRIZ X
            Dim PIF_MX As Long
            Dim PFF_MX As Long
            Dim PIC_MX As Long
            Dim PFC_MX As Long

            .Cells(4, 5).Select()
            PIF_MX = .ActiveCell.Row
            PFF_MX = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            PIC_MX = .Cells(4, 5).Column
            PFC_MX = .Cells(4, 5).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column

            ''UBICACION DE DE LA MATRIZ BETAS Y DE LAS BETAS
            '
            Dim PIF_MB As Long
            Dim PFF_MB As Long
            Dim PC_MB As Long

            Dim n1 As Long
            n1 = .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Column

            .Cells(.Rows.Count, n1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()
            PFF_MB = .ActiveCell.Row
            PIF_MB = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            PC_MB = n1


            'DETERMINACION DE LA MATRIZ DE Y ESTIMADA
            .Cells(4, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Offset(4, 0).Select()
            Dim PIF_YE As String
            PIF_YE = .ActiveCell.Address

            Dim PFF_YE As String
            PFF_YE = .ActiveCell.Offset(.Cells(1, 2).Value - 1, 0).Address
            .Range(PIF_YE & ":" & PFF_YE).Select()

            .Selection.FormulaArray = "=MMULT(R" & PIF_MX & "C" & PIC_MX & ":R" & PFF_MX & "C" & PFC_MX & ",R" _
            & PIF_MB & "C" & PC_MB & ":R" & PFF_MB & "C" & PC_MB & ")"

            .Range(PIF_YE).Select()
            Dim i As Long
            Dim i2 As Long

            i = .ActiveCell.Row
            i2 = 4
            Do While CStr(.Cells(i, PC_MY).Value) <> ""
                .Cells(i, PC_MY).Offset(0, 2).FormulaR1C1 = "=(R" & i2 & "C2-R" & i & "C2)^2"
                i = i + 1
                i2 = i2 + 1
            Loop

            .Cells(1, 6).Select()
            .ActiveCell.FormulaR1C1 = "=AVERAGE(R" & PIF_MY & "C2:R" & PFF_MY & "C2)"

            .Range(PIF_YE).Select()
            i = .ActiveCell.Row
            i2 = 4

            Do While CStr(.Cells(i, PC_MY).Value) <> ""
                .Cells(i, PC_MY).Offset(0, 4).FormulaR1C1 = "=(R" & i2 & "C2-R1C6)^2"
                i = i + 1
                i2 = i2 + 1
            Loop

            'UBICACION DE DE LA MATRIZ SRC
            Dim PIF_MSRC As Long
            Dim PFF_MSRC As Long
            Dim PC_MSRC As Long

            PC_MSRC = 4

            .Cells(.Rows.Count, PC_MSRC).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()
            PFF_MSRC = .ActiveCell.Row
            PIF_MSRC = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            .Cells(2, 6).Select()
            .ActiveCell.FormulaR1C1 = "=SUM(R" & PIF_MSRC & "C" & PC_MSRC & ":R" & PFF_MSRC & "C" & PC_MSRC & ")"

            'UBICACION DE DE LA MATRIZ STC
            Dim PIF_MSTC As Long
            Dim PFF_MSTC As Long
            Dim PC_MSTC As Long
            PC_MSTC = 6

            .Cells(.Rows.Count, PC_MSRC).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()
            PFF_MSTC = .ActiveCell.Row
            PIF_MSTC = .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            .Cells(1, 8).Select()
            .ActiveCell.FormulaR1C1 = "=SUM(R" & PIF_MSTC & "C" & PC_MSTC & ":R" & PFF_MSTC & "C" & PC_MSTC & ")"
        End With
        AJUSTE_03()
    End Sub

    Public Sub M_FALTANTES()
        'On Error Resume Next
        With Globals.ThisAddIn.Application

            Dim fila As Long
            Dim col As Long
            Dim n As Long
            Dim i As Long
            Dim i2 As Long
            Dim i3 As Long
            Dim i4 As Long

            .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Select()
            n = .ActiveCell.Column
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()

            fila = .ActiveCell.Row
            col = .ActiveCell.Column

            i = 0
            i2 = 0
            i3 = 0


            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(4, 0).Select()
            For i3 = 0 To .Cells(2, 2).Value - 1 Step 1

                For i = 0 To .Cells(2, 2).Value - 1 Step 1
                    .ActiveCell.Offset(0, i).FormulaR1C1 = "=R2C14*R" & fila & "C" & col + i2
                    i2 = i2 + 1
                Next i
                .ActiveCell.Offset(1, 0).Select()
                i2 = 0
                fila = fila + 1
            Next i3
            '#######################################################################
            Dim aux As Excel.Range
            Dim aux2 As Excel.Range
            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()
            fila = .ActiveCell.Row
            col = .ActiveCell.Column
            aux = .ActiveCell
            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(4, 0).Select()
            aux2 = .ActiveCell

            i = 0
            i2 = 0
            i3 = 0

            Dim rgVarianza As String
            For i = 0 To .Cells(2, 2).Value - 1 Step 1
                aux.Select()
                .ActiveCell.Offset(i2, i3).Select()
                rgVarianza = .ActiveCell.Address

                aux2.Select()
                .ActiveCell.Offset(i3, 0).Select()
                .ActiveCell.Value = "=" & rgVarianza
                i2 = i2 + 1
                i3 = i3 + 1
            Next i
            '#######################################################################
            Dim n2 As Long
            n2 = .Cells(2, 2).Value
            n2 = n2 + 3
            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()

            Do While CStr(.ActiveCell.Value) <> ""
                .ActiveCell.Offset(n2, 0).FormulaR1C1 = "=SQRT(R[-" & n2 & "]C)"
                .ActiveCell.Offset(1, 0).Select()
            Loop
            '#######################################################################
            .Cells(4, .Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Select()
            .ActiveCell.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Select()
            fila = .ActiveCell.Row
            col = .ActiveCell.Column

            i = 0
            i2 = 0
            i3 = 0

            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()

            Do While CStr(.ActiveCell.Value) <> ""
                .ActiveCell.Offset(n2, 0).FormulaR1C1 = "=R" & fila + i & "C" & col & "/R[-" & n2 & "]C"
                i = i + 1
                .ActiveCell.Offset(1, 0).Select()
            Loop

            '#######################################################################

            .Cells(.Rows.Count, n).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Select()

            Do While CStr(.ActiveCell.Value) <> ""
                .ActiveCell.Offset(0, 1).FormulaR1C1 = "=T.DIST.2T(ABS(RC[-1]),R1C4)"
                .ActiveCell.Offset(1, 0).Select()
            Loop

        End With

        AJUSTE_04()

    End Sub
End Module
