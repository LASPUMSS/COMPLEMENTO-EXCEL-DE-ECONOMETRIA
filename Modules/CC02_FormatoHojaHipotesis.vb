Imports Microsoft.Office.Interop
Module CC02_FormatoHojaHipotesis

    Public Sub formatoTitulosHipotesis()

        Dim hojaActiva As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim nFil As Long = Globals.ThisAddIn.Application.ActiveSheet.Cells(Globals.ThisAddIn.Application.ActiveSheet.Rows.Count, 2).End(Excel.XlDirection.xlUp).Row
        Dim rangoB As Excel.Range = hojaActiva.Range(Globals.ThisAddIn.Application.ActiveSheet.Cells(1, 2), Globals.ThisAddIn.Application.ActiveSheet.Cells(nFil, 2))

        'MsgBox(CStr(nFil) + Chr(13) + rangoB.Address + Chr(13) + hojaActiva.Name)
        hojaActiva.Columns("A:A").ColumnWidth = 16.71
        hojaActiva.Columns("B:B").ColumnWidth = 16.71
        rangoB.NumberFormat = "#,##0.00"

        Globals.ThisAddIn.Application.Range("E1:N1").Select()
        With Globals.ThisAddIn.Application.Selection

            .Merge()
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlBottom
            .Font.Bold = True
            .Font.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            .Font.TintAndShade = 0
            .Font.Size = 12

        End With

        Dim i As Integer
        Dim celdas(9) As String

        celdas(0) = "A1"
        celdas(1) = "A2"
        celdas(2) = "A8"
        celdas(3) = "A9"
        celdas(4) = "A13"
        celdas(5) = "A16"
        celdas(6) = "A20"
        celdas(7) = "A23"
        celdas(8) = "A28"
        celdas(9) = "A29"

        For i = 0 To celdas.Length
            With Globals.ThisAddIn.Application.Range(celdas(i))
                .Font.Bold = True
                .Font.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = 0
            End With
        Next


    End Sub

End Module
