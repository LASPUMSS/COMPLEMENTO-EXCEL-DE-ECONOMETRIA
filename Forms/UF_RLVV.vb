Public Class UF_RLVV
    Private Sub UF_RLVV_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        ComB_Hojas.Items.Clear()

        For Each hoja As Excel.Worksheet In Globals.ThisAddIn.Application.Worksheets
            ComB_Hojas.Items.Add(hoja.Name)
        Next
    End Sub

    Private Sub ComB_Hojas_TextChanged(sender As Object, e As EventArgs) Handles ComB_Hojas.TextChanged
        Dim hojaActiva As String
        hojaActiva = Globals.ThisAddIn.Application.ActiveSheet.Name

        If hojaActiva <> ComB_Hojas.Text Then

            For Each hoja As Excel.Worksheet In Globals.ThisAddIn.Application.Worksheets
                hoja.Activate()
                If hoja.Name = ComB_Hojas.Text Then
                    Exit For
                End If

            Next

        End If
    End Sub

    Private Sub txt_RgY_DoubleClick(sender As Object, e As EventArgs) Handles txt_RgY.DoubleClick
        Dim RangoY As Excel.Range

        If ComB_Hojas.Text = "" Then
            MsgBox("Debes definir la hoja donde estan los datos", MsgBoxStyle.Exclamation, "ALERTA")
        Else
            RangoY = Globals.ThisAddIn.Application.InputBox(Prompt:="Introduce el rango de Y", Type:=8, HelpFile:="Selecciona un rango de celdas")
            txt_RgY.Text = RangoY.Address
            Me.Activate()
        End If
    End Sub

    Private Sub txt_RgX_DoubleClick(sender As Object, e As EventArgs) Handles txt_RgX.DoubleClick
        Dim RangoX As Excel.Range

        If ComB_Hojas.Text = "" Then
            MsgBox("Debes definir la hoja donde estan los datos", MsgBoxStyle.Exclamation, "ALERTA")
        Else
            RangoX = Globals.ThisAddIn.Application.InputBox(Prompt:="Introduce el rango de X", Type:=8, HelpFile:="Selecciona un rango de celdas")
            txt_RgX.Text = RangoX.Address
            Me.Activate()
        End If
    End Sub

    Private Sub btn_Aceptar_Click(sender As Object, e As EventArgs) Handles btn_Aceptar.Click
        metodoPrincipalRLVV(txt_RgX.Text, txt_RgY.Text)
        Me.Close()

    End Sub
End Class