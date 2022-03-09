Public Class UF_RLUV

    Private Sub UF_RLUV_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        ComBox_Hojas.Items.Clear()

        For Each hoja As Excel.Worksheet In Globals.ThisAddIn.Application.Worksheets
            ComBox_Hojas.Items.Add(hoja.Name)
        Next
    End Sub

    Private Sub ComBox_Hojas_TextChanged(sender As Object, e As EventArgs) Handles ComBox_Hojas.TextChanged
        Dim hojaActiva As String
        hojaActiva = Globals.ThisAddIn.Application.ActiveSheet.Name

        If hojaActiva <> ComBox_Hojas.Text Then

            For Each hoja As Excel.Worksheet In Globals.ThisAddIn.Application.Worksheets
                hoja.Activate()
                If hoja.Name = ComBox_Hojas.Text Then
                    Exit For
                End If

            Next

        End If

    End Sub

    Private Sub txt_RangoY_DoubleClick(sender As Object, e As EventArgs) Handles txt_RangoY.DoubleClick

        Dim RangoY As Excel.Range

        If ComBox_Hojas.Text = "" Then
            MsgBox("Debes definir la hoja donde estan los datos", MsgBoxStyle.Exclamation, "ALERTA")
        Else
            RangoY = Globals.ThisAddIn.Application.InputBox(Prompt:="Introduce el rango de Y", Type:=8, HelpFile:="Selecciona un rango de celdas")
            txt_RangoY.Text = RangoY.Address
            Me.Activate()
        End If


    End Sub

    Private Sub txt_RangoX_DoubleClick(sender As Object, e As EventArgs) Handles txt_RangoX.DoubleClick

        Dim RangoX As Excel.Range


        If ComBox_Hojas.Text = "" Then
            MsgBox("Debes definir la hoja donde estan los datos", MsgBoxStyle.Exclamation, "ALERTA")
        Else
            RangoX = Globals.ThisAddIn.Application.InputBox(Prompt:="Introduce el rango de X", Type:=8, HelpFile:="Selecciona un rango de celdas")
            txt_RangoX.Text = RangoX.Address
            Me.Activate()
        End If

    End Sub

    Private Sub btn_Aceptar_Click(sender As Object, e As EventArgs) Handles btn_Aceptar.Click


        If ComBox_Hojas.Text <> "" And txt_RangoY.Text <> "Haz doble click para seleccionar los datos." And txt_RangoX.Text <> "Haz doble click para seleccionar los datos." Then

            preparaHoja(ComBox_Hojas.Text, txt_RangoY.Text, txt_RangoX.Text, txt_SigY.Text, txt_SigX.Text)

            Me.Close()

        Else
            MsgBox("Debes Elegir La Hoja Donde Estan Los Datos " + Chr(13) + "o no has elegido un rango de celdas para las variables")

        End If


    End Sub
End Class