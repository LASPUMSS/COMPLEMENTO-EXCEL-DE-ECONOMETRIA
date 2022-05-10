Public Class UF_HIPOTESIS
    Private Sub UF_HIPOTESIS_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        ComB_Betas.Items.Clear()
        ComB_Betas.Items.Add("BETA 1")
        ComB_Betas.Items.Add("BETA 2")
    End Sub

    Private Sub btn_ProbarHipotesis_Click(sender As Object, e As EventArgs) Handles btn_ProbarHipotesis.Click
        On Error GoTo Etiqueta
        If ComB_Betas.Text <> "" And txt_NvSg.Text <> "" And txt_Hp.Text <> "" Then
            llamadaHipotesis(txt_Hipotesis.Text, txt_Hp.Text, txt_NvSg.Text, ComB_Betas.Text)
            Me.Close()
        Else
            MsgBox("Todos los datos deben estar dados ")
        End If
        Exit Sub
Etiqueta:
        MsgBox("Ingrese de manera correcta los datos")
    End Sub
End Class