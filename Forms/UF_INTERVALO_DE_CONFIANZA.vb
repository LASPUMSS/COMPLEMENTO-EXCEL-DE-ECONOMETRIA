Public Class UF_INTERVALO_DE_CONFIANZA
    Private Sub UF_INTERVALO_DE_CONFIANZA_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        ComB_Betas.Items.Clear()
        ComB_Betas.Items.Add("BETA 1")
        ComB_Betas.Items.Add("BETA 2")
    End Sub

    Private Sub btn_Aceptar_Click(sender As Object, e As EventArgs) Handles btn_Aceptar.Click
        intervaloDeConfianzaMet(txt_NvSf.Text, ComB_Betas.Text)
        Me.Close()
    End Sub
End Class