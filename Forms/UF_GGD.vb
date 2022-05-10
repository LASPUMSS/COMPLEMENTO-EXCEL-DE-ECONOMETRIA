Public Class UF_GGD

    Private Sub UF_GGD_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        ComboBoxHojasActivas.Items.Clear()

        For Each hoja As Excel.Worksheet In Globals.ThisAddIn.Application.Worksheets
            ComboBoxHojasActivas.Items.Add(hoja.Name)
        Next

    End Sub
    Private Sub ComboBoxHojas_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxHojasActivas.TextChanged
        Dim hojaActiva As String
        hojaActiva = Globals.ThisAddIn.Application.ActiveSheet.Name

        If hojaActiva <> ComboBoxHojasActivas.Text Then

            For Each hoja As Excel.Worksheet In Globals.ThisAddIn.Application.Worksheets
                hoja.Activate()
                If hoja.Name = ComboBoxHojasActivas.Text Then
                    Exit For
                End If

            Next

        End If
    End Sub

    Private Sub txt_RgY_DoubleClick(sender As Object, e As EventArgs) Handles txt_RgY.DoubleClick
        Dim RangoY As Excel.Range

        If ComboBoxHojasActivas.Text = "" Then
            MsgBox("Debes definir la hoja donde estan los datos", MsgBoxStyle.Exclamation, "ALERTA")
        Else
            RangoY = Globals.ThisAddIn.Application.InputBox(Prompt:="Introduce el rango de Y", Type:=8, HelpFile:="Selecciona un rango de celdas")
            txt_RgY.Text = RangoY.Address
            Me.Activate()
        End If
    End Sub

    Private Sub txt_RgX_DoubleClick(sender As Object, e As EventArgs) Handles txt_RgX.DoubleClick
        Dim RangoX As Excel.Range

        If ComboBoxHojasActivas.Text = "" Then
            MsgBox("Debes definir la hoja donde estan los datos", MsgBoxStyle.Exclamation, "ALERTA")
        Else
            RangoX = Globals.ThisAddIn.Application.InputBox(Prompt:="Introduce el rango de X", Type:=8, HelpFile:="Selecciona un rango de celdas")
            txt_RgX.Text = RangoX.Address
            Me.Activate()
        End If
    End Sub

    Private Sub btn_Aceptar_Click(sender As Object, e As EventArgs) Handles btn_Aceptar.Click
        If ComboBoxHojasActivas.Text <> "" And txt_RgY.Text <> "Haz doble click para seleccionar los datos." And txt_RgX.Text <> "Haz doble click para seleccionar los datos." Then

            graficoDispersionMetodo(txt_SigY.Text, txt_SigX.Text, txt_RgY.Text, txt_RgX.Text, ComboBoxHojasActivas.Text)

            Me.Close()

        Else
            MsgBox("Debes Elegir La Hoja Donde Estan Los Datos " + Chr(13) + "o no has elegido un rango de celdas para las variables")

        End If
    End Sub
End Class