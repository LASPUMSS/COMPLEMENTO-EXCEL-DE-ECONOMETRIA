<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UF_RLUV
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btn_Aceptar = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txt_SigX = New System.Windows.Forms.TextBox()
        Me.txt_SigY = New System.Windows.Forms.TextBox()
        Me.txt_RangoX = New System.Windows.Forms.TextBox()
        Me.txt_RangoY = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComBox_Hojas = New System.Windows.Forms.ComboBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_Aceptar
        '
        Me.btn_Aceptar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Aceptar.Location = New System.Drawing.Point(410, 12)
        Me.btn_Aceptar.Name = "btn_Aceptar"
        Me.btn_Aceptar.Size = New System.Drawing.Size(97, 43)
        Me.btn_Aceptar.TabIndex = 0
        Me.btn_Aceptar.Text = "ACEPTAR"
        Me.btn_Aceptar.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txt_SigX)
        Me.GroupBox1.Controls.Add(Me.txt_SigY)
        Me.GroupBox1.Controls.Add(Me.txt_RangoX)
        Me.GroupBox1.Controls.Add(Me.txt_RangoY)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ComBox_Hojas)
        Me.GroupBox1.Location = New System.Drawing.Point(25, 37)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(365, 272)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "VARIABLES"
        '
        'txt_SigX
        '
        Me.txt_SigX.Location = New System.Drawing.Point(129, 208)
        Me.txt_SigX.Name = "txt_SigX"
        Me.txt_SigX.Size = New System.Drawing.Size(217, 20)
        Me.txt_SigX.TabIndex = 9
        Me.txt_SigX.Text = "N/A"
        '
        'txt_SigY
        '
        Me.txt_SigY.Location = New System.Drawing.Point(129, 165)
        Me.txt_SigY.Name = "txt_SigY"
        Me.txt_SigY.Size = New System.Drawing.Size(217, 20)
        Me.txt_SigY.TabIndex = 8
        Me.txt_SigY.Text = "N/A"
        '
        'txt_RangoX
        '
        Me.txt_RangoX.Location = New System.Drawing.Point(129, 122)
        Me.txt_RangoX.Name = "txt_RangoX"
        Me.txt_RangoX.Size = New System.Drawing.Size(217, 20)
        Me.txt_RangoX.TabIndex = 7
        Me.txt_RangoX.Text = "Haz doble click para seleccionar los datos."
        '
        'txt_RangoY
        '
        Me.txt_RangoY.Location = New System.Drawing.Point(129, 79)
        Me.txt_RangoY.Name = "txt_RangoY"
        Me.txt_RangoY.Size = New System.Drawing.Size(217, 20)
        Me.txt_RangoY.TabIndex = 6
        Me.txt_RangoY.Text = "Haz doble click para seleccionar los datos."
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(20, 210)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 13)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Significado X:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(20, 167)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Significado Y:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(20, 124)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(52, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Rango X:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 81)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Rango Y:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Hoja Datos:"
        '
        'ComBox_Hojas
        '
        Me.ComBox_Hojas.FormattingEnabled = True
        Me.ComBox_Hojas.Location = New System.Drawing.Point(129, 35)
        Me.ComBox_Hojas.Name = "ComBox_Hojas"
        Me.ComBox_Hojas.Size = New System.Drawing.Size(138, 21)
        Me.ComBox_Hojas.TabIndex = 0
        '
        'UF_RLUV
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(519, 332)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btn_Aceptar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "UF_RLUV"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "REGRESION POR MCO PARA UNA VARIABLE INDEPENDIENTE"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btn_Aceptar As Windows.Forms.Button
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents ComBox_Hojas As Windows.Forms.ComboBox
    Friend WithEvents txt_SigX As Windows.Forms.TextBox
    Friend WithEvents txt_SigY As Windows.Forms.TextBox
    Friend WithEvents txt_RangoX As Windows.Forms.TextBox
    Friend WithEvents txt_RangoY As Windows.Forms.TextBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
End Class
