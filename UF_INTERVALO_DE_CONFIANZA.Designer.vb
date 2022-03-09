<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UF_INTERVALO_DE_CONFIANZA
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_NvSf = New System.Windows.Forms.TextBox()
        Me.ComB_Betas = New System.Windows.Forms.ComboBox()
        Me.btn_Aceptar = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txt_NvSf)
        Me.GroupBox1.Controls.Add(Me.ComB_Betas)
        Me.GroupBox1.Location = New System.Drawing.Point(27, 25)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 132)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "DATOS"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "NIVEL DE SIGNIFICANCIA"
        '
        'txt_NvSf
        '
        Me.txt_NvSf.Location = New System.Drawing.Point(19, 92)
        Me.txt_NvSf.Name = "txt_NvSf"
        Me.txt_NvSf.Size = New System.Drawing.Size(159, 20)
        Me.txt_NvSf.TabIndex = 1
        '
        'ComB_Betas
        '
        Me.ComB_Betas.FormattingEnabled = True
        Me.ComB_Betas.Location = New System.Drawing.Point(19, 28)
        Me.ComB_Betas.Name = "ComB_Betas"
        Me.ComB_Betas.Size = New System.Drawing.Size(159, 21)
        Me.ComB_Betas.TabIndex = 0
        '
        'btn_Aceptar
        '
        Me.btn_Aceptar.Location = New System.Drawing.Point(62, 163)
        Me.btn_Aceptar.Name = "btn_Aceptar"
        Me.btn_Aceptar.Size = New System.Drawing.Size(114, 38)
        Me.btn_Aceptar.TabIndex = 1
        Me.btn_Aceptar.Text = "ACEPTAR"
        Me.btn_Aceptar.UseVisualStyleBackColor = True
        '
        'UF_INTERVALO_DE_CONFIANZA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(260, 213)
        Me.Controls.Add(Me.btn_Aceptar)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "UF_INTERVALO_DE_CONFIANZA"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "INTERVALO DE CONFIANZA"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents txt_NvSf As Windows.Forms.TextBox
    Friend WithEvents ComB_Betas As Windows.Forms.ComboBox
    Friend WithEvents btn_Aceptar As Windows.Forms.Button
End Class
