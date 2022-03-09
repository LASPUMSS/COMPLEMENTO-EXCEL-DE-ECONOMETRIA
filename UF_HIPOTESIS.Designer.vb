<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UF_HIPOTESIS
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_Hipotesis = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txt_Hp = New System.Windows.Forms.TextBox()
        Me.txt_NvSg = New System.Windows.Forms.TextBox()
        Me.ComB_Betas = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btn_ProbarHipotesis = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(37, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "HIPOTESIS:"
        '
        'txt_Hipotesis
        '
        Me.txt_Hipotesis.Location = New System.Drawing.Point(123, 30)
        Me.txt_Hipotesis.Name = "txt_Hipotesis"
        Me.txt_Hipotesis.Size = New System.Drawing.Size(363, 20)
        Me.txt_Hipotesis.TabIndex = 1
        Me.txt_Hipotesis.Text = "N/A"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txt_Hp)
        Me.GroupBox1.Controls.Add(Me.txt_NvSg)
        Me.GroupBox1.Controls.Add(Me.ComB_Betas)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(26, 76)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(475, 167)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "DATOS"
        '
        'txt_Hp
        '
        Me.txt_Hp.Location = New System.Drawing.Point(119, 96)
        Me.txt_Hp.Name = "txt_Hp"
        Me.txt_Hp.Size = New System.Drawing.Size(126, 20)
        Me.txt_Hp.TabIndex = 5
        '
        'txt_NvSg
        '
        Me.txt_NvSg.Location = New System.Drawing.Point(119, 53)
        Me.txt_NvSg.Name = "txt_NvSg"
        Me.txt_NvSg.Size = New System.Drawing.Size(126, 20)
        Me.txt_NvSg.TabIndex = 4
        '
        'ComB_Betas
        '
        Me.ComB_Betas.FormattingEnabled = True
        Me.ComB_Betas.Location = New System.Drawing.Point(327, 53)
        Me.ComB_Betas.Name = "ComB_Betas"
        Me.ComB_Betas.Size = New System.Drawing.Size(115, 21)
        Me.ComB_Betas.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(267, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Para:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(22, 99)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(18, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "H:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(22, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Significancia:"
        '
        'btn_ProbarHipotesis
        '
        Me.btn_ProbarHipotesis.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_ProbarHipotesis.Location = New System.Drawing.Point(181, 259)
        Me.btn_ProbarHipotesis.Name = "btn_ProbarHipotesis"
        Me.btn_ProbarHipotesis.Size = New System.Drawing.Size(159, 28)
        Me.btn_ProbarHipotesis.TabIndex = 3
        Me.btn_ProbarHipotesis.Text = "PROBAR HIPOTESIS"
        Me.btn_ProbarHipotesis.UseVisualStyleBackColor = True
        '
        'UF_HIPOTESIS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(536, 305)
        Me.Controls.Add(Me.btn_ProbarHipotesis)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txt_Hipotesis)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "UF_HIPOTESIS"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "HIPOTESIS SIMPLES"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents txt_Hipotesis As Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents txt_Hp As Windows.Forms.TextBox
    Friend WithEvents txt_NvSg As Windows.Forms.TextBox
    Friend WithEvents ComB_Betas As Windows.Forms.ComboBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents btn_ProbarHipotesis As Windows.Forms.Button
End Class
