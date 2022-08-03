<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UF_GGD
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
        Me.txt_SigX = New System.Windows.Forms.TextBox()
        Me.txt_SigY = New System.Windows.Forms.TextBox()
        Me.txt_RgX = New System.Windows.Forms.TextBox()
        Me.txt_RgY = New System.Windows.Forms.TextBox()
        Me.ComboBoxHojasActivas = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_Aceptar = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txt_SigX)
        Me.GroupBox1.Controls.Add(Me.txt_SigY)
        Me.GroupBox1.Controls.Add(Me.txt_RgX)
        Me.GroupBox1.Controls.Add(Me.txt_RgY)
        Me.GroupBox1.Controls.Add(Me.ComboBoxHojasActivas)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(27, 34)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(392, 207)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "VARIABLES"
        '
        'txt_SigX
        '
        Me.txt_SigX.Location = New System.Drawing.Point(130, 151)
        Me.txt_SigX.Name = "txt_SigX"
        Me.txt_SigX.Size = New System.Drawing.Size(240, 22)
        Me.txt_SigX.TabIndex = 9
        Me.txt_SigX.Text = "NA"
        '
        'txt_SigY
        '
        Me.txt_SigY.Location = New System.Drawing.Point(130, 120)
        Me.txt_SigY.Name = "txt_SigY"
        Me.txt_SigY.Size = New System.Drawing.Size(240, 22)
        Me.txt_SigY.TabIndex = 8
        Me.txt_SigY.Text = "NA"
        '
        'txt_RgX
        '
        Me.txt_RgX.Font = New System.Drawing.Font("Times New Roman", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_RgX.Location = New System.Drawing.Point(130, 92)
        Me.txt_RgX.Name = "txt_RgX"
        Me.txt_RgX.Size = New System.Drawing.Size(240, 21)
        Me.txt_RgX.TabIndex = 7
        Me.txt_RgX.Text = "Haz doble click para seleccionar los datos."
        '
        'txt_RgY
        '
        Me.txt_RgY.Font = New System.Drawing.Font("Times New Roman", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_RgY.Location = New System.Drawing.Point(130, 60)
        Me.txt_RgY.Name = "txt_RgY"
        Me.txt_RgY.Size = New System.Drawing.Size(240, 21)
        Me.txt_RgY.TabIndex = 6
        Me.txt_RgY.Text = "Haz doble click para seleccionar los datos."
        '
        'ComboBoxHojasActivas
        '
        Me.ComboBoxHojasActivas.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxHojasActivas.FormattingEnabled = True
        Me.ComboBoxHojasActivas.Location = New System.Drawing.Point(130, 29)
        Me.ComboBoxHojasActivas.Name = "ComboBoxHojasActivas"
        Me.ComboBoxHojasActivas.Size = New System.Drawing.Size(126, 23)
        Me.ComboBoxHojasActivas.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(19, 124)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(70, 14)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Significado Y:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(19, 155)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 14)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Signidicado X:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(19, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 14)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Rango X:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(19, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 14)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Rango Y:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(19, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Hoja Datos:"
        '
        'btn_Aceptar
        '
        Me.btn_Aceptar.Font = New System.Drawing.Font("Times New Roman", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Aceptar.Location = New System.Drawing.Point(437, 12)
        Me.btn_Aceptar.Name = "btn_Aceptar"
        Me.btn_Aceptar.Size = New System.Drawing.Size(110, 39)
        Me.btn_Aceptar.TabIndex = 1
        Me.btn_Aceptar.Text = "ACEPTAR"
        Me.btn_Aceptar.UseVisualStyleBackColor = True
        '
        'UF_GGD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(559, 257)
        Me.Controls.Add(Me.btn_Aceptar)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "UF_GGD"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "GRAFICO DE DISPERSION"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents txt_SigX As Windows.Forms.TextBox
    Friend WithEvents txt_SigY As Windows.Forms.TextBox
    Friend WithEvents txt_RgX As Windows.Forms.TextBox
    Friend WithEvents txt_RgY As Windows.Forms.TextBox
    Friend WithEvents ComboBoxHojasActivas As Windows.Forms.ComboBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btn_Aceptar As Windows.Forms.Button
End Class
