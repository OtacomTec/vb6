<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnCaptura = New System.Windows.Forms.Button
        Me.chkIncluiFormulario = New System.Windows.Forms.CheckBox
        Me.picCopiaTela = New System.Windows.Forms.PictureBox
        CType(Me.picCopiaTela, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCaptura
        '
        Me.btnCaptura.Location = New System.Drawing.Point(2, 468)
        Me.btnCaptura.Name = "btnCaptura"
        Me.btnCaptura.Size = New System.Drawing.Size(232, 29)
        Me.btnCaptura.TabIndex = 0
        Me.btnCaptura.Text = "Capturar a tela"
        Me.btnCaptura.UseVisualStyleBackColor = True
        '
        'chkIncluiFormulario
        '
        Me.chkIncluiFormulario.AutoSize = True
        Me.chkIncluiFormulario.Location = New System.Drawing.Point(465, 468)
        Me.chkIncluiFormulario.Name = "chkIncluiFormulario"
        Me.chkIncluiFormulario.Size = New System.Drawing.Size(125, 17)
        Me.chkIncluiFormulario.TabIndex = 1
        Me.chkIncluiFormulario.Text = "Incluir este formulário"
        Me.chkIncluiFormulario.UseVisualStyleBackColor = True
        '
        'picCopiaTela
        '
        Me.picCopiaTela.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picCopiaTela.Location = New System.Drawing.Point(2, 12)
        Me.picCopiaTela.Name = "picCopiaTela"
        Me.picCopiaTela.Size = New System.Drawing.Size(588, 450)
        Me.picCopiaTela.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picCopiaTela.TabIndex = 2
        Me.picCopiaTela.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(596, 499)
        Me.Controls.Add(Me.picCopiaTela)
        Me.Controls.Add(Me.chkIncluiFormulario)
        Me.Controls.Add(Me.btnCaptura)
        Me.Name = "Form1"
        Me.Text = "Capturando Telas"
        CType(Me.picCopiaTela, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCaptura As System.Windows.Forms.Button
    Friend WithEvents chkIncluiFormulario As System.Windows.Forms.CheckBox
    Friend WithEvents picCopiaTela As System.Windows.Forms.PictureBox

End Class
