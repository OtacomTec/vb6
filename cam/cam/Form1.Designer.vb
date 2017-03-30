<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lstVideoSources = New System.Windows.Forms.ListBox
        Me.Source = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnStartRecording = New System.Windows.Forms.Button
        Me.btnStopRecording = New System.Windows.Forms.Button
        Me.btnStopCamera = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lstVideoSources
        '
        Me.lstVideoSources.FormattingEnabled = True
        Me.lstVideoSources.Location = New System.Drawing.Point(9, 53)
        Me.lstVideoSources.Name = "lstVideoSources"
        Me.lstVideoSources.Size = New System.Drawing.Size(682, 147)
        Me.lstVideoSources.TabIndex = 0
        '
        'Source
        '
        Me.Source.AutoSize = True
        Me.Source.Location = New System.Drawing.Point(13, 34)
        Me.Source.Name = "Source"
        Me.Source.Size = New System.Drawing.Size(41, 13)
        Me.Source.TabIndex = 1
        Me.Source.Text = "Source"
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(9, 237)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(678, 390)
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 221)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Vídeo"
        '
        'btnStartRecording
        '
        Me.btnStartRecording.Location = New System.Drawing.Point(262, 635)
        Me.btnStartRecording.Name = "btnStartRecording"
        Me.btnStartRecording.Size = New System.Drawing.Size(91, 32)
        Me.btnStartRecording.TabIndex = 4
        Me.btnStartRecording.Text = "Start recording"
        Me.btnStartRecording.UseVisualStyleBackColor = True
        '
        'btnStopRecording
        '
        Me.btnStopRecording.Location = New System.Drawing.Point(359, 635)
        Me.btnStopRecording.Name = "btnStopRecording"
        Me.btnStopRecording.Size = New System.Drawing.Size(161, 32)
        Me.btnStopRecording.TabIndex = 5
        Me.btnStopRecording.Text = "Stop Recording"
        Me.btnStopRecording.UseVisualStyleBackColor = True
        '
        'btnStopCamera
        '
        Me.btnStopCamera.Location = New System.Drawing.Point(526, 635)
        Me.btnStopCamera.Name = "btnStopCamera"
        Me.btnStopCamera.Size = New System.Drawing.Size(161, 32)
        Me.btnStopCamera.TabIndex = 6
        Me.btnStopCamera.Text = "Stop Camera"
        Me.btnStopCamera.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(9, 635)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(91, 32)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "conf"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(106, 635)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(62, 32)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "conf2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(174, 635)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(62, 32)
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "FOTO"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'PictureBox2
        '
        Me.PictureBox2.Location = New System.Drawing.Point(708, 53)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(524, 574)
        Me.PictureBox2.TabIndex = 10
        Me.PictureBox2.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1283, 679)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnStopCamera)
        Me.Controls.Add(Me.btnStopRecording)
        Me.Controls.Add(Me.btnStartRecording)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Source)
        Me.Controls.Add(Me.lstVideoSources)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstVideoSources As System.Windows.Forms.ListBox
    Friend WithEvents Source As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnStartRecording As System.Windows.Forms.Button
    Friend WithEvents btnStopRecording As System.Windows.Forms.Button
    Friend WithEvents btnStopCamera As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox

End Class
