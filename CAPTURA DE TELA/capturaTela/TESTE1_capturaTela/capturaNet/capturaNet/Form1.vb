Public Class Form1

    Private Sub btnCaptura_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCaptura.Click
        picCopiaTela.Image = capturaTela()
    End Sub
    Private Sub chkIncluiFormulario_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncluiFormulario.CheckedChanged
        If (chkIncluiFormulario.Checked = True) Then
            Me.Opacity = 1.0
        Else
            Me.Opacity = 0.99
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ' ----- Selecionando a imagem
        Dim locateFile As New OpenFileDialog

        ' ----- Solicita o arquivo inicial
        locateFile.Filter = "JPG files (*.jpg)|*.jpg"
        locateFile.Multiselect = False

        If (locateFile.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            ' ----- Mostra a imagem selecionada
            picCopiaTela.Load(locateFile.FileName)
            Me.AutoScroll = True
        Else
            ' ----- sai do programa
            Me.Close()
        End If

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' ----- Selecionando a imagem
        'Dim locateFile As New OpenFileDialog

        '' ----- Solicita o arquivo inicial
        'locateFile.Filter = "JPG files (*.jpg)|*.jpg"
        'locateFile.Multiselect = False

        'If (locateFile.ShowDialog() = Windows.Forms.DialogResult.OK) Then
        '    ' ----- Mostra a imagem selecionada
        '    picCopiaTela.Load(locateFile.FileName)
        '    Me.AutoScroll = True
        'Else
        '    ' ----- sai do programa
        '    Me.Close()
        'End If

    End Sub
End Class
