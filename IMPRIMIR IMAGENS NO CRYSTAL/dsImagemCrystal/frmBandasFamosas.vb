Imports System.io
Public Class frmImagem
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)

        If disposing Then

            If Not (components Is Nothing) Then
                components.Dispose()
            End If

        End If

        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents crystalReportViewer1 As _
            CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnCarregaImagem As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmImagem))
        Me.crystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnCarregaImagem = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'crystalReportViewer1
        '
        Me.crystalReportViewer1.ActiveViewIndex = -1
        Me.crystalReportViewer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crystalReportViewer1.Location = New System.Drawing.Point(0, 8)
        Me.crystalReportViewer1.Name = "crystalReportViewer1"
        Me.crystalReportViewer1.ReportSource = Nothing
        Me.crystalReportViewer1.Size = New System.Drawing.Size(616, 280)
        Me.crystalReportViewer1.TabIndex = 0
        '
        'btnCarregaImagem
        '
        Me.btnCarregaImagem.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.btnCarregaImagem.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCarregaImagem.Image = CType(resources.GetObject("btnCarregaImagem.Image"), System.Drawing.Image)
        Me.btnCarregaImagem.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnCarregaImagem.Location = New System.Drawing.Point(0, 222)
        Me.btnCarregaImagem.Name = "btnCarregaImagem"
        Me.btnCarregaImagem.Size = New System.Drawing.Size(616, 72)
        Me.btnCarregaImagem.TabIndex = 1
        Me.btnCarregaImagem.Text = "Carregar Imagem do Disco"
        Me.btnCarregaImagem.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmImagem
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(616, 294)
        Me.Controls.Add(Me.btnCarregaImagem)
        Me.Controls.Add(Me.crystalReportViewer1)
        Me.Name = "frmImagem"
        Me.Text = "Carregando Imagens do disco no Crystal Reports"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnLoadImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCarregaImagem.Click

        Dim ofdImagem As New OpenFileDialog

        If ofdImagem.ShowDialog = DialogResult.OK Then

            Try

                'Cria uma nova instância do dataset
                Dim dsBandas As New dsBandasFamosas

                'Cria uma nova instância do relatório
                Dim rptBandas As New rptBandasFamosas

                'Cria uma nova linha no dataset
                Dim dr As dsBandasFamosas.ImageRow = dsBandas.Image.NewImageRow


                'Define os elementos da linha
                dr.Name = ofdImagem.FileName
                'preenche o campo com o byte array
                dr.Photo = carregaImagem(ofdImagem.FileName)

                'Inclui a nova linha no dataset
                dsBandas.Image.Rows.Add(dr)

                'Usa o dataset como fonte de dados para o relatório
                rptBandas.SetDataSource(dsBandas)

                'Exibe o relatório no reportviewer
                Me.crystalReportViewer1.ReportSource = rptBandas


            Catch ex As Exception
                'captura alguma exceção 
                MessageBox.Show("Alguma coisa deu errado, verifique... : " & ex.Message)

            End Try

        End If

    End Sub

    Private Function carregaImagem(ByVal fileName As String) As Byte()

        'Método para carregar uma imagem do disco e retorná-la como um byteStream
        Dim fs As FileStream = New FileStream(fileName, FileMode.Open, FileAccess.Read)
        Dim br As BinaryReader = New BinaryReader(fs)
        Return (br.ReadBytes(Convert.ToInt32(br.BaseStream.Length)))

    End Function

End Class
