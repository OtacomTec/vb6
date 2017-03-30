VERSION 5.00
Begin VB.Form FrmTira_Teima 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tira Teima"
   ClientHeight    =   5190
   ClientLeft      =   1725
   ClientTop       =   900
   ClientWidth     =   10965
   Icon            =   "FrmTira_Teima.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3090
      Top             =   90
   End
   Begin VB.TextBox txtPreco_Unitario 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3420
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4290
      Width           =   3165
   End
   Begin VB.TextBox txtPreco_Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   7290
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4260
      Width           =   3345
   End
   Begin VB.TextBox txtCodigo_Produto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   180
      MaxLength       =   14
      TabIndex        =   1
      Top             =   1620
      Width           =   3225
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   -90
      MaxLength       =   40
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8580
      Width           =   10875
   End
   Begin VB.TextBox txtQuantidade_Produto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   300
      TabIndex        =   3
      Top             =   4290
      Width           =   2325
   End
   Begin VB.TextBox txtDescricao_Produto 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   270
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2820
      Width           =   6285
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6810
      TabIndex        =   12
      Top             =   4380
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Preço Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7260
      TabIndex        =   11
      Top             =   3690
      Width           =   1785
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Preço Unitário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   10
      Top             =   3720
      Width           =   1980
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2850
      TabIndex        =   9
      Top             =   4380
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   8
      Top             =   2250
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   525
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   3285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código (F2 Consulta)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   1140
      Width           =   2940
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   525
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   1650
      Width           =   3285
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   6
      Top             =   3720
      Width           =   1620
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2670
      Width           =   6345
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   4140
      Width           =   2535
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   4140
      Width           =   3465
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   7200
      Shape           =   4  'Rounded Rectangle
      Top             =   4110
      Width           =   3675
   End
   Begin VB.Image imgProduto 
      Height          =   2475
      Left            =   3570
      Stretch         =   -1  'True
      Top             =   60
      Width           =   2955
   End
   Begin VB.Image imgLogo_Empresa 
      Height          =   2055
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   60
      Width           =   2955
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   10770
      Top             =   8580
      Width           =   525
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   11280
      Top             =   8580
      Width           =   525
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   10860
      Shape           =   3  'Circle
      Top             =   8610
      Width           =   225
   End
   Begin VB.Image imgInd_pouco_papel 
      Height          =   255
      Left            =   11430
      Picture         =   "FrmTira_Teima.frx":1782
      Stretch         =   -1  'True
      Top             =   8610
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   915
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4170
      Width           =   2535
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   3150
      Shape           =   4  'Rounded Rectangle
      Top             =   4140
      Width           =   3435
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   945
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   4140
      Width           =   3675
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   915
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2730
      Width           =   6315
   End
End
Attribute VB_Name = "FrmTira_Teima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Cadastro Base                                                  '
' Objetivo...............: Consulta Tira Teima                                            '
' Data de Criação........: 16/01/2005                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strTamanho As String
Dim strNomes As String
Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Dim booAlterar As Boolean
Public strSql As String
Dim conexao As New DLLConexao_Sistema.conexao
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim log As New DLLSystemManager.log

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       frmConsulta_Produto.Show
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
  
Private Sub Form_Load()
    On Error GoTo Erro
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Consulta Tira Teima"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando o cadastro de Parâmetros EOF"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    On Error Resume Next
    
    Dim rstBusca_Logo As New ADODB.Recordset
    strSql = Empty
    strSql = "SELECT * FROM TBEmpresa WHERE PKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " "
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Logo, "Otica", Me)
    
    imgLogo_Empresa.Picture = LoadPicture(rstBusca_Logo.Fields("DFPath_logomarca_TBEmpresa"))
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Set log = Nothing
    
    strCombo = Empty
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub Timer1_Timer()
    Call Objetos.Limpa_TXT(Me)
    Timer1.Enabled = False
    txtDescricao_Produto.Enabled = True
    txtQuantidade_Produto.Enabled = True
    txtPreco_Unitario.Enabled = True
    txtPreco_Total.Enabled = True
    imgProduto.Picture = LoadPicture()
    txtCodigo_Produto.SetFocus
End Sub

Private Sub txtCodigo_Produto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Produto_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_Produto_LostFocus()
    If txtCodigo_Produto.Text <> Empty Then
       Call Consulta
    End If
End Sub

Private Function Consulta()
    
    On Error Resume Next
    
    Dim strDigito_Peso_Variavel As String
    Dim strDigito_Produto_Digitado As String
    Dim strCodigo_Produto_Etiqueta As String
    Dim strID_Produto As String
    Dim strPreco_Peso_Parametro As String
    Dim strTabela_Vigente As String
    Dim rstBusca_Preco As New ADODB.Recordset
    Dim rstBusca_Paramentros As New ADODB.Recordset
    Dim rstBusca_Produto As New ADODB.Recordset
    Dim strPreco_Tabela As String
    Dim strTotal As String
    Dim strPreco_Peso As String
    Dim strDecimal As String
    Dim strQuantidade As String
       
    strSql = Empty
    strSql = "SELECT * FROM TBParametros_ecf WHERE FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ""
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Paramentros, "Otica", Me)
    
    strDigito_Peso_Variavel = rstBusca_Paramentros.Fields("DFCodigo_inicial_peso_variavel_TBParametros_ecf")
    If rstBusca_Paramentros.Fields("DFPreco_peso_balanca_TBParametros_ecf") = False Then
       strPreco_Peso_Parametro = 0
    Else
       strPreco_Peso_Parametro = 1
    End If
    Set rstBusca_Paramentros = Nothing
     
    strSql = Empty
    strSql = "SELECT * FROM TBParametros_venda WHERE IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ""
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Paramentros, "Otica", Me)
    
    strTabela_Vigente = rstBusca_Paramentros.Fields("DFNumero_tabela_vigente_TBParametros_venda")
    Set rstBusca_Paramentros = Nothing
        
    If Len(txtCodigo_Produto.Text) > 6 Then
       strDigito_Produto_Digitado = Left(txtCodigo_Produto.Text, 1)
       If strDigito_Peso_Variavel = strDigito_Produto_Digitado Then
          strCodigo_Produto_Etiqueta = Mid(txtCodigo_Produto.Text, 2, 4)
          strPreco_Peso = Mid(txtCodigo_Produto.Text, 6, 7)
          strSql = Empty
          strSql = "SELECT * FROM TBProduto WHERE IXCodigo_TBproduto = " & strCodigo_Produto_Etiqueta & " "
          Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Produto, "Otica", Me)
          If rstBusca_Produto.RecordCount = 0 Then
             MsgBox "Produto não Cadastrado, Verifique.", vbCritical, "Only Tech"
             txtCodigo_Produto.SetFocus
             Exit Function
          End If
          txtDescricao_Produto.Text = rstBusca_Produto.Fields("DFDescricao_resumida_TBProduto")
          imgProduto.Picture = LoadPicture(rstBusca_Produto.Fields("DFPath_imagem_TBProduto"))
          strID_Produto = Funcoes_Gerais.Localiza_ID("PKId_TBProduto", "IXCodigo_TBProduto", strCodigo_Produto_Etiqueta, "TBProduto", "Otica", Me, "BDRetaguarda")
          Set rstBusca_Produto = Nothing
          strSql = Empty
          strSql = "SELECT TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco " & _
                   "FROM TBItens_tabela_preco " & _
                   "WHERE FKCodigo_TBTabela_preco = " & strTabela_Vigente & " AND " & _
                   "FKId_TBProduto = " & strID_Produto & ""
          Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Preco, "Otica", Me)
          If rstBusca_Preco.RecordCount = 0 Then
             MsgBox "Produto não cadastrado na tabela de preço vigente.Verifique.", vbCritical, "Only Tech"
             txtCodigo_Produto.SetFocus
             Exit Function
          End If
          strPreco_Tabela = Format(rstBusca_Preco.Fields("DFPreco_varejo_TBItens_tabela_preco"), "#,###0.00")
          Set rstBusca_Preco = Nothing
          If strPreco_Peso_Parametro = 0 Then
             strPreco_Peso = Mid(txtCodigo_Produto.Text, 6, 5)
             strDecimal = Mid(txtCodigo_Produto.Text, 11, 2)
             strPreco_Peso = strPreco_Peso & "," & strDecimal
             strPreco_Peso = Format(strPreco_Peso, "#,###0.00")
             strQuantidade = CDbl(strPreco_Peso) / CDbl(strPreco_Tabela)
             strQuantidade = Format(strQuantidade, "#,###0.000")
             txtQuantidade_Produto.Text = strQuantidade
             txtPreco_Unitario.Text = strPreco_Tabela
             strTotal = CDbl(strPreco_Tabela) * CDbl(strQuantidade)
             strTotal = Format(strTotal, "#,###0.00")
             txtPreco_Total.Text = strTotal
          Else
             strPreco_Peso = Format(strPreco_Peso, "#,###0.000")
             strTotal = strPreco_Peso * strPreco_Tabela
             txtQuantidade_Produto.Text = strPreco_Peso
             txtPreco_Unitario.Text = Format(strPreco_Tabela, "#,###0.00")
             txtPreco_Total.Text = Format(strTotal, "#,###0.00")
          End If
       Else
          strID_Produto = Funcoes_Gerais.Localiza_ID("FKId_TBProduto", "IXCodigo_TBCodigo_barras", txtCodigo_Produto.Text, "TBCodigo_barras", "Otica", Me, "BDRetaguarda")
          strSql = Empty
          strSql = "SELECT TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco, " & _
                   "TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                   "FROM TBItens_tabela_preco " & _
                   "INNER JOIN TBProduto ON TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                   "WHERE FKCodigo_TBTabela_preco = " & strTabela_Vigente & " AND " & _
                   "FKId_TBProduto = " & strID_Produto & ""
          Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Preco, "Otica", Me)
          If rstBusca_Preco.RecordCount = 0 Then
             MsgBox "Produto não cadastrado na tabela de preço vigente.Verifique.", vbCritical, "Only Tech"
             txtCodigo_Produto.SetFocus
             Exit Function
          End If
          strPreco_Tabela = Format(rstBusca_Preco.Fields("DFPreco_varejo_TBItens_tabela_preco"), "#,###0.00")
          txtQuantidade_Produto.Text = 1
          txtPreco_Unitario.Text = strPreco_Tabela
          txtDescricao_Produto.Text = rstBusca_Preco.Fields("DFDescricao_resumida_TBProduto")
          imgProduto.Picture = LoadPicture(rstBusca_Preco.Fields("DFPath_imagem_TBProduto"))
          strTotal = CDbl(txtQuantidade_Produto.Text) * CDbl(txtPreco_Unitario.Text)
          strTotal = Format(strTotal, "#,###0.00")
          txtPreco_Total.Text = strTotal
          Set rstBusca_Preco = Nothing
       End If
    Else
       strID_Produto = Funcoes_Gerais.Localiza_ID("PKId_TBProduto", "IXCodigo_TBProduto", txtCodigo_Produto.Text, "TBProduto", "Otica", Me, "BDRetaguarda")
       strSql = Empty
       strSql = "SELECT TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco, " & _
                "TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                "FROM TBItens_tabela_preco " & _
                "INNER JOIN TBProduto ON TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                "WHERE FKCodigo_TBTabela_preco = " & strTabela_Vigente & " AND " & _
                "FKId_TBProduto = " & strID_Produto & ""
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Preco, "Otica", Me)
       If rstBusca_Preco.RecordCount = 0 Then
          MsgBox "Produto não cadastrado na tabela de preço vigente.Verifique.", vbCritical, "Only Tech"
          txtCodigo_Produto.SetFocus
          Exit Function
       End If
       strPreco_Tabela = Format(rstBusca_Preco.Fields("DFPreco_varejo_TBItens_tabela_preco"), "#,###0.00")
       txtPreco_Unitario.Text = strPreco_Tabela
       txtQuantidade_Produto.Text = 1
       txtDescricao_Produto.Text = rstBusca_Preco.Fields("DFDescricao_resumida_TBProduto")
       imgProduto.Picture = LoadPicture(rstBusca_Preco.Fields("DFPath_imagem_TBProduto"))
       strTotal = CDbl(txtQuantidade_Produto.Text) * CDbl(txtPreco_Unitario.Text)
       strTotal = Format(strTotal, "#,###0.00")
       txtPreco_Total.Text = strTotal
       Set rstBusca_Preco = Nothing
    End If
    Timer1.Enabled = True
    txtDescricao_Produto.Enabled = False
    txtQuantidade_Produto.Enabled = False
    txtPreco_Unitario.Enabled = False
    txtPreco_Total.Enabled = False
End Function
Private Sub txtPreco_Unitario_LostFocus()
   txtPreco_Unitario.Text = Format(txtPreco_Unitario.Text, "#,###0.00")
End Sub
