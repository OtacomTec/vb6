VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C5014412-BD55-402F-8335-07C273732964}#1.1#0"; "AplicativoUsuário.ocx"
Begin VB.Form frmRelatorio_Bombas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Bombas"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelatorio_Bombas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6225
   Begin VB.Frame Frame3 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   90
      TabIndex        =   5
      Top             =   660
      Width           =   6045
      Begin VB.TextBox txtBomba 
         Height          =   360
         Left            =   105
         TabIndex        =   1
         ToolTipText     =   "Código da Bomba"
         Top             =   585
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo dtcBomba 
         Height          =   360
         Left            =   1335
         TabIndex        =   2
         Top             =   585
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bomba"
         Height          =   240
         Left            =   105
         TabIndex        =   6
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Cancel          =   -1  'True
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Visualiza Impressão"
      Top             =   1890
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Limpa os Filtros"
      Top             =   1890
      Width           =   1245
   End
   Begin AplicativoUsuárioOCX.AplicativoUsuário ocxUsuario 
      Left            =   6840
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6930
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Bombas.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Bombas.frx":17E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Bombas.frx":183E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Bombas.frx":189C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Bombas.frx":18FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Bombas.frx":1958
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   30
      Width           =   1290
   End
End
Attribute VB_Name = "frmRelatorio_Bombas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Compras                                                        '
' Objetivo...............: Relatório de Bombas                               '
' Data de Criação........: 20/05/2004                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim booAlterar As Boolean
Public strSql As String
Dim log As New DLLSystemManager.log
'Declaração da variavel do intercomunicador de mensagens
Private Cliente_mensagem_exe As VetorDeMensagens.ClienteDeMensagens

Private Sub cmdCancelar_Click()
    Call Objetos.Limpa_TXT(Me)
    txtBomba.SetFocus
End Sub

Private Sub dtcBomba_GotFocus()
    Call Movimentacoes.Verifica_DataCombo(dtcBomba)
End Sub

Private Sub dtcBomba_LostFocus()
    txtBomba.Text = dtcBomba.BoundText
    If IsNumeric(txtBomba.Text) = False Or dtcBomba.Text = Empty Then txtBomba.Text = Empty: Exit Sub
End Sub

Private Sub cmdImprimir_Click()
    frmAguarde.Show
    DoEvents
    Call Impressao
    Unload frmAguarde
End Sub

Private Sub dtcEmpresa_LostFocus()
    If dtcEmpresa.BoundText <> Empty Then
       strSql = "SELECT IXCodigo_Bomba,DFDescricao_TBBomba FROM TBBomba " & _
                "INNER JOIN TBBomba_bico ON TBBomba_bico.FKId_TBBomba = TBBomba.PKId_TBBomba " & _
                "INNER JOIN TBTanque ON TBTanque.PKId_TBTanque = TBBomba_bico.FKId_TBTanque " & _
                "WHERE TBTanque.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
                "GROUP BY TBBomba.IXCodigo_Bomba,TBBomba.DFDescricao_TBBomba"
                
       Movimentacoes.Movimenta_DataCombo "IXCodigo_Bomba", "DFDescricao_TBBomba", dtcBomba, strSql, "BDRetaguarda", "Otica", Me
    Else
       strSql = "SELECT IXCodigo_Bomba,DFDescricao_TBBomba FROM TBBomba "
       
       Movimentacoes.Movimenta_DataCombo "IXCodigo_Bomba", "DFDescricao_TBBomba", dtcBomba, strSql, "BDRetaguarda", "Otica", Me
    End If
    txtBomba.Text = Empty
    dtcEmpresa.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = "113" Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
  
Private Sub Form_Load()
    On Error GoTo erro
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Relatório de Bombas"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Relatório de Bombas"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
     
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
       
    strSql = "SELECT IXCodigo_Bomba,DFDescricao_TBBomba FROM TBBomba " & _
             "INNER JOIN TBBomba_bico ON TBBomba_bico.FKId_TBBomba = TBBomba.PKId_TBBomba " & _
             "INNER JOIN TBTanque ON TBTanque.PKId_TBTanque = TBBomba_bico.FKId_TBTanque " & _
             "WHERE TBTanque.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
             "GROUP BY TBBomba.IXCodigo_Bomba,TBBomba.DFDescricao_TBBomba"
             
    Movimentacoes.Movimenta_DataCombo "IXCodigo_Bomba", "DFDescricao_TBBomba", dtcBomba, strSql, "BDRetaguarda", "Otica", Me
                 
    Exit Sub
erro:
    Call erro.erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando Relatório de Bombas"
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Exit Sub
erro:
    Call erro.erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Function Cancelar()
    On Error GoTo erro
    
    Call Objetos.Limpa_TXT(Me)
            
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento Relatório de Bombas"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Sub txtBomba_Change()
    dtcBomba.BoundText = txtBomba.Text
    If IsNumeric(txtBomba.Text) = False Then txtBomba.Text = Empty: Exit Sub
End Sub

Private Sub txtBomba_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Function Impressao()
    strSql = Empty
    
    strSql = "SELECT TBBomba.IXCodigo_Bomba," & _
            "TBBomba.DFDescricao_TBBomba," & _
            "TBBomba.DFNumero_bicos_TBBomba," & _
            "TBBomba_bico.PKId_TBBomba_bico," & _
            "TBBomba_bico.FKId_TBProduto," & _
            "TBBomba_bico.FKId_TBTanque," & _
            "TBBomba_bico.IXCodigo_TBBomba_bico," & _
            "TBBomba_bico.DFUltimo_encerrante_TBBomba_bico," & _
            "TBBomba_bico.DFNumero_maximo_encerrante_TBBomba_bico," & _
            "TBBomba_bico.DFTipo_preco_TBBomba_bico," & _
            "TBProduto.IXCodigo_TBProduto," & _
            "TBProduto.DFDescricao_TBProduto," & _
            "TBItens_tabela_preco.DFPreco_avista_TBItens_tabela_preco," & _
            "TBItens_tabela_preco.DFPreco_promocao_TBItens_tabela_preco," & _
            "TBItens_tabela_preco.DFPreco_revenda_TBItens_tabela_preco," & _
            "TBItens_tabela_preco.DFPreco_especial_TBItens_tabela_preco," & _
            "TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco " & _
            "FROM TBBomba_bico " & _
            "INNER JOIN TBBomba " & _
            "ON TBBomba_bico.FKId_TBBomba = TBBomba.PKId_TBBomba " & _
            "INNER JOIN TBProduto " & _
            "ON TBBomba_bico.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
            "INNER JOIN TBEmpresa " & _
            "ON TBProduto.IXCodigo_TBEmpresa  = TBEmpresa.PKCodigo_TBEmpresa " & _
            "INNER JOIN TBParametros_venda "

   strSql = strSql + "ON TBEmpresa.PKCodigo_TBEmpresa = TBParametros_venda.IXCodigo_TBEmpresa " & _
            "INNER JOIN TBTabela_preco " & _
            "ON  DFNumero_tabela_vigente_TBParametros_venda = TBTabela_preco.PKCodigo_TBTabela_preco " & _
            "INNER JOIN TBItens_tabela_preco " & _
            "ON TBProduto.PKId_TBProduto = TBItens_tabela_preco.FKId_TBProduto " & _
            "WHERE TBItens_tabela_preco.FKCodigo_TBTabela_preco = DFNumero_tabela_vigente_TBParametros_venda "

   If dtcEmpresa.BoundText <> "" Then
      strSql = strSql + " AND TBEmpresa.PKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
   End If
   
   If dtcBomba.BoundText <> "" Then
      strSql = strSql + " AND TBBomba.IXCodigo_Bomba = " & dtcBomba.BoundText & " "
   End If
   
   strSql = strSql + " ORDER BY  IXCodigo_Bomba"
   
   Call frmConsole_Relatorio_Bombas.Show
   
End Function
