VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C5014412-BD55-402F-8335-07C273732964}#1.1#0"; "AplicativoUsuário.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatorio_Encerrante 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Encerrantes"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelatorio_Encerrante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6255
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
      TabIndex        =   7
      Top             =   660
      Width           =   6045
      Begin VB.TextBox txtOperador 
         Height          =   360
         Left            =   105
         TabIndex        =   1
         ToolTipText     =   "Código do Operador"
         Top             =   585
         Width           =   1155
      End
      Begin MSDataListLib.DataCombo dtcOperador 
         Height          =   360
         Left            =   1305
         TabIndex        =   2
         Top             =   585
         Width           =   4635
         _ExtentX        =   8176
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
         Caption         =   "Operador"
         Height          =   240
         Left            =   105
         TabIndex        =   8
         Top             =   330
         Width           =   810
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
      TabIndex        =   5
      ToolTipText     =   "Visualiza Impressão"
      Top             =   2010
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
      TabIndex        =   6
      ToolTipText     =   "Limpa os Filtros"
      Top             =   2010
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
            Picture         =   "frmRelatorio_Encerrante.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Encerrante.frx":17E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Encerrante.frx":183E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Encerrante.frx":189C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Encerrante.frx":18FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRelatorio_Encerrante.frx":1958
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicial 
      Height          =   360
      Left            =   60
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   8388608
      Format          =   50528257
      CurrentDate     =   37881
   End
   Begin MSComCtl2.DTPicker dtpFinal 
      Height          =   360
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   8388608
      Format          =   50528257
      CurrentDate     =   37881
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
      TabIndex        =   11
      Top             =   30
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   240
      Left            =   90
      TabIndex        =   10
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "até"
      Height          =   240
      Left            =   1590
      TabIndex        =   9
      Top             =   2160
      Width           =   270
   End
End
Attribute VB_Name = "frmRelatorio_Encerrante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Faturamento                                                    '
' Objetivo...............: Relatório Relação de Encerrantes                               '
' Data de Criação........: 30/04/04                                                       '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim booAlterar As Boolean
Public strSQL As String
Dim log As New DLLSystemManager.log
'Declaração da variavel do intercomunicador de mensagens
Private Cliente_mensagem_exe As VetorDeMensagens.ClienteDeMensagens

Private Sub cmdCancelar_Click()
    Call Cancelar
End Sub

Private Sub cmdImprimir_Click()

    If dtpInicial.Value > dtpFinal.Value Then
       MsgBox "Data final menor que data Inicial.Verifique!", vbInformation, "Only Tech"
       Exit Sub
    End If
    
    frmAguarde.Show
    DoEvents
    Call Impressao
    Unload frmAguarde
End Sub

Private Sub dtcEmpresa_LostFocus()
    If dtcEmpresa.BoundText <> Empty Then
        strSQL = "SELECT PKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf FROM TBOperadores_ecf WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
        Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSQL, "BDRetaguarda", "Otica", Me
    Else
        strSQL = "SELECT PKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf FROM TBOperadores_ecf "
        Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSQL, "BDRetaguarda", "Otica", Me
    End If
    txtOperador.Text = Empty
    dtcEmpresa.Enabled = False
End Sub

Private Sub dtcOperador_GotFocus()
    If txtOperador.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcOperador.Text)
    End If
End Sub

Private Sub dtcOperador_LostFocus()
    txtOperador.Text = dtcOperador.BoundText
    If IsNumeric(txtOperador.Text) = False Or dtcOperador.Text = Empty Then txtOperador.Text = Empty: Exit Sub
End Sub

Private Sub dtpFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

Private Sub dtpInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = "113" Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
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
    log.Programa = "Relatório de Encerrantes"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
        Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando o Relatório Relação de Entregas"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    'Montando os datacombo de tela
    strSQL = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSQL, "BDRetaguarda", "Otica", Me
    
    'dtcCodigo_empresa.boundtext = ---- Inserir aqui informações da DLLIntercomunicador de EXE's
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    strSQL = "SELECT PKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf FROM TBOperadores_ecf WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSQL, "BDRetaguarda", "Otica", Me

    dtpInicial.Value = Date - 10
    dtpFinal.Value = Date

    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando a Relação de Encerrantes"
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    'Call Limpa_Combos
            
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento do Relatório de Encerrantes"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    dtpInicial.Value = Date
    dtpFinal.Value = Date + 10
    txtOperador.SetFocus
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Impressao()

    strSQL = Empty
    strSQL = "SELECT TBEncerrante.PKId_TBEncerrante," & _
             "TBEncerrante.FKCodigo_TBPdv," & _
             "TBEncerrante.FKCodigo_TBOperadores_ecf," & _
             "TBEncerrante.DFData_TBEncerrante," & _
             "TBEncerrante.DFHora_TBEncerrante," & _
             "TBEncerrante.DFAbertura_fechamento_TBEncerrante," & _
             "TBEncerrante_Bomba.PKId_TBEncerrante_Bomba," & _
             "TBEncerrante_Bomba.FKId_TBBomba_bico," & _
             "TBEncerrante_Bomba.FKId_TBEncerrante," & _
             "TBEncerrante_Bomba.DFEncerrante_TBEncerrante_Bomba," & _
             "TBOperadores_ecf.PKCodigo_TBOperadores_ecf," & _
             "TBOperadores_ecf.DFNome_TBOperadores_ecf," & _
             "TBBomba.IXCodigo_Bomba," & _
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
             "TBItens_tabela_preco.DFPreco_avista_TBItens_tabela_preco,"
            
    strSQL = strSQL + "TBItens_tabela_preco.DFPreco_promocao_TBItens_tabela_preco," & _
            "TBItens_tabela_preco.DFPreco_revenda_TBItens_tabela_preco," & _
            "TBItens_tabela_preco.DFPreco_especial_TBItens_tabela_preco," & _
            "TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco," & _
            "Cast(TBEncerrante_Bomba.FKId_TBBomba_bico AS varchar(128))+ '-' + Cast (TBEncerrante.DFAbertura_fechamento_TBEncerrante AS varchar(128)) as Quebra " & _
            "FROM TBEncerrante " & _
            "INNER JOIN TBEncerrante_Bomba " & _
            "ON TBEncerrante.PKId_TBEncerrante = TBEncerrante_Bomba.FKId_TBEncerrante " & _
            "INNER JOIN TBOperadores_ecf " & _
            "ON TBEncerrante.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
            "INNER JOIN TBBomba_bico " & _
            "ON TBEncerrante_Bomba.FKId_TBBomba_bico = TBBomba_bico.FKId_TBBomba " & _
            "INNER JOIN TBBomba " & _
            "ON TBBomba_bico.FKId_TBBomba = TBBomba.PKId_TBBomba " & _
            "INNER JOIN TBProduto " & _
            "ON TBBomba_bico.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
            "INNER JOIN TBEmpresa " & _
            "ON  TBProduto.IXCodigo_TBEmpresa  = TBEmpresa.PKCodigo_TBEmpresa " & _
            "INNER JOIN TBParametros_venda " & _
            "ON TBEmpresa.PKCodigo_TBEmpresa = TBParametros_venda.IXCodigo_TBEmpresa " & _
            "INNER JOIN TBTabela_preco " & _
            "ON  DFNumero_tabela_vigente_TBParametros_venda = TBTabela_preco.PKCodigo_TBTabela_preco " & _
            "INNER JOIN TBItens_tabela_preco " & _
            "ON TBProduto.PKId_TBProduto = TBItens_tabela_preco.FKId_TBProduto " & _
            "WHERE TBItens_tabela_preco.FKCodigo_TBTabela_preco = DFNumero_tabela_vigente_TBParametros_venda "
            
    strSQL = strSQL + "AND DFData_TBEncerrante >= '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
            "AND DFData_TBEncerrante <=  '" & Format(dtpFinal, "YYYYMMDD") & "' "
   
    If dtcEmpresa.BoundText <> "" Then
       strSQL = strSQL + " AND TBEmpresa.PKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
    End If
            
    If dtcOperador.BoundText <> "" Then
       strSQL = strSQL + " AND FKCodigo_TBOperadores_ecf = " & dtcOperador.BoundText & " "
    End If

    strSQL = strSQL + " ORDER BY  Quebra"
            
    Call frmConsole_Relatorio_Encerrantes.Show
    
End Function

Private Sub txtOperador_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtOperador_Change()
    dtcOperador.BoundText = txtOperador.Text
    If IsNumeric(txtOperador.Text) = False Then txtOperador.Text = Empty: Exit Sub
End Sub

