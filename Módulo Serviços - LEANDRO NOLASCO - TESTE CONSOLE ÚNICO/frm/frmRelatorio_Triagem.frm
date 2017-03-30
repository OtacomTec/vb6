VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatorio_Triagem 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Triagem"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelatorio_Triagem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6360
   Begin MSComCtl2.DTPicker dtpPeriodo_Lancamento_Inicio 
      Height          =   375
      Left            =   90
      TabIndex        =   18
      ToolTipText     =   "Período de Lançamento (Início)"
      Top             =   5340
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49807361
      CurrentDate     =   38797
   End
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
      Height          =   3720
      Left            =   90
      TabIndex        =   24
      Top             =   660
      Width           =   6165
      Begin MSComCtl2.DTPicker dtpCompetencia_Ano 
         Height          =   360
         Left            =   4650
         TabIndex        =   11
         ToolTipText     =   "Competência (Ano)"
         Top             =   3210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   49807363
         CurrentDate     =   38797
      End
      Begin AutoCompletar.CbCompleta cbbCompetencia_Mes 
         Height          =   360
         Left            =   2490
         TabIndex        =   10
         ToolTipText     =   "Competência (Mês)"
         Top             =   3210
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
      End
      Begin VB.TextBox txtLote 
         Height          =   360
         Left            =   120
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "Código Lote"
         Top             =   3210
         Width           =   2325
      End
      Begin VB.TextBox txtFabricante 
         Height          =   360
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Código Fabricante"
         Top             =   2550
         Width           =   1365
      End
      Begin VB.TextBox txtInsumo 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Código Insumo"
         Top             =   1890
         Width           =   1365
      End
      Begin VB.TextBox txtRamo_Atividade 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Código Ramo Atividade (Convênio)"
         Top             =   570
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo_Cliente 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Código Cliente"
         Top             =   1230
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   360
         Left            =   1530
         TabIndex        =   4
         ToolTipText     =   "Descrição Cliente"
         Top             =   1230
         Width           =   4485
         _ExtentX        =   7911
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
      Begin MSDataListLib.DataCombo dtcRamo_Atividade 
         Height          =   360
         Left            =   1530
         TabIndex        =   2
         ToolTipText     =   "Descrição Ramo Atividade (Convênio)"
         Top             =   570
         Width           =   4485
         _ExtentX        =   7911
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
      Begin MSDataListLib.DataCombo dtcInsumo 
         Height          =   360
         Left            =   1530
         TabIndex        =   6
         ToolTipText     =   "Descrição Insumo"
         Top             =   1890
         Width           =   4485
         _ExtentX        =   7911
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
      Begin MSDataListLib.DataCombo dtcFabricante 
         Height          =   360
         Left            =   1530
         TabIndex        =   8
         ToolTipText     =   "Descrição Fabricante"
         Top             =   2550
         Width           =   4485
         _ExtentX        =   7911
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   2970
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Competência Mês / Ano"
         Height          =   240
         Left            =   2520
         TabIndex        =   30
         Top             =   2970
         Width           =   2040
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fabricante"
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   2310
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Insumo"
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   1650
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ramo Atividade(Convênio)"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   330
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   990
         Width           =   585
      End
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
      Left            =   5010
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5340
      Width           =   1245
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
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5340
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ativos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   90
      TabIndex        =   23
      Top             =   4410
      Width           =   3080
      Begin VB.OptionButton optAtivos_Sim 
         Caption         =   "Sim"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   885
      End
      Begin VB.OptionButton optAtivos_Nao 
         Caption         =   "Não"
         Height          =   240
         Left            =   1140
         TabIndex        =   13
         Top             =   330
         Width           =   675
      End
      Begin VB.OptionButton optAtivos_Todos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   2100
         TabIndex        =   14
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Bloqueados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3180
      TabIndex        =   22
      Top             =   4410
      Width           =   3080
      Begin VB.OptionButton optBloqueados_Nao 
         Caption         =   "Não"
         Height          =   240
         Left            =   1140
         TabIndex        =   16
         Top             =   330
         Width           =   645
      End
      Begin VB.OptionButton optBloqueados_Sim 
         Caption         =   "Sim"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   330
         Width           =   885
      End
      Begin VB.OptionButton optBloqueados_Todos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   2100
         TabIndex        =   17
         Top             =   330
         Width           =   855
      End
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   6165
      _ExtentX        =   10874
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
   Begin MSComCtl2.DTPicker dtpPeriodo_Lancamento_Fim 
      Height          =   375
      Left            =   2130
      TabIndex        =   19
      ToolTipText     =   "Período de Lançamento (Fim)"
      Top             =   5340
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49807361
      CurrentDate     =   38797
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Até"
      Height          =   240
      Left            =   1680
      TabIndex        =   33
      Top             =   5490
      Width           =   285
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Lançamento"
      Height          =   240
      Left            =   90
      TabIndex        =   32
      Top             =   5100
      Width           =   1035
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      Height          =   240
      Left            =   90
      TabIndex        =   27
      Top             =   30
      Width           =   1290
   End
End
Attribute VB_Name = "frmRelatorio_Triagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviço                                                        '
' Objetivo...............: Relatório Triagem                                              '
' Data de Criação........: 21/03/2006                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........: Rodrigo Santos                                                 '
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
    txtCodigo_Cliente.SetFocus
End Sub

Private Sub dtcEmpresa_LostFocus()
    If Not IsNumeric(dtcEmpresa.BoundText) Then dtcEmpresa.Text = Empty
    If IsNumeric(dtcEmpresa.Text) Then dtcEmpresa.Text = Empty

    dtcEmpresa.Enabled = False: txtCodigo_Cliente.SetFocus
End Sub

Private Sub dtcCliente_GotFocus()
    If txtCodigo_Cliente.Text = Empty Then
        Call Movimentacoes.Verifica_DataCombo(dtcCliente)
    End If
End Sub

Private Sub dtcCliente_LostFocus()
    txtCodigo_Cliente.Text = dtcCliente.BoundText
    If IsNumeric(txtCodigo_Cliente.Text) = False Or dtcCliente.Text = Empty Then txtCodigo_Cliente.Text = Empty: Exit Sub
End Sub

Private Sub cmdImprimir_Click()
    frmAguarde.Show
    DoEvents
    If dtpPeriodo_Lancamento_Inicio.Value > dtpPeriodo_Lancamento_Fim Then
       MsgBox "Data Início maior que a Data Final. Verifique.", vbInformation, "OnlyTech"
       dtpPeriodo_Lancamento_Inicio.SetFocus
       Exit Sub
    End If
    Call Impressao
    Unload frmAguarde
End Sub

Private Sub dtcFabricante_GotFocus()
    If txtFabricante.Text = Empty Then
        Call Movimentacoes.Verifica_DataCombo(dtcFabricante)
    End If
End Sub

Private Sub dtcFabricante_LostFocus()
    txtFabricante.Text = dtcFabricante.BoundText
    If IsNumeric(txtFabricante.Text) = False Or dtcFabricante.Text = Empty Then txtFabricante.Text = Empty: Exit Sub
End Sub

Private Sub dtcInsumo_GotFocus()
    If txtInsumo.Text = Empty Then
        Call Movimentacoes.Verifica_DataCombo(dtcInsumo)
    End If
End Sub

Private Sub dtcInsumo_LostFocus()
    txtInsumo.Text = dtcInsumo.BoundText
    If IsNumeric(txtInsumo.Text) = False Or dtcInsumo.Text = Empty Then txtInsumo.Text = Empty: Exit Sub
End Sub

Private Sub dtcRamo_Atividade_GotFocus()
    If txtRamo_Atividade.Text = Empty Then
        Call Movimentacoes.Verifica_DataCombo(dtcRamo_Atividade)
    End If
End Sub

Private Sub dtcRamo_Atividade_LostFocus()
    txtRamo_Atividade.Text = dtcRamo_Atividade.BoundText
    
    If dtcRamo_Atividade.BoundText <> Empty Then
       strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente " & _
                "WHERE FKCodigo_TBRamo_atividade = '" & dtcRamo_Atividade.BoundText & "'"
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
    Else
       strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente "
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
    End If
    
    If IsNumeric(txtRamo_Atividade.Text) = False Or dtcRamo_Atividade.Text = Empty Then txtRamo_Atividade.Text = Empty: Exit Sub
    
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
    On Error GoTo Erro
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Relatório Triagem"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Relatório Triagem"
    'Gravando o log
    log.Gravar_log "Otica", Me
   
    optAtivos_Todos.Value = True
    optBloqueados_Todos.Value = True
    
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
            
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    dtpPeriodo_Lancamento_Inicio.Value = Date - 15
    dtpPeriodo_Lancamento_Fim.Value = Date
    Call Monta_DataCombo
    Call Monta_Combo
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando o Relatorio Triagem"
    
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
    log.Descricao = "Cancelamento da Relação de Triagem"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    optAtivos_Todos.Value = True
    optBloqueados_Todos.Value = True
        
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
            
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    dtpPeriodo_Lancamento_Inicio.Value = Date - 15
    dtpPeriodo_Lancamento_Fim.Value = Date
    Call Monta_DataCombo
    Call Monta_Combo
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Sub txtCodigo_Cliente_Change()
    dtcCliente.BoundText = txtCodigo_Cliente.Text
    If IsNumeric(txtCodigo_Cliente.Text) = False Then txtCodigo_Cliente.Text = Empty: Exit Sub
End Sub

Private Sub txtCodigo_Cliente_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Cliente_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtFabricante_Change()
    dtcFabricante.BoundText = txtFabricante.Text
    If IsNumeric(txtFabricante.Text) = False Then txtFabricante.Text = Empty: Exit Sub
End Sub

Private Sub txtFabricante_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFabricante_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtInsumo_Change()
    dtcInsumo.BoundText = txtInsumo.Text
    If IsNumeric(txtInsumo.Text) = False Then txtInsumo.Text = Empty: Exit Sub
End Sub

Private Sub txtInsumo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtInsumo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtLote_LostFocus()
    If txtLote.Text <> Empty Then txtLote.Text = UCase(txtLote.Text)
End Sub

Private Sub txtRamo_Atividade_Change()
    dtcRamo_Atividade.BoundText = txtRamo_Atividade.Text
    If IsNumeric(txtRamo_Atividade.Text) = False Then txtRamo_Atividade.Text = Empty: Exit Sub
End Sub

Function Impressao()
    
    strSql = "SELECT PKId_TBTriagem," & _
             "FKCodigo_TBFabricante," & _
             "FKCodigo_TBInsumo," & _
             "FKId_TBCliente," & _
             "DFData_lancamento_TBTriagem," & _
             "DFData_fabricacao_TBTriagem," & _
             "DFLote_TBTriagem," & _
             "DFMes_ano_competencia_TBTriagem," & _
             "DFIntegrado_TBTriagem," & _
             "DFData_validade_TBTriagem," & _
             "DFAno_competencia_TBTriagem," & _
             "TBCliente.IXCodigo_TBCliente," & _
             "TBCliente.DFNome_TBCliente," & _
             "TBCliente.IXCodigo_TBEmpresa," & _
             "TBCliente.FKCodigo_TBRamo_atividade," & _
             "TBCliente.DFBloqueado_TBCliente,TBCliente.DFInativo_TBCliente," & _
             "TBCliente.DFintegrado_TBCliente,TBInsumo.DFDescricao_TBInsumo," & _
             "TBRamo_atividade.DFDescricao_TBRamo_atividade " & _
             "FROM TBTriagem " & _
             "INNER JOIN TBFabricante ON TBTriagem.FKCodigo_TBFabricante = TBFabricante.PKCodigo_TBFabricante " & _
             "INNER JOIN TBInsumo ON TBTriagem.FKCodigo_TBInsumo = TBInsumo.PKCodigo_TBInsumo " & _
             "INNER JOIN TBCliente ON TBTriagem.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
             "INNER JOIN TBRamo_atividade ON TBCliente.FKCodigo_TBRamo_atividade = TBRamo_atividade.PKCodigo_TBRamo_atividade " & _
             "WHERE TBCliente.IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "' " & _
             "AND DFData_lancamento_TBTriagem BETWEEN '" & Format(dtpPeriodo_Lancamento_Inicio.Value, "yyyymmdd") & "' AND '" & Format(dtpPeriodo_Lancamento_Fim.Value, "yyyymmdd") & "' "
   
    If dtcCliente.BoundText <> Empty Then
       strSql = strSql & "AND TBCliente.IXCodigo_TBCliente = '" & dtcCliente.BoundText & "' "
    End If
    
    If dtcRamo_Atividade.BoundText <> Empty Then
       strSql = strSql & "AND TBCliente.FKCodigo_TBRamo_atividade = '" & dtcRamo_Atividade.BoundText & "' "
    End If
    
    If dtcInsumo.BoundText <> Empty Then
       strSql = strSql & "AND TBTriagem.FKCodigo_TBInsumo = '" & dtcInsumo.BoundText & "' "
    End If
    
    If dtcFabricante.BoundText <> Empty Then
       strSql = strSql & "AND TBTriagem.FKCodigo_TBFabricante = '" & dtcFabricante.BoundText & "' "
    End If
    
    If txtLote.Text <> Empty Then
       strSql = strSql & "AND DFLote_TBTriagem = '" & txtLote.Text & "' "
    End If
    
    If cbbCompetencia_Mes.Text <> Empty Then
       strSql = strSql & "AND DFMes_ano_competencia_TBTriagem = '" & cbbCompetencia_Mes.Text & "' "
    End If
    
    If dtpCompetencia_Ano.Value <> Empty Then
       strSql = strSql & "AND DFAno_competencia_TBTriagem = '" & Format(dtpCompetencia_Ano.Value, "yyyy") & "' "
    End If
    
    If optAtivos_Sim.Value = True Then
       strSql = strSql & "AND TBCliente.DFInativo_TBCliente = '1' "
    ElseIf optAtivos_Nao.Value = True Then
       strSql = strSql & "AND TBCliente.DFInativo_TBCliente = '0' "
    End If

    If optBloqueados_Sim.Value = True Then
       strSql = strSql & "AND TBCliente.DFBloqueado_TBCliente = '1' "
    ElseIf optBloqueados_Sim.Value = True Then
       strSql = strSql & "AND TBCliente.DFBloqueado_TBCliente = '0' "
    End If
          
    Call frmConsole_Relatorio_Triagem.Show
End Function

Private Sub txtRamo_Atividade_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Monta_DataCombo()
    
     strSql = "SELECT PKCodigo_TBRamo_atividade,DFDescricao_TBRamo_atividade FROM TBRamo_atividade"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBRamo_atividade", "DFDescricao_TBRamo_atividade", dtcRamo_Atividade, strSql, "BDRetaguarda", "Otica", Me
            
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBInsumo,DFDescricao_TBInsumo FROM TBInsumo"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBInsumo", "DFDescricao_TBInsumo", dtcInsumo, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBFabricante,DFNome_TBFabricante FROM TBFabricante"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBFabricante", "DFNome_TBFabricante", dtcFabricante, strSql, "BDRetaguarda", "Otica", Me
    
End Function

Private Function Monta_Combo()

    cbbCompetencia_Mes.Clear
    cbbCompetencia_Mes.AddItem ("Janeiro")
    cbbCompetencia_Mes.AddItem ("Fevereiro")
    cbbCompetencia_Mes.AddItem ("Março")
    cbbCompetencia_Mes.AddItem ("Abril")
    cbbCompetencia_Mes.AddItem ("Maio")
    cbbCompetencia_Mes.AddItem ("Junho")
    cbbCompetencia_Mes.AddItem ("Julho")
    cbbCompetencia_Mes.AddItem ("Agosto")
    cbbCompetencia_Mes.AddItem ("Setembro")
    cbbCompetencia_Mes.AddItem ("Outubro")
    cbbCompetencia_Mes.AddItem ("Novembro")
    cbbCompetencia_Mes.AddItem ("Dezembro")

End Function

Private Sub txtRamo_Atividade_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub
