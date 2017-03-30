VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatorio_Pendencias 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Pendências"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelatorio_Pendenicas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   7020
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
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
      TabIndex        =   45
      Top             =   6810
      Width           =   3390
      Begin VB.OptionButton Option1 
         Caption         =   "Gráfico"
         Height          =   240
         Left            =   2400
         TabIndex        =   24
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optAnalitico 
         Caption         =   "Analítico"
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   330
         Width           =   1095
      End
      Begin VB.OptionButton optSintetico 
         Caption         =   "Sintético"
         Height          =   240
         Left            =   1260
         TabIndex        =   23
         Top             =   330
         Width           =   1065
      End
   End
   Begin MSComCtl2.DTPicker dtpPeriodo_Lancamento_Inicio 
      Height          =   375
      Left            =   90
      TabIndex        =   27
      ToolTipText     =   "Período de Lançamento (Início)"
      Top             =   7740
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
      Height          =   5040
      Left            =   90
      TabIndex        =   34
      Top             =   660
      Width           =   6795
      Begin VB.TextBox txtMenu 
         Height          =   360
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Código Menu"
         Top             =   2550
         Width           =   1500
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Código Cliente"
         Top             =   1230
         Width           =   1500
      End
      Begin VB.TextBox txtFuncionario 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Código Funcionário"
         Top             =   570
         Width           =   1500
      End
      Begin VB.TextBox txtPrograma 
         Height          =   360
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Código Programa"
         Top             =   3210
         Width           =   1500
      End
      Begin VB.TextBox txtTipo_Servico 
         Height          =   360
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Código Tipo Serviço"
         Top             =   3870
         Width           =   1500
      End
      Begin VB.TextBox txtStatus 
         Height          =   360
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Código Status"
         Top             =   4530
         Width           =   1500
      End
      Begin VB.TextBox txtPrioridade 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Código Prioridade Serviço"
         Top             =   1890
         Width           =   1500
      End
      Begin MSDataListLib.DataCombo dtcMenu 
         Height          =   360
         Left            =   1680
         TabIndex        =   7
         ToolTipText     =   "Descrição Menu"
         Top             =   2550
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   360
         Left            =   1680
         TabIndex        =   3
         ToolTipText     =   "Nome Cliente"
         Top             =   1230
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcFuncionario 
         Height          =   360
         Left            =   1680
         TabIndex        =   1
         ToolTipText     =   "Nome Funcionário"
         Top             =   570
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcPrograma 
         Height          =   360
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   "Descrição Programa"
         Top             =   3210
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcTipo_servico 
         Height          =   360
         Left            =   1680
         TabIndex        =   11
         ToolTipText     =   "Descrição Tipo Serviço"
         Top             =   3870
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcStatus 
         Height          =   360
         Left            =   1680
         TabIndex        =   13
         ToolTipText     =   "Descrição Status"
         Top             =   4530
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcPrioridade 
         Height          =   360
         Left            =   1680
         TabIndex        =   5
         ToolTipText     =   "Descrição da Prioridade do Serviço"
         Top             =   1890
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
         Height          =   240
         Left            =   120
         TabIndex        =   44
         Top             =   2310
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   240
         Left            =   120
         TabIndex        =   43
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário"
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   330
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programa"
         Height          =   240
         Left            =   120
         TabIndex        =   41
         Top             =   2970
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Serviço"
         Height          =   240
         Left            =   120
         TabIndex        =   40
         Top             =   3630
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   240
         Left            =   120
         TabIndex        =   39
         Top             =   4290
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prioridade"
         Height          =   240
         Left            =   120
         TabIndex        =   38
         Top             =   1650
         Width           =   870
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
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7740
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
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7740
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenar"
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
      Left            =   3480
      TabIndex        =   33
      Top             =   6810
      Width           =   3420
      Begin VB.OptionButton optOrdenar_Codigo 
         Caption         =   "Código"
         Height          =   240
         Left            =   2340
         TabIndex        =   26
         Top             =   330
         Width           =   885
      End
      Begin VB.OptionButton optOrdenar_Alfabetico 
         Caption         =   "Alfabético"
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   330
         Width           =   1155
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Classificar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   90
      TabIndex        =   32
      Top             =   5730
      Width           =   6795
      Begin VB.OptionButton optData 
         Caption         =   "Data"
         Height          =   240
         Left            =   5490
         TabIndex        =   21
         Top             =   690
         Width           =   1185
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Status"
         Height          =   240
         Left            =   2280
         TabIndex        =   19
         Top             =   690
         Width           =   855
      End
      Begin VB.OptionButton optTipo_Servico 
         Caption         =   "Tipo Serviço"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   690
         Width           =   1365
      End
      Begin VB.OptionButton optMenu 
         Caption         =   "Menu"
         Height          =   240
         Left            =   3900
         TabIndex        =   16
         Top             =   330
         Width           =   765
      End
      Begin VB.OptionButton optPrograma 
         Caption         =   "Programa"
         Height          =   240
         Left            =   3900
         TabIndex        =   20
         Top             =   690
         Width           =   1185
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Cliente"
         Height          =   240
         Left            =   2280
         TabIndex        =   15
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optFuncionario 
         Caption         =   "Funcionário"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   330
         Width           =   1305
      End
      Begin VB.OptionButton optPrioridade 
         Caption         =   "Prioridade"
         Height          =   240
         Left            =   5490
         TabIndex        =   17
         Top             =   330
         Width           =   1185
      End
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   360
      Left            =   90
      TabIndex        =   31
      Top             =   270
      Width           =   6795
      _ExtentX        =   11986
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
      Left            =   2400
      TabIndex        =   28
      ToolTipText     =   "Período de Lançamento (Fim)"
      Top             =   7740
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
      Left            =   1830
      TabIndex        =   37
      Top             =   7890
      Width           =   285
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Lançamento"
      Height          =   240
      Left            =   90
      TabIndex        =   36
      Top             =   7500
      Width           =   1035
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      Height          =   240
      Left            =   90
      TabIndex        =   35
      Top             =   30
      Width           =   1290
   End
End
Attribute VB_Name = "frmRelatorio_Pendencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviço                                                        '
' Objetivo...............: Relatório Pendências                                              '
' Data de Criação........: 27/04/2006                                                     '
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
    
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
            
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    dtpPeriodo_Lancamento_Inicio.Value = Date - 30
    dtpPeriodo_Lancamento_Fim.Value = Date
    optAnalitico.Value = True
    optOrdenar_Alfabetico.Value = True
    optFuncionario.Value = True
    Call Monta_Data_Combos
   
    
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
    dtpPeriodo_Lancamento_Inicio.Value = Date - 30
    dtpPeriodo_Lancamento_Fim.Value = Date
    optAnalitico.Value = True
    optOrdenar_Alfabetico.Value = True
    optFuncionario.Value = True
    Call Monta_Data_Combos
        
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Function Impressao()
    
    strSql = "SELECT PKID_TBPendencia_servico,DFNumero_Relatorio_TBPendencia_servico,DFData_Cadastro_TBPendencia_servico," & _
             "TBPendencia_Servicos.FKCodigo_TBFuncionario,DFNome_TBFuncionario," & _
             "TBPendencia_Servicos.FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico,DFDescricao__TBPrioridade_pendencia_servico," & _
             "IXCodigo_TBCliente,DFNome_TBCliente," & _
             "TBPendencia_Servicos.FKID_TBMenu,DFDescricao_TBMenu,FKID_TBProgramas,DFDescricao_TBProgramas," & _
             "TBPendencia_Servicos.FKCodigo_TBTipo_servico_Pendencia_Servico,DFDescricao_TBTipo_Pendencia_servico," & _
             "TBPendencia_Servicos.FKCodigo_TBStatus_Pendencia_Servico,DFDescricao_TBStatus_pendencia_servico," & _
             "DFData_Inicio_TBPendencia_servico,DFData_fim_TBPendencia_servico," & _
             "CONVERT(char,DFHora_Inicio_TBPendencia_servico,108)Hora_Inicio,CONVERT(char,DFHora_Fim_TBPendencia_servico,108)Hora_Fim," & _
             "CONVERT(char,DFHora_Cadastro_TBPendencia_servico,108)Hora_Cadastrado,DFObservacao_TBPendencia_servico,PKCodigo_TBEmpresa," & _
             "DFRazao_Social_TBEmpresa " & _
             "FROM TBPendencia_Servicos " & _
             "INNER JOIN TBFuncionario ON TBPendencia_Servicos.FKCodigo_TBFuncionario = TBFuncionario.PKCodigo_TBFuncionario " & _
             "INNER JOIN TBCliente ON TBPendencia_Servicos.FKID_TBCliente = TBCliente.PKId_TBCliente "
    
    strSql = strSql & "INNER JOIN TBEmpresa ON TBPendencia_Servicos.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
                      "INNER JOIN TBProgramas ON TBPendencia_Servicos.FKID_TBProgramas = TBProgramas.PKId_TBProgramas " & _
                      "INNER JOIN TBStatus_Pendencia_servico " & _
                      "ON TBPendencia_Servicos.FKCodigo_TBStatus_Pendencia_Servico = TBStatus_Pendencia_servico.PKCodigo_TBStatus_pendencia_servico " & _
                      "INNER JOIN TBPrioridade_Pendencia_Servico " & _
                      "ON TBPendencia_Servicos.FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico = TBPrioridade_Pendencia_Servico.PKCodigo__TBPrioridade_pendencia_servico " & _
                      "INNER JOIN TBTipo_servico_Pendencia_Servico " & _
                      "ON TBPendencia_Servicos.FKCodigo_TBTipo_servico_Pendencia_Servico = TBTipo_servico_Pendencia_Servico.PKCodigo_Prioridade_TBTipo_Pendencia_servico " & _
                      "INNER JOIN TBMenu ON TBPendencia_Servicos.FKID_TBMenu = TBMenu.PKId_TBMenu " & _
                      "WHERE TBPendencia_Servicos.FKCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "' " & _
                      "AND DFData_Cadastro_TBPendencia_servico BETWEEN '" & Format(dtpPeriodo_Lancamento_Inicio.Value, "YYYYMMDD") & "' AND '" & Format(dtpPeriodo_Lancamento_Fim.Value, "YYYYMMDD") & "' "

           
    If dtcFuncionario.BoundText <> Empty Then
       strSql = strSql & " AND FKCodigo_TBFuncionario = '" & dtcFuncionario.BoundText & "' "
    End If
    If dtcPrioridade.BoundText <> Empty Then
       strSql = strSql & " AND FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico = '" & dtcPrioridade.BoundText & "' "
    End If
    If dtcCliente.BoundText <> Empty Then
       strSql = strSql & " AND IXCodigo_TBCliente = '" & dtcCliente.BoundText & "' "
    End If
    If dtcMenu.BoundText <> Empty Then
       strSql = strSql & " AND FKID_TBMenu = '" & dtcMenu.BoundText & "' "
    End If
    If dtcPrograma.BoundText <> Empty Then
       strSql = strSql & " AND FKID_TBProgramas = '" & dtcPrograma.BoundText & "' "
    End If
    If dtcTipo_servico.BoundText <> Empty Then
       strSql = strSql & " AND FKCodigo_TBTipo_servico_Pendencia_Servico = '" & dtcTipo_servico.BoundText & "' "
    End If
    If dtcStatus.BoundText <> Empty Then
       strSql = strSql & " AND FKCodigo_TBStatus_Pendencia_Servico = '" & dtcStatus.BoundText & "' "
    End If
                
    If optOrdenar_Alfabetico.Value = True Then
       strSql = strSql & " ORDER BY DFNumero_Relatorio_TBPendencia_servico"
    ElseIf cmdOrdenar.Caption = "A" Then
       strSql = strSql & " ORDER BY TBFuncionario.DFNome_TBFuncionario"
    End If
          
    Call frmConsole_Relatorio_Pendencias.Show
    
End Function

Private Function Monta_Data_Combos()
    
    strSql = "SELECT PKCodigo_TBFuncionario,DFNome_TBFuncionario FROM TBFuncionario " & _
             "WHERE FKCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBFuncionario", "DFNome_TBFuncionario", dtcFuncionario, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo__TBPrioridade_pendencia_servico,DFDescricao__TBPrioridade_pendencia_servico " & _
             "FROM TBPrioridade_Pendencia_servico"
    Movimentacoes.Movimenta_DataCombo "PKCodigo__TBPrioridade_pendencia_servico", "DFDescricao__TBPrioridade_pendencia_servico", dtcPrioridade, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente " & _
             "WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKId_TBMenu,DFDescricao_TBMenu FROM TBMenu"
    Movimentacoes.Movimenta_DataCombo "PKId_TBMenu", "DFDescricao_TBMenu", dtcMenu, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKId_TBProgramas,DFDescricao_TBProgramas FROM TBProgramas"
    Movimentacoes.Movimenta_DataCombo "PKId_TBProgramas", "DFDescricao_TBProgramas", dtcPrograma, strSql, "BDRetaguarda", "Otica", Me

    strSql = "SELECT PKCodigo_Prioridade_TBTipo_Pendencia_servico,DFDescricao_TBTipo_Pendencia_servico " & _
             "FROM TBTipo_servico_Pendencia_servico"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_Prioridade_TBTipo_Pendencia_servico", "DFDescricao_TBTipo_Pendencia_servico", dtcTipo_servico, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBStatus_pendencia_servico,DFDescricao_TBStatus_pendencia_servico " & _
             "FROM TBStatus_Pendencia_servico"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBStatus_pendencia_servico", "DFDescricao_TBStatus_pendencia_servico", dtcStatus, strSql, "BDRetaguarda", "Otica", Me

End Function

Private Sub txtFuncionario_Change()
    dtcFuncionario.BoundText = txtFuncionario.Text
    If IsNumeric(txtFuncionario.Text) = False Then txtFuncionario.Text = Empty: Exit Sub
End Sub

Private Sub txtFuncionario_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFuncionario_LostFocus()
    If dtcFuncionario.Text = Empty Then txtFuncionario.Text = Empty
End Sub

Private Sub txtMenu_Change()
    dtcMenu.BoundText = txtMenu.Text
    If IsNumeric(txtMenu.Text) = False Then txtMenu.Text = Empty: Exit Sub
End Sub

Private Sub txtMenu_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMenu_LostFocus()
    If dtcMenu.Text = Empty Then txtMenu.Text = Empty
End Sub

Private Sub txtPrioridade_Change()
    dtcPrioridade.BoundText = txtPrioridade.Text
    If IsNumeric(txtPrioridade.Text) = False Then txtPrioridade.Text = Empty: Exit Sub
End Sub

Private Sub txtPrioridade_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrioridade_LostFocus()
    If dtcPrioridade.Text = Empty Then txtPrioridade.Text = Empty
End Sub

Private Sub txtPrograma_Change()
    dtcPrograma.BoundText = txtPrograma.Text
    If IsNumeric(txtPrograma.Text) = False Then txtPrograma.Text = Empty: Exit Sub
End Sub

Private Sub txtPrograma_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrograma_LostFocus()
    If dtcPrograma.Text = Empty Then txtPrograma.Text = Empty
End Sub

Private Sub txtStatus_Change()
    dtcStatus.BoundText = txtStatus.Text
    If IsNumeric(txtStatus.Text) = False Then txtStatus.Text = Empty: Exit Sub
End Sub

Private Sub txtStatus_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtStatus_LostFocus()
    If dtcStatus.Text = Empty Then txtStatus.Text = Empty
End Sub

Private Sub txtTipo_Servico_Change()
    dtcTipo_servico.BoundText = txtTipo_Servico.Text
    If IsNumeric(txtTipo_Servico.Text) = False Then txtTipo_Servico.Text = Empty: Exit Sub
End Sub

Private Sub txtTipo_Servico_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTipo_Servico_LostFocus()
    If dtcTipo_servico.Text = Empty Then txtTipo_Servico.Text = Empty
End Sub

Private Sub txtCliente_Change()
    dtcCliente.BoundText = txtCliente.Text
    If IsNumeric(txtCliente.Text) = False Then txtCliente.Text = Empty: Exit Sub
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCliente_LostFocus()
    If dtcCliente.Text = Empty Then txtCliente.Text = Empty
End Sub

Private Sub dtcStatus_GotFocus()
    If Me.txtStatus.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcStatus.Text)
    End If
End Sub

Private Sub dtcStatus_LostFocus()
    txtStatus.Text = dtcStatus.BoundText
End Sub

Private Sub dtcTipo_servico_GotFocus()
    If Me.txtTipo_Servico.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcTipo_servico.Text)
    End If
End Sub

Private Sub dtcTipo_servico_LostFocus()
    txtTipo_Servico.Text = dtcTipo_servico.BoundText
End Sub

Private Sub dtcCliente_GotFocus()
    If Me.txtCliente.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCliente.Text)
    End If
End Sub

Private Sub dtcCliente_LostFocus()
    txtCliente.Text = dtcCliente.BoundText
End Sub

Private Sub dtcEmpresa_Change()
    Call Monta_Data_Combos
End Sub

Private Sub dtcEmpresa_LostFocus()
    dtcEmpresa.Enabled = False
End Sub

Private Sub dtcFuncionario_GotFocus()
    If Me.txtFuncionario.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFuncionario.Text)
    End If
End Sub

Private Sub dtcFuncionario_LostFocus()
    txtFuncionario.Text = dtcFuncionario.BoundText
End Sub

Private Sub dtcMenu_GotFocus()
    If Me.txtMenu.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcMenu.Text)
    End If
End Sub

Private Sub dtcMenu_LostFocus()
    txtMenu.Text = dtcMenu.BoundText
    
    If dtcMenu.BoundText <> Empty Then
       strSql = "SELECT PKId_TBProgramas,DFDescricao_TBProgramas FROM TBProgramas " & _
                "WHERE FKID_Menu = '" & dtcMenu.BoundText & "'"
       Movimentacoes.Movimenta_DataCombo "PKId_TBProgramas", "DFDescricao_TBProgramas", dtcPrograma, strSql, "BDRetaguarda", "Otica", Me
    Else
       strSql = "SELECT PKId_TBProgramas,DFDescricao_TBProgramas FROM TBProgramas "
       Movimentacoes.Movimenta_DataCombo "PKId_TBProgramas", "DFDescricao_TBProgramas", dtcPrograma, strSql, "BDRetaguarda", "Otica", Me
    End If
    
    dtcPrograma.BoundText = txtPrograma.Text
    
    If Not IsNumeric(dtcPrograma.BoundText) Then txtPrograma.Text = Empty
    If txtMenu.Text = Empty Then txtPrograma.Text = Empty
End Sub

Private Sub dtcPrioridade_GotFocus()
    If Me.txtPrioridade.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcPrioridade.Text)
    End If
End Sub

Private Sub dtcPrioridade_LostFocus()
    txtPrioridade.Text = dtcPrioridade.BoundText
End Sub

Private Sub dtcPrograma_GotFocus()
    If Me.txtPrograma.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcPrograma.Text)
    End If
End Sub

Private Sub dtcPrograma_LostFocus()
    txtPrograma.Text = dtcPrograma.BoundText
    
    If dtcMenu.BoundText = Empty And txtPrograma.Text <> Empty Then
       Dim rstBusca_Modulo As New ADODB.Recordset
    
       strSql = "SELECT PKId_TBMenu " & _
                "FROM TBMenu, TBProgramas " & _
                "WHERE TBProgramas.FKId_Menu = TBMenu.PKId_TBMenu " & _
                "AND PKId_TBProgramas = '" & txtPrograma.Text & "' "
       
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstBusca_Modulo, "Otica", Me
       
       If rstBusca_Modulo.EOF = False Then
          txtMenu.Text = rstBusca_Modulo!PKId_TBMenu
       End If
       
       Set rstBusca_Modulo = Nothing
    End If
End Sub

