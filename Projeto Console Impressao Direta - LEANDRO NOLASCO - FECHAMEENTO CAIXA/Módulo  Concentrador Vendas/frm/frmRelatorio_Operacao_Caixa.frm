VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatorio_Operacao_Caixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fechamento Operação de Caixa"
   ClientHeight    =   4920
   ClientLeft      =   1800
   ClientTop       =   1845
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
   Icon            =   "frmRelatorio_Operacao_Caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6255
   Begin VB.Frame Frame2 
      Caption         =   "Canceladas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2550
      TabIndex        =   30
      Top             =   3480
      Width           =   3585
      Begin VB.OptionButton optCancelada_Todos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   2580
         TabIndex        =   21
         Top             =   330
         Width           =   975
      End
      Begin VB.OptionButton optCancelada_Nao 
         Caption         =   "Não"
         Height          =   240
         Left            =   1410
         TabIndex        =   8
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optCancelada_Sim 
         Caption         =   "Sim"
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   675
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Tipo de Relatório"
      Top             =   3480
      Width           =   2415
      Begin VB.OptionButton optAnalitico 
         Caption         =   "Analítico"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   330
         Width           =   1095
      End
      Begin VB.OptionButton optSintetico 
         Caption         =   "Sintético"
         Height          =   240
         Left            =   1230
         TabIndex        =   19
         Top             =   330
         Width           =   1095
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
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Visualiza Impressão"
      Top             =   4410
      Width           =   1185
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
      TabIndex        =   12
      ToolTipText     =   "Limpa os Filtros"
      Top             =   4410
      Width           =   1185
   End
   Begin VB.Frame freOrdenar 
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
      Height          =   1005
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Ordenar Impressão"
      Top             =   2430
      Width           =   6015
      Begin VB.OptionButton optFinalizadora_Periodo 
         Caption         =   "Finaliz. Período"
         Enabled         =   0   'False
         Height          =   240
         Left            =   4230
         TabIndex        =   29
         Top             =   660
         Width           =   1665
      End
      Begin VB.OptionButton optOperador 
         Caption         =   "Operador"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   1125
      End
      Begin VB.OptionButton optFinalizadora 
         Caption         =   "Finalizadora"
         Height          =   240
         Left            =   2190
         TabIndex        =   17
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton optPDV 
         Caption         =   "PDV"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   660
         Width           =   675
      End
      Begin VB.OptionButton optFaixa_Horaria 
         Caption         =   "Faixa Hora"
         Height          =   240
         Left            =   4230
         TabIndex        =   15
         Top             =   330
         Width           =   1665
      End
      Begin VB.OptionButton optData 
         Caption         =   "Data"
         Height          =   240
         Left            =   2190
         TabIndex        =   14
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1725
      Left            =   120
      TabIndex        =   22
      Top             =   660
      Width           =   6015
      Begin VB.TextBox txtOperador 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Código do Operador"
         Top             =   570
         Width           =   1155
      End
      Begin VB.TextBox txtPdv 
         Height          =   360
         Left            =   4680
         TabIndex        =   3
         ToolTipText     =   "Código do Ponto de Venda"
         Top             =   570
         Width           =   1185
      End
      Begin VB.TextBox txtFinalizadora 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Código da Finalizadora"
         Top             =   1215
         Width           =   1155
      End
      Begin MSDataListLib.DataCombo dtcFinalizadora 
         Height          =   360
         Left            =   1320
         TabIndex        =   5
         Top             =   1215
         Width           =   4575
         _ExtentX        =   8070
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
      Begin MSDataListLib.DataCombo dtcOperador 
         Height          =   360
         Left            =   1320
         TabIndex        =   2
         Top             =   570
         Width           =   3315
         _ExtentX        =   5847
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
         Caption         =   "Operador"
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PDV"
         Height          =   240
         Left            =   4710
         TabIndex        =   27
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Finalizadora"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   975
         Width           =   1035
      End
   End
   Begin MSComCtl2.DTPicker dtpInicial 
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   1515
      _ExtentX        =   2672
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
      Format          =   20643841
      CurrentDate     =   37881
   End
   Begin MSComCtl2.DTPicker dtpFinal 
      Height          =   360
      Left            =   2070
      TabIndex        =   10
      Top             =   4440
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   20643841
      CurrentDate     =   37881
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   270
      Width           =   6015
      _ExtentX        =   10610
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
   Begin VB.Label lblAuditoria 
      AutoSize        =   -1  'True
      Caption         =   "Auditoria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5250
      TabIndex        =   31
      Top             =   30
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "até"
      Height          =   240
      Left            =   1695
      TabIndex        =   25
      Top             =   4560
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   240
      Left            =   120
      TabIndex        =   24
      Top             =   4200
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      Height          =   240
      Left            =   120
      TabIndex        =   23
      Top             =   30
      Width           =   1290
   End
End
Attribute VB_Name = "frmRelatorio_Operacao_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Estatística de Resumo Diário de Vendas                         '
' Data de Criação........: 22/06/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public strSQL As String
Dim log As New DLLSystemManager.log
'Declaração da variavel do intercomunicador de mensagens
Private Cliente_mensagem_exe As VetorDeMensagens.ClienteDeMensagens

Private Sub dtcEmpresa_LostFocus()
    dtcEmpresa.Enabled = False
End Sub

Private Sub dtcFinalizadora_GotFocus()
    If txtFinalizadora.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora.Text)
    End If
End Sub

Private Sub dtcFinalizadora_LostFocus()
    txtFinalizadora.Text = dtcFinalizadora.BoundText
    If IsNumeric(txtFinalizadora.Text) = False Or dtcFinalizadora.Text = Empty Then txtFinalizadora.Text = Empty: Exit Sub
End Sub

Private Sub dtcOperador_GotFocus()
    If txtOperador.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcOperador.Text)
    End If
End Sub

Private Sub dtcOperador_LostFocus()
    If txtOperador.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcOperador.Text)
    End If
    txtOperador.Text = dtcOperador.BoundText
    If IsNumeric(txtOperador.Text) = False Or dtcOperador.Text = Empty Then txtOperador.Text = Empty: Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    Call Cancelar
End Sub

Private Sub cmdImprimir_Click()

    If dtpInicial.Value > dtpFinal.Value Then
       MsgBox "Data Final menor que Data Inicial. Verifique!", vbInformation, "Only Tech"
       Exit Sub
    End If

    frmAguarde.Show
    DoEvents
    Call Impressao
    Unload frmAguarde
    
End Sub

Private Sub dtpFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

Private Sub dtpInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = "113" And booAlterar = False Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
    
    'Verifica se foi preciona do F12 e muda a interface  para orçamento
    If KeyCode = 123 Then
       lblAuditoria.Visible = False
    End If
    
    If KeyCode = 122 Then
       lblAuditoria.Visible = True
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
    log.Programa = "Relatorio de Fechamento Diário de Vendas"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
        Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Relatorio de Fechamento Diário de Vendas"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    'Montando os datacombo de tela
    strSQL = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSQL, "BDRetaguarda", "Otica", Me

    strSQL = "SELECT IXCodigo_TBFinalizadora, DFDescricao_TBFinalizadora FROM TBFinalizadora "
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT PKCodigo_TBOperadores_ecf, DFNome_TBOperadores_ecf FROM TBOperadores_ecf "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSQL, "BDRetaguarda", "Otica", Me
     
    'dtcCodigo_empresa.boundtext = ---- Inserir aqui informações da DLLIntercomunicador de EXE's
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
 
    dtpInicial.Value = Date
    dtpFinal.Value = Date
    
    optOperador.Value = True
    optAnalitico.Value = True
    optCancelada_Todos.Value = True
    Exit Sub
    
Erro:

    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Saindo de Relatorio de Fechamento Diário de Vendas"
    
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
            
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Relatorio de Operação de Caixa"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    dtpInicial.Value = Date
    dtpInicial.Day = 1
    dtpFinal = Date
    txtOperador.SetFocus
    optOperador.Value = True
    optAnalitico.Value = True
    optCancelada_Todos.Value = True
    
    Exit Function
    
Erro:

    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
    
End Function

Private Function Impressao()
    
    If optFinalizadora.Value = True And optSintetico.Value = True Then
       strSQL = "SELECT DISTINCT TBFinalizadora.IXCodigo_TBFinalizadora AS COD_FINALIZADORA, " & _
                "TBFinalizadora.DFDescricao_TBFinalizadora AS DESCR_FINALIZADORA, " & _
                "DFData_TBOperacao_caixa AS DATA, " & _
                "TBFinalizadora.DFDebito_credito_TBFinalizadora as CREDITO_DEBITO, " & _
                "SUM(DFValor_TBOperacao_caixa) AS VALOR_OPERACAO, " & _
                "DFTipo_operacao_TBOperacao_caixa AS TIPO, " & _
                "DFStatus_aberto_fechado_TBOperacao_caixa As Status "
                
       If Me.lblAuditoria.Visible = True Then
          strSQL = strSQL & ",DFCodigo_cupom_impressora_TBCupom,DFGrant_total_impressora_TBOperacao_caixa "
       End If
       
       strSQL = strSQL & "FROM TBOPERACAO_CAIXA " & _
                "INNER JOIN TBFinalizadora " & _
                "ON TBOPERACAO_CAIXA.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                "INNER JOIN TBOperadores_ecf " & _
                "ON TBOPERACAO_CAIXA.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
                "INNER JOIN TBCUPOM ON TBOPERACAO_CAIXA.DFNumero_Cupom_TBOperacao_caixa = TBCupom.DFNumero_TBCupom " & _
                "WHERE TBOPERACAO_CAIXA.FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
                "AND (TBOPERACAO_CAIXA.DFData_TBOperacao_caixa BETWEEN '" & Format(Me.dtpInicial.Value, "YYYYMMDD") & "' " & _
                "AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "' ) "
    ElseIf optFinalizadora_Periodo.Value = True And optSintetico.Value = True Then
       strSQL = "SELECT DISTINCT TBFinalizadora.IXCodigo_TBFinalizadora AS COD_FINALIZADORA," & _
                "TBFinalizadora.DFDescricao_TBFinalizadora AS DESCR_FINALIZADORA," & _
                "TBFinalizadora.DFDebito_credito_TBFinalizadora as CREDITO_DEBITO," & _
                "SUM(DFValor_TBOperacao_caixa) AS VALOR_OPERACAO," & _
                "DFTipo_operacao_TBOperacao_caixa AS TIPO," & _
                "DFStatus_aberto_fechado_TBOperacao_caixa As Status "
                
                If Me.lblAuditoria.Visible = True Then
                   strSQL = strSQL & ",DFCodigo_cupom_impressora_TBCupom,DFGrant_total_impressora_TBOperacao_caixa "
                End If
                
       strSQL = strSQL & "FROM TBOPERACAO_CAIXA " & _
                "INNER JOIN TBFinalizadora " & _
                "ON TBOPERACAO_CAIXA.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                "INNER JOIN TBOperadores_ecf " & _
                "ON TBOPERACAO_CAIXA.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
                "INNER JOIN TBCUPOM ON TBOPERACAO_CAIXA.DFNumero_Cupom_TBOperacao_caixa = TBCupom.DFNumero_TBCupom " & _
                "WHERE TBOPERACAO_CAIXA.FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
                "AND (TBOPERACAO_CAIXA.DFData_TBOperacao_caixa BETWEEN '" & Format(Me.dtpInicial.Value, "YYYYMMDD") & "' " & _
                "AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "' ) "
    Else
       strSQL = "SELECT DISTINCT FKCodigo_TBPdv AS PDV,TBFinalizadora.IXCodigo_TBFinalizadora AS COD_FINALIZADORA, " & _
                "TBFinalizadora.DFDescricao_TBFinalizadora AS DESCR_FINALIZADORA, " & _
                "TBOperadores_ecf.PKCodigo_TBOperadores_ecf AS COD_OPERADOR, " & _
                "TBOperadores_ecf.DFNome_TBOperadores_ecf AS OPERADOR,DFData_TBOperacao_caixa AS DATA,DFHora_TBOperacao_caixa AS HORA,TBFinalizadora.DFDebito_credito_TBFinalizadora as CREDITO_DEBITO, " & _
                "DFValor_TBOperacao_caixa AS VALOR_OPERACAO,DFTipo_operacao_TBOperacao_caixa AS TIPO,DFStatus_aberto_fechado_TBOperacao_caixa AS STATUS "
                
                If Me.lblAuditoria.Visible = True Then
                   strSQL = strSQL & ",DFCodigo_cupom_impressora_TBCupom,DFGrant_total_impressora_TBOperacao_caixa "
                End If
                
       strSQL = strSQL & ",DFObservacao_TBOperacao_caixa As Observacao FROM TBoperacao_caixa " & _
                "INNER JOIN TBFinalizadora " & _
                "ON TBOPERACAO_CAIXA.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                "INNER JOIN TBOperadores_ecf " & _
                "ON TBOPERACAO_CAIXA.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
                "LEFT JOIN TBCUPOM ON TBOPERACAO_CAIXA.DFNumero_Cupom_TBOperacao_caixa = TBCupom.DFNumero_TBCupom " & _
                "WHERE TBOPERACAO_CAIXA.FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
                "AND (TBOPERACAO_CAIXA.DFData_TBOperacao_caixa BETWEEN '" & Format(Me.dtpInicial.Value, "YYYYMMDD") & "' " & _
                "AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "' ) "
    End If
    
    ' Finalizadora
    If dtcFinalizadora.Text <> "" Then
       strSQL = strSQL & " AND TBFinalizadora.IXCodigo_TBFinalizadora = " & dtcFinalizadora.BoundText & " "
    End If
    
    ' PDV
    If txtPdv.Text <> "" Then
       strSQL = strSQL & " AND FKCodigo_TBPdv = " & txtPdv.Text & " "
    End If
    
    ' Operador
    If dtcOperador.BoundText <> "" Then
       strSQL = strSQL & " AND TBOperadores_ecf.PKCodigo_TBOperadores_ecf = " & dtcOperador.BoundText & " "
    End If
    
    If optCancelada_Sim.Value = True Then
       strSQL = strSQL & " AND DFCancelado_TBCupom = 1 "
    End If
    
    If optCancelada_Nao.Value = True Then
       strSQL = strSQL & " AND DFCancelado_TBCupom = 0 "
    End If
    
    If optFinalizadora.Value = True And optSintetico.Value = True Then
       strSQL = strSQL & "GROUP BY TBFinalizadora.IXCodigo_TBFinalizadora," & _
                         "TBFinalizadora.DFDescricao_TBFinalizadora," & _
                         "DFData_TBOperacao_caixa," & _
                         "TBFinalizadora.DFDebito_credito_TBFinalizadora," & _
                         "DFTipo_operacao_TBOperacao_caixa," & _
                         "DFStatus_aberto_fechado_TBOperacao_caixa "
       If Me.lblAuditoria.Visible = True Then
          strSQL = strSQL & ",DFCodigo_cupom_impressora_TBCupom,DFGrant_total_impressora_TBOperacao_caixa "
       End If
       
    ElseIf optFinalizadora_Periodo.Value = True And optSintetico.Value = True Then
    
       strSQL = strSQL & "GROUP BY TBFinalizadora.IXCodigo_TBFinalizadora," & _
                         "TBFinalizadora.DFDescricao_TBFinalizadora," & _
                         "TBFinalizadora.DFDebito_credito_TBFinalizadora," & _
                         "DFTipo_operacao_TBOperacao_caixa," & _
                         "DFStatus_aberto_fechado_TBOperacao_caixa "

        If Me.lblAuditoria.Visible = True Then
           strSQL = strSQL & ",DFCodigo_cupom_impressora_TBCupom,DFGrant_total_impressora_TBOperacao_caixa "
        End If
                             
    ElseIf optOperador.Value = True And optAnalitico.Value = True Then
       strSQL = strSQL & "GROUP BY FKCodigo_TBPdv ,TBFinalizadora.IXCodigo_TBFinalizadora ," & _
                         "TBFinalizadora.DFDescricao_TBFinalizadora ," & _
                         "TBOperadores_ecf.PKCodigo_TBOperadores_ecf," & _
                         "TBOperadores_ecf.DFNome_TBOperadores_ecf ,DFData_TBOperacao_caixa ," & _
                         "DFHora_TBOperacao_caixa ,TBFinalizadora.DFDebito_credito_TBFinalizadora ," & _
                         "DFValor_TBOperacao_caixa ,DFTipo_operacao_TBOperacao_caixa ," & _
                         "DFStatus_aberto_fechado_TBOperacao_caixa ,DFObservacao_TBOperacao_caixa ," & _
                         "TBOPERACAO_CAIXA.PKId_TBOperacao_caixa "
       If Me.lblAuditoria.Visible = True Then
           strSQL = strSQL & ",DFCodigo_cupom_impressora_TBCupom,DFGrant_total_impressora_TBOperacao_caixa "
       End If
                         
    End If
    
    ' Ordenacao do relatorio
    If optFaixa_Horaria.Value = True Then
        strSQL = strSQL & " ORDER BY DFHora_TBOperacao_caixa, dfdata_tboperacao_caixa, dfhora_tboperacao_caixa "
    ElseIf optData.Value = True Then
        strSQL = strSQL & " ORDER BY DFData_TBOperacao_caixa, dfdata_tboperacao_caixa, dfhora_tboperacao_caixa "
    ElseIf optPDV.Value = True Then
        strSQL = strSQL & " ORDER BY FKCodigo_TBPdv, dfdata_tboperacao_caixa, dfhora_tboperacao_caixa "
    ElseIf optFinalizadora.Value = True Or optFinalizadora_Periodo.Value = True Then
        strSQL = strSQL & " ORDER BY TBFinalizadora.IXCodigo_TBFinalizadora, dfdata_tboperacao_caixa, dfhora_tboperacao_caixa "
    ElseIf optOperador.Value = True Then
        strSQL = strSQL & " ORDER BY TBOperadores_ecf.PKCodigo_TBOperadores_ecf, dfdata_tboperacao_caixa, dfhora_tboperacao_caixa "
    End If
    
    Call frmConsole_Relatorio_Fechamento_Operacao_Caixa.Show
    
End Function

Private Sub optAnalitico_Click()
    If optFinalizadora_Periodo.Value = True Then optSintetico.Value = True
End Sub

Private Sub optFinalizadora_Periodo_Click()
    optSintetico.Value = True
End Sub

Private Sub optSintetico_Click()
    If optFinalizadora.Value = False And optFinalizadora_Periodo.Value = False Then
       optAnalitico.Value = True
    End If
End Sub

Private Sub txtFinalizadora_Change()
    dtcFinalizadora.BoundText = txtFinalizadora.Text
    If IsNumeric(txtFinalizadora.Text) = False Then txtFinalizadora.Text = Empty: Exit Sub
End Sub

Private Sub txtFinalizadora_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_Change()
    dtcOperador.BoundText = txtOperador.Text
    If IsNumeric(txtOperador.Text) = False Then txtOperador.Text = Empty: Exit Sub
End Sub

Private Sub txtOperador_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_LostFocus()
    dtcOperador.BoundText = txtOperador.Text
End Sub

Private Sub txtPdv_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub
