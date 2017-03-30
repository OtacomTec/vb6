VERSION 5.00
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatorio_Solicitacao_Visita 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Solicitação de Visitas"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frmRelatorio_Solicitacao_Visita.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5370
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
      Height          =   1695
      Left            =   90
      TabIndex        =   9
      Top             =   30
      Width           =   5175
      Begin AutoCompletar.CbCompleta cbbStatus 
         Height          =   360
         Left            =   105
         TabIndex        =   12
         Top             =   1230
         Width           =   3525
         _ExtentX        =   6218
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
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   570
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   22675457
         CurrentDate     =   37881
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   360
         Left            =   1980
         TabIndex        =   14
         Top             =   570
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   22675457
         CurrentDate     =   37881
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "até"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1605
         TabIndex        =   15
         Top             =   690
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Chamada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   10
         Top             =   990
         Width           =   540
      End
   End
   Begin VB.Frame freOrdenar 
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
      Height          =   675
      Left            =   90
      TabIndex        =   5
      Top             =   1770
      Width           =   5175
      Begin VB.OptionButton optStatus 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   855
      End
      Begin VB.OptionButton optData 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1620
         TabIndex        =   7
         Top             =   330
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optPrevisao_atendimento 
         Caption         =   "Previsão Atendimento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2910
         TabIndex        =   6
         Top             =   330
         Width           =   2175
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
      Left            =   90
      TabIndex        =   2
      Top             =   2490
      Width           =   2505
      Begin VB.OptionButton optAnalitico 
         Caption         =   "Analítico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optSintetico 
         Caption         =   "Sintético"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1290
         TabIndex        =   3
         Top             =   330
         Width           =   1095
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
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
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
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1245
   End
End
Attribute VB_Name = "frmRelatorio_Solicitacao_Visita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                         '
' Módulo.................: Compras                                                        '
' Objetivo...............: Relatório de Fornecedores                                      '
' Data de Criação........: 02/01/2004                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim log As New DLLSystemManager.log
Dim booAlterar As Boolean
Public strSQL As String
'Declaração da variavel do intercomunicador de mensagens
Private Cliente_mensagem_exe As VetorDeMensagens.ClienteDeMensagens

Private Sub cmdCancelar_Click()
    Call Objetos.Limpa_TXT(Me)
    txtCidade.SetFocus
End Sub

Private Sub cmdImprimir_Click()
    frmAguarde.Show
    DoEvents
    Call Impressao
    Unload frmAguarde
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 71: MsgBox a
                       Case 67: MsgBox b
                       Case 83: Unload Me  'CTRL+S
                End Select
    End Select
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
    log.Programa = "Relação de Solicitações de Visita"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Relatório de Solicitação de visita - Geral"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    cbbStatus.Clear
    cbbStatus.AddItem ("1 - A Visitar")
    cbbStatus.AddItem ("2 - Visitado - Aguardando Confirmação")
    cbbStatus.AddItem ("3 - Visitado - Fechado")
    cbbStatus.AddItem ("4 - Visitado - Não Finalizado")
    
    Me.optAnalitico.Value = True
    Me.optStatus.Value = True
    dtpData_chamado.Value = Date
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando a Relação de Solicitação de Visita"
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub tlbbotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Unload Me ' Gravar
           Case 2: Unload Me ' Cancelar
           Case 4: Unload Me
    
    End Select
End Sub

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    'Call Limpa_Combos
            
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento da Relação de Solicitação de Visita"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Function Impressao()
    Dim strStatus As String

    If cbbStatus.Text = "1 - A Visitar" Then
       strStatus = "1"
    ElseIf cbbStatus.Text = "2 - Visitado - Aguardando Confirmação" Then
       strStatus = "2"
    ElseIf cbbStatus.Text = "3 - Visitado - Fechado" Then
       strStatus = "3"
    ElseIf cbbStatus.Text = "4 - Visitado - Não Finalizado" Then
       strStatus = "4"
    End If

    strSQL = Empty
    strSQL = "SELECT DFData_TBSolicitacao_visita,DFHora_TBSolicitacao_visita,DFTelefone_TBAgenda_solicitacao_visita, " & _
             "DFContato_TBAgenda_solicitacao_visita,DFEndereco_TBSolicitacao_visita,DFNumero_TBSolicitacao_visita, " & _
             "DFComplemento_TBSolicitacao_visita,DFBairro_TBSolicitacao_visita,DFObservacao_TBSolicitacao_visita, " & _
             "TBAgenda_solicitacao_visita.DFContato_TBAgenda_solicitacao_visita, " & _
             "TBAgenda_solicitacao_visita.DFTelefone_TBAgenda_solicitacao_visita " & _
             "FROM TBSolicitacao_visita " & _
             "INNER JOIN TBCidade_otica " & _
             "ON TBSolicitacao_visita.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
             "INNER JOIN TBAgenda_solicitacao_visita " & _
             "ON TBAgenda_solicitacao_visita.FKId_TBSolicitacao_visita = TBSolicitacao_visita.PKId_TBSolicitacao_visita " & _
             "INNER JOIN TBAtendimento_solicitacao_visita " & _
             "ON TBAtendimento_solicitacao_visita.FKId_TBSolicitacao_visita = TBSolicitacao_visita.PKId_TBSolicitacao_visita "

    If cbbStatus.Text <> Empty Then
       strSQL = strSQL & "WHERE TBSolicitacao_visita.DFStatus_TBSolicitacao_visita = " & strStatus & ""
    End If
    
    If optData.Value = True Then
       strSQL = strSQL & " WHERE TBContrato_veiculo.DFData_inicio_TBContrato_veiculo >= '" & Format(dtpInicio.Value, "YYYYMMDD") & "' " & _
                         " AND TBContrato_veiculo.DFData_inicio_TBContrato_veiculo <= '" & Format(dtpFim.Value, "YYYYMMDD") & "' "
    ElseIf optStatus.Caption = True Then
       strSQL = strSQL & " ORDER BY TBSolicitacao_visita.DFStatus_TBSolicitacao_visita"
    ElseIf optPrevisao_atendimento.Caption = True Then
       strSQL = strSQL & " ORDER BY TBAtendimento_solicitacao_visita.DFData_previsao_TBAtendimento_solicitacao_visita"
    End If
       
    Call frmConsole_Relatorio_Solicitacao_Visitas.Show
End Function

Private Sub txtCidade_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtFim_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFim_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtInicio_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtInicio_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub
