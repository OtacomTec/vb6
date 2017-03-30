VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorio_Senha 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Senha"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelatorio_Senha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6345
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
      Left            =   90
      TabIndex        =   21
      Top             =   3150
      Width           =   3270
      Begin VB.OptionButton optBloqueados_Todos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   2280
         TabIndex        =   12
         Top             =   330
         Width           =   855
      End
      Begin VB.OptionButton optBloqueados_Sim 
         Caption         =   "Sim"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   330
         Width           =   885
      End
      Begin VB.OptionButton optBloqueados_Nao 
         Caption         =   "Não"
         Height          =   240
         Left            =   1260
         TabIndex        =   11
         Top             =   330
         Width           =   645
      End
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
      TabIndex        =   20
      Top             =   2460
      Width           =   3270
      Begin VB.OptionButton optAtivos_Todos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   2220
         TabIndex        =   7
         Top             =   330
         Width           =   855
      End
      Begin VB.OptionButton optAtivos_Nao 
         Caption         =   "Não"
         Height          =   240
         Left            =   1170
         TabIndex        =   6
         Top             =   330
         Width           =   675
      End
      Begin VB.OptionButton optAtivos_Sim 
         Caption         =   "Sim"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   330
         Width           =   885
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
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3390
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3390
      Width           =   1245
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
      Height          =   1770
      Left            =   90
      TabIndex        =   16
      Top             =   660
      Width           =   6165
      Begin VB.TextBox txtCodigo_Cliente 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   1230
         Width           =   1365
      End
      Begin VB.TextBox txtRamo_Atividade 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   360
         Left            =   1530
         TabIndex        =   4
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ramo Atividade(Convênio)"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   330
         Width           =   2265
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   3360
      TabIndex        =   15
      Top             =   2460
      Width           =   2910
      Begin VB.OptionButton optCodigo_Ordenar 
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   885
      End
      Begin VB.OptionButton optAlfabetica_Ordenar 
         Caption         =   "Alfabética"
         Height          =   240
         Left            =   1590
         TabIndex        =   9
         Top             =   330
         Width           =   1185
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
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      Height          =   240
      Left            =   90
      TabIndex        =   19
      Top             =   30
      Width           =   1290
   End
End
Attribute VB_Name = "frmRelatorio_Senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviço                                                        '
' Objetivo...............: Relatório Senha                                                '
' Data de Criação........: 07/03/2006                                                     '
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
    Call Impressao
    Unload frmAguarde
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
    log.Programa = "Relatório Senha"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Relatório Senha"
    'Gravando o log
    log.Gravar_log "Otica", Me
   
    optCodigo_Ordenar.Value = True
    optAtivos_Todos.Value = True
    optBloqueados_Todos.Value = True
        
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
            
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    strSql = "SELECT PKCodigo_TBRamo_atividade,DFDescricao_TBRamo_atividade FROM TBRamo_atividade"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBRamo_atividade", "DFDescricao_TBRamo_atividade", dtcRamo_Atividade, strSql, "BDRetaguarda", "Otica", Me
            
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
    
                               
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando o Relatorio Senha"
    
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
    log.Descricao = "Cancelamento da Relação de Senha"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    optCodigo_Ordenar.Value = True
    optAtivos_Todos.Value = True
    optBloqueados_Todos.Value = True
    
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

Private Sub txtRamo_Atividade_Change()
    dtcRamo_Atividade.BoundText = txtRamo_Atividade.Text
    If IsNumeric(txtRamo_Atividade.Text) = False Then txtRamo_Atividade.Text = Empty: Exit Sub
End Sub

Function Impressao()
    
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente,DFSenha_TBCliente FROM TBCliente " & _
             "WHERE IXCodigo_TBCliente is not null "
    
    
    If dtcCliente.BoundText <> Empty Then
       strSql = strSql & "AND IXCodigo_TBCliente = '" & dtcCliente.BoundText & "' "
    End If
    
    If dtcRamo_Atividade.BoundText <> Empty Then
       strSql = strSql & "AND FKCodigo_TBRamo_atividade = '" & dtcRamo_Atividade.BoundText & "' "
    End If
    
    If optAtivos_Sim.Value = True Then
       strSql = strSql & "AND DFInativo_TBCliente = '1' "
    ElseIf optAtivos_Nao.Value = True Then
       strSql = strSql & "AND DFInativo_TBCliente = '0' "
    End If

    If optBloqueados_Sim.Value = True Then
       strSql = strSql & "AND DFBloqueado_TBCliente = '1' "
    ElseIf optBloqueados_Sim.Value = True Then
       strSql = strSql & "AND DFBloqueado_TBCliente = '0' "
    End If
       
    If optAlfabetica_Ordenar.Value = True Then
       strSql = strSql & "ORDER BY DFNome_TBCliente "
    Else
       strSql = strSql & "ORDER BY IXCodigo_TBCliente "
    End If
      
    Call frmConsole_Relatorio_Senha.Show
End Function

Private Sub txtRamo_Atividade_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub


Private Sub txtRamo_Atividade_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub
