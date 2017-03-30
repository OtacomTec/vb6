VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorio_Impressao_Carteirinha_Gerente 
   Caption         =   "Impressão de Carteirinha Gerente"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   Icon            =   "frmRelatorio_Impressao_Carteirinha_Gerente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbbFuncao 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   150
      TabIndex        =   3
      Top             =   1890
      Width           =   5895
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
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Limpa os Filtros"
      Top             =   2490
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
      Left            =   3540
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Visualiza Impressão"
      Top             =   2490
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
      Height          =   1785
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   6045
      Begin VB.TextBox txtFuncionario 
         Height          =   345
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Código do Operador"
         Top             =   630
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dtcFuncionario 
         Height          =   360
         Left            =   1200
         TabIndex        =   2
         Top             =   630
         Width           =   4785
         _ExtentX        =   8440
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Função"
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
         Left            =   75
         TabIndex        =   9
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Funcionário"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   5
         Top             =   390
         Width           =   990
      End
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   360
      Left            =   90
      TabIndex        =   7
      Top             =   240
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
      Left            =   90
      TabIndex        =   8
      Top             =   0
      Width           =   1290
   End
End
Attribute VB_Name = "frmRelatorio_Impressao_Carteirinha_Gerente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador_Vendas                                                    '
' Objetivo...............: Relatório Impressão de Carteirinha Gerente                               '
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
 Call Impressao
End Sub

Private Sub dtcEmpresa_LostFocus()
    If dtcEmpresa.BoundText <> Empty Then
        strSQL = "SELECT PKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf FROM TBOperadores_ecf WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
        Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcEmpresa, strSQL, "BDRetaguarda", "Otica", Me
    Else
        strSQL = "SELECT PKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf FROM TBOperadores_ecf "
        Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcFuncionario, strSQL, "BDRetaguarda", "Otica", Me
    End If
    txtFuncionario.Text = Empty
    dtcEmpresa.Enabled = False
End Sub

Private Sub Form_Load()
    On Error GoTo Erro
    

    
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Relatório Impressão de Carterinha Gerente"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
        Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando o Relatório Impressão de Carteirinha Gerente"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    'Montando os datacombo de tela
    strSQL = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSQL, "BDRetaguarda", "Otica", Me
    
    'dtcCodigo_empresa.boundtext = ---- Inserir aqui informações da DLLIntercomunicador de EXE's
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    strSQL = "SELECT PKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf FROM TBOperadores_ecf WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcFuncionario, strSQL, "BDRetaguarda", "Otica", Me
    
    cbbFuncao.Clear
    cbbFuncao.AddItem ("CAIXA")
    cbbFuncao.AddItem ("FISCAL")
    cbbFuncao.AddItem ("SUPERVISOR")
    cbbFuncao.AddItem ("SUB-GERENTE")
    cbbFuncao.AddItem ("GERENTE")
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando a Impressão de Carteirinha Gerente"
    
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
    log.Descricao = "Cancelamento do Relatório de Impressão de Carteirinha para Gerente"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Impressao()

    Dim strImpressao As String
    
    If cbbFuncao.Text = "CAIXA" Then
       strNivel = 1
    ElseIf cbbFuncao.Text = "FISCAL" Then
       strNivel = 2
    ElseIf cbbFuncao.Text = "SUPERVISOR" Then
       strNivel = 3
    ElseIf cbbFuncao.Text = "SUB-GERENTE" Then
       strNivel = 4
    Else
       strNivel = 5
    End If
       

    strSQL = Empty
    strSQL = "SELECT TBOperadores_ecf.PKCodigo_TBOperadores_ecf," & _
             "TBOperadores_ecf.DFNome_TBOperadores_ecf," & _
             "TBOperadores_ecf.DFNivel_TBOperadores_ecf," & _
             "TBOperadores_ecf.DFNumero_cartao_TBOperadores_ecf," & _
             "TBOperadores_ecf.DFSenha_TBOperadores_ecf," & _
             "TBOperadores_ecf.FKCodigo_TBEmpresa," & _
             "TBOperadores_ecf.DFCodigo_Identificador_TBOperadores_ecf" & _
             "FROM TBOperadores_ecf  " & _
             "INNER JOIN TBEmpresa " & _
             "ON TBOperadores_ecf.FKCodigo_TBEmpresa  = TBEmpresa.PKCodigo_TBEmpresa"

   
    If dtcEmpresa.BoundText <> "" Then
       strSQL = strSQL + " AND TBEmpresa.PKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
    End If
            
    If dtcFuncionario.BoundText <> "" Then
       strSQL = strSQL + " AND PKCodigo_TBOperadores_ecf = " & dtcFuncionario.BoundText & " "
    End If
    
    If cbbFuncao.Text <> "" Then
       strSQL = strSQL + " AND DFNivel_TBOperadores_ecf = " & strNivel & " "
    End If




            
    Call frmConsole_Relatorio_Impressao_Carteirinha_Gerente.Show
    
End Function

Private Sub dtcFuncionario_LostFocus()
    txtFuncionario.Text = dtcFuncionario.BoundText
    If IsNumeric(txtFuncionario.Text) = False Or dtcFuncionario.Text = Empty Then txtFuncionario.Text = Empty: Exit Sub
End Sub

Private Sub dtcFuncionario_GotFocus()
    If Me.txtFuncionario.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFuncionario.Text)
    End If
End Sub

Private Sub txtFuncionario_Change()
    dtcFuncionario.BoundText = txtFuncionario.Text
    If IsNumeric(txtFuncionario.Text) = False Then
       txtFuncionario.Text = Empty
       Exit Sub
    End If
End Sub

Private Sub txtFuncionario_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFuncionario_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
  
