VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMovimentacoes_Processamento_Senha_Portal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processamento de Senha Portal"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimentacoes_Processamento_Senha_Portal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5520
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1005
      Left            =   1650
      TabIndex        =   11
      Top             =   2760
      Width           =   3765
      Begin VB.Label lblAreprocessar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2910
         TabIndex        =   15
         Top             =   660
         Width           =   705
      End
      Begin VB.Label lblAprocessar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2910
         TabIndex        =   14
         Top             =   330
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Qtde. Registros a Reprocessar.:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   630
         Width           =   2745
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Qtde. Registros a Processar.....:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   2760
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reprocessar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1005
      Left            =   90
      TabIndex        =   10
      Top             =   2760
      Width           =   1545
      Begin VB.OptionButton optReprocessar_Nao 
         Caption         =   "Não"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   675
      End
      Begin VB.OptionButton optReprocessar_Sim 
         Caption         =   "Sim"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   330
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Cliente"
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
      Left            =   90
      TabIndex        =   6
      Top             =   990
      Width           =   5325
      Begin VB.TextBox txtRamo_Atividade 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   570
         Width           =   1230
      End
      Begin VB.TextBox txtCodigo_Cliente 
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   1230
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   360
         Left            =   1410
         TabIndex        =   3
         Top             =   1230
         Width           =   3795
         _ExtentX        =   6694
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
         Left            =   1410
         TabIndex        =   1
         Top             =   570
         Width           =   3795
         _ExtentX        =   6694
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
         Caption         =   "Ramo Atividade(Convênio)"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   330
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   990
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Confirmar"
            Object.ToolTipText     =   "Gravar registro - CTRL+G"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sair"
            Object.ToolTipText     =   "Sair - CTRL+S"
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   390
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
            Picture         =   "frmMovimentacoes_Processamento_Senha_Portal.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Processamento_Senha_Portal.frx":1A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Processamento_Senha_Portal.frx":1DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Processamento_Senha_Portal.frx":2150
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Processamento_Senha_Portal.frx":24EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Processamento_Senha_Portal.frx":2804
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   360
      Left            =   90
      TabIndex        =   16
      Top             =   600
      Width           =   5325
      _ExtentX        =   9393
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
      TabIndex        =   17
      Top             =   360
      Width           =   1290
   End
End
Attribute VB_Name = "frmMovimentacoes_Processamento_Senha_Portal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Movimentações de Planos de Serviços                            '
' Data de Criação........: 07/03/2006                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........: Rodrigo Santos                                                 '
' Data última manutenção.: 07/03/2006                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim conexao As New DLLConexao_Sistema.conexao
Dim I As Integer
Dim booPrivilegio_Incluir As Boolean
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log
Dim rstTodos As New ADODB.Recordset
Dim intcontareprocessar As Integer
Dim booResumo As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 71: Call Gravar   'CTRL+G
                       Case 67: Cancelar 'CTRL+C
                       Case 83: Unload Me  'CTRL+S
                End Select
    End Select
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
    log.Programa = "Gera Processamento de Senha do Portal"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Movimentação Gera Processamento de Senha do Portal"
    'Gravando o log
    log.Gravar_log "Otica", Me

        
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
            
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    strSql = "SELECT PKCodigo_TBRamo_atividade,DFDescricao_TBRamo_atividade FROM TBRamo_atividade"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBRamo_atividade", "DFDescricao_TBRamo_atividade", dtcRamo_Atividade, strSql, "BDRetaguarda", "Otica", Me
            
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
    
    optReprocessar_Nao.Value = True
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    strEvento_log = "Unload"
    
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
        
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Unload")
    Exit Sub
End Sub

Private Sub optReprocessar_Nao_Click()
    
    If booResumo = False Then Exit Sub
    
    booResumo = False

    'CARREGANDO A LABEL
    strSql = "SELECT PKId_TBCliente,DFSenha_TBCliente " & _
             "FROM TBCliente " & _
             "WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "' "
    
    If dtcRamo_Atividade.BoundText <> Empty Then
       strSql = strSql & "AND FKCodigo_TBRamo_atividade = '" & dtcRamo_Atividade.BoundText & "' "
    End If
    
    If dtcCliente.BoundText <> Empty Then
       strSql = strSql & "AND IXCodigo_TBCliente = '" & dtcCliente.BoundText & "' "
    End If
    
    If optReprocessar_Nao.Value = True Then
       strSql = strSql & "AND DFSenha_TBCliente = '' "
    End If
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstTodos, "Otica", Me
    
    Do While rstTodos.EOF = False
       If rstTodos!DFSenha_TBCliente <> Empty Then
          intcontareprocessar = intcontareprocessar + 1
       End If
       rstTodos.MoveNext
    Loop
    
    lblAprocessar.Caption = rstTodos.RecordCount
    lblAreprocessar = intcontareprocessar
    
    intcontareprocessar = 0
    Set rstTodos = Nothing
    
    booResumo = True
End Sub

Private Sub optReprocessar_Sim_Click()
    If booResumo = False Then Exit Sub

    booResumo = False
    
    'CARREGANDO A LABEL
    strSql = "SELECT PKId_TBCliente,DFSenha_TBCliente " & _
             "FROM TBCliente " & _
             "WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "' "
    
    If dtcRamo_Atividade.BoundText <> Empty Then
       strSql = strSql & "AND FKCodigo_TBRamo_atividade = '" & dtcRamo_Atividade.BoundText & "' "
    End If
    
    If dtcCliente.BoundText <> Empty Then
       strSql = strSql & "AND IXCodigo_TBCliente = '" & dtcCliente.BoundText & "' "
    End If
    
    If optReprocessar_Nao.Value = True Then
       strSql = strSql & "AND DFSenha_TBCliente = '' "
    End If
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstTodos, "Otica", Me
    
    Do While rstTodos.EOF = False
       If rstTodos!DFSenha_TBCliente <> Empty Then
          intcontareprocessar = intcontareprocessar + 1
       End If
       rstTodos.MoveNext
    Loop
    
    lblAprocessar.Caption = rstTodos.RecordCount
    lblAreprocessar = intcontareprocessar
    
    intcontareprocessar = 0
    Set rstTodos = Nothing
    
    booResumo = True
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Gravar
           Case 2: Call Cancelar
           Case 4: Unload Me
    End Select
End Sub

Function Gravar()

    Dim strCampo As String
    Dim strValores As String
    Dim strDescricao As String
    Dim strSenha As String
    Dim intResp As Integer
    Dim rstAplicacao As New ADODB.Recordset
    
    If optReprocessar_Sim.Value = True Then
       intResp = MsgBox("Esta operação irá alterar a senha de todos os usuários selecionados! Deseja prosseguir?", vbOKCancel + vbExclamation, "Only Tech")
       If intResp = 2 Then Exit Function
    End If
    
    frmAguarde.Show
    
    'Abrindo conexao
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    On Error GoTo Erro_transacao

    log.Evento = "Incluir Novo"
    
    strSql = "SELECT PKId_TBCliente " & _
             "FROM TBCliente " & _
             "WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "' "
    
    If dtcRamo_Atividade.BoundText <> Empty Then
       strSql = strSql & "AND FKCodigo_TBRamo_atividade = '" & dtcRamo_Atividade.BoundText & "' "
    End If
    
    If dtcCliente.BoundText <> Empty Then
       strSql = strSql & "AND IXCodigo_TBCliente = '" & dtcCliente.BoundText & "' "
    End If
    
    If optReprocessar_Nao.Value = True Then
       strSql = strSql & "AND DFSenha_TBCliente = '' "
    End If
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    strSenha = Empty
    intContador = 0
    
    Do While rstAplicacao.EOF = False
       strSenha = Funcoes_Gerais.GerarChave(6)
       
       strSql = "UPDATE TBCliente " & _
                "SET DFSenha_TBCliente = '" & strSenha & "', " & _
                "DFintegrado_TBCliente = '" & 0 & "' " & _
                "WHERE PKId_TBCliente = '" & rstAplicacao!PKId_TBCliente & "' "

       conexao.CNConexao.Execute strSql
       
       rstAplicacao.MoveNext
       intContador = intContador + 1
    Loop

    Set rstAplicacao = Nothing
    
    log.Descricao = "Alterando " & intContador & " registros de senha pelo usuário: " + MDIPrincipal.OCXUsuario.Nome
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando log
    log.Gravar_log "Otica", Me

    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao

    txtRamo_Atividade.Text = Empty
    txtCodigo_Cliente.Text = Empty
    txtRamo_Atividade.SetFocus
    
    Unload frmAguarde
    
    MsgBox "" & intContador & " registro(s) gerado(s) com sucesso.", vbInformation, "Only Tech"
    
    optReprocessar_Nao.Value = True
    
    Exit Function
    
Erro_transacao:
    Unload frmAguarde
    'cancelando as alteracoes
    conexao.CNConexao.RollbackTrans
    'fechando conexao
    conexao.Fechar_conexao
Erro:
    Call Erro.Erro(Me, "Otica", "Gravar")
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)

    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir

    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    cbbControle.Text = Empty
    cbbPeriodo.Text = Empty

    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Cancelar")
    Exit Function
End Function

Private Sub dtcEmpresa_LostFocus()
    If Not IsNumeric(dtcEmpresa.BoundText) Then dtcEmpresa.Text = Empty
    If IsNumeric(dtcEmpresa.Text) Then dtcEmpresa.Text = Empty

    dtcEmpresa.Enabled = False: txtCodigo_Cliente.SetFocus
End Sub

Private Sub dtcCliente_GotFocus()
    booResumo = False
    If txtCodigo_Cliente.Text = Empty Then
        Call Movimentacoes.Verifica_DataCombo(dtcCliente)
    End If
End Sub

Private Sub dtcCliente_LostFocus()
    
    txtCodigo_Cliente.Text = dtcCliente.BoundText
    
    strSql = "SELECT PKId_TBCliente,DFSenha_TBCliente " & _
             "FROM TBCliente " & _
             "WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "' "
    
    If dtcRamo_Atividade.BoundText <> Empty Then
       strSql = strSql & "AND FKCodigo_TBRamo_atividade = '" & dtcRamo_Atividade.BoundText & "' "
    End If
    
    If dtcCliente.BoundText <> Empty Then
       strSql = strSql & "AND IXCodigo_TBCliente = '" & dtcCliente.BoundText & "' "
    End If
    
    If optReprocessar_Nao.Value = True Then
       strSql = strSql & "AND (DFSenha_TBCliente = '' OR DFSenha_TBCliente IS NULL)  "
    End If
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstTodos, "Otica", Me
    
    If rstTodos.RecordCount <> 0 Then
       Do While rstTodos.EOF = False
          If IsNull(rstTodos!DFSenha_TBCliente) = False And rstTodos!DFSenha_TBCliente <> Empty Then
             intcontareprocessar = intcontareprocessar + 1
          End If
          rstTodos.MoveNext
       Loop
    End If
    
    lblAprocessar.Caption = rstTodos.RecordCount
    lblAreprocessar = intcontareprocessar
    
    intcontareprocessar = 0
    Set rstTodos = Nothing
        
    booResumo = True
    
    If IsNumeric(txtCodigo_Cliente.Text) = False Or dtcCliente.Text = Empty Then txtCodigo_Cliente.Text = Empty: Exit Sub
End Sub

Private Sub dtcRamo_Atividade_GotFocus()
    booResumo = False
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
    
    'CARREGANDO A LABEL
    strSql = "SELECT PKId_TBCliente,DFSenha_TBCliente " & _
             "FROM TBCliente " & _
             "WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "' "
    
    If dtcRamo_Atividade.BoundText <> Empty Then
       strSql = strSql & "AND FKCodigo_TBRamo_atividade = '" & dtcRamo_Atividade.BoundText & "' "
    End If
    
    If dtcCliente.BoundText <> Empty Then
       strSql = strSql & "AND IXCodigo_TBCliente = '" & dtcCliente.BoundText & "' "
    End If
    
    If optReprocessar_Nao.Value = True Then
       strSql = strSql & "AND (DFSenha_TBCliente = '' OR DFSenha_TBCliente IS NULL) "
    End If
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstTodos, "Otica", Me
    
    If rstTodos.RecordCount <> 0 Then
       Do While rstTodos.EOF = False
          If IsNull(rstTodos!DFSenha_TBCliente) = False And rstTodos!DFSenha_TBCliente <> Empty Then
             intcontareprocessar = intcontareprocessar + 1
          End If
          rstTodos.MoveNext
       Loop
    End If
    
    lblAprocessar.Caption = rstTodos.RecordCount
    lblAreprocessar = intcontareprocessar
    
    intcontareprocessar = 0
    Set rstTodos = Nothing
    
    booResumo = True
    
    If IsNumeric(txtRamo_Atividade.Text) = False Or dtcRamo_Atividade.Text = Empty Then txtRamo_Atividade.Text = Empty: Exit Sub
    
End Sub

Private Sub txtCodigo_Cliente_Change()
    dtcCliente.BoundText = txtCodigo_Cliente.Text
    If IsNumeric(txtCodigo_Cliente.Text) = False Then txtCodigo_Cliente.Text = Empty: Exit Sub
End Sub

Private Sub txtCodigo_Cliente_GotFocus()
    booResumo = False
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Cliente_LostFocus()
    booResumo = False
    Call dtcCliente_LostFocus
End Sub

Private Sub txtRamo_Atividade_Change()
    dtcRamo_Atividade.BoundText = txtRamo_Atividade.Text
    If IsNumeric(txtRamo_Atividade.Text) = False Then txtRamo_Atividade.Text = Empty: Exit Sub
End Sub

Private Sub txtRamo_Atividade_GotFocus()
    booResumo = False
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtRamo_Atividade_LostFocus()
    booResumo = False
    Call dtcRamo_Atividade_LostFocus
End Sub
