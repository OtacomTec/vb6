VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmMovimentacoes_Gera_Plano_Completo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gera Plano Completo"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimentacoes_Gera_Plano_Completo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5250
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Plano"
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
      TabIndex        =   5
      Top             =   360
      Width           =   5055
      Begin VB.TextBox txtCodigo 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox txtDescricao 
         Height          =   360
         Left            =   1140
         MaxLength       =   40
         TabIndex        =   1
         Top             =   570
         Width           =   3795
      End
      Begin VB.TextBox txtLimite 
         Height          =   360
         Left            =   2310
         TabIndex        =   3
         Top             =   1230
         Width           =   1245
      End
      Begin AutoCompletar.CbCompleta cbbControle 
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   1230
         Width           =   2145
         _ExtentX        =   3784
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
      Begin AutoCompletar.CbCompleta cbbPeriodo 
         Height          =   360
         Left            =   3600
         TabIndex        =   4
         Top             =   1230
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   330
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   1140
         TabIndex        =   9
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Controle"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lblControle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite"
         Height          =   240
         Left            =   2310
         TabIndex        =   7
         Top             =   990
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Período"
         Height          =   240
         Left            =   3630
         TabIndex        =   6
         Top             =   990
         Width           =   645
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
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
            Picture         =   "frmMovimentacoes_Gera_Plano_Completo.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Plano_Completo.frx":1A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Plano_Completo.frx":1DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Plano_Completo.frx":2150
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Plano_Completo.frx":24EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Plano_Completo.frx":2804
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMovimentacoes_Gera_Plano_Completo"
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
' Data de Criação........: 25/08/2003                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.: 07/11/2005                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strsql As String
Dim conexao As New DLLConexao_Sistema.conexao
Dim I As Integer
Dim booPrivilegio_Incluir As Boolean
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log

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
    log.Programa = "Gera Plano Completo"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Movimentação Gera Plano Completo"
    'Gravando o log
    log.Gravar_log "Otica", Me

    Call Monta_Combo
    
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
    
    'Verifica se os campos necessarios para gravar não estão nulos
    If txtCodigo.Text = Empty Then
       MsgBox "O campo Código do Serviço não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtCodigo.SetFocus
       Exit Function
    ElseIf txtDescricao.Text = Empty Then
       MsgBox "O campo Descrição do Serviço não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       dtcServico.SetFocus
       Exit Function
    End If

    frmAguarde.Show
    
    strCampo = "PKCodigo_TBPlano_servico,DFDescricao_TBPlano_servico"

    strValores = "" & txtCodigo.Text & ",'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'"
    
    'Abrindo conexao
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    On Error GoTo Erro_transacao

    log.Evento = "Incluir Novo"

    strsql = "INSERT INTO TBPlano_servico (" & strCampo & ") " & _
             "SELECT " & strValores & " "
    
    conexao.CNConexao.Execute strsql
    
    If cbbControle.Text = "Valor Contrato" Then
       intControle = 1
    ElseIf cbbControle.Text = "Serviços" Then
       intControle = 2
    ElseIf cbbControle.Text = "Grupo Serviços" Then
       intControle = 3
    End If
    
    strsql = "INSERT INTO TBPlano_servico_servico_laboratorio " & _
             "(FKCodigo_TBServico_laboratorio," & _
             "FKCodigo_TBPlano_servico,DFQuantidade_TBPlano_servico_servico_laboratorio," & _
             "DFControle_TBPlano_servico_servico_laboratorio," & _
             "DFPeriodo_TBPlano_servico_servico_laboratorio) " & _
             "SELECT PKCodigo_TBServico_laboratorio,'" & txtCodigo.Text & "','" & txtLimite.Text & "'," & _
             "'" & intControle & "','" & cbbPeriodo.Text & "' " & _
             "FROM TBServico_laboratorio "

    conexao.CNConexao.Execute strsql

    log.Descricao = "Incluíndo o registro: " + txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando log
    log.Gravar_log "Otica", Me

    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao

    strDescricao = txtDescricao.Text
    
    txtLimite.Text = Empty
    txtCodigo.Text = Empty
    txtDescricao.Text = Empty
    cbbPeriodo.Text = Empty
    cbbControle.Text = Empty
    
    Unload frmAguarde
    
    MsgBox "Plano " & strDescricao & " gerado com sucesso.", vbInformation, "Only Tech"
    
    strDescricao = Empty
    
    Unload Me
    
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

Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> Empty Then
       Movimentacoes.Verifica_Numero "PKCodigo_TBPlano_servico", "TBPlano_servico", txtCodigo, "OTICA", Me
    End If
End Sub

Private Sub txtDescricao_LostFocus()
    txtDescricao.Text = UCase(txtDescricao.Text)
End Sub

Private Sub txtLimite_LostFocus()
    If Left(txtLimite.Text, 1) = "0" Then
       txtLimite.Text = Right(txtLimite.Text, 1)
    End If
End Sub

Private Sub txtLimite_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Function Monta_Combo()

    cbbControle.Clear
    cbbControle.AddItem ("Valor Contrato")
    cbbControle.AddItem ("Serviços")
    cbbControle.AddItem ("Grupo Serviços")
    
    cbbPeriodo.Clear
    cbbPeriodo.AddItem ("1")
    cbbPeriodo.AddItem ("2")
    cbbPeriodo.AddItem ("3")
    cbbPeriodo.AddItem ("4")
    cbbPeriodo.AddItem ("5")
    cbbPeriodo.AddItem ("6")
    cbbPeriodo.AddItem ("7")
    cbbPeriodo.AddItem ("8")
    cbbPeriodo.AddItem ("9")
    cbbPeriodo.AddItem ("10")
    cbbPeriodo.AddItem ("11")
    cbbPeriodo.AddItem ("12")
    
End Function

Private Sub txtCodigo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDescricao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtLimite_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

