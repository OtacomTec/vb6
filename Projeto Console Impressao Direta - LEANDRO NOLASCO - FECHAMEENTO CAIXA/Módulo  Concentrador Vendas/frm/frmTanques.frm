VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmTanques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tanques"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTanques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6405
   Begin TabDlg.SSTab sstTanque 
      Height          =   2680
      Left            =   0
      TabIndex        =   8
      Top             =   330
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   4736
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Geral"
      TabPicture(0)   =   "frmTanques.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtcEmpresa"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCapacidade"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDescricao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmTanques.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdTanque_Consulta_Empresa"
      Tab(1).Control(1)=   "txtConsulta"
      Tab(1).Control(2)=   "cmdRefresh"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdConsulta"
      Tab(1).Control(4)=   "hfgTanque"
      Tab(1).Control(5)=   "cbbCampos"
      Tab(1).Control(6)=   "cbbConsulta"
      Tab(1).Control(7)=   "Label6"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmdTanque_Consulta_Empresa 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69870
         Picture         =   "frmTanques.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtCodigo 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Código do Tanque"
         Top             =   1440
         Width           =   945
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   1110
         MaxLength       =   40
         TabIndex        =   6
         ToolTipText     =   "Descrição do Tanque"
         Top             =   1440
         Width           =   3525
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -73110
         TabIndex        =   1
         Top             =   780
         Width           =   3165
      End
      Begin VB.CommandButton cmdRefresh 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69090
         Picture         =   "frmTanques.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69480
         Picture         =   "frmTanques.frx":383E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtCapacidade 
         Height          =   375
         Left            =   4680
         MaxLength       =   8
         TabIndex        =   7
         ToolTipText     =   "Capacidade do Tanque"
         Top             =   1440
         Width           =   1575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgTanque 
         Height          =   1365
         Left            =   -74880
         TabIndex        =   3
         Top             =   1200
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   2408
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   -74880
         TabIndex        =   0
         Top             =   780
         Width           =   1725
         _ExtentX        =   3043
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
      Begin AutoCompletar.CbCompleta cbbConsulta 
         Height          =   360
         Left            =   -73110
         TabIndex        =   10
         Top             =   780
         Width           =   3165
         _ExtentX        =   5583
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
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   780
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa [ F2 ]"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   1110
         TabIndex        =   13
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   12
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Capacidade"
         Height          =   240
         Left            =   4680
         TabIndex        =   11
         Top             =   1200
         Width           =   990
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alt + N"
            Description     =   "Novo"
            Object.ToolTipText     =   "Novo registro - CTRL+N"
            ImageIndex      =   4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Confirmar"
            Object.ToolTipText     =   "Gravar registro - CTRL+G"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Excluir"
            Object.ToolTipText     =   "Excluir registro - CTRL+E"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Imprimir"
            Object.ToolTipText     =   "Imprimir - CTRL+I"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sair"
            Object.ToolTipText     =   "Sair - CTRL+S"
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Integração"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6660
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTanques.frx":5538
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTanques.frx":5852
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTanques.frx":5B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTanques.frx":5F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTanques.frx":62A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTanques.frx":65BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTanques.frx":68D4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTanques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Transporte                                                     '
' Objetivo...............: Cadastro Pedágio                                               '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Peixoto                                                  '
' Data de Criação........: 04/03/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public strCodigo_Empresa_Consulta As String
Dim strTamanho As String
Dim strNomes As String
Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim strID_Tanque As String
Public strSql As String
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
Dim log As New DLLSystemManager.log

Function Imprimir()
    On Error GoTo erro
    'Tratamento de Erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    'Call frmConsole_Relatorio_Pedagio.Show
    
    Unload frmAguarde
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty
    cbbConsulta.Text = Empty
    
    If cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       cbbConsulta.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Cobra Eixo Suspenso" Then
       cbbConsulta.Visible = True
       txtConsulta.Visible = False
       cbbConsulta.SetFocus
    Else
       cbbConsulta.Visible = False
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdTanque_Consulta_Empresa_Click()
    'STRING QUE COLETA DADOS RELATIVOS A ACESSIBILIDADE DO USUARIO
    Dim rstAcesso_Consulta_Empresa As New ADODB.Recordset
    
    strSql = Empty
    strSql = "SELECT  DFNivel_TBUsuario FROM TBUsuario " & _
             "WHERE DFNome_TBUsuario = '" & MDIPrincipal.ocxUsuario.Nome & "'"
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAcesso_Consulta_Empresa, "Otica", Me
    
    If rstAcesso_Consulta_Empresa!DFNivel_TBUsuario < 5 Then
       Exit Sub
    End If
    
    Set rstAcesso_Consulta_Empresa = Nothing
    
    Unload frmTanques_Consulta_Empresa
    frmAguarde.Show
    DoEvents
    frmTanques_Consulta_Empresa.Show
    Unload frmAguarde
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub dtcEmpresa_LostFocus()
    dtcEmpresa.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 78: If booPrivilegio_Incluir = True Then Call Novo     'CTRL+N
                       Case 71: If booPrivilegio_Incluir = True Then Call Gravar   'CTRL+G
                       Case 67: If booPrivilegio_Incluir = True Then Call Cancelar 'CTRL+C
                       Case 69: If booPrivilegio_Excluir = True Then Call Excluir  'CTRL+E
                       Case 83: Unload Me  'CTRL+S
                End Select
    End Select
    
    If KeyCode = "113" And booAlterar = False Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
  
Private Sub Form_Load()
    On Error GoTo erro
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.ocxUsuario.Nome
    log.Programa = "Cadastro de Pedágios"
    log.Estacao = MDIPrincipal.ocxUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstTanque, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.ocxUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o cadastro de Pedágios"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    Call Reposicao
    
    sstTanque.TabEnabled(0) = False
    sstTanque.Tab = 1
     
    strCodigo_Empresa_Consulta = "USUARIO"
    
    'INTEGRAÇÃO PORTAL E FILIAIS
   booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.ocxUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.ocxUsuario.Empresa))
   booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.ocxUsuario.Empresa, Me)
      
    Exit Sub
erro:
    Call erro.erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando o cadastro de Pedagio"
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Set log = Nothing
    
    strCombo = Empty
    
    If frmIntegracao.Visible = True Then
        Unload frmIntegracao
    End If
    
    Exit Sub
erro:
    Call erro.erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgTanque_Click()

    If hfgTanque.Col = 0 And hfgTanque.Text <> Empty Then
      
        On Error Resume Next

       'Novo
       tlbBotoes.Buttons.Item(1).Enabled = False
       'Gravar
       tlbBotoes.Buttons.Item(2).Enabled = booPrivilegio_Alterar
       'Cancelar
       tlbBotoes.Buttons.Item(3).Enabled = booPrivilegio_Alterar
       'Excluir
       tlbBotoes.Buttons.Item(4).Enabled = booPrivilegio_Excluir
       'Imprimir
       tlbBotoes.Buttons.Item(5).Enabled = False
       'Integração
       If booIntegra_Portal = True Then
          tlbBotoes.Buttons.Item(9).Enabled = True
       End If
        
       frmAguarde.Show
       DoEvents
        
       txtCodigo.Text = hfgTanque.TextArray((hfgTanque.Row * hfgTanque.Cols + hfgTanque.Col + 1))
       txtDescricao.Text = hfgTanque.TextArray((hfgTanque.Row * hfgTanque.Cols + hfgTanque.Col + 2))
       txtCapacidade.Text = Format(hfgTanque.TextArray((hfgTanque.Row * hfgTanque.Cols + hfgTanque.Col + 3)), "#,###0.00")
       dtcEmpresa.Text = hfgTanque.TextArray((hfgTanque.Row * hfgTanque.Cols + hfgTanque.Col + 5))
       strID_Tanque = hfgTanque.TextArray((hfgTanque.Row * hfgTanque.Cols + hfgTanque.Col + 6))
       
       booAlterar = True
       
       sstTanque.TabEnabled(0) = True
       sstTanque.Tab = 0
       txtCodigo.Enabled = False
       dtcEmpresa.Enabled = False
       Me.txtDescricao.SetFocus
   End If
   
   Unload frmAguarde
   
End Sub

Private Sub hfgTanque_DblClick()
    hfgTanque.Sort = 1
End Sub

Private Sub hfgTanque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgTanque_Click
    End If
End Sub

Private Sub sstTanque_Click(PreviousTab As Integer)
    If sstTanque.Tab = 0 Then
       If txtCodigo.Enabled = True Then
          txtCodigo.SetFocus
       Else
          txtDescricao.SetFocus
       End If
    End If
    If sstTanque.Tab = 1 Then
        If frmIntegracao.Visible = True Then
            Unload frmIntegracao
        End If
        If strCombo <> Empty And strCombo <> "Todos" And strCombo <> "Cobra Eixo Suspenso" Then
           cbbCampos.Text = strCombo
           txtConsulta.SetFocus
        ElseIf strCombo = "Todos" Then
           hfgTanque.Row = 1
           hfgTanque.Col = 0
           hfgTanque.SetFocus
        End If
    End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstTanque.Tab <> 1: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
    End Select
End Sub

Function Gravar()
    On Error GoTo erro
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    
    Call Objetos.Retira_Espaco_Lateral(Me)
    Call Objetos.Maiusculo_TXT(Me)
    
    If txtCodigo.Text = Empty Then
       MsgBox "O Código do Tanque não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtCodigo.SetFocus
       Exit Function
    End If
    
    If dtcEmpresa.Text = Empty Then
       MsgBox "O campo Empresa não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       dtcEmpresa.Enabled = True
       dtcEmpresa.SetFocus
       Exit Function
    End If
    
    strCampo = "IXCodigo_TBEmpresa,IXCodigo_TBTanque,DFDescricao_TBTanque,DFCapacidade_TBTanque," & _
               "DFData_alteracao_TBTanque,DFIntegrado_filiais_TBTanque"
               
    If booIntegra_Portal = True Then
       strCampo = strCampo & ",DFIntegrado_portal_TBTanque"
    End If
               
    strValores = "" & dtcEmpresa.BoundText & "," & txtCodigo.Text & "," & _
                 "'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & Funcoes_Gerais.Grava_Moeda(txtCapacidade.Text) & "," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0"
     
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
    End If
    
     If booAlterar = True Then
        log.Evento = "Alterar"
        strSet = "SET IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ",IXCodigo_TBTanque = " & txtCodigo.Text & ", " & _
                 "    DFDescricao_TBTanque = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & " ', " & _
                 "    DFCapacidade_TBTanque = " & Funcoes_Gerais.Grava_Moeda(txtCapacidade.Text) & "," & _
                 "    DFData_alteracao_TBTanque = '" & Format(Date, "YYYYMMDD") & "'," & _
                 "    DFIntegrado_filiais_TBTanque = 0 "
                 
        If booIntegra_Portal = True Then
           strSet = strSet & ",DFIntegrado_portal_TBTanque = 0"
        End If
                 
        Call funcoes_banco.Alterar("TBTanque", strSet, "PKId_TBTanque", strID_Tanque, "OTICA", Me, "BDRetaguarda")
        log.Descricao = "Alterando o registro de ID: " + "strId_Tanque"
        log.Tipo = 1
        log.Hora = Format(Now, "hh:mm:ss")
        'Gravando log
        log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBTanque", strCampo, strValores, "Otica", Me, "BDRetaguarda")
       log.Descricao = "Gravando o registro de ID: " + "strId_Tanque"
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me

    End If
    
    Call Objetos.Limpa_TXT(Me)
        
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgTanque.Visible = False
    End If
    
    sstTanque.TabEnabled(0) = False
    sstTanque.Tab = 1
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBTanque", "PKId_TBTanque", strID_Tanque, "Otica", Me, "BDRetaguarda")
    
    'Iniciando conexao
    Conexao.Initial_Catalog = "BDRetaguarda"
    Conexao.Abrir_conexao ("Otica")
    
    Conexao.CNconexao.BeginTrans
    
    strSql = "DELETE FROM TBTanque WHERE PKId_TBTanque = " & strID_Tanque & " AND IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
    
    Conexao.CNconexao.Execute strSql
    
    Conexao.CNconexao.CommitTrans
    Conexao.Fechar_conexao
    
    Call Objetos.Limpa_TXT(Me)
               
    'Novo
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = False
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = False
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = False
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgTanque.Visible = False
    End If
        
    sstTanque.TabEnabled(0) = False
    sstTanque.Tab = 1
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Excluir")
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo erro
    
    Call Objetos.Limpa_TXT(Me)
    
    'Novo
     tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = False
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = False
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = booPrivilegio_Excluir
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgTanque.Visible = False
    End If
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstTanque.TabEnabled(0) = False
    sstTanque.Tab = 1
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo erro
    
    Dim rstBusca_Parametro As New ADODB.Recordset
    
    sstTanque.TabEnabled(0) = True
    sstTanque.Tab = 0
    
    Call Objetos.Limpa_TXT(Me)
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")

    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    'Novo
    tlbBotoes.Buttons.Item(1).Enabled = False
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = booPrivilegio_Incluir
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = booPrivilegio_Incluir
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = False
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = False
    
    sstTanque.TabEnabled(0) = True
    sstTanque.Tab = 0
    
    'dtcCodigo_empresa.boundtext = ---- Inserir aqui informações da DLLIntercomunicador de EXE's
    dtcEmpresa.BoundText = MDIPrincipal.ocxUsuario.Empresa
    dtcEmpresa.Enabled = False
   
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    booAlterar = False
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    On Error GoTo erro
    
    strNomes = "Código,Descrição,Capacidade,Empresa,Nome,ID"
    strTamanho = "1000,2500,1500,1200,2500,0"
    
    Movimentacoes.Monta_HFlex_Grid hfgTanque, strTamanho, strNomes, 6, "Otica", Me
    
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    dtcEmpresa.BoundText = MDIPrincipal.ocxUsuario.Empresa
    
    Call Monta_Combo
          
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Reposicao")
    Resume Next
End Function

Private Sub txtCodigo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> Empty And booAlterar = False Then
       Movimentacoes.Verifica_Numero "IXCodigo_TBTanque", "TBTanque", txtCodigo, "OTICA", Me, "IXCodigo_TBEmpresa", dtcEmpresa.BoundText
    End If
End Sub

Private Sub txtDescricao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDescricao_LostFocus()
    txtDescricao.Text = UCase(txtDescricao.Text)
End Sub

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtCapacidade_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCapacidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código")
    cbbCampos.AddItem ("Descrição")
    cbbCampos.AddItem ("Capacidade")
    cbbCampos.AddItem ("Código Empresa")
    cbbCampos.AddItem ("Nome Empresa")
End Function

Private Function Consulta()

    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbInformation, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
          
    strSql = "SELECT TBTanque.IXCodigo_TBTanque, " & _
             "TBTanque.DFDescricao_TBTanque, " & _
             "TBTanque.DFCapacidade_TBTanque, " & _
             "IXCodigo_TBEmpresa,DFRazao_Social_TBEmpresa,PKId_TBTanque " & _
             "FROM TBTanque " & _
             "INNER JOIN TBEmpresa " & _
             "ON TBTanque.IXCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa "
                     
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código" Then
          strSql = strSql & " WHERE IXCodigo_TBTanque = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Descrição" Then
          strSql = strSql & " WHERE convert(nvarchar,DFDescricao_TBTanque) LIKE '%" & txtConsulta.Text & "%'"
       ElseIf cbbCampos.Text = "Capacidade" Then
          If Not IsNumeric(txtConsulta.Text) Then
               txtConsulta.Text = Empty
               txtConsulta.SetFocus
               Exit Function
          End If
          txtConsulta.Text = Format(txtConsulta.Text, "#,###0.00")
          strSql = strSql & " WHERE convert(money,DFCapacidade_TBTanque) = " & Funcoes_Gerais.Grava_Moeda(txtConsulta) & ""
       ElseIf cbbCampos.Text = "Código Empresa" Then
          strSql = strSql & " WHERE IXCodigo_TBEmpresa = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Nome Empresa" Then
          strSql = strSql & " WHERE convert(nvarchar,TBEmpresa.DFRazao_Social_TBEmpresa) LIKE '%" & txtConsulta.Text & "%'"
       End If
       If IsNumeric(strCodigo_Empresa_Consulta) = False Then
          strSql = strSql & " AND IXCodigo_TBEmpresa = '" & MDIPrincipal.ocxUsuario.Empresa & "' "
       ElseIf IsNumeric(strCodigo_Empresa_Consulta) = True And CDbl(strCodigo_Empresa_Consulta) <> 0 Then
          strSql = strSql & " AND IXCodigo_TBEmpresa = '" & strCodigo_Empresa_Consulta & "' "
       End If
    Else
       If IsNumeric(strCodigo_Empresa_Consulta) = False Then
          strSql = strSql & " WHERE IXCodigo_TBEmpresa = '" & MDIPrincipal.ocxUsuario.Empresa & "' "
       ElseIf IsNumeric(strCodigo_Empresa_Consulta) = True And CDbl(strCodigo_Empresa_Consulta) <> 0 Then
          strSql = strSql & " WHERE IXCodigo_TBEmpresa = '" & strCodigo_Empresa_Consulta & "' "
       End If
    End If
                             
    frmAguarde.Show
    DoEvents
    
    strSql = strSql & " ORDER BY TBTanque.IXCodigo_TBTanque"
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgTanque, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    hfgTanque.Row = 1
    hfgTanque.Col = 0
    If hfgTanque.Text = Empty Then
       hfgTanque.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgTanque, strTamanho, strNomes, 6, "Otica", Me
    End If
    
    Unload frmAguarde
    hfgTanque.SetFocus
    
End Function

Private Sub txtCapacidade_LostFocus()
    txtCapacidade.Text = Format(txtCapacidade.Text, "#,###0.00")
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("IXCodigo_TBTanque", txtCodigo.Text, "DFIntegrado_filiais_TBTanque", "TBTanque", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBTanque", Me.Top, Me.Left, Me.width, Me.Height, "Tanque")
    
End Function



