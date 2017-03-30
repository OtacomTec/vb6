VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmOperador_Ecf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operador ECF"
   ClientHeight    =   3015
   ClientLeft      =   3255
   ClientTop       =   2670
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOperador_Ecf.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6090
   Begin TabDlg.SSTab sstOperador 
      Height          =   2685
      Left            =   0
      TabIndex        =   10
      Top             =   330
      Width           =   6075
      _ExtentX        =   10716
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
      TabPicture(0)   =   "frmOperador_Ecf.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label14"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dtcEmpresa"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbbNivel_Operador"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCodigo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtNome"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtNumero_Cartao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtSenha"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmOperador_Ecf.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtConsulta"
      Tab(1).Control(1)=   "hfgOperador"
      Tab(1).Control(2)=   "cbbCampos"
      Tab(1).Control(3)=   "cmdRefresh"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdConsulta"
      Tab(1).Control(5)=   "Label6"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtSenha 
         Alignment       =   2  'Center
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3720
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   2130
         Width           =   2205
      End
      Begin VB.TextBox txtNumero_Cartao 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1680
         TabIndex        =   8
         Top             =   2130
         Width           =   1995
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72870
         TabIndex        =   1
         Top             =   720
         Width           =   2925
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1440
         Width           =   4665
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Código do Operador"
         Top             =   1440
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgOperador 
         Height          =   1365
         Left            =   -74880
         TabIndex        =   3
         Top             =   1200
         Width           =   5835
         _ExtentX        =   10292
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
         Top             =   720
         Width           =   1965
         _ExtentX        =   3466
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
      Begin AutoCompletar.CbCompleta cbbNivel_Operador 
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   2130
         Width           =   1515
         _ExtentX        =   2672
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
         Top             =   810
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
         Left            =   -69450
         Picture         =   "frmOperador_Ecf.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   720
         Width           =   405
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
         Left            =   -69870
         Picture         =   "frmOperador_Ecf.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Consultar"
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         Height          =   240
         Left            =   3750
         TabIndex        =   19
         Top             =   1890
         Width           =   540
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cartão"
         Height          =   240
         Left            =   1710
         TabIndex        =   17
         Top             =   1890
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nível Operador"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   240
         Left            =   1260
         TabIndex        =   12
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6090
      _ExtentX        =   10742
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
      Left            =   6240
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
            Picture         =   "frmOperador_Ecf.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador_Ecf.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador_Ecf.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador_Ecf.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador_Ecf.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador_Ecf.frx":5578
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperador_Ecf.frx":5892
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOperador_Ecf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Cadastro Operador de ECF                                       '
' Data de Criação........: 17/01/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim I As Integer
Dim strTamanho As String
Dim strNomes As String
Dim strCombo As String
Dim strConsulta As String
Dim booAlterar As Boolean
Public strSQL As String
Dim log As New DLLSystemManager.log
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
Option Explicit

Function Imprimir()
    On Error GoTo Erro
    'Tratamento de erro
    If strSQL = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       'cbbCampos.SetFocus
       Me.txtConsulta.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Relatorio_Operador_Ecf.Show
        
    Unload frmAguarde
        
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty
    
    If cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    Else
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
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
    log.Programa = "Cadastro de Operador de ECF"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstOperador, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando cadastro de Operador de ECF"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    
    sstOperador.TabEnabled(0) = False
    sstOperador.Tab = 1
        
    Call Reposicao
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando cadastro de Seção"
        
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    strCombo = Empty
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgOperador_Click()
    If hfgOperador.Col = 0 Then
        
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
       
       txtCodigo.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 1))
       txtNome.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 2))
       cbbNivel_Operador.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 3))
       txtNumero_Cartao.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 4))
       txtSenha.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 5))
       dtcEmpresa.BoundText = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 6))
            
       booAlterar = True
       txtConsulta.Text = Empty
       sstOperador.TabEnabled(0) = True
       sstOperador.Tab = 0
       txtCodigo.Enabled = False
       txtSenha.Enabled = False
       txtNome.SetFocus
   End If
   Unload frmAguarde
End Sub

Private Sub hfgOperador_DblClick()
    hfgOperador.Sort = 1
End Sub

Private Sub hfgOperador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgOperador_Click
    End If
End Sub

Private Sub sstOperador_Click(PreviousTab As Integer)
    If sstOperador.Tab = 0 Then
       txtNome.SetFocus
    ElseIf sstOperador.Tab = 1 Then
       If frmIntegracao.Visible = True Then
          Unload frmIntegracao
       End If
       If strCombo <> Empty And strCombo <> "Todos" Then
          cbbCampos.Text = strCombo
          txtConsulta.SetFocus
       ElseIf strCombo = "Todos" Then
          hfgOperador.Row = 1
          hfgOperador.Col = 0
          hfgOperador.SetFocus
       End If
    End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
        End Select
End Sub

Function Gravar()
    On Error GoTo Erro
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim strNivel As String
        
    If txtCodigo.Text = Empty Then
       MsgBox "Código não pode ser nulo.", vbInformation, "Only Tech"
       txtCodigo.SetFocus
       Exit Function
    End If
    
    If cbbNivel_Operador.Text = "Caixa" Then
       strNivel = 1
    ElseIf cbbNivel_Operador.Text = "Fiscal" Then
       strNivel = 2
    ElseIf cbbNivel_Operador.Text = "Supervisor" Then
       strNivel = 3
    ElseIf cbbNivel_Operador.Text = "Sub-Gerente" Then
       strNivel = 4
    Else
       strNivel = 5
    End If
       
    Call Objetos.Retira_Espaco_Lateral(Me)
    Call Objetos.Maiusculo_TXT(Me)
    
    strCampo = "PKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf," & _
               "DFNivel_TBOperadores_ecf,DFNumero_cartao_TBOperadores_ecf," & _
               "DFSenha_TBOperadores_ecf,FKCodigo_TBEmpresa,DFData_alteracao_TBOperadores_ecf," & _
               "DFIntegrado_filiais_TBOperadores_ecf"
               
    If booIntegra_Portal = True Then
       strCampo = strCampo & ",DFIntegrado_portal_TBOperadores_ecf"
    End If
    
    strValores = "" & txtCodigo.Text & ",'" & Funcoes_Gerais.Grava_String(txtNome.Text) & "'," & _
                 "" & strNivel & "," & txtNumero_Cartao.Text & "," & _
                 "'" & Funcoes_Gerais.Grava_String(txtSenha.Text) & "'," & dtcEmpresa.BoundText & "," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
    End If
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFNome_TBOperadores_ecf = '" & Funcoes_Gerais.Grava_String(txtNome.Text) & "'," & _
                "    DFNivel_TBOperadores_ecf = " & strNivel & "," & _
                "    DFNumero_cartao_TBOperadores_ecf = " & txtNumero_Cartao.Text & "," & _
                "    DFSenha_TBOperadores_ecf = " & txtSenha.Text & "," & _
                "    FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & "," & _
                "    DFData_alteracao_TBOperadores_ecf = '" & Format(Date, "YYYYMMDD") & "'," & _
                "    DFIntegrado_filiais_TBOperadores_ecf = 0 "
                
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBOperadores_ecf = 0"
       End If
       
       Call funcoes_banco.Alterar("TBOperadores_ecf", strSet, "PKCodigo_TBOperadores_ecf", txtCodigo.Text, "OTICA", Me, "BDRetaguarda")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBOperadores_ecf", strCampo, strValores, "OTICA", Me, "BDRetaguarda")
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
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
       hfgOperador.Visible = False
    End If
    
    txtSenha.Enabled = False
    sstOperador.TabEnabled(0) = False
    sstOperador.Tab = 1
    hfgOperador.Refresh
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBOperadores_ecf", "PKCodigo_TBOperadores_ecf", txtCodigo.Text, "OTICA", Me, "BDRetaguarda")
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
        
    'Gravando log
    log.Gravar_log "OTICA", Me
           
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
       hfgOperador.Visible = False
    End If
            
    sstOperador.TabEnabled(0) = False
    sstOperador.Tab = 1
    
    Exit Function
Erro:
     Call Erro.Erro(Me, "OTICA", "Excluir")
     Exit Function
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
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
       hfgOperador.Visible = False
    End If
    
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Operação com Registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    txtSenha.Enabled = False
    sstOperador.TabEnabled(0) = False
    sstOperador.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
    Call Reposicao
    
    Call Objetos.Limpa_TXT(Me)
    
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
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
    
    sstOperador.TabEnabled(0) = True
    sstOperador.Tab = 0
    
    txtSenha.Enabled = True
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    booAlterar = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Sub txtCodigo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> Empty And booAlterar = False Then
       Movimentacoes.Verifica_Numero "PKCodigo_TBOperadores_ecf", "TBOperadores_ecf", txtCodigo, "OTICA", Me
    End If
End Sub

Private Function Reposicao()
    On Error GoTo Erro
          
    strTamanho = "1000,2000,1300,1300,0,1000,2000"
    strNomes = "Operador,Nome,Nível Operador,Nº Cartão,Senha,Empresa,Nome"
    
    Movimentacoes.Monta_HFlex_Grid hfgOperador, strTamanho, strNomes, 7, "OTICA", Me
    
    Call Monta_Combo
    Call Monta_DataCombo
              
    hfgOperador.Refresh
    Exit Function
Erro:
   Call Erro.Erro(Me, "OTICA", "Reposicao")
   Resume Next
End Function

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Sub txtNome_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNome_LostFocus()
    txtNome.Text = UCase(txtNome.Text)
End Sub

Private Function Consulta()
    
    Dim strNivel As String
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
    
    If cbbCampos.Text = "Nível do Operador" Then
       If txtConsulta.Text = "CAIXA" Then
          strNivel = 1
       ElseIf txtConsulta.Text = "FISCAL" Then
          strNivel = 2
       ElseIf txtConsulta.Text = "SUPERVISOR" Then
          strNivel = 3
       ElseIf txtConsulta.Text = "SUB-GERENTE" Then
          strNivel = 4
       Else
          strNivel = 5
       End If
    End If
            
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
           
    txtConsulta.Text = Funcoes_Gerais.Grava_String(txtConsulta.Text)
    
    strSQL = "SELECT PKCodigo_TBOperadores_ecf," & _
             "DFNome_TBOperadores_ecf," & _
             "DFNivel_TBOperadores_ecf," & _
             "DFNumero_cartao_TBOperadores_ecf," & _
             "DFSenha_TBOperadores_ecf," & _
             "TBOperadores_ecf.FKCodigo_TBEmpresa," & _
             "TBEmpresa.DFRazao_Social_TBEmpresa " & _
             "FROM TBOperadores_ecf " & _
             "INNER JOIN TBEmpresa ON TBOperadores_ecf.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa"
           
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código do Operador" Then
          strSQL = strSQL & " WHERE convert(nvarchar,PKCodigo_TBOperadores_ecf) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Nome do Operador" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFNome_TBOperadores_ecf) LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Nível do Operador" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFNivel_TBOperadores_ecf) = " & strNivel & ""
       ElseIf cbbCampos.Text = "Nº do Cartão" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFNumero_cartao_TBOperadores_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Código da Empresa" Then
          strSQL = strSQL & " WHERE convert(nvarchar,FKCodigo_TBEmpresa) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Nome da Empresa" Then
          strSQL = strSQL & " WHERE TBEmpresa.DFRazao_Social_TBEmpresa LIKE '%" & txtConsulta.Text & "%' "
       End If
    End If
    
    frmAguarde.Show
    DoEvents
     
    strSQL = strSQL & " ORDER BY TBOperadores_ecf.PKCodigo_TBOperadores_ecf"
    
    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgOperador, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    If hfgOperador.Rows > 1 Then
       For I = 1 To hfgOperador.Rows - 1
           hfgOperador.Row = I
           hfgOperador.Col = 3
           If hfgOperador.Text = "1" Then
              hfgOperador.Text = "Caixa"
           ElseIf hfgOperador.Text = "2" Then
              hfgOperador.Text = "Fiscal"
           ElseIf hfgOperador.Text = "3" Then
              hfgOperador.Text = "Supervisor"
           ElseIf hfgOperador.Text = "4" Then
              hfgOperador.Text = "Sub-Gerente"
           ElseIf hfgOperador.Text = "5" Then
              hfgOperador.Text = "Gerente"
           End If
       Next I
    End If
                
    Unload frmAguarde
    hfgOperador.Refresh
    hfgOperador.Row = 1
    hfgOperador.Col = 0
    hfgOperador.SetFocus
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código do Operador")
    cbbCampos.AddItem ("Nome do Operador")
    cbbCampos.AddItem ("Nível do Operador")
    cbbCampos.AddItem ("Nº do Cartão")
    cbbCampos.AddItem ("Código da Empresa")
    cbbCampos.AddItem ("Nome da Empresa")
    
    cbbNivel_Operador.Clear
    cbbNivel_Operador.AddItem ("Caixa")
    cbbNivel_Operador.AddItem ("Fiscal")
    cbbNivel_Operador.AddItem ("Supervisor")
    cbbNivel_Operador.AddItem ("Sub-Gerente")
    cbbNivel_Operador.AddItem ("Gerente")
    
End Function
Private Function Monta_DataCombo()
    
    strSQL = Empty
    strSQL = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSQL, "BDRetaguarda", "Otica", Me
    
End Function
Private Sub txtNumero_Cartao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_Cartao_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo_TBOperadores_ecf", txtCodigo.Text, "DFIntegrado_filiais_TBOperadores_ecf", "TBOperadores_ecf", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBOperadores_ecf", Me.Top, Me.Left, Me.width, Me.Height, "Operador ECF")
    
End Function
