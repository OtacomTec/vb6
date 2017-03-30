VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmFinalizadora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finalizadora"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinalizadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5550
   Begin TabDlg.SSTab sstFinalizadora 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmFinalizadora.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtPercentual"
      Tab(0).Control(1)=   "txtDescricao"
      Tab(0).Control(2)=   "txtCodigo_Finalizadora"
      Tab(0).Control(3)=   "cbbModalidade"
      Tab(0).Control(4)=   "cbbAcres_Desc"
      Tab(0).Control(5)=   "cbbDeb_Cred"
      Tab(0).Control(6)=   "cbbTroco"
      Tab(0).Control(7)=   "cbbTipo_Finalizadora"
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(12)=   "Label3"
      Tab(0).Control(13)=   "Label7"
      Tab(0).Control(14)=   "Label8"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmFinalizadora.frx":179E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cbbCampos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgFinalizadora"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdOrdenar"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdConsulta"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdRefresh"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtConsulta"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtPercentual 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   -71160
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1440
         Width           =   1545
      End
      Begin VB.TextBox txtDescricao 
         Height          =   360
         Left            =   -73800
         MaxLength       =   40
         TabIndex        =   2
         Top             =   780
         Width           =   4185
      End
      Begin VB.TextBox txtCodigo_Finalizadora 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   -74880
         MaxLength       =   6
         TabIndex        =   1
         Top             =   781
         Width           =   1035
      End
      Begin VB.TextBox txtConsulta 
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
         Left            =   1920
         TabIndex        =   9
         Top             =   720
         Width           =   2145
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
         Left            =   5010
         Picture         =   "frmFinalizadora.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   4590
         Picture         =   "frmFinalizadora.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Consultar"
         Top             =   720
         Width           =   405
      End
      Begin VB.CommandButton cmdOrdenar 
         Caption         =   "A"
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
         Left            =   4170
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ordenar: (A) Alfabética/ (C) Código "
         Top             =   720
         Width           =   405
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgFinalizadora 
         Height          =   1395
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2461
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
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1755
         _ExtentX        =   3096
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
      Begin AutoCompletar.CbCompleta cbbModalidade 
         Height          =   360
         Left            =   -74880
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
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
      Begin AutoCompletar.CbCompleta cbbAcres_Desc 
         Height          =   360
         Left            =   -73020
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
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
      Begin AutoCompletar.CbCompleta cbbDeb_Cred 
         Height          =   360
         Left            =   -74880
         TabIndex        =   6
         Top             =   2100
         Width           =   1785
         _ExtentX        =   3149
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
      Begin AutoCompletar.CbCompleta cbbTroco 
         Height          =   360
         Left            =   -73050
         TabIndex        =   7
         Top             =   2100
         Width           =   1155
         _ExtentX        =   2037
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
      Begin AutoCompletar.CbCompleta cbbTipo_Finalizadora 
         Height          =   360
         Left            =   -71850
         TabIndex        =   22
         Top             =   2100
         Width           =   2265
         _ExtentX        =   3995
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Finalizadora"
         Height          =   240
         Left            =   -71850
         TabIndex        =   23
         Top             =   1860
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Troco"
         Height          =   240
         Left            =   -73050
         TabIndex        =   21
         Top             =   1860
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Débito/Crédito"
         Height          =   240
         Left            =   -74880
         TabIndex        =   20
         Top             =   1860
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Acréscimo/Desconto"
         Height          =   240
         Left            =   -73020
         TabIndex        =   19
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Percentual"
         Height          =   240
         Left            =   -71160
         TabIndex        =   18
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Finalizadora"
         Height          =   240
         Left            =   -74880
         TabIndex        =   17
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   435
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Modalidade"
         Height          =   240
         Left            =   -74880
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   330
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
            Picture         =   "frmFinalizadora.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":5578
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
      EndProperty
   End
End
Attribute VB_Name = "frmFinalizadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Cadastro Base                                                  '
' Objetivo...............: Cadastro Finalizadora                                          '
' Data de Criação........: 21/01/2005                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes,Sérgio    '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strID_Finalizadora As String
Dim I As Integer
Dim strTamanho As String
Dim strNomes As String
Dim strCombo As String
Dim strConsulta As String
Dim booAlterar As Boolean
Public strSql As String
Dim log As New DLLSystemManager.log
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Option Explicit

Function Imprimir()
    On Error GoTo Erro
    'Tratamento de erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       'cbbCampos.SetFocus
       Me.txtConsulta.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Relatorio_Finalizadora.Show
        
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

Private Sub cmdOrdenar_Click()
    If cmdOrdenar.Caption = "A" Then
       cmdOrdenar.Caption = "C"
    Else
       cmdOrdenar.Caption = "A"
    End If
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
    log.Programa = "Cadastro de Finalizadora"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstFinalizadora, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando cadastro de Finalizadora"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    sstFinalizadora.TabEnabled(0) = False
    sstFinalizadora.Tab = 1
        
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
    log.Descricao = "Finalizando cadastro de Finalizadora"
        
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    strCombo = Empty
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgFinalizadora_Click()
    If hfgFinalizadora.Col = 0 Then
        
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
           
       frmAguarde.Show
       DoEvents
       
       strID_Finalizadora = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 1))
       txtCodigo_Finalizadora.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 2))
       txtDescricao.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 3))
       cbbModalidade.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 4))
       cbbAcres_Desc.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 5))
       txtPercentual.Text = Format(hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 6)), "#,###0.00")
       cbbDeb_Cred.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 7))
       cbbTroco.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 8))
       cbbTipo_Finalizadora.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 9))
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstFinalizadora.TabEnabled(0) = True
       sstFinalizadora.Tab = 0
       txtCodigo_Finalizadora.Enabled = False
       txtDescricao.SetFocus
   End If
   Unload frmAguarde
End Sub

Private Sub hfgFinalizadora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgFinalizadora_Click
    End If
End Sub

Private Sub sstFinalizadora_Click(PreviousTab As Integer)
    If sstFinalizadora.Tab = 0 Then
       txtDescricao.SetFocus
    ElseIf sstFinalizadora.Tab = 1 Then
       If strCombo <> Empty And strCombo <> "Todos" Then
          cbbCampos.Text = strCombo
          txtConsulta.SetFocus
       ElseIf strCombo = "Todos" Then
          hfgFinalizadora.Row = 1
          hfgFinalizadora.Col = 0
          hfgFinalizadora.SetFocus
       End If
    End If
End Sub

Private Sub tlbbotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
        End Select
End Sub

Function Gravar()
    On Error GoTo Erro
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim strModalidade As String
    Dim strTroco As String
    Dim strDeb_Cred As String
    Dim strAcres_Desc As String
    Dim strTipo_Finalizadora As String
             
    Call Objetos.Maiusculo_TXT(Me)
    
    If cbbModalidade.Text = "Banco" Then
       strModalidade = 1
    ElseIf cbbModalidade.Text = "Carteira" Then
       strModalidade = 2
    Else
       strModalidade = 3
    End If
    
    If cbbTroco.Text = "Sim" Then
       strTroco = 1
    Else
       strTroco = 0
    End If
    
    If cbbTipo_Finalizadora.Text = "Controle" Then
       strTipo_Finalizadora = 0
    Else
       strTipo_Finalizadora = 1
    End If
    
    If cbbDeb_Cred.Text = "Débito" Then
       strDeb_Cred = 0
    Else
       strDeb_Cred = 1
    End If
    
    If cbbAcres_Desc.Text = "Acréscimo" Then
       strAcres_Desc = 1
    Else
       strAcres_Desc = 0
    End If
             
    strCampo = "IXCodigo_TBFinalizadora," & _
               "DFDescricao_TBFinalizadora," & _
               "DFModalidade_TBFinalizadora," & _
               "DFAcrescimo_desconto_TBFinalizadora," & _
               "DFPercentual_TBFinalizadora," & _
               "DFDebito_credito_TBFinalizadora," & _
               "DFTroco_TBFinalizadora," & _
               "DFControle_venda_TBFinalizadora"

    strValores = "" & txtCodigo_Finalizadora.Text & "," & _
                 "'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                 "" & strModalidade & "," & _
                 "" & strAcres_Desc & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtPercentual.Text) & "," & _
                 "" & strDeb_Cred & "," & _
                 "" & strTroco & "," & _
                 "" & strTipo_Finalizadora & ""

    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFDescricao_TBFinalizadora = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                "    DFModalidade_TBFinalizadora =" & strModalidade & "," & _
                "    DFAcrescimo_desconto_TBFinalizadora = " & strAcres_Desc & "," & _
                "    DFPercentual_TBFinalizadora = " & Funcoes_Gerais.Grava_Moeda(txtPercentual.Text) & "," & _
                "    DFDebito_credito_TBFinalizadora = " & strDeb_Cred & "," & _
                "    DFTroco_TBFinalizadora = " & strTroco & "," & _
                "    DFControle_venda_TBFinalizadora = " & strTipo_Finalizadora & ""
       Call funcoes_banco.Alterar("TBFinalizadora", strSet, "PKId_TBFinalizadora", strID_Finalizadora, "Otica", Me, "BDRetaguarda")
       log.Descricao = "Alterando o registro: " + txtCodigo_Finalizadora.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBFinalizadora", strCampo, strValores, "OTICA", Me, "BDRetaguarda")
       log.Descricao = "Gravando o registro: " + txtCodigo_Finalizadora.Text
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
    
    If booPrivilegio_Consultar = False Then
       hfgFinalizadora.Visible = False
    End If
    
    sstFinalizadora.TabEnabled(0) = False
    sstFinalizadora.Tab = 1
    hfgFinalizadora.Refresh
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBFinalizadora", "PKId_TBFinalizadora", strID_Finalizadora, "OTICA", Me, "BDRetaguarda")
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtCodigo_Finalizadora.Text
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
    
    If booPrivilegio_Consultar = False Then
       hfgFinalizadora.Visible = False
    End If
            
    sstFinalizadora.TabEnabled(0) = False
    sstFinalizadora.Tab = 1
    
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
    
    If booPrivilegio_Consultar = False Then
       hfgFinalizadora.Visible = False
    End If
    
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Operação com Registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    txtCodigo_Finalizadora.Enabled = False
    sstFinalizadora.TabEnabled(0) = False
    sstFinalizadora.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
          
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
    
    sstFinalizadora.TabEnabled(0) = True
    sstFinalizadora.Tab = 0
    txtCodigo_Finalizadora.Enabled = True
    txtCodigo_Finalizadora.SetFocus
    booAlterar = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Sub txtCodigo_Finalizadora_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Finalizadora_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_Finalizadora_LostFocus()
    If txtCodigo_Finalizadora.Text <> Empty And booAlterar = False Then
       Movimentacoes.Verifica_Numero "IXCodigo_TBFinalizadora", "TBFinalizadora", txtCodigo_Finalizadora, "OTICA", Me
    End If
End Sub

Private Function Reposicao()
    On Error GoTo Erro
          
    strTamanho = "0,1000,2000,1300,1500,1000,1300,1000,1300"
    strNomes = "ID,Finalizadora,Descrição,Modalidade,Acréscimo/Desc.,Percentual,Débito/Crédito,Troco,Tipo Finalizadora"
    
    Movimentacoes.Monta_HFlex_Grid hfgFinalizadora, strTamanho, strNomes, 9, "OTICA", Me
    
    Call Monta_Combo
                  
    hfgFinalizadora.Refresh
    Exit Function
Erro:
   Call Erro.Erro(Me, "OTICA", "Reposicao")
   Resume Next
End Function

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Function Consulta()
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
      
    Dim strModalidade As String
    Dim strTroco As String
    Dim strDeb_Cred As String
    Dim strAcres_Desc As String
    Dim strTipo_Finalizadora As String
    
    If cbbCampos.Text = "Modalidade" Then
       If txtConsulta.Text = 1 Then
          strModalidade = 1
       ElseIf txtConsulta.Text = 2 Then
          strModalidade = 2
       Else
          strModalidade = 3
       End If
    End If
          
    If cbbCampos.Text = "Acréscimo/Desconto" Then
       If txtConsulta.Text = "ACRÉSCIMO" Then
          strAcres_Desc = 1
       Else
          strAcres_Desc = 0
       End If
    End If
    
    If cbbCampos.Text = "Débito/Crédito" Then
       If txtConsulta.Text = "DÉBITO" Then
          strDeb_Cred = 0
       Else
          strDeb_Cred = 1
       End If
    End If
    
    If cbbCampos.Text = "Troco" Then
       If txtConsulta.Text = "Sim" Then
          strTroco = 1
       Else
          strTroco = 0
       End If
    End If
    
    If cbbCampos.Text = "Tipo Finalizadora" Then
       If txtConsulta.Text = "VENDA" Then
          strTipo_Finalizadora = 1
       Else
          strTipo_Finalizadora = 0
       End If
    End If
   
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
           
    txtConsulta.Text = Funcoes_Gerais.Grava_String(txtConsulta.Text)
    
    strSql = "SELECT TBFinalizadora.PKId_TBFinalizadora," & _
             "TBFinalizadora.IXCodigo_TBFinalizadora," & _
             "TBFinalizadora.DFDescricao_TBFinalizadora," & _
             "TBFinalizadora.DFModalidade_TBFinalizadora," & _
             "TBFinalizadora.DFAcrescimo_desconto_TBFinalizadora," & _
             "TBFinalizadora.DFPercentual_TBFinalizadora," & _
             "TBFinalizadora.DFDebito_credito_TBFinalizadora," & _
             "TBFinalizadora.DFTroco_TBFinalizadora," & _
             "TBFinalizadora.DFControle_venda_TBFinalizadora " & _
             "FROM TBFinalizadora "

    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código Finalizadora" Then
          strSql = strSql & " WHERE convert(nvarchar,IXCodigo_TBFinalizadora) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Descrição Finalizadora" Then
          strSql = strSql & " WHERE convert(nvarchar,DFDescricao_TBFinalizadora) LIKE '%" & txtConsulta.Text & "%'"
       ElseIf cbbCampos.Text = "Modalidade" Then
          strSql = strSql & " WHERE convert(nvarchar,DFModalidade_TBFinalizadora) = '" & strModalidade & "'"
       ElseIf cbbCampos.Text = "Acréscimo/Desconto" Then
          strSql = strSql & " WHERE convert(nvarchar,DFAcrescimo_desconto_TBFinalizadora) = " & strAcres_Desc & ""
       ElseIf cbbCampos.Text = "Percentual" Then
          strSql = strSql & " WHERE convert(money,DFPercentual_TBFinalizadora) =  " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Débito/Crédito" Then
          strSql = strSql & " WHERE convert(nvarchar,DFDebito_credito_TBFinalizadora) = '" & strDeb_Cred & "'"
       ElseIf cbbCampos.Text = "Troco" Then
          strSql = strSql & " WHERE convert(nvarchar,DFTroco_TBFinalizadora) = '" & strTroco & "'"
       ElseIf cbbCampos.Text = "Troco" Then
          strSql = strSql & " WHERE convert(nvarchar,DFControle_venda_TBFinalizadora) = '" & strTipo_Finalizadora & "'"
       End If
    End If
    
    frmAguarde.Show
    DoEvents
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgFinalizadora, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    If hfgFinalizadora.Rows > 1 Then
       For I = 1 To hfgFinalizadora.Rows - 1
           hfgFinalizadora.Row = I
           hfgFinalizadora.Col = 4
           If hfgFinalizadora.Text = "1" Then
              hfgFinalizadora.Text = "Banco"
           ElseIf hfgFinalizadora.Text = "2" Then
              hfgFinalizadora.Text = "Carteira"
           Else
              hfgFinalizadora.Text = "Pré-Datado"
           End If
           hfgFinalizadora.Col = 5
           If hfgFinalizadora.Text = "1" Then
              hfgFinalizadora.Text = "Acréscimo"
           Else
              hfgFinalizadora.Text = "Desconto"
           End If
           hfgFinalizadora.Col = 7
           If hfgFinalizadora.Text = "Sim" Then
              hfgFinalizadora.Text = "Crédito"
           Else
              hfgFinalizadora.Text = "Débito"
           End If
           hfgFinalizadora.Col = 9
           If hfgFinalizadora.Text = "Sim" Then
              hfgFinalizadora.Text = "Venda"
           Else
              hfgFinalizadora.Text = "Controle"
           End If
       Next I
    End If
    
    Unload frmAguarde
    hfgFinalizadora.Refresh
    hfgFinalizadora.Row = 1
    hfgFinalizadora.Col = 0
    hfgFinalizadora.SetFocus
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código Finalizadora")
    cbbCampos.AddItem ("Descrição Finalizadora")
    cbbCampos.AddItem ("Modalidade")
    cbbCampos.AddItem ("Acréscimo/Desconto")
    cbbCampos.AddItem ("Percentual")
    cbbCampos.AddItem ("Débito/Crédito")
    cbbCampos.AddItem ("Troco")
        
    cbbAcres_Desc.Clear
    cbbAcres_Desc.AddItem ("Acréscimo")
    cbbAcres_Desc.AddItem ("Desconto")
    
    cbbModalidade.Clear
    cbbModalidade.AddItem ("Banco") ' - 1
    cbbModalidade.AddItem ("Carteira") ' - 2
    cbbModalidade.AddItem ("Pré-Datado") ' - 3
    
    cbbDeb_Cred.Clear
    cbbDeb_Cred.AddItem ("Débito")
    cbbDeb_Cred.AddItem ("Crédito")
    
    cbbTipo_Finalizadora.Clear
    cbbTipo_Finalizadora.AddItem ("Controle")
    cbbTipo_Finalizadora.AddItem ("Venda")
    
    cbbTroco.Clear
    cbbTroco.AddItem ("Sim")
    cbbTroco.AddItem ("Não")
    
End Function

Private Sub txtFinalizadora_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFinalizadora_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPercentual_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPercentual_LostFocus()
    txtPercentual.Text = Format(txtPercentual.Text, "#,###0.00")
End Sub
