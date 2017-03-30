VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmFinalizadora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finalizadora"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
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
   ScaleHeight     =   4260
   ScaleWidth      =   6015
   Begin TabDlg.SSTab sstFinalizadora 
      Height          =   3915
      Left            =   0
      TabIndex        =   17
      Top             =   330
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6906
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
      TabPicture(0)   =   "frmFinalizadora.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtcPlano_Pagamento"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cbbImpDireto"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbbTipo_Finalizadora"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbbTroco"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cbbDeb_Cred"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cbbAcres_Desc"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cbbModalidade"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCodigo_Finalizadora"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDescricao"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtPercentual"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtCodificacao_impressora"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCodAsc"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtPlano_Pagamento"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdInformacoes_Adicionais"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmFinalizadora.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "cbbCampos"
      Tab(1).Control(2)=   "hfgFinalizadora"
      Tab(1).Control(3)=   "cmdConsulta"
      Tab(1).Control(4)=   "cmdRefresh"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtConsulta"
      Tab(1).Control(6)=   "cbbConsulta"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdInformacoes_Adicionais 
         Height          =   360
         Left            =   5520
         Picture         =   "frmFinalizadora.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1440
         Width           =   345
      End
      Begin VB.TextBox txtPlano_Pagamento 
         Height          =   360
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Código do Plano de Pagamento"
         Top             =   1440
         Width           =   1275
      End
      Begin VB.TextBox txtCodAsc 
         Height          =   360
         Left            =   3960
         MaxLength       =   1
         TabIndex        =   16
         Top             =   3390
         Width           =   1875
      End
      Begin VB.ComboBox cbbConsulta 
         Height          =   360
         Left            =   -73080
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtCodificacao_impressora 
         Height          =   360
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2100
         Width           =   1875
      End
      Begin VB.TextBox txtPercentual 
         Height          =   360
         Left            =   3960
         TabIndex        =   13
         Top             =   2760
         Width           =   1905
      End
      Begin VB.TextBox txtDescricao 
         Height          =   360
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   5
         Top             =   780
         Width           =   4425
      End
      Begin VB.TextBox txtCodigo_Finalizadora 
         Height          =   360
         Left            =   120
         MaxLength       =   6
         TabIndex        =   4
         ToolTipText     =   "Código da Finalizadora"
         Top             =   780
         Width           =   1275
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -73080
         TabIndex        =   18
         Top             =   720
         Width           =   3075
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
         Left            =   -69510
         Picture         =   "frmFinalizadora.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   -69930
         Picture         =   "frmFinalizadora.frx":2B86
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Consultar"
         Top             =   720
         Width           =   405
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgFinalizadora 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   3
         Top             =   1140
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4683
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
         Left            =   120
         TabIndex        =   8
         Top             =   2100
         Width           =   1875
         _ExtentX        =   3307
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
         Left            =   2040
         TabIndex        =   12
         Top             =   2760
         Width           =   1875
         _ExtentX        =   3307
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
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1875
         _ExtentX        =   3307
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
         Left            =   2040
         TabIndex        =   15
         Top             =   3390
         Width           =   1875
         _ExtentX        =   3307
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
         Left            =   120
         TabIndex        =   14
         Top             =   3390
         Width           =   1875
         _ExtentX        =   3307
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
      Begin AutoCompletar.CbCompleta cbbImpDireto 
         Height          =   360
         Left            =   3960
         TabIndex        =   10
         Top             =   2100
         Width           =   1905
         _ExtentX        =   3360
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
      Begin MSDataListLib.DataCombo dtcPlano_Pagamento 
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         Top             =   1440
         Width           =   4065
         _ExtentX        =   7170
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Plano Pagamento"
         Height          =   240
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label11 
         Caption         =   "Tecla de Atalho"
         Height          =   315
         Left            =   3960
         TabIndex        =   31
         Top             =   3150
         Width           =   1725
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Finalizadora"
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   3150
         Width           =   1470
      End
      Begin VB.Label Label10 
         Caption         =   "Impr. Direto"
         Height          =   315
         Left            =   3960
         TabIndex        =   29
         Top             =   1860
         Width           =   1305
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Impressora"
         Height          =   240
         Left            =   2040
         TabIndex        =   28
         Top             =   1860
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Troco"
         Height          =   240
         Left            =   2040
         TabIndex        =   27
         Top             =   3150
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Débito/Crédito"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Acréscimo/Desconto"
         Height          =   240
         Left            =   2040
         TabIndex        =   25
         Top             =   2520
         Width           =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Percentual"
         Height          =   240
         Left            =   3960
         TabIndex        =   24
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Finalizadora"
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   21
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
         Left            =   120
         TabIndex        =   20
         Top             =   1860
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":4880
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":4B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":4EB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":524E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":55E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":5902
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":5C1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
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
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Cadastro Finalizadora                                          '
' Data de Criação........: 21/01/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strID_Finalizadora As String
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
    
    Call frmConsole_Relatorio_Finalizadora.Show
        
    Unload frmAguarde
        
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty
    cbbConsulta.Text = Empty
    
    If cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       cbbConsulta.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Imp. Direto" Or cbbCampos.Text = "Troco" _
           Or cbbCampos.Text = "Débito/Crédito" Or cbbCampos.Text = "Acréscimo/Desconto" _
           Or cbbCampos.Text = "Modalidade" Then
              txtConsulta.Visible = False
              cbbConsulta.Visible = True
              cbbConsulta.SetFocus
    ElseIf cbbCampos.Text = "Código Finalizadora" Or cbbCampos.Text = "Descrição Finalizadora" _
           Or cbbCampos.Text = "Percentual" Or cbbCampos.Text = "Codificação Impressora" _
           Or cbbCampos.Text = "Código Plano" Or cbbCampos.Text = "Descrição Plano" Then
              txtConsulta.Visible = True
              cbbConsulta.Visible = False
              txtConsulta.SetFocus
    Else
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cbbCampos_LostFocus()

'Monta a Combo Consulta de acordo com o campo selecionado RVS
    cbbConsulta.Clear
    If cbbCampos.Text = "Imp. Direto" Or cbbCampos.Text = "Troco" Then
        cbbConsulta.AddItem "Sim"
        cbbConsulta.AddItem "Não"
    ElseIf cbbCampos.Text = "Acréscimo/Desconto" Then
        cbbConsulta.AddItem "Acréscimo"
        cbbConsulta.AddItem "Desconto"
    ElseIf cbbCampos.Text = "Débito/Crédito" Then
        cbbConsulta.AddItem "Débito"
        cbbConsulta.AddItem "Crédito"
    ElseIf cbbCampos.Text = "Modalidade" Then
        cbbConsulta.AddItem "Banco"
        cbbConsulta.AddItem "Carteira"
        cbbConsulta.AddItem "Pré-Datado"
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

Private Sub cmdInformacoes_Adicionais_Click()
    If Me.dtcPlano_pagamento.Text <> " " Then
       Unload frmInformacoes_Adicionais_Plano_Pagamento
       Call frmInformacoes_Adicionais_Plano_Pagamento.Info_Plano(dtcPlano_pagamento.BoundText, MDIPrincipal.OCXUsuario.Empresa, "Otica", "BDRetaguarda", Me.Top, Me.Left, Me.width, Me.Height)
    End If
End Sub

Private Sub dtcPlano_Pagamento_GotFocus()
    If Me.txtPlano_pagamento.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(txtPlano_pagamento.Text)
    End If
End Sub

Private Sub dtcPlano_Pagamento_LostFocus()
    txtPlano_pagamento.Text = dtcPlano_pagamento.BoundText
    If IsNumeric(txtPlano_pagamento.Text) = False Or dtcPlano_pagamento.Text = Empty Then
        txtPlano_pagamento.Text = Empty
    End If
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
    
    strSQL = "SELECT IXCodigo_TBPlano_pagamento,DFDescricao_TBPlano_pagamento " & _
             "FROM TBPlano_pagamento WHERE IXCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "'"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBPlano_pagamento", "DFDescricao_TBPlano_pagamento", dtcPlano_pagamento, strSQL, "BDRetaguarda", "Otica", Me
    
    log.Descricao = "Inicializando cadastro de Finalizadora"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    
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
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgFinalizadora_Click()
    If hfgFinalizadora.Col = 0 And hfgFinalizadora.Text <> Empty Then
        
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
       
       strID_Finalizadora = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 1))
       txtCodigo_Finalizadora.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 2))
       txtDescricao.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 3))
       cbbModalidade.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 4))
       cbbAcres_Desc.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 5))
       txtPercentual.Text = Format(hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 6)), "#,###0.00")
       cbbDeb_Cred.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 7))
       cbbTroco.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 8))
       cbbTipo_Finalizadora.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 9))
       txtCodificacao_impressora.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 10))
       cbbImpDireto.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 11))
       'Definindo o tamanho do text de acordo com o texto grid
       txtCodAsc.MaxLength = Len(hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 12)))
       txtCodAsc.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 12))
       txtPlano_pagamento.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 13))
       dtcPlano_pagamento.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 14))
       
       'KeyCodeConstants = ValueConstants.tbrPressed
       booAlterar = True
       txtConsulta.Text = Empty
       sstFinalizadora.TabEnabled(0) = True
       sstFinalizadora.Tab = 0
       txtCodigo_Finalizadora.Enabled = False
       txtDescricao.SetFocus
   End If
   Unload frmAguarde
End Sub

Private Sub hfgFinalizadora_DblClick()
    hfgFinalizadora.Sort = 1
End Sub

Private Sub hfgFinalizadora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgFinalizadora_Click
    End If
End Sub

Private Sub Image2_Click()

End Sub

Private Sub sstFinalizadora_Click(PreviousTab As Integer)
    If sstFinalizadora.Tab = 0 Then
       txtDescricao.SetFocus
    ElseIf sstFinalizadora.Tab = 1 Then
       If frmIntegracao.Visible = True Then
          Unload frmIntegracao
       End If
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
    Dim strModalidade As String
    Dim strTroco As String
    Dim strDeb_Cred As String
    Dim strAcres_Desc As String
    Dim strTipo_Finalizadora As String
    Dim strImpDireto As String
    Dim intAsc As Integer
    Dim intId_Plano As Integer
    
    Call Objetos.Retira_Espaco_Lateral(Me)
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
    
    If cbbImpDireto.Text = "Sim" Then
       strImpDireto = 1
    Else
       strImpDireto = 0
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
    
    If txtCodAsc.Text = "ESPAÇO" Then
       intAsc = 32
    ElseIf txtCodAsc.Text = "F1" Then
       intAsc = 112
    ElseIf txtCodAsc.Text = "F2" Then
       intAsc = 113
    ElseIf txtCodAsc.Text = "F3" Then
       intAsc = 114
    ElseIf txtCodAsc.Text = "F4" Then
       intAsc = 115
    ElseIf txtCodAsc.Text = "F5" Then
       intAsc = 116
    ElseIf txtCodAsc.Text = "F6" Then
       intAsc = 117
    ElseIf txtCodAsc.Text = "F7" Then
       intAsc = 118
    ElseIf txtCodAsc.Text = "F8" Then
        intAsc = 119
    ElseIf txtCodAsc.Text = "F9" Then
       intAsc = 120
    ElseIf txtCodAsc.Text = "F10" Then
       intAsc = 121
    ElseIf txtCodAsc.Text = "F11" Then
       intAsc = 122
    ElseIf txtCodAsc.Text = "F12" Then
       intAsc = 123
    Else
       If txtCodAsc.Text <> Empty Then
          intAsc = Asc(txtCodAsc.Text)
       Else
          intAsc = 0
       End If
    End If
    
    intId_Plano = Funcoes_Gerais.Localiza_ID("PKId_TBPlano_pagamento", "IXCodigo_TBPlano_pagamento", dtcPlano_pagamento.BoundText, "TBPlano_pagamento", "Otica", Me, "BDRetaguarda")

    strCampo = "IXCodigo_TBFinalizadora," & _
               "DFDescricao_TBFinalizadora," & _
               "DFModalidade_TBFinalizadora," & _
               "DFAcrescimo_desconto_TBFinalizadora," & _
               "DFPercentual_TBFinalizadora," & _
               "DFDebito_credito_TBFinalizadora," & _
               "DFTroco_TBFinalizadora," & _
               "DFControle_venda_TBFinalizadora," & _
               "DFCodificacao_impressora_fiscal_TBFinalizadora," & _
               "DFImprime_direto_TBFinalizadora," & _
               "DFCodigo_asc_TBFinalizadora," & _
               "DFData_alteracao_TBFinalizadora," & _
               "DFIntegrado_filiais_TBFinalizadora," & _
               "DFCodigo_link_plano_pagamento_Identificador_TBFinalizadora"
               
    If booIntegra_Portal = True Then
       strCampo = strCampo & ",DFIntegrado_portal_TBFinalizadora"
    End If

    strValores = "" & txtCodigo_Finalizadora.Text & "," & _
                 "'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                 "" & strModalidade & "," & _
                 "" & strAcres_Desc & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtPercentual.Text) & "," & _
                 "" & strDeb_Cred & "," & _
                 "" & strTroco & "," & _
                 "" & strTipo_Finalizadora & "," & _
                 "'" & txtCodificacao_impressora.Text & "'," & _
                 "" & strImpDireto & "," & _
                 "'" & intAsc & "'," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0," & intId_Plano & ""
                 
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
    End If

    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFDescricao_TBFinalizadora = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                "    DFModalidade_TBFinalizadora =" & strModalidade & "," & _
                "    DFAcrescimo_desconto_TBFinalizadora = " & strAcres_Desc & "," & _
                "    DFPercentual_TBFinalizadora = " & Funcoes_Gerais.Grava_Moeda(txtPercentual.Text) & "," & _
                "    DFDebito_credito_TBFinalizadora = " & strDeb_Cred & "," & _
                "    DFTroco_TBFinalizadora = " & strTroco & "," & _
                "    DFControle_venda_TBFinalizadora = " & strTipo_Finalizadora & "," & _
                "    DFCodificacao_impressora_fiscal_TBFinalizadora =  '" & txtCodificacao_impressora.Text & "'," & _
                "    DFImprime_direto_TBFinalizadora =  '" & strImpDireto & "'," & _
                "    DFCodigo_asc_TBFinalizadora =  '" & intAsc & "'," & _
                "    DFData_alteracao_TBFinalizadora = '" & Format(Date, "YYYYMMDD") & "'," & _
                "    DFIntegrado_filiais_TBFinalizadora = 0 ," & _
                "    DFCodigo_link_plano_pagamento_Identificador_TBFinalizadora = '" & intId_Plano & "'"
                
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBFinalizadora = 0"
       End If

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
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
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
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
  
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
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
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
          
    
    Call Limpa_Combos
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

Private Sub txtCodAsc_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodAsc_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 112 Then
       txtCodAsc.MaxLength = 2
       txtCodAsc.Text = "F1"
    ElseIf KeyCode = 113 Then
       txtCodAsc.MaxLength = 2
       txtCodAsc.Text = "F2"
    ElseIf KeyCode = 114 Then
       txtCodAsc.MaxLength = 2
       txtCodAsc.Text = "F3"
    ElseIf KeyCode = 115 Then
       txtCodAsc.MaxLength = 2
       txtCodAsc.Text = "F4"
    ElseIf KeyCode = 116 Then
       txtCodAsc.MaxLength = 2
       txtCodAsc.Text = "F5"
    ElseIf KeyCode = 117 Then
       txtCodAsc.MaxLength = 2
       txtCodAsc.Text = "F6"
    ElseIf KeyCode = 118 Then
       txtCodAsc.MaxLength = 2
       txtCodAsc.Text = "F7"
    ElseIf KeyCode = 119 Then
       txtCodAsc.MaxLength = 2
       txtCodAsc.Text = "F8"
    ElseIf KeyCode = 120 Then
       txtCodAsc.MaxLength = 2
       txtCodAsc.Text = "F9"
    ElseIf KeyCode = 121 Then
       txtCodAsc.MaxLength = 3
       txtCodAsc.Text = "F10"
    ElseIf KeyCode = 122 Then
       txtCodAsc.MaxLength = 3
       txtCodAsc.Text = "F11"
    ElseIf KeyCode = 123 Then
       txtCodAsc.MaxLength = 3
       txtCodAsc.Text = "F12"
    End If
    
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodAsc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       txtCodAsc.MaxLength = 6
       txtCodAsc.Text = "ESPAÇO"
    Else
       txtCodAsc.MaxLength = 1
    End If
    
End Sub

Private Sub txtCodAsc_LostFocus()
    txtCodAsc.Text = UCase(txtCodAsc)
End Sub

Private Sub txtCodificacao_impressora_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

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
          
    strTamanho = "0,1000,2000,1300,1500,1000,1300,1000,1300,1500,1150,1200,1200,2500"
    strNomes = "ID,Finalizadora,Descrição,Modalidade,Acréscimo/Desc.,Percentual,Débito/Crédito,Troco,Tipo Finalizadora,Cod. Impressora,Imp. Direto,Tecla Atalho,Código,Plano Pagamento"
    
    Movimentacoes.Monta_HFlex_Grid hfgFinalizadora, strTamanho, strNomes, 14, "OTICA", Me
    
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
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty And cbbConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbInformation, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
      
    Dim strModalidade As String
    Dim strTroco As String
    Dim strDeb_Cred As String
    Dim strAcres_Desc As String
    Dim strTipo_Finalizadora As String
    Dim strImpDireto As String
    Dim intId_Plano_Consulta As Integer
    
    If cbbCampos.Text = "Modalidade" Then
       If cbbConsulta.Text = "Banco" Then
          strModalidade = 1
       ElseIf cbbConsulta.Text = "Carteira" Then
          strModalidade = 2
       Else
          strModalidade = 3
       End If
    End If
          
    If cbbCampos.Text = "Acréscimo/Desconto" Then
       If cbbConsulta.Text = "Acréscimo" Then
          strAcres_Desc = 1
       Else
          strAcres_Desc = 0
       End If
    End If
    
    If cbbCampos.Text = "Débito/Crédito" Then
       If cbbConsulta.Text = "Débito" Then
          strDeb_Cred = 0
       Else
          strDeb_Cred = 1
       End If
    End If
    
    If cbbCampos.Text = "Troco" Then
       If cbbConsulta.Text = "Sim" Then
          strTroco = 1
       Else
          strTroco = 0
       End If
    End If
    
    If cbbCampos.Text = "Imp. Direto" Then
       If cbbConsulta.Text = "Sim" Then
          strImpDireto = 1
       Else
          strImpDireto = 0
       End If
    End If
    
    If cbbCampos.Text = "Tipo Finalizadora" Then
       If cbbConsulta.Text = "Venda" Then
          strTipo_Finalizadora = 1
       Else
          strTipo_Finalizadora = 0
       End If
    End If
    
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
           
    txtConsulta.Text = Funcoes_Gerais.Grava_String(txtConsulta.Text)
    
    strSQL = "SELECT TBFinalizadora.PKId_TBFinalizadora,TBFinalizadora.IXCodigo_TBFinalizadora," & _
             "TBFinalizadora.DFDescricao_TBFinalizadora,TBFinalizadora.DFModalidade_TBFinalizadora," & _
             "TBFinalizadora.DFAcrescimo_desconto_TBFinalizadora,TBFinalizadora.DFPercentual_TBFinalizadora," & _
             "TBFinalizadora.DFDebito_credito_TBFinalizadora,TBFinalizadora.DFTroco_TBFinalizadora," & _
             "TBFinalizadora.DFControle_venda_TBFinalizadora, TBFinalizadora.DFCodificacao_impressora_fiscal_TBFinalizadora," & _
             "TBFinalizadora.DFImprime_direto_TBFinalizadora,TBFinalizadora.DFCodigo_asc_TBFinalizadora," & _
             "TBPlano_pagamento.IXCodigo_TBPlano_pagamento , TBPlano_pagamento.DFDescricao_TBPlano_pagamento " & _
             "FROM TBFinalizadora " & _
             "LEFT JOIN TBPlano_pagamento ON " & _
             "TBFinalizadora.DFCodigo_link_plano_pagamento_Identificador_TBFinalizadora = TBPlano_pagamento.PKId_TBPlano_pagamento "


    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código Finalizadora" Then
          strSQL = strSQL & " WHERE convert(nvarchar,IXCodigo_TBFinalizadora) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Descrição Finalizadora" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFDescricao_TBFinalizadora) LIKE '%" & txtConsulta.Text & "%'"
       ElseIf cbbCampos.Text = "Modalidade" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFModalidade_TBFinalizadora) = '" & strModalidade & "'"
       ElseIf cbbCampos.Text = "Acréscimo/Desconto" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFAcrescimo_desconto_TBFinalizadora) = " & strAcres_Desc & ""
       ElseIf cbbCampos.Text = "Percentual" Then
          strSQL = strSQL & " WHERE convert(money,DFPercentual_TBFinalizadora) =  " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Débito/Crédito" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFDebito_credito_TBFinalizadora) = '" & strDeb_Cred & "'"
       ElseIf cbbCampos.Text = "Troco" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFTroco_TBFinalizadora) = '" & strTroco & "'"
       ElseIf cbbCampos.Text = "Troco" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFControle_venda_TBFinalizadora) = '" & strTipo_Finalizadora & "'"
       ElseIf cbbCampos.Text = "Codificação Impressora" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFCodificacao_impressora_fiscal_TBFinalizadora) LIKE '%" & txtConsulta.Text & "%'"
       ElseIf cbbCampos.Text = "Imp. Direto" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFImprime_direto_TBFinalizadora) = '" & strImpDireto & "'"
       ElseIf cbbCampos.Text = "Tecla Atalho" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFCodigo_asc_TBFinalizadora) =  '" & Asc(txtConsulta.Text) & "'"
       ElseIf cbbCampos.Text = "Código Plano" Then
          strSQL = strSQL & " WHERE convert(nvarchar,TBPlano_pagamento.IXCodigo_TBPlano_pagamento) =  '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Descrição Plano" Then
          strSQL = strSQL & " WHERE convert(nvarchar,TBPlano_pagamento.DFDescricao_TBPlano_pagamento) LIKE  '%" & txtConsulta.Text & "%'"
       End If
    End If
    
    frmAguarde.Show
    DoEvents
    
    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgFinalizadora, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
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
           hfgFinalizadora.Col = 12
           If hfgFinalizadora.Text <> Empty And hfgFinalizadora.Text <> " " Then
              If hfgFinalizadora.Text = 32 Then
                 hfgFinalizadora.Text = "ESPAÇO"
              ElseIf hfgFinalizadora = 112 Then
                 hfgFinalizadora.Text = "F1"
              ElseIf hfgFinalizadora = 113 Then
                 hfgFinalizadora.Text = "F2"
              ElseIf hfgFinalizadora = 114 Then
                 hfgFinalizadora.Text = "F3"
              ElseIf hfgFinalizadora = 115 Then
                 hfgFinalizadora.Text = "F4"
              ElseIf hfgFinalizadora = 116 Then
                 hfgFinalizadora.Text = "F5"
              ElseIf hfgFinalizadora = 117 Then
                 hfgFinalizadora.Text = "F6"
              ElseIf hfgFinalizadora = 118 Then
                 hfgFinalizadora.Text = "F7"
              ElseIf hfgFinalizadora = 119 Then
                 hfgFinalizadora.Text = "F8"
              ElseIf hfgFinalizadora = 120 Then
                 hfgFinalizadora.Text = "F9"
              ElseIf hfgFinalizadora = 121 Then
                 hfgFinalizadora.Text = "F10"
              ElseIf hfgFinalizadora = 122 Then
                 hfgFinalizadora.Text = "F11"
              ElseIf hfgFinalizadora = 123 Then
                 hfgFinalizadora.Text = "F12"
              Else
                 hfgFinalizadora.Text = Chr(hfgFinalizadora.Text)
              End If
              
           End If
       Next I
    End If
    
    hfgFinalizadora.Row = 1
    hfgFinalizadora.Col = 0
    If hfgFinalizadora.Text = Empty Then
       hfgFinalizadora.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgFinalizadora, strTamanho, strNomes, 14, "Otica", Me
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
    cbbCampos.AddItem ("Codificação Impressora")
    cbbCampos.AddItem ("Imp. Direto")
    cbbCampos.AddItem ("Tecla Atalho")
    cbbCampos.AddItem ("Código Plano")
    cbbCampos.AddItem ("Descrição Plano")
    
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
    
    cbbImpDireto.Clear
    cbbImpDireto.AddItem ("Sim")
    cbbImpDireto.AddItem ("Não")
    
End Function

Private Sub txtDescricao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDescricao_LostFocus()
    txtDescricao.Text = UCase(txtDescricao.Text)
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

Private Function Limpa_Combos()
    cbbAcres_Desc.Text = Empty
    cbbTipo_Finalizadora.Text = Empty
    cbbModalidade.Text = Empty
    cbbTroco.Text = Empty
    cbbDeb_Cred.Text = Empty
    cbbImpDireto.Text = Empty
End Function

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKId_TBFinalizadora", strID_Finalizadora, "DFIntegrado_filiais_TBFinalizadora", "TBFinalizadora", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBFinalizadora", Me.Top, Me.Left, Me.width, Me.Height, "Finalizadora")
    
End Function

Private Sub txtPlano_pagamento_Change()
     dtcPlano_pagamento.BoundText = txtPlano_pagamento.Text
End Sub

Private Sub txtPlano_pagamento_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPlano_Pagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub
