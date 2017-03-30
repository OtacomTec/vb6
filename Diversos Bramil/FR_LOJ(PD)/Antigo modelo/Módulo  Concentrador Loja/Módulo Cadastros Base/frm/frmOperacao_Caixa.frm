VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOperacao_Caixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operação Caixa"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOperacao_Caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7395
   Begin TabDlg.SSTab sstOperador 
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   5847
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
      TabPicture(0)   =   "frmOperacao_Caixa.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbbPdv"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbbStatus"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtcOperadores_Ecf"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtpHora"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dtpData"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dtcFinalizadora"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtFinalizadora"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtOperadores_Ecf"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtValor"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkTipo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtObservacao"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmOperacao_Caixa.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "lblA"
      Tab(1).Control(2)=   "dtpFin"
      Tab(1).Control(3)=   "dtpIni"
      Tab(1).Control(4)=   "cbbCampos"
      Tab(1).Control(5)=   "hfgOperador"
      Tab(1).Control(6)=   "txtConsulta"
      Tab(1).Control(7)=   "cmdRefresh"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdConsulta"
      Tab(1).Control(9)=   "cmdOrdenar"
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtObservacao 
         Height          =   360
         Left            =   120
         MaxLength       =   100
         TabIndex        =   30
         Top             =   2790
         Width           =   7125
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Tipo Operação"
         Height          =   240
         Left            =   5700
         TabIndex        =   24
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   19
         Top             =   780
         Width           =   945
      End
      Begin VB.TextBox txtOperadores_Ecf 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   120
         MaxLength       =   6
         TabIndex        =   16
         Top             =   2100
         Width           =   1335
      End
      Begin VB.TextBox txtFinalizadora 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   120
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
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
         Left            =   -68970
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ordenar: (A) Alfabética/ (C) Código "
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
         Left            =   -68550
         Picture         =   "frmOperacao_Caixa.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Consultar"
         Top             =   720
         Width           =   405
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
         Left            =   -68130
         Picture         =   "frmOperacao_Caixa.frx":34B4
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   720
         Width           =   405
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72870
         TabIndex        =   1
         Top             =   720
         Width           =   3825
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgOperador 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   2
         Top             =   1140
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   3625
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
         TabIndex        =   3
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
      Begin MSDataListLib.DataCombo dtcFinalizadora 
         Height          =   360
         Left            =   1530
         TabIndex        =   9
         Top             =   1440
         Width           =   5730
         _ExtentX        =   10107
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
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   22740993
         CurrentDate     =   37949
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   360
         Left            =   1530
         TabIndex        =   12
         Top             =   780
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   22740994
         CurrentDate     =   37858
      End
      Begin MSDataListLib.DataCombo dtcOperadores_Ecf 
         Height          =   360
         Left            =   1530
         TabIndex        =   17
         Top             =   2100
         Width           =   5730
         _ExtentX        =   10107
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
      Begin AutoCompletar.CbCompleta cbbStatus 
         Height          =   360
         Left            =   3750
         TabIndex        =   21
         Top             =   780
         Width           =   1125
         _ExtentX        =   1984
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
      Begin AutoCompletar.CbCompleta cbbPdv 
         Height          =   360
         Left            =   4920
         TabIndex        =   26
         Top             =   780
         Width           =   675
         _ExtentX        =   1191
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
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   360
         Left            =   -72870
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   22740993
         CurrentDate     =   37949
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   360
         Left            =   -71190
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   22740993
         CurrentDate     =   37949
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
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
         Left            =   -71400
         TabIndex        =   29
         Top             =   870
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   240
         Left            =   3750
         TabIndex        =   22
         Top             =   540
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   240
         Left            =   2760
         TabIndex        =   20
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Operadores ECF"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   1860
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   540
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PDV"
         Height          =   240
         Left            =   4920
         TabIndex        =   14
         Top             =   540
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   240
         Left            =   1530
         TabIndex        =   13
         Top             =   540
         Width           =   405
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Finalizadora"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   435
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
            Picture         =   "frmOperacao_Caixa.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperacao_Caixa.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperacao_Caixa.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperacao_Caixa.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperacao_Caixa.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperacao_Caixa.frx":5578
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   7395
      _ExtentX        =   13044
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
Attribute VB_Name = "frmOperacao_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Cadastro Base                                                  '
' Objetivo...............: Cadastro Operação Caixa                                        '
' Data de Criação........: 17/01/2005                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes,Sérgio    '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strID_Operacao_Caixa As String
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
    
    Call frmConsole_Relatorio_Operacao_Caixa.Show
        
    Unload frmAguarde
        
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty
    
    If cbbCampos.Text = "Todos" Then
       dtpIni.Visible = False
       dtpFin.Visible = False
       lblA.Visible = False
       txtConsulta.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
       Exit Sub
    End If
    
    If cbbCampos.Text = "Data da Operação" Or _
       cbbCampos.Text = "Hora da Operação" Then
       dtpIni.Visible = True
       dtpFin.Visible = True
       lblA.Visible = True
       txtConsulta.Visible = False
       dtpIni.SetFocus
       If cbbCampos.Text = "Hora da Operação" Then
          dtpIni.Format = dtpTime
          dtpFin.Format = dtpTime
       Else
          dtpIni.Format = dtpShortDate
          dtpFin.Format = dtpShortDate
       End If
       Exit Sub
    Else
       dtpIni.Visible = False
       dtpFin.Visible = False
       lblA.Visible = False
       txtConsulta.Visible = True
       txtConsulta.SetFocus
       Exit Sub
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

Private Sub dtcFinalizadora_GotFocus()
    If txtFinalizadora.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora.Text)
    End If
End Sub

Private Sub dtcFinalizadora_LostFocus()
    txtFinalizadora.Text = dtcFinalizadora.BoundText
End Sub

Private Sub dtcOperadores_Ecf_GotFocus()
    If txtOperadores_Ecf.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcOperadores_Ecf.Text)
    End If
End Sub

Private Sub dtcOperadores_Ecf_LostFocus()
    txtOperadores_Ecf.Text = dtcOperadores_Ecf.BoundText
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
    log.Programa = "Cadastro de Operação Caixa"
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
    
    log.Descricao = "Inicializando cadastro de Operação Caixa"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
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
    log.Descricao = "Finalizando cadastro de Operação Caixa"
        
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    strCombo = Empty
    
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
           
       frmAguarde.Show
       DoEvents
       
       strID_Operacao_Caixa = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 1))
       cbbPdv.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 2))
       txtFinalizadora.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 3))
       txtOperadores_Ecf.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 5))
       dtpData.Value = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 7))
       dtpHora.Value = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 8))
       txtValor.Text = Format(hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 9)), "#,###0.00")
       If hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 10)) = "Sim" Then
          chkTipo.Value = 1
       Else
          chkTipo.Value = 0
       End If
       cbbStatus.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 11))
       txtObservacao.Text = hfgOperador.TextArray((hfgOperador.Row * hfgOperador.Cols + hfgOperador.Col + 12))
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstOperador.TabEnabled(0) = True
       sstOperador.Tab = 0
       dtpData.SetFocus
   End If
   Unload frmAguarde
End Sub

Private Sub hfgOperador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgOperador_Click
    End If
End Sub

Private Sub sstOperador_Click(PreviousTab As Integer)
    If sstOperador.Tab = 0 Then
       dtpData.SetFocus
    ElseIf sstOperador.Tab = 1 Then
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
    Dim strID_Finalizadora As String
    Dim strStatus As String
    Dim strTipo As String
            
    If txtOperadores_Ecf.Text = Empty Or cbbPdv.Text = Empty Or txtFinalizadora.Text = Empty Then
       MsgBox "Os campos Operadores ECF, PDV e Finalizadora não podem ser nulos. Verifique!", vbInformation, "Only Tech"
       dtpData.SetFocus
       Exit Function
    End If
       
    Call Objetos.Maiusculo_TXT(Me)
    
    If cbbStatus.Text = "Aberto" Then
       strStatus = 0
    Else
       strStatus = 1
    End If
    
    If chkTipo.Value = 1 Then
       strTipo = 1
    Else
       strTipo = 0
    End If
    
    strID_Finalizadora = Funcoes_Gerais.Localiza_ID("PKId_TBFinalizadora", "IXCodigo_TBFinalizadora", txtFinalizadora.Text, "TBFinalizadora", "Otica", Me, "BDRetaguarda")
         
    strCampo = "FKCodigo_TBPdv," & _
               "FKId_TBFinalizadora," & _
               "FKCodigo_TBOperadores_ecf," & _
               "DFData_TBOperacao_caixa," & _
               "DFHora_TBOperacao_caixa," & _
               "DFValor_TBOperacao_caixa," & _
               "DFTipo_operacao_TBOperacao_caixa," & _
               "DFStatus_aberto_fechado_TBOperacao_caixa," & _
               "DFObservacao_TBOperacao_caixa"
    
    strValores = "" & cbbPdv.Text & "," & _
                 "" & strID_Finalizadora & "," & _
                 "" & txtOperadores_Ecf.Text & "," & _
                 "'" & Format(dtpData.Value, "YYYYMMDD") & "'," & _
                 "'" & Format(dtpHora.Value, "hh:mm:ss") & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtValor.Text) & "," & _
                 "" & strTipo & "," & _
                 "" & strStatus & "," & _
                 "'" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "'"
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET FKCodigo_TBPdv = " & cbbPdv.Text & "," & _
                "    FKId_TBFinalizadora =" & strID_Finalizadora & "," & _
                "    FKCodigo_TBOperadores_ecf = " & txtOperadores_Ecf.Text & "," & _
                "    DFData_TBOperacao_caixa = '" & Format(dtpData.Value, "YYYYMMDD") & "'," & _
                "    DFHora_TBOperacao_caixa = '" & Format(dtpHora.Value, "hh:mm:ss") & "'," & _
                "    DFValor_TBOperacao_caixa = " & Funcoes_Gerais.Grava_Moeda(txtValor.Text) & "," & _
                "    DFTipo_operacao_TBOperacao_caixa = " & strTipo & "," & _
                "    DFStatus_aberto_fechado_TBOperacao_caixa = " & strStatus & "," & _
                "    DFObservacao_TBOperacao_caixa = '" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "'"
       Call funcoes_banco.Alterar("TBOperacao_caixa", strSet, "PKId_TBOperacao_caixa", strID_Operacao_Caixa, "Otica", Me, "BDRetaguarda")
       log.Descricao = "Alterando o registro: " + strID_Finalizadora
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBOperacao_caixa", strCampo, strValores, "OTICA", Me, "BDRetaguarda")
       log.Descricao = "Gravando o registro: " + strID_Finalizadora
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
       hfgOperador.Visible = False
    End If
    
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
    Call funcoes_banco.Excluir("TBOperacao_caixa", "PKId_TBOperacao_caixa", strID_Operacao_Caixa, "OTICA", Me, "BDRetaguarda")
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + strID_Operacao_Caixa
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
    
    chkTipo.Value = 0
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

    dtpData.Value = Date
    dtpHora.Value = Format(Now, "hh:mm:ss")
    dtpData.SetFocus
    booAlterar = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function
Private Function Reposicao()
    On Error GoTo Erro
          
    strTamanho = "0,900,1180,2000,1000,2000,1100,1100,1000,1100,1000,2000"
    strNomes = "ID,PDV,Finalizadora,Descrição,Operador,Nome,Data,Hora,Valor,Tipo Op.,Status,Observação"
    
    Movimentacoes.Monta_HFlex_Grid hfgOperador, strTamanho, strNomes, 12, "OTICA", Me
    
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

Private Function Consulta()
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
     
    Dim strData_Inicial As String
    Dim strdata_Final As String
    Dim strStatus As String
    
    If cbbCampos.Text = "Data da Operação" Then
       strData_Inicial = Format(dtpIni.Value, "YYYYMMDD")
       strdata_Final = Format(dtpFin.Value, "YYYYMMDD")
    ElseIf cbbCampos.Text = "Hora da Operação" Then
       strData_Inicial = Format(dtpIni.Value, "hh:mm:ss")
       strdata_Final = Format(dtpFin.Value, "hh:mm:ss")
    End If
    
    If cbbCampos.Text = "Status da Operação" Then
       If txtConsulta.Text = "ABERTO" Then
          strStatus = 1
       Else
          strStatus = 0
       End If
    End If
    
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
           
    txtConsulta.Text = Funcoes_Gerais.Grava_String(txtConsulta.Text)
    
    strSql = "SELECT TBOperacao_caixa.PKId_TBOperacao_caixa," & _
             "TBOperacao_caixa.FKCodigo_TBPdv," & _
             "TBFinalizadora.IXCodigo_TBFinalizadora," & _
             "TBFinalizadora.DFDescricao_TBFinalizadora," & _
             "TBOperacao_caixa.FKCodigo_TBOperadores_ecf," & _
             "TBOperadores_ecf.DFNome_TBOperadores_ecf," & _
             "TBOperacao_caixa.DFData_TBOperacao_caixa," & _
             "TBOperacao_caixa.DFHora_TBOperacao_caixa," & _
             "TBOperacao_caixa.DFValor_TBOperacao_caixa," & _
             "TBOperacao_caixa.DFTipo_operacao_TBOperacao_caixa," & _
             "TBOperacao_caixa.DFStatus_aberto_fechado_TBOperacao_caixa," & _
             "TBOperacao_caixa.DFObservacao_TBOperacao_caixa " & _
             "FROM TBOperacao_caixa " & _
             "INNER JOIN TBFinalizadora ON TBOperacao_caixa.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
             "INNER JOIN TBOperadores_ecf ON TBOperacao_caixa.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf"
           
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Data da Operação" Then
          strSql = strSql & " WHERE TBOperacao_caixa.DFData_TBOperacao_caixa >= '" & strData_Inicial & "' AND" & _
                            " TBOperacao_caixa.DFData_TBOperacao_caixa <= '" & strdata_Final & "'"
       ElseIf cbbCampos.Text = "Hora da Operação" Then
          strSql = strSql & " WHERE TBOperacao_caixa.DFHora_TBOperacao_caixa >= '" & strData_Inicial & "' AND" & _
                            " TBOperacao_caixa.DFHora_TBOperacao_caixa <= '" & strdata_Final & "'"
       ElseIf cbbCampos.Text = "Valor da Operação" Then
          strSql = strSql & " WHERE convert(money,DFValor_TBOperacao_caixa) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Status da Operação" Then
          strSql = strSql & " WHERE convert(nvarchar,DFStatus_aberto_fechado_TBOperacao_caixa) = " & strStatus & ""
       ElseIf cbbCampos.Text = "PDV" Then
          strSql = strSql & " WHERE convert(nvarchar,FKCodigo_TBPdv) =  " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Código da Finalizadora" Then
          strSql = strSql & " WHERE convert(nvarchar,FKCodigo_TBOperadores_ecf) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Descrição da Finalizadora" Then
          strSql = strSql & " WHERE convert(nvarchar,DFDescricao_TBOperadores_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Código do Operador" Then
          strSql = strSql & " WHERE convert(nvarchar,FKCodigo_TBOperadores_ecf) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Nome do Operador" Then
          strSql = strSql & " WHERE convert(nvarchar,DFNome_TBOperadores_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Observação" Then
          strSql = strSql & " WHERE convert(nvarchar,DFObservacao_TBOperacao_caixa) = '" & txtConsulta.Text & "'"
       End If
    End If
    
    frmAguarde.Show
    DoEvents
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgOperador, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    If hfgOperador.Rows > 1 Then
       For I = 1 To hfgOperador.Rows - 1
           hfgOperador.Row = I
           hfgOperador.Col = 11
           If hfgOperador.Text = "Não" Then
              hfgOperador.Text = "Aberto"
           Else
              hfgOperador.Text = "Fechado"
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
    cbbCampos.AddItem ("Data da Operação")
    cbbCampos.AddItem ("Hora da Operação")
    cbbCampos.AddItem ("Valor da Operação")
    cbbCampos.AddItem ("Status da Operação")
    cbbCampos.AddItem ("PDV")
    cbbCampos.AddItem ("Código da Finalizadora")
    cbbCampos.AddItem ("Descrição da Finalizadora")
    cbbCampos.AddItem ("Código do Operador")
    cbbCampos.AddItem ("Nome do Operador")
    cbbCampos.AddItem ("Observação")
    
    cbbStatus.Clear
    cbbStatus.AddItem ("Aberto")
    cbbStatus.AddItem ("Fechado")
End Function
Private Function Monta_DataCombo()
    Dim rstBusca_PDV As New ADODB.Recordset
    
    strSql = Empty
    strSql = "SELECT TBFinalizadora.IXCodigo_TBFinalizadora,TBFinalizadora.DFDescricao_TBFinalizadora FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSql, "BDRetaguarda", "Otica", Me

    strSql = Empty
    strSql = "SELECT TBOperadores_ecf.PKCodigo_TBOperadores_ecf,TBOperadores_ecf.DFNome_TBOperadores_ecf FROM TBOperadores_ecf"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperadores_Ecf, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = Empty
    strSql = "SELECT PKCofigo_TBPdv FROM TBPdv"
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_PDV, "Otica", Me)
    
    cbbPdv.Clear
    Do While rstBusca_PDV.EOF = False
       cbbPdv.AddItem (rstBusca_PDV.Fields("PKCofigo_TBPdv"))
       rstBusca_PDV.MoveNext
    Loop
    Set rstBusca_PDV = Nothing
End Function

Private Sub txtFinalizadora_Change()
    dtcFinalizadora.BoundText = txtFinalizadora.Text
End Sub

Private Sub txtFinalizadora_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFinalizadora_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtObservacao_LostFocus()
    txtObservacao.Text = UCase(txtObservacao.Text)
End Sub

Private Sub txtOperadores_Ecf_Change()
    dtcOperadores_Ecf.BoundText = txtOperadores_Ecf.Text
End Sub

Private Sub txtOperadores_Ecf_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperadores_Ecf_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_LostFocus()
    txtValor.Text = Format(txtValor.Text, "#,###0.00")
End Sub
