VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEncerrante 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encerrante"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEncerrante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8670
   Begin TabDlg.SSTab sstEncerrante 
      Height          =   5505
      Left            =   0
      TabIndex        =   21
      Top             =   330
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9710
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Geral"
      TabPicture(0)   =   "frmEncerrante.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label33"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbbOperacao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dtpData"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dtpHora"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtcOperador"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "hfgEncerrante_Bomba"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtcPdv"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCodigo_Pdv"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtOperador"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "frmPercentual"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmEncerrante.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label29"
      Tab(1).Control(1)=   "lblA"
      Tab(1).Control(2)=   "dtpFinal"
      Tab(1).Control(3)=   "dtpInicial"
      Tab(1).Control(4)=   "cbbcampos"
      Tab(1).Control(5)=   "hfgEncerrante"
      Tab(1).Control(6)=   "cmdConsulta"
      Tab(1).Control(7)=   "cmdRefresh"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtConsulta"
      Tab(1).ControlCount=   9
      Begin VB.Frame frmPercentual 
         Caption         =   "Bombas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   30
         Top             =   1920
         Width           =   8385
         Begin VB.TextBox txtBomba 
            Enabled         =   0   'False
            Height          =   360
            Left            =   5760
            MaxLength       =   40
            TabIndex        =   14
            Top             =   570
            Width           =   1095
         End
         Begin VB.TextBox txtValor 
            Enabled         =   0   'False
            Height          =   360
            Left            =   4650
            MaxLength       =   40
            TabIndex        =   13
            Top             =   570
            Width           =   1065
         End
         Begin VB.TextBox txtBico 
            Height          =   360
            Left            =   120
            TabIndex        =   11
            Top             =   570
            Width           =   885
         End
         Begin VB.TextBox txtEncerrante 
            Height          =   360
            Left            =   6900
            MaxLength       =   40
            TabIndex        =   15
            Top             =   570
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "Incluir"
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
            Left            =   5790
            TabIndex        =   16
            ToolTipText     =   "Incluir"
            Top             =   1080
            Width           =   1185
         End
         Begin VB.CommandButton cmdRemover 
            Caption         =   "Remover"
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
            Left            =   7050
            TabIndex        =   17
            ToolTipText     =   "Remover"
            Top             =   1080
            Width           =   1185
         End
         Begin MSDataListLib.DataCombo dtcBico 
            Height          =   360
            Left            =   1050
            TabIndex        =   12
            Top             =   570
            Width           =   3555
            _ExtentX        =   6271
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bomba"
            Height          =   240
            Left            =   5760
            TabIndex        =   35
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   240
            Left            =   4650
            TabIndex        =   33
            Top             =   330
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Bico"
            Height          =   240
            Left            =   120
            TabIndex        =   32
            Top             =   330
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Encerrante"
            Height          =   240
            Left            =   6900
            TabIndex        =   31
            Top             =   330
            Width           =   930
         End
      End
      Begin VB.TextBox txtOperador 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Código do Operador"
         Top             =   780
         Width           =   1065
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72840
         TabIndex        =   1
         Top             =   780
         Width           =   5505
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   -66870
         Picture         =   "frmEncerrante.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   -67260
         Picture         =   "frmEncerrante.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtCodigo_Pdv 
         Height          =   360
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Código do Ponto de Venda"
         Top             =   1440
         Width           =   1065
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgEncerrante 
         Height          =   4155
         Left            =   -74880
         TabIndex        =   3
         Top             =   1200
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   7329
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin AutoCompletar.CbCompleta cbbcampos 
         Height          =   360
         Left            =   -74880
         TabIndex        =   0
         Top             =   780
         Width           =   1995
         _ExtentX        =   3519
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
      Begin MSDataListLib.DataCombo dtcPdv 
         Height          =   360
         Left            =   1230
         TabIndex        =   9
         Top             =   1440
         Width           =   4650
         _ExtentX        =   8202
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgEncerrante_Bomba 
         Height          =   1755
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   3096
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo dtcOperador 
         Height          =   360
         Left            =   1230
         TabIndex        =   5
         Top             =   780
         Width           =   4650
         _ExtentX        =   8202
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
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   360
         Left            =   7380
         TabIndex        =   7
         Top             =   780
         Width           =   1125
         _ExtentX        =   1984
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
         Format          =   20709378
         CurrentDate     =   38412
      End
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   360
         Left            =   5940
         TabIndex        =   6
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   37858
      End
      Begin AutoCompletar.CbCompleta cbbOperacao 
         Height          =   360
         Left            =   5940
         TabIndex        =   10
         Top             =   1440
         Width           =   2565
         _ExtentX        =   4524
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
         Left            =   -72840
         TabIndex        =   19
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   37858
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   360
         Left            =   -70620
         TabIndex        =   20
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   37858
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         Height          =   240
         Left            =   -71130
         TabIndex        =   34
         Top             =   930
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Operação"
         Height          =   240
         Left            =   5940
         TabIndex        =   29
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   240
         Left            =   5940
         TabIndex        =   28
         Top             =   540
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora"
         Height          =   240
         Left            =   7380
         TabIndex        =   27
         Top             =   540
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Operador"
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   24
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "PDV"
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   345
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8730
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEncerrante.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEncerrante.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEncerrante.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEncerrante.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEncerrante.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEncerrante.frx":5578
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEncerrante.frx":5892
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEncerrante.frx":66E4
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
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            Object.ToolTipText     =   "Consulta Detalhada ICMS"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Integração"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEncerrante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador                                                            '
' Objetivo...............: Cadastro de Encerrante                                         '
' Data de Criação........: 18/08/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........: Jones                                                           '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strTamanho As String
Dim strNomes As String
Dim strTamanho_encerrante As String
Dim strNomes_encerrante As String
Public strCombo As String
Public strConsulta As String
Dim strCasas_Decimais As Integer
Dim strClique_encerrante As String
Dim strID_Encerrante As String
Public strSql As String
Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim I As Integer
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean

Function Imprimir()
    On Error GoTo erro
    'Tratamento de Erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Logicx"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    'Call frmConsole_Relatorio_Encerrante.Show
    
    Unload frmAguarde
    
    Exit Function
erro:
        Call erro.erro(Me, "OTICA", "Gravar")
        Exit Function
End Function

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty
    
    If cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       lblA.Visible = False
       cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Data" Then
       dtpInicial.Visible = True
       dtpFinal.Visible = True
       lblA.Visible = True
       txtConsulta.Visible = False
    Else
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       txtConsulta.Visible = True
       lblA.Visible = False
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cbbOperacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdIncluir_Click()
    Dim strIndice As String
    Dim intContador As Integer

    If txtBico.Text = Empty Then
       MsgBox "Bomba inválida. Verifique!", vbCritical, "Only Tech"
       txtBico.SetFocus
       Exit Sub
    ElseIf txtEncerrante.Text = Empty Then
       MsgBox "Defina um Encerrante antes de acrescentar uma Bomba.", vbInformation, "Only Tech"
       txtEncerrante.SetFocus
       Exit Sub
    ElseIf txtOperador.Text = Empty Then
       MsgBox "Defina um Operador antes de acrescentar uma Bomba.", vbInformation, "Only Tech"
       txtOperador.SetFocus
       Exit Sub
    ElseIf txtCodigo_Pdv.Text = Empty Then
       MsgBox "Defina um PDV antes de acrescentar uma Bomba.", vbInformation, "Only Tech"
       txtCodigo_Pdv.SetFocus
       Exit Sub
    End If

    'Verificar se o item está no grid de itens do pedido
    intContador = 1

    Do While intContador <= hfgEncerrante_Bomba.Rows - 1
        hfgEncerrante_Bomba.Row = intContador
        hfgEncerrante_Bomba.Col = 1
        If cmdIncluir.Caption = "Alterar" Then
           If hfgEncerrante_Bomba.Text = txtBico.Text And hfgEncerrante_Bomba.Row <> strClique_encerrante Then
              MsgBox "O Bico alterado já pertence a outro item neste cadastro. Verifique.", vbInformation, "Only Tech"
              txtBico.SetFocus
              Exit Sub
           End If
        Else
           If hfgEncerrante_Bomba.Text = txtBico.Text Then
              MsgBox "Bico já incluído nesse cadastro. Verifique.", vbInformation, "Only Tech"
              'Limpando os campos dos Itens
              txtBico.SetFocus
              Exit Sub
           End If
        End If
        intContador = intContador + 1
    Loop
    
    hfgEncerrante_Bomba.Row = hfgEncerrante_Bomba.TopRow
    If cmdIncluir.Caption = "Incluir" Then
       If hfgEncerrante_Bomba.Text <> Empty Then
          strIndice = intContador
          hfgEncerrante_Bomba.Rows = hfgEncerrante_Bomba.Rows + 1
       Else
          strIndice = intContador - 1
       End If
    Else
       strIndice = strClique_encerrante
    End If
    
    hfgEncerrante_Bomba.Row = strIndice
    
    hfgEncerrante_Bomba.Col = 0
    hfgEncerrante_Bomba.ColWidth(0) = 500
    hfgEncerrante_Bomba.Font.Name = "Tahoma"
    hfgEncerrante_Bomba.CellFontSize = 7
    hfgEncerrante_Bomba.CellFontBold = False
    hfgEncerrante_Bomba.CellBackColor = &H80FFFF
    hfgEncerrante_Bomba.Text = strIndice
    
    hfgEncerrante_Bomba.Col = 1
    hfgEncerrante_Bomba.Text = txtBico.Text
    
    hfgEncerrante_Bomba.Col = 2
    hfgEncerrante_Bomba.Text = dtcBico.Text
    
    hfgEncerrante_Bomba.Col = 3
    hfgEncerrante_Bomba.Text = txtValor.Text
    
    hfgEncerrante_Bomba.Col = 4
    hfgEncerrante_Bomba.Text = txtBomba.Text
    
    hfgEncerrante_Bomba.Col = 5
    hfgEncerrante_Bomba.Text = txtEncerrante.Text
    
    hfgEncerrante_Bomba.Col = 6
    hfgEncerrante_Bomba.Text = Funcoes_Gerais.Localiza_ID("PKId_TBBomba_bico", "IXCodigo_TBBomba_bico", txtBico.Text, "TBBomba_bico", "Otica", Me)
    
    'capturando id do produto para uso futuro
    strSql = "SELECT FKId_TBProduto FROM TBBomba_bico WHERE PKId_TBBomba_bico = " & hfgEncerrante_Bomba.Text & ""
    
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    hfgEncerrante_Bomba.Col = 7
    If rstAplicacao.RecordCount <> 0 Then
       hfgEncerrante_Bomba.Text = rstAplicacao.Fields("FKId_TBProduto")
    Else
       MsgBox "Ocorreram problemas na inserção desse Bico no cadastro. Reveja o cadastro da Bomba referente.", vbInformation, "Onlytech"
       hfgEncerrante_Bomba.RemoveItem (hfgEncerrante_Bomba.Row)
       
       Do While intContador <= hfgEncerrante_Bomba.Rows - 1
         hfgEncerrante_Bomba.Row = intContador
         hfgEncerrante_Bomba.Text = intContador
         intContador = intContador + 1
       Loop
       Set rstAplicacao = Nothing
       Exit Sub
    End If
    Set rstAplicacao = Nothing
    
    Do While intContador <= hfgEncerrante_Bomba.Rows - 1
       hfgEncerrante_Bomba.Row = intContador
       hfgEncerrante_Bomba.Text = intContador
       intContador = intContador + 1
    Loop
       
    hfgEncerrante_Bomba.Refresh
    
    txtBico.Text = Empty
    txtValor.Text = Empty
    txtEncerrante.Text = Empty
    txtBomba.Text = Empty
    
    hfgEncerrante_Bomba.Col = 0
    
    txtCodigo_Pdv.Enabled = False
    dtcPdv.Enabled = False
    txtOperador.Enabled = False
    dtcOperador.Enabled = False
    cbbOperacao.Enabled = False
    dtpHora.Enabled = False
    dtpData.Enabled = False
    
    cmdIncluir.Caption = "Incluir"
    
    txtBico.SetFocus
    
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub cmdRemover_Click()
    Dim intContador As Integer
    
    hfgEncerrante_Bomba.Col = 0
    If hfgEncerrante_Bomba.Text = Empty Then
       MsgBox "Não há Bomba selecionada para este Encerrante.", vbInformation, "Only Tech"
       txtBico.SetFocus
       Exit Sub
    End If
    
    If hfgEncerrante_Bomba.Rows <= 2 Then
       hfgEncerrante_Bomba.Clear
       Movimentacoes.Monta_HFlex_Grid hfgEncerrante_Bomba, strTamanho_encerrante, strNomes_encerrante, 7, "Otica", Me
    Else
       hfgEncerrante_Bomba.RemoveItem (hfgEncerrante_Bomba.Row)
       hfgEncerrante_Bomba.Col = 0
       intContador = 1
       Do While intContador <= hfgEncerrante_Bomba.Rows - 1
          hfgEncerrante_Bomba.Row = intContador
          hfgEncerrante_Bomba.Text = intContador
          intContador = intContador + 1
       Loop
    End If
    
    hfgEncerrante_Bomba.Refresh
    
    hfgEncerrante_Bomba.Col = 0
    hfgEncerrante_Bomba.Row = 1
    If booAlterar = False And hfgEncerrante_Bomba.Rows <= 2 And hfgEncerrante_Bomba.Text = Empty Then
       txtOperador.Enabled = True
       dtcOperador.Enabled = True
       dtpHora.Enabled = True
       dtpData.Enabled = True
       txtCodigo_Pdv.Enabled = True
       cbbOperacao.Enabled = True
       dtcPdv.Enabled = True
    End If
    
    cmdIncluir.Caption = "Incluir"
    
    txtBico.Text = Empty
    txtValor.Text = Empty
    txtBomba.Text = Empty
    txtEncerrante.Text = Empty
    
    txtBico.SetFocus
    hfgEncerrante_Bomba.Col = 0
    hfgEncerrante_Bomba.Row = 0
End Sub

Private Sub dtcBico_LostFocus()
    txtBico.Text = dtcBico.BoundText
    If IsNumeric(txtBico.Text) = False Or dtcBico.Text = Empty Then txtBico.Text = Empty And txtBomba.Text = Empty
    If txtBico.Text <> Empty Then
       Call txtBico_LostFocus
    End If
End Sub

Private Sub dtcOperador_GotFocus()
   If txtOperador.Text = Empty Then
      Call Movimentacoes.Verifica_DataCombo(dtcOperador.Text)
   End If
End Sub

Private Sub dtcOperador_LostFocus()
    txtOperador.Text = dtcOperador.BoundText
    If IsNumeric(txtOperador.Text) = False Or dtcOperador.Text = Empty Then txtOperador.Text = Empty: Exit Sub
End Sub

Private Sub dtcPdv_GotFocus()
   If txtCodigo_Pdv.Text = Empty Then
      Call Movimentacoes.Verifica_DataCombo(dtcPdv.Text)
   End If
End Sub

Private Sub dtcPdv_LostFocus()
    txtCodigo_Pdv.Text = dtcPdv.BoundText
    If IsNumeric(txtCodigo_Pdv.Text) = False Or dtcPdv.Text = Empty Then txtCodigo_Pdv.Text = Empty: Exit Sub
End Sub

Private Sub dtpData_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos no dataPicker pelo ENTER
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
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
    On Error GoTo erro
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.ocxUsuario.Nome
    log.Programa = "Cadastro de Encerrante"
    log.Estacao = MDIPrincipal.ocxUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstEncerrante, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.ocxUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando Cadastro de Estados para ICMS"
    'Gravando o log
    log.Gravar_log "Otica", Me
        
    sstEncerrante.TabEnabled(0) = False
    sstEncerrante.Tab = 1
    strClique_encerrante = 0
    Call Reposicao
    
    dtpInicial.Value = Date
    dtpFinal.Value = Date
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.ocxUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.ocxUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.ocxUsuario.Empresa, Me)
    
    'ABASTECENDO VARIÁVEL PARA CASAS DECIMAIS
    strSql = "SELECT DFNumero_decimais_TBParametros_ecf FROM TBParametros_ecf " & _
             "WHERE FKCodigo_TBEmpresa = " & MDIPrincipal.ocxUsuario.Empresa & ""
                
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "OTICA", Me
    
    If Not IsNull(rstAplicacao.Fields("DFNumero_decimais_TBParametros_ecf")) And rstAplicacao.RecordCount <> 0 Then
        strCasas_Decimais = rstAplicacao.Fields("DFNumero_decimais_TBParametros_ecf")
    End If
    Set rstAplicacao = Nothing
    
    Exit Sub
erro:
    Call erro.erro(Me, "Otica", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro
    
    strEvento_log = "Unload"
    
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    
    strCombo = Empty
    strConsulta = Empty
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
        
    Exit Sub
erro:
    Call erro.erro(Me, "Otica", "Unload")
    Exit Sub
End Sub

Private Sub hfgEncerrante_Click()
    If hfgEncerrante.Col = 0 And hfgEncerrante.Text <> Empty Then
    
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
           tlbBotoes.Buttons.Item(11).Enabled = True
        End If
        
        frmAguarde.Show
        DoEvents
     
       txtOperador.Text = hfgEncerrante.TextArray((hfgEncerrante.Row * hfgEncerrante.Cols + hfgEncerrante.Col + 1))
       dtpData.Value = hfgEncerrante.TextArray((hfgEncerrante.Row * hfgEncerrante.Cols + hfgEncerrante.Col + 3))
       dtpHora.Value = hfgEncerrante.TextArray((hfgEncerrante.Row * hfgEncerrante.Cols + hfgEncerrante.Col + 4))
       txtCodigo_Pdv.Text = hfgEncerrante.TextArray((hfgEncerrante.Row * hfgEncerrante.Cols + hfgEncerrante.Col + 5))
       cbbOperacao.Text = hfgEncerrante.TextArray((hfgEncerrante.Row * hfgEncerrante.Cols + hfgEncerrante.Col + 7))
       
       strID_Encerrante = hfgEncerrante.TextArray((hfgEncerrante.Row * hfgEncerrante.Cols + hfgEncerrante.Col + 8))
              
'''''''Abastecendo grid
       strSql = "SELECT IXCodigo_TBBomba_bico," & _
                "DFDescricao_TBProduto," & _
                "DFTipo_preco_TBBomba_bico,IXCodigo_Bomba," & _
                "DFEncerrante_TBEncerrante_Bomba," & _
                "FKId_TBBomba_bico,FKId_TBProduto " & _
                "FROM TBEncerrante_Bomba " & _
                "INNER JOIN TBBomba_bico " & _
                "ON TBBomba_bico.PKId_TBBomba_bico = TBEncerrante_Bomba.FKId_TBBomba_bico " & _
                "INNER JOIN TBBomba " & _
                "ON TBBomba_bico.FKId_TBBomba = TBBomba.PKId_TBBomba " & _
                "INNER JOIN TBProduto " & _
                "ON TBProduto.PKId_TBProduto = TBBomba_bico.FKId_TBProduto " & _
                "WHERE FKId_TBEncerrante = " & strID_Encerrante & " "
       
       Call Movimentacoes.Movimenta_HFlex_Grid(strSql, hfgEncerrante_Bomba, strTamanho_encerrante, strNomes_encerrante, "BDRetaguarda", "Otica", Me)
        
       hfgEncerrante_Bomba.Col = 1
       hfgEncerrante_Bomba.Row = 1
       If hfgEncerrante_Bomba.Text = Empty Then
          hfgEncerrante_Bomba.Rows = 2
          Movimentacoes.Monta_HFlex_Grid hfgEncerrante_Bomba, strTamanho_encerrante, strNomes_encerrante, 7, "Otica", Me
       Else
           intContador = 1
           hfgEncerrante_Bomba.Col = 3
           
           Dim rstItem_tabela As New ADODB.Recordset
           
           Do While intContador <= hfgEncerrante_Bomba.Rows - 1
              hfgEncerrante_Bomba.Row = intContador
              
              If hfgEncerrante_Bomba.Text = 1 Then
                 hfgEncerrante_Bomba.Col = 6
                  
                 strSql = "SELECT DFPreco_avista_TBItens_tabela_preco FROM TBItens_tabela_preco " & _
                          "WHERE FKId_TBProduto = " & hfgEncerrante_Bomba.Text & ""
                 Select_geral strSql, "BDRetaguarda", rstItem_tabela, "Otica", Me
                 
                 hfgEncerrante_Bomba.Col = 3
                 hfgEncerrante_Bomba.Text = rstItem_tabela("DFPreco_avista_TBItens_tabela_preco")
              ElseIf hfgEncerrante_Bomba.Text = 2 Then
                 hfgEncerrante_Bomba.Col = 6
                  
                 strSql = "SELECT DFPreco_promocao_TBItens_tabela_preco FROM TBItens_tabela_preco " & _
                          "WHERE FKId_TBProduto = " & hfgEncerrante_Bomba.Text & ""
                 Select_geral strSql, "BDRetaguarda", rstItem_tabela, "Otica", Me
    
                 hfgEncerrante_Bomba.Col = 3
                 hfgEncerrante_Bomba.Text = rstItem_tabela("DFPreco_promocao_TBItens_tabela_preco")
                 
              ElseIf hfgEncerrante_Bomba.Text = 3 Then
                 hfgEncerrante_Bomba.Col = 6
                  
                 strSql = "SELECT DFPreco_revenda_TBItens_tabela_preco FROM TBItens_tabela_preco " & _
                          "WHERE FKId_TBProduto = " & hfgEncerrante_Bomba.Text & ""
                 Select_geral strSql, "BDRetaguarda", rstItem_tabela, "Otica", Me
                 
                 hfgEncerrante_Bomba.Col = 3
                 hfgEncerrante_Bomba.Text = rstItem_tabela("DFPreco_revenda_TBItens_tabela_preco")
                 
              ElseIf hfgEncerrante_Bomba.Text = 4 Then
                 hfgEncerrante_Bomba.Col = 6
                  
                 strSql = "SELECT DFPreco_especial_TBItens_tabela_preco FROM TBItens_tabela_preco " & _
                          "WHERE FKId_TBProduto = " & hfgEncerrante_Bomba.Text & ""
                 Select_geral strSql, "BDRetaguarda", rstItem_tabela, "Otica", Me
                 
                 hfgEncerrante_Bomba.Col = 3
                 hfgEncerrante_Bomba.Text = rstItem_tabela("DFPreco_especial_TBItens_tabela_preco")
                 
              ElseIf hfgEncerrante_Bomba.Text = 5 Then
                 hfgEncerrante_Bomba.Col = 6
                  
                 strSql = "SELECT DFPreco_varejo_TBItens_tabela_preco FROM TBItens_tabela_preco " & _
                          "WHERE FKId_TBProduto = " & hfgEncerrante_Bomba.Text & ""
                 Select_geral strSql, "BDRetaguarda", rstItem_tabela, "Otica", Me
                 
                 hfgEncerrante_Bomba.Col = 3
                 hfgEncerrante_Bomba.Text = rstItem_tabela("DFPreco_varejo_TBItens_tabela_preco")
                 
              End If
              Set rstItem_tabela = Nothing
              intContador = intContador + 1
            Loop
       End If
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstEncerrante.TabEnabled(0) = True
       sstEncerrante.Tab = 0
       txtOperador.Enabled = False
       cbbOperacao.Enabled = False
       dtcOperador.Enabled = False
       txtCodigo_Pdv.Enabled = False
       dtcPdv.Enabled = False
       dtpHora.Enabled = False
       dtpData.Enabled = False
       txtBico.SetFocus
    End If
    Unload frmAguarde
End Sub

Private Sub hfgEncerrante_DblClick()
    hfgEncerrante.Sort = 1
End Sub

Private Sub hfgEncerrante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgEncerrante_Click
    End If
End Sub

Private Sub hfgEncerrante_Bomba_Click()
    If hfgEncerrante_Bomba.Col = 0 And hfgEncerrante_Bomba.Text <> Empty And hfgEncerrante_Bomba.Row <> strClique_encerrante Then
        txtBico.Text = Empty
        txtEncerrante.Text = Empty
        txtValor.Text = Empty
        txtBomba.Text = Empty
        cmdIncluir.Caption = "Incluir"
    End If
End Sub

Private Sub hfgEncerrante_Bomba_DblClick()
   If hfgEncerrante_Bomba.Col = 0 And hfgEncerrante_Bomba.Text <> Empty Then
       strClique_encerrante = hfgEncerrante_Bomba.Row
       cmdIncluir.Caption = "Alterar"
       txtBico.Text = hfgEncerrante_Bomba.TextArray((hfgEncerrante_Bomba.Row * hfgEncerrante_Bomba.Cols + hfgEncerrante_Bomba.Col + 1))
       txtValor.Text = hfgEncerrante_Bomba.TextArray((hfgEncerrante_Bomba.Row * hfgEncerrante_Bomba.Cols + hfgEncerrante_Bomba.Col + 3))
       txtBomba.Text = hfgEncerrante_Bomba.TextArray((hfgEncerrante_Bomba.Row * hfgEncerrante_Bomba.Cols + hfgEncerrante_Bomba.Col + 4))
       txtEncerrante.Text = hfgEncerrante_Bomba.TextArray((hfgEncerrante_Bomba.Row * hfgEncerrante_Bomba.Cols + hfgEncerrante_Bomba.Col + 5))
    End If
    hfgEncerrante_Bomba.SetFocus
End Sub

Private Sub sstEncerrante_Click(PreviousTab As Integer)
    If sstEncerrante.Tab = 0 Then
       If booAlterar = False And txtOperador.Enabled = True Then
          txtOperador.SetFocus
       Else
          txtBico.SetFocus
       End If
    ElseIf sstEncerrante.Tab = 1 Then
        If frmIntegracao.Visible = True Then
           Unload frmIntegracao
        End If
        If strCombo <> Empty And strCombo <> "Todos" Then
           cbbCampos.Text = strCombo
           txtConsulta.SetFocus
        ElseIf strCombo = "Todos" Then
           hfgEncerrante.Row = 1
           hfgEncerrante.Col = 0
           hfgEncerrante.SetFocus
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
           Case 9: frmConsulta_Detalhada_ICMS.Show
           Case 11: Call Integracao
    End Select
End Sub

Function Gravar()
    On Error GoTo erro
            
    'Verifica se os campos necessarios para gravar não estão nulos
    If txtOperador.Text = Empty Then
       MsgBox "Não há Operador definido. Verifique.", vbInformation, "Only Tech"
       txtCodigo_Pdv.SetFocus
       Exit Function
    ElseIf dtcPdv.Text = Empty Then
       MsgBox "Não há PDV definido. Verifique.", vbInformation, "Only Tech"
       dtcPdv.SetFocus
       Exit Function
    End If
    
    hfgEncerrante_Bomba.Col = 0
    hfgEncerrante_Bomba.Row = 1
    If hfgEncerrante_Bomba.Text = Empty Then
       MsgBox "Não há Bomba a ser cadastrada. Verifique.", vbInformation, "Only Tech"
       Exit Function
    End If

    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim strOperacao As String
    Dim strId_novo As String
    Dim strId_Bomba As String
    Dim intContador As Integer
    Dim strData As String
    Dim strHora As String
    
    strData = Format(dtpData.Value, "YYYYMMDD")
    strHora = Format(dtpHora.Value, "hh:mm:ss")
    
    If cbbOperacao.Text = "Abertura" Then
       strOperacao = 0
    Else
       strOperacao = 1
    End If
    
    strCampo = "FKCodigo_TBPdv,FKCodigo_TBOperadores_ecf,DFData_TBEncerrante,DFHora_TBEncerrante," & _
               "DFAbertura_fechamento_TBEncerrante,DFData_alteracao_TBEncerrante," & _
               "DFIntegrado_filiais_TBEncerrante"
               
    If booIntegra_Portal = True Then
       strCampo = strCampo & ",DFIntegrado_portal_TBEncerrante"
    End If
    
    strValores = "" & txtCodigo_Pdv.Text & "," & txtOperador.Text & ",'" & strData & "','" & strHora & "'," & _
                 "" & strOperacao & ",'" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
    End If
    
    If booAlterar = True Then
                                            
       On Error GoTo Erro_alteracao

       'abrindo conexao
       Conexao.Abrir_conexao "Otica"
       Conexao.CNconexao.BeginTrans
       
       log.Evento = "Alterar"
       strSql = "UPDATE TBEncerrante " & _
                "SET FKCodigo_TBPdv = " & txtCodigo_Pdv.Text & ", " & _
                "    FKCodigo_TBOperadores_ecf = " & txtOperador.Text & "," & _
                "    DFData_TBEncerrante  = '" & strData & "'," & _
                "    DFHora_TBEncerrante = '" & strHora & "'," & _
                "    DFAbertura_fechamento_TBEncerrante = '" & strOperacao & "', " & _
                "    DFData_alteracao_TBEncerrante = '" & Format(Date, "YYYYMMDD") & "'," & _
                "    DFIntegrado_filiais_TBEncerrante = 0"
                
       If booIntegra_Portal = True Then
          strSql = strSql & ",DFIntegrado_portal_TBEncerrante = 0"
       End If
                
       strSql = strSql & "    WHERE PKId_TBEncerrante = " & strID_Encerrante & " "
                
       Conexao.CNconexao.Execute strSql
       
       'Deletando registros antes da nova gravacao'
       strSql = "DELETE FROM TBEncerrante_Bomba WHERE FKId_TBEncerrante = " & strID_Encerrante & ""
       
       Conexao.CNconexao.Execute strSql
           
        intContador = 1
        Do While intContador <= hfgEncerrante_Bomba.Rows - 1
        
            hfgEncerrante_Bomba.Row = intContador

            hfgEncerrante_Bomba.Col = 6
            strId_Bomba = hfgEncerrante_Bomba.Text
            
            hfgEncerrante_Bomba.Col = 5

            strSql = Empty
            strSql = "INSERT INTO TBEncerrante_Bomba (FKId_TBBomba_bico,FKId_TBEncerrante," & _
                     "DFEncerrante_TBEncerrante_Bomba,DFData_alteracao_TBEncerrante_Bomba," & _
                     "DFIntegrado_filiais_TBEncerrante_Bomba"
                     
            If booIntegra_Portal = True Then
               strSql = strSql & ",DFIntegrado_portal_TBEncerrante_Bomba) "
            Else
               strSql = strSql & ") "
            End If
                    
            strSql = strSql & "SELECT '" & strId_Bomba & "','" & strID_Encerrante & "'," & _
                              "" & Funcoes_Gerais.Grava_Moeda(hfgEncerrante_Bomba.Text) & "," & _
                              "'" & Format(Date, "YYYYMMDD") & "',0"
                              
            If booIntegra_Portal = True Then
               strSql = strSql & ",0 "
            End If
            
            Conexao.CNconexao.Execute strSql

            intContador = intContador + 1
         Loop
       
       'fechando conexao
       Conexao.CNconexao.CommitTrans
       Conexao.Fechar_conexao

       log.Descricao = "Alterando o registro de ID: " + strID_Encerrante
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "Otica", Me
       
    Else

        'VERIFICANDO SE HÁ OUTRA ABERTURA PARA ESSA DATA E HORA
        strSql = "SELECT PKId_TBEncerrante " & _
                 "FROM TBEncerrante WHERE FKCodigo_TBPdv = " & txtCodigo_Pdv.Text & " AND DFData_TBEncerrante = '" & strData & "' AND " & _
                 "DFHora_TBEncerrante = '" & strHora & "'"
    
        Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstAplicacao, "Otica", Me)
        
        If rstAplicacao.RecordCount <> 0 Then
           MsgBox "Esta Data, Hora e PDV já pertencem a outro Encerrante. Verifique.", vbInformation, "Only Tech"
           Set rstAplicacao = Nothing
           Exit Function
        End If
         
       Set rstAplicacao = Nothing
    
       On Error GoTo Erro_inclusao
        
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBEncerrante", strCampo, strValores, "Otica", Me, "BDRetaguarda")
       
       'localizando a ID gravada
       strSql = "SELECT MAX(PKId_TBEncerrante) as IdEncerrante FROM TBEncerrante"
    
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstAplicacao, "OTICA", Me)
       strId_novo = rstAplicacao.Fields("IdEncerrante")
     
       Set rstAplicacao = Nothing
       
       'abrindo conexao
       Conexao.Abrir_conexao "Otica"
       Conexao.CNconexao.BeginTrans
       
        intContador = 1
        Do While intContador <= hfgEncerrante_Bomba.Rows - 1
        
            hfgEncerrante_Bomba.Row = intContador

            hfgEncerrante_Bomba.Col = 6
            strId_Bomba = hfgEncerrante_Bomba.Text
            
            hfgEncerrante_Bomba.Col = 5

            strSql = Empty
            strSql = "INSERT INTO TBEncerrante_Bomba (FKId_TBBomba_bico,FKId_TBEncerrante," & _
                     "DFEncerrante_TBEncerrante_Bomba,DFData_alteracao_TBEncerrante_Bomba," & _
                     "DFIntegrado_filiais_TBEncerrante_Bomba"
                     
            If booIntegra_Portal = True Then
               strSql = strSql & ",DFIntegrado_portal_TBEncerrante_Bomba) "
            End If
            
            strSql = strSql & "SELECT '" & strId_Bomba & "','" & strId_novo & "'," & _
                              "" & Funcoes_Gerais.Grava_Moeda(hfgEncerrante_Bomba.Text) & "," & _
                              "'" & Format(Date, "YYYYMMDD") & "',0"
                              
            If booIntegra_Portal = True Then
               strSql = strSql & ",0"
            End If
                              
            Conexao.CNconexao.Execute strSql

            intContador = intContador + 1
         Loop
         
        'fechando conexao
        Conexao.CNconexao.CommitTrans
        Conexao.Fechar_conexao
    
       log.Descricao = "Gravando o registro de ID: " + strId_novo
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    End If
    
       
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
           
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    tlbBotoes.Buttons.Item(11).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       Me.hfgEncerrante.Visible = False
    End If
    
    sstEncerrante.TabEnabled(0) = False
    sstEncerrante.Tab = 1

    Exit Function
    
Erro_alteracao:
    'cancelando as alteracoes
    Conexao.CNconexao.RollbackTrans
    'fechando conexao
    Conexao.Fechar_conexao
    
    Call erro.erro(Me, "Otica", "Gravar")
    Exit Function
Erro_inclusao:

    'cancelando as alteracoes
    Conexao.CNconexao.RollbackTrans
    'fechando conexao
    Conexao.Fechar_conexao
    
    Call funcoes_banco.Excluir("TBEncerrante", "PKId_TBEncerrante", strId_novo, "Otica", Me, "BDRetaguarda")
    
    Call erro.erro(Me, "Otica", "Gravar")
    Exit Function
erro:

    Call erro.erro(Me, "Otica", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro de ID: " + strID_Encerrante
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "Otica", Me
           
    'Excluindo Registro
    'Iniciando conexao
    Conexao.Initial_Catalog = "BDRetaguarda"
    Conexao.Abrir_conexao ("Otica")
    
    Conexao.CNconexao.BeginTrans
    
    strSql = "DELETE FROM TBEncerrante_Bomba WHERE FKId_TBEncerrante = " & strID_Encerrante & " "
    
    Conexao.CNconexao.Execute strSql
    
    strSql = "DELETE FROM TBEncerrante WHERE PKId_TBEncerrante = " & strID_Encerrante & " "
    
    Conexao.CNconexao.Execute strSql
    
    Conexao.CNconexao.CommitTrans
    Conexao.Fechar_conexao
    
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos

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
    tlbBotoes.Buttons.Item(11).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       Me.hfgEncerrante.Visible = False
    End If
    
    sstEncerrante.TabEnabled(0) = False
    sstEncerrante.Tab = 1
        
    Exit Function
erro:
    Conexao.CNconexao.RollbackTrans
    Conexao.Fechar_conexao
    
    Call erro.erro(Me, "Otica", "Excluir")
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo erro
    
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
                      
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
    tlbBotoes.Buttons.Item(11).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgEncerrante.Visible = False
    End If
    
    sstEncerrante.TabEnabled(0) = False
    sstEncerrante.Tab = 1
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    
    sstEncerrante.Tab = 1
    
    Exit Function
erro:
    Call erro.erro(Me, "Otica", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo erro
    
    sstEncerrante.TabEnabled(0) = True
    sstEncerrante.Tab = 0
    
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
           
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
        
    'Gravando Log
    log.Gravar_log "Otica", Me
            
    tlbBotoes.Buttons.Item(1).Enabled = False
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = booPrivilegio_Incluir
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = booPrivilegio_Incluir
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = False
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = False

    txtCodigo_Pdv.Enabled = True
    dtcPdv.Enabled = True
    txtOperador.Enabled = True
    dtcOperador.Enabled = True
    cbbOperacao.Enabled = True
    dtpHora.Enabled = True
    dtpData.Enabled = True
    dtpHora.Value = Now
    dtpData.Value = Date
    txtOperador.SetFocus
    
    hfgEncerrante_Bomba.Rows = 2
    Call Movimentacoes.Monta_HFlex_Grid(hfgEncerrante_Bomba, strTamanho_encerrante, strNomes_encerrante, 7, "Otica", Me)
    
    cbbOperacao.Text = Empty
    
    booAlterar = False
    
    Exit Function
erro:
    Call erro.erro(Me, "Otica", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    On Error GoTo erro
    
    strTamanho = "1500,3000,1500,1500,1500,1800,1900,0"
    strNomes = "Cod. Operador,Operador,Data,Hora,PDV,Endereço IP,Operação,ID"
    
    Movimentacoes.Monta_HFlex_Grid hfgEncerrante, strTamanho, strNomes, 8, "Otica", Me

    strTamanho_encerrante = "1200,3300,1600,1200,1600,0,0"
    strNomes_encerrante = "Bico,Produto,Valor,Bomba,Encerrante,ID_Bomba,ID_Produto"
    
    Movimentacoes.Monta_HFlex_Grid hfgEncerrante_Bomba, strTamanho_encerrante, strNomes_encerrante, 7, "Otica", Me
         
    strSql = "SELECT TBOperadores_ecf.PKCodigo_TBOperadores_ecf,TBOperadores_ecf.DFNome_TBOperadores_ecf FROM TBOperadores_ecf"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT TBPdv.PKCodigo_TBPdv,TBPdv.DFEndereco_ip_TBPdv FROM TBPdv"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBPdv", "DFEndereco_ip_TBPdv", dtcPdv, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_TBBomba_bico,DFDescricao_TBProduto FROM TBBomba_bico " & _
             "INNER JOIN TBProduto ON TBBomba_bico.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
             "INNER JOIN TBEmpresa ON TBEmpresa.PKCodigo_TBEmpresa = TBProduto.IXCodigo_TBEmpresa " & _
             "WHERE TBEmpresa.PKCodigo_TBEmpresa = " & MDIPrincipal.ocxUsuario.Empresa & ""
                
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBBomba_bico", "DFDescricao_TBProduto", dtcBico, strSql, "BDRetaguarda", "Otica", Me

    Call Monta_Combo
    
    strSql = Empty
    
    Exit Function
erro:
    Call erro.erro(Me, "Otica", "Reposicao")
    Resume Next
End Function

Private Sub txtBico_Change()
    dtcBico.BoundText = txtBico.Text
    If IsNumeric(txtBico.Text) = False Then
       txtBico.Text = Empty
       txtBomba.Text = Empty
       Exit Sub
    End If
End Sub

Private Sub txtBico_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtBico_LostFocus()
    If txtBico.Text <> Empty Then
    
       strSql = "SELECT TBBomba_bico.FKId_TBProduto,DFTipo_preco_TBBomba_bico,DFPreco_avista_TBItens_tabela_preco, " & _
                "DFPreco_promocao_TBItens_tabela_preco,DFPreco_revenda_TBItens_tabela_preco," & _
                "DFPreco_especial_TBItens_tabela_preco,DFPreco_varejo_TBItens_tabela_preco,IXCodigo_Bomba " & _
                "FROM TBBomba_bico " & _
                "INNER JOIN TBProduto ON TBBomba_bico.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                "INNER JOIN TBEmpresa ON TBEmpresa.PKCodigo_TBEmpresa = TBProduto.IXCodigo_TBEmpresa " & _
                "INNER JOIN TBItens_tabela_preco ON TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                "INNER JOIN TBBomba ON TBBomba_bico.FKId_TBBomba = TBBomba.PKId_TBBomba " & _
                "WHERE IXCodigo_TBBomba_bico = " & txtBico.Text & " " & _
                "AND TBEmpresa.PKCodigo_TBEmpresa = " & MDIPrincipal.ocxUsuario.Empresa & ""
                
        Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstAplicacao, "OTICA", Me)
        
        If rstAplicacao.RecordCount <> 0 Then
           If rstAplicacao.Fields("DFTipo_preco_TBBomba_bico") = 1 Then
              txtValor.Text = rstAplicacao.Fields("DFPreco_avista_TBItens_tabela_preco")
           ElseIf rstAplicacao.Fields("DFTipo_preco_TBBomba_bico") = 2 Then
              txtValor.Text = rstAplicacao.Fields("DFPreco_promocao_TBItens_tabela_preco")
           ElseIf rstAplicacao.Fields("DFTipo_preco_TBBomba_bico") = 3 Then
              txtValor.Text = rstAplicacao.Fields("DFPreco_revenda_TBItens_tabela_preco")
           ElseIf rstAplicacao.Fields("DFTipo_preco_TBBomba_bico") = 4 Then
              txtValor.Text = rstAplicacao.Fields("DFPreco_especial_TBItens_tabela_preco")
           ElseIf rstAplicacao.Fields("DFTipo_preco_TBBomba_bico") = 5 Then
              txtValor.Text = rstAplicacao.Fields("DFPreco_varejo_TBItens_tabela_preco")
           End If
           txtBomba.Text = rstAplicacao.Fields("IXCodigo_Bomba")
        Else
           txtBico.Text = Empty
           txtValor.Text = Empty
        End If
        Set rstAplicacao = Nothing
     Else
        txtValor.Text = Empty
        txtBomba.Text = Empty
     End If
End Sub

Private Sub txtCodigo_Pdv_Change()
    dtcPdv.BoundText = txtCodigo_Pdv.Text
    If IsNumeric(txtCodigo_Pdv.Text) = False Then txtCodigo_Pdv.Text = Empty: Exit Sub
End Sub

Private Sub txtCodigo_Pdv_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Limpa_Combos()
    dtcPdv.Text = Empty
    cbbCampos.Text = Empty
End Function

Public Function Consulta()
    Dim strOperacao As Integer
    Dim intContador As Integer
    
    If cbbCampos.Text = Empty Then
       MsgBox "Selecione um campo para consulta.", vbCritical, "Logicx"
       cbbCampos.SetFocus
       Exit Function
    End If
     
    If cbbCampos.Text = "Operação" Then
       If txtConsulta.Text = "ABERTURA" Then
          strOperacao = 0
       Else
          strOperacao = 1
       End If
    End If
    
    strSql = "SELECT FKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf,DFData_TBEncerrante," & _
             "DFHora_TBEncerrante," & _
             "FKCodigo_TBPdv,DFEndereco_ip_TBPdv,DFAbertura_fechamento_TBEncerrante,PKId_TBEncerrante " & _
             "FROM TBEncerrante " & _
             "INNER JOIN TBPdv ON TBEncerrante.FKCodigo_TBPdv = TBPdv.PKCodigo_TBPdv " & _
             "INNER JOIN TBOperadores_ecf ON TBEncerrante.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf "
             
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    If cbbCampos.Text <> "Todos" Then
        If cbbCampos.Text = "Cod. Operador" Then
           strSql = strSql & " WHERE FKCodigo_TBOperadores_ecf = '" & txtConsulta.Text & "'"
        ElseIf cbbCampos.Text = "Operador" Then
            strSql = strSql & " WHERE DFNome_TBOperadores_ecf LIKE '%" & txtConsulta.Text & "%'"
        ElseIf cbbCampos.Text = "Cod. PDV" Then
            strSql = strSql & " WHERE FKCodigo_TBPdv LIKE '%" & txtConsulta.Text & "%'"
        ElseIf cbbCampos.Text = "Endereço de IP" Then
            strSql = strSql & " WHERE convert(nvarchar,DFEndereco_ip_TBPdv) LIKE '%" & txtConsulta.Text & "%'"
        ElseIf cbbCampos.Text = "Data" Then
            strSql = strSql & " WHERE TBEncerrante.DFData_TBEncerrante >= '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                              " AND TBEncerrante.DFData_TBEncerrante <= '" & Format(dtpFinal.Value, "YYYYMMDD") & "' "
        ElseIf cbbCampos.Text = "Hora" Then
            strSql = strSql & " WHERE convert(nvarchar,DFHora_TBEncerrante) LIKE '%" & txtConsulta.Text & "%'"
        ElseIf cbbCampos.Text = "Operação" Then
            strSql = strSql & " WHERE convert(bit,DFAbertura_fechamento_TBEncerrante) = '" & strOperacao & "'"
        End If
    End If

    strSql = strSql & " Order by PKId_TBEncerrante"
'    strSql = strSql + "GROUP BY FKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf," & _
'                    "TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto "

    frmAguarde.Show
    DoEvents
        
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgEncerrante, strTamanho, strNomes, "BDRetaguarda", "Otica", Me

    hfgEncerrante.Row = 1
    hfgEncerrante.Col = 0
    If hfgEncerrante.Text = Empty Then
       hfgEncerrante.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgEncerrante, strTamanho, strNomes, 8, "Otica", Me
    Else
       hfgEncerrante.Col = 7
       intContador = 1
       Do While intContador <= hfgEncerrante.Rows - 1
          hfgEncerrante.Row = intContador
          If hfgEncerrante.Text = "Não" Then
             hfgEncerrante.Text = "Abertura"
          Else
             hfgEncerrante.Text = "Fechamento"
          End If
          intContador = intContador + 1
       Loop
    End If
    Unload frmAguarde
    hfgEncerrante.Row = 1
    hfgEncerrante.Col = 0
    hfgEncerrante.SetFocus
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Cod. Operador")
    cbbCampos.AddItem ("Cod. PDV")
    cbbCampos.AddItem ("Endereço de IP")
    cbbCampos.AddItem ("Data")
    cbbCampos.AddItem ("Hora")
    cbbCampos.AddItem ("Operação")
    
    cbbOperacao.Clear
    cbbOperacao.AddItem ("Abertura")
    cbbOperacao.AddItem ("Fechamento")
    
End Function

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Sub txtEncerrante_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtEncerrante_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 44 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtEncerrante_LostFocus()
    If strCasas_Decimais = 2 Then
       txtEncerrante.Text = Format(txtEncerrante, "#,###0.00")
    ElseIf strCasas_Decimais = 3 Then
       txtEncerrante.Text = Format(txtEncerrante, "#,###0.000")
    Else
       txtEncerrante.Text = Format(txtEncerrante.Text, "#,##0.00")
    End If
End Sub

Private Sub txtOperador_Change()
    dtcOperador.BoundText = txtOperador.Text
    If IsNumeric(txtOperador.Text) = False Then txtOperador.Text = Empty: Exit Sub
End Sub

Private Sub txtOperador_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKId_TBEncerrante", strID_Encerrante, "DFIntegrado_filiais_TBEncerrante", "TBEncerrante", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBEncerrante", Me.Top, Me.Left, Me.width, Me.Height, "Encerrante")
    


End Function
