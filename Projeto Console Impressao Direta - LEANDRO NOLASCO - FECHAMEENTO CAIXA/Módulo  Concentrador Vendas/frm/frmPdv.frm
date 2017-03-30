VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmPdv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDV"
   ClientHeight    =   4830
   ClientLeft      =   1830
   ClientTop       =   2040
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPdv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7035
   Begin TabDlg.SSTab sstPdv 
      Height          =   4485
      Left            =   0
      TabIndex        =   17
      Top             =   330
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   7911
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
      TabPicture(0)   =   "frmPdv.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "freInformacoes_adicionais"
      Tab(0).Control(1)=   "txtEndereco_ip"
      Tab(0).Control(2)=   "txtNumero_ECF"
      Tab(0).Control(3)=   "txtCodigo"
      Tab(0).Control(4)=   "dtcEmpresa"
      Tab(0).Control(5)=   "Label18"
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(8)=   "Label7"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmPdv.frx":179E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cbbConsulta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cbbCampos"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "hfgPdv"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtConsulta"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdRefresh"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdConsulta"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.Frame freInformacoes_adicionais 
         Caption         =   "Periféricos Acoplados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2415
         Left            =   -74880
         TabIndex        =   24
         Top             =   1920
         Width           =   6735
         Begin VB.TextBox txtCaminho_Impressora 
            Height          =   360
            Left            =   120
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1890
            Width           =   6465
         End
         Begin VB.TextBox txtPorta 
            Height          =   360
            Left            =   3420
            MaxLength       =   2
            TabIndex        =   10
            Top             =   570
            Width           =   1515
         End
         Begin VB.TextBox txtImpressora 
            Height          =   360
            Left            =   2010
            TabIndex        =   13
            Top             =   1230
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo dtcImpressora 
            Height          =   360
            Left            =   3390
            TabIndex        =   14
            Top             =   1230
            Width           =   3225
            _ExtentX        =   5689
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
         Begin AutoCompletar.CbCompleta cbbGaveta_integrada 
            Height          =   360
            Left            =   120
            TabIndex        =   8
            Top             =   570
            Width           =   1845
            _ExtentX        =   3254
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
         Begin AutoCompletar.CbCompleta cbbImpressoes_utilizadas 
            Height          =   360
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Fiscal/Não Fiscal/Ambas"
            Top             =   1230
            Width           =   1845
            _ExtentX        =   3254
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
         Begin AutoCompletar.CbCompleta cbbLeitor 
            Height          =   360
            Left            =   2010
            TabIndex        =   9
            Top             =   570
            Width           =   1365
            _ExtentX        =   2408
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
         Begin AutoCompletar.CbCompleta cbbTipo_Impressora 
            Height          =   360
            Left            =   4980
            TabIndex        =   11
            ToolTipText     =   "Inativo"
            Top             =   570
            Width           =   1635
            _ExtentX        =   2884
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho Impressora"
            Height          =   240
            Left            =   120
            TabIndex        =   32
            Top             =   1650
            Width           =   1785
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Impressora"
            Height          =   240
            Left            =   4980
            TabIndex        =   30
            Top             =   330
            Width           =   1410
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Porta COM"
            Height          =   240
            Left            =   3420
            TabIndex        =   29
            Top             =   330
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Leitor Serial"
            Height          =   240
            Left            =   2010
            TabIndex        =   28
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impressora"
            Height          =   240
            Left            =   2010
            TabIndex        =   27
            Top             =   990
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gaveta Integrada"
            Height          =   240
            Left            =   120
            TabIndex        =   26
            Top             =   330
            Width           =   1470
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Imp. Utilizadas"
            Height          =   240
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "Impressoras Utilizadas"
            Top             =   990
            Width           =   1290
         End
      End
      Begin VB.TextBox txtEndereco_ip 
         Height          =   375
         Left            =   -70320
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
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
         Left            =   6120
         Picture         =   "frmPdv.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
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
         Left            =   6510
         Picture         =   "frmPdv.frx":34B4
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   2190
         TabIndex        =   1
         Top             =   780
         Width           =   3855
      End
      Begin VB.TextBox txtNumero_ECF 
         Height          =   375
         Left            =   -73230
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1440
         Width           =   2865
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         ToolTipText     =   "Código do Ponto de Venda"
         Top             =   1440
         Width           =   1605
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgPdv 
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   5530
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
         TabIndex        =   0
         Top             =   780
         Width           =   2025
         _ExtentX        =   3572
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
         Left            =   2190
         TabIndex        =   16
         Top             =   780
         Width           =   3855
         _ExtentX        =   6800
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
         Left            =   -74880
         TabIndex        =   4
         Top             =   780
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa [ F2 ]"
         Height          =   240
         Left            =   -74880
         TabIndex        =   31
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Endereço de IP"
         Height          =   240
         Left            =   -70320
         TabIndex        =   23
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número ECF"
         Height          =   240
         Left            =   -73230
         TabIndex        =   19
         Top             =   1200
         Width           =   1065
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
         Left            =   -74880
         TabIndex        =   18
         Top             =   1200
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
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
      Left            =   7230
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
            Picture         =   "frmPdv.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPdv.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPdv.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPdv.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPdv.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPdv.frx":5578
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPdv.frx":5892
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Cadastro de PDV                                                '
' Data de Criação........: 17/01/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
Option Explicit

Function Imprimir()

    On Error GoTo erro
    'Tratamento de erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       'cbbCampos.SetFocus
       Me.txtConsulta.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Relatorio_Pdv.Show
        
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
    ElseIf cbbCampos.Text = "Gaveta Integrada" Then
       txtConsulta.Visible = False
       cbbConsulta.Visible = True
       cbbConsulta.SetFocus
    ElseIf cbbCampos.Text = "Leitor Serial" Then
       txtConsulta.Visible = False
       cbbConsulta.Visible = True
       cbbConsulta.SetFocus
    Else
       txtConsulta.Visible = True
       cbbConsulta.Visible = False
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cbbLeitor_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub cbbTipo_Impressora_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub dtcEmpresa_LostFocus()
    dtcEmpresa.Enabled = False
End Sub

Private Sub dtcImpressora_GotFocus()
    If txtImpressora.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcImpressora.Text)
    End If
End Sub

Private Sub dtcImpressora_LostFocus()
    txtImpressora.Text = dtcImpressora.BoundText
    If IsNumeric(txtImpressora.Text) = False Or dtcImpressora.Text = Empty Then txtImpressora.Text = Empty: Exit Sub
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
    
    'If KeyCode = "113" Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
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
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de PDV"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstPdv, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    
    log.Descricao = "Inicializando cadastro de PDV"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    sstPdv.TabEnabled(0) = False
    sstPdv.Tab = 1
    Call Reposicao

    Exit Sub
erro:
    Call erro.erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro
    
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
erro:
    Call erro.erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgPdv_Click()
    If hfgPdv.Col = 0 Then
        
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
       
       txtCodigo.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 1))
       txtNumero_ECF.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 2))
       txtImpressora.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 3))
       txtEndereco_ip.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 5))
       cbbGaveta_integrada.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 6))
       cbbImpressoes_utilizadas.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 7))
       cbbLeitor.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 8))
       txtPorta.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 9))
       cbbTipo_Impressora.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 10))
       txtCaminho_Impressora.Text = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 11))
       dtcEmpresa.BoundText = hfgPdv.TextArray((hfgPdv.Row * hfgPdv.Cols + hfgPdv.Col + 12))
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstPdv.TabEnabled(0) = True
       sstPdv.Tab = 0
       txtNumero_ECF.SetFocus
   End If
   Unload frmAguarde
End Sub

Private Sub hfgPdv_DblClick()
    hfgPdv.Sort = 1
End Sub

Private Sub hfgPdv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgPdv_Click
    End If
End Sub

Private Sub sstPdv_Click(PreviousTab As Integer)
    If sstPdv.Tab = 0 Then
       txtNumero_ECF.SetFocus
    ElseIf sstPdv.Tab = 1 Then
       If frmIntegracao.Visible = True Then
          Unload frmIntegracao
       End If
       If strCombo <> Empty And strCombo <> "Todos" And strCombo <> "Gaveta Integrada" And strCombo <> "Leitor Serial" Then
          cbbCampos.Text = strCombo
          txtConsulta.SetFocus
       ElseIf strCombo = "Todos" Then
          hfgPdv.Row = 1
          hfgPdv.Col = 0
          hfgPdv.SetFocus
       ElseIf strCombo = "Gaveta Integrada" Then
          cbbCampos.Text = strCombo
          cbbConsulta.SetFocus
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
    On Error GoTo erro
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim strGaveta As String
    Dim strImpressora As String
    Dim strLeitor As String
    Dim strTipo_Impressora As String
    
    If txtCodigo.Text = Empty Then
       MsgBox "Código não pode ser nulo.", vbInformation, "Only Tech"
       txtCodigo.SetFocus
       Exit Function
    End If
    
    If cbbGaveta_integrada.Text = "Sim" Then
       strGaveta = 1
    Else
       strGaveta = 0
    End If
    
    If cbbLeitor.Text = "Sim" Then
       strLeitor = 1
    Else
       strLeitor = 0
    End If
    
    If cbbImpressoes_utilizadas.Text = "CF" Then
       strImpressora = "0"
    ElseIf cbbImpressoes_utilizadas.Text = "CNF" Then
       strImpressora = "1"
    Else
       strImpressora = "2"
    End If
    
    If cbbTipo_Impressora.Text = "Comum" Then
       strTipo_Impressora = 1
    ElseIf cbbTipo_Impressora.Text = "Não Fiscal" Then
       strTipo_Impressora = 0
    End If
    
    Call Objetos.Retira_Espaco_Lateral(Me)
    Call Objetos.Maiusculo_TXT(Me)
    
    strCampo = "PKCodigo_TBPdv,FKCodigo_TBImpressoras_ecf,DFNumero_ecf_TBPdv,DFEndereco_ip_TBPdv," & _
               "DFGaveta_integrada_TBPdv,DFImpressoes_suportadas_TBPdv,DFLeitor_Serial_integrado," & _
               "DFPorta_com_leitor_serial,DFTipo_impressora_orcamento_balcao_TBpdv,IXCodigo_TBEmpresa," & _
               "DFCaminho_Impressora_Comum,DFData_alteracao_TBPdv,DFIntegrado_filiais_TBPdv"
                         
    If booIntegra_Portal = True Then
       strCampo = strCampo & ",DFIntegrado_portal_TBPdv"
    End If
    
    strValores = "" & txtCodigo & ",'" & txtImpressora.Text & "'," & txtNumero_ECF.Text & "," & _
                 "'" & txtEndereco_ip.Text & "'," & strGaveta & "," & strImpressora & "," & _
                 "" & strLeitor & ",'" & txtPorta.Text & "','" & strTipo_Impressora & "'," & _
                 "" & dtcEmpresa.BoundText & ",'" & txtCaminho_Impressora.Text & "','" & _
                 "" & Format(Date, "YYYYMMDD") & "',0"
    
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
    End If
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFNumero_ecf_TBPdv = " & txtNumero_ECF.Text & "," & _
                "    FKCodigo_TBImpressoras_ecf = '" & txtImpressora.Text & "'," & _
                "    DFEndereco_ip_TBPdv = '" & txtEndereco_ip.Text & "'," & _
                "    DFGaveta_integrada_TBPdv = '" & strGaveta & "', " & _
                "    DFImpressoes_suportadas_TBPdv = '" & strImpressora & "'," & _
                "    DFLeitor_Serial_integrado= '" & strLeitor & "'," & _
                "    DFPorta_com_leitor_serial = '" & txtPorta.Text & "'," & _
                "    DFTipo_impressora_orcamento_balcao_TBpdv = '" & strTipo_Impressora & "'," & _
                "    IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & "," & _
                "    DFCaminho_Impressora_Comum = '" & txtCaminho_Impressora.Text & "'," & _
                "    DFData_alteracao_TBPdv = '" & Format(Date, "YYYYMMDD") & "'," & _
                "    DFIntegrado_filiais_TBPdv = 0"
                
       If booIntegra_Portal = True Then
          strSet = strSet & ", DFIntegrado_portal_TBPdv = 0"
       End If
       
       Call funcoes_banco.Alterar("TBPdv", strSet, "PKCodigo_TBPdv", txtCodigo.Text, "OTICA", Me, "BDRetaguarda")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBPdv", strCampo, strValores, "OTICA", Me, "BDRetaguarda")
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
       hfgPdv.Visible = False
    End If
    
    txtCodigo.Enabled = False
    sstPdv.TabEnabled(0) = False
    sstPdv.Tab = 1
    hfgPdv.Refresh
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo erro
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBPdv", "PKCodigo_TBPdv", Me.txtCodigo.Text, "OTICA", Me, "BDRetaguarda")
    
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
       hfgPdv.Visible = False
    End If
            
    sstPdv.TabEnabled(0) = False
    sstPdv.Tab = 1
    
    Exit Function
erro:
     Call erro.erro(Me, "OTICA", "Excluir")
     Exit Function
End Function

Private Function Cancelar()
    On Error GoTo erro
    
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
    
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
       hfgPdv.Visible = False
    End If
    
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Operação com Registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    cbbGaveta_integrada.Text = Empty
    sstPdv.TabEnabled(0) = False
    sstPdv.Tab = 1
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo erro
           
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
    
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
    
    sstPdv.TabEnabled(0) = True
    sstPdv.Tab = 0
    
    'dtcCodigo_empresa.boundtext = ---- Inserir aqui informações da DLLIntercomunicador de EXE's
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    dtcEmpresa.Enabled = False
   
    cbbGaveta_integrada.Text = Empty
    cbbImpressoes_utilizadas.Text = Empty
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    booAlterar = False
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Novo")
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
       Movimentacoes.Verifica_Numero "PKCodigo_TBPdv", "TBPdv", txtCodigo, "OTICA", Me
       Call txtEndereco_ip_LostFocus
    End If
End Sub

Private Function Reposicao()
   On Error GoTo erro
          
    Movimentacoes.Monta_HFlex_Grid hfgPdv, "1000,1000,1200,2000,1700,1700,1700,1500,1500,1500,2500,1200,2200", "Código,Nº ECF,Impressora,Nome,Endereço IP,Gaveta Integrada,Imp. Utilizadas,Leitor Serial,Porta COM,Tipo Impressora,Caminho Impressora,Empresa,Nome", 13, "OTICA", Me
     
    Call Monta_Combo
    Call Monta_DataCombo
              
    hfgPdv.Refresh

    Exit Function
erro:
   Call erro.erro(Me, "OTICA", "Reposicao")
   Resume Next
End Function

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Function Consulta()
    Dim strGaveta As String
    Dim contBens As String
    Dim strImpressoes As String
    Dim strLeitor As String
    Dim strTipo_Impressora As String
    
    If cbbCampos.Text <> "Todos" And cbbConsulta.Visible = False Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    ElseIf cbbConsulta.Visible = True Then
       If cbbConsulta.Text = Empty Then
          MsgBox "Selecione uma opção para consulta.", vbCritical, "Only Tech"
          cbbConsulta.SetFocus
          Exit Function
       End If
    End If
            
    If cbbCampos.Text = "Gaveta Integrada" Then
        If cbbConsulta.Text = "Sim" Then
           strGaveta = 1
        Else
           strGaveta = 0
        End If
    ElseIf cbbCampos.Text = "Leitor Serial" Then
       If cbbConsulta.Text = "Sim" Then
          strLeitor = 1
       Else
          strLeitor = 0
       End If
    ElseIf cbbCampos.Text = "Impressões Utilizadas" Then
       If txtConsulta.Text = "CF" Then
          strImpressoes = 0
       ElseIf txtConsulta.Text = "CNF" Then
          strImpressoes = 1
       ElseIf txtConsulta.Text = "AMBAS" Then
          strImpressoes = 2
       End If
    ElseIf cbbCampos.Text = "Tipo Impressora" Then
        If txtConsulta.Text = "NAO FISCAL" Or txtConsulta.Text = "NÃO FISCAL" Then
          strTipo_Impressora = 0
       ElseIf txtConsulta.Text = "Comum" Or txtConsulta.Text = "COMUM" Then
          strTipo_Impressora = 1
       End If
    End If
    
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
           
    txtConsulta.Text = Funcoes_Gerais.Grava_String(txtConsulta.Text)
    
    strSql = "SELECT TBPdv.PKCodigo_TBPdv," & _
             "TBPdv.DFNumero_ecf_TBPdv," & _
             "TBPdv.FKCodigo_TBImpressoras_ecf," & _
             "TBImpressoras_ecf.DFNome_TBImpressoras_ecf,DFEndereco_ip_TBPdv,DFGaveta_integrada_TBPdv," & _
             "TBPdv.DFImpressoes_suportadas_TBPdv,TBPdv.DFLeitor_Serial_integrado," & _
             "TBPdv.DFPorta_com_leitor_serial, " & _
             "DFTipo_impressora_orcamento_balcao_TBpdv,DFCaminho_Impressora_Comum,PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa " & _
             "FROM TBPdv " & _
             "INNER JOIN TBImpressoras_ecf ON TBImpressoras_ecf.PKCodigo_TBImpressoras_ecf = TBPdv.FKCodigo_TBImpressoras_ecf " & _
             "INNER JOIN TBEmpresa ON TBEmpresa.PKCodigo_TBEmpresa = TBPdv.IXCodigo_TBEmpresa"
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código do PDV" Then
          strSql = strSql & " WHERE convert(nvarchar,PKCodigo_TBPdv) = " & txtConsulta.Text & " "
       ElseIf cbbCampos.Text = "Número ECF" Then
          strSql = strSql & " WHERE convert(nvarchar,DFNumero_ecf_TBPdv) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Código da Impressora" Then
          strSql = strSql & " WHERE convert(nvarchar,FKCodigo_TBImpressoras_ecf) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Endereço IP" Then
          strSql = strSql & " WHERE convert(nvarchar,DFEndereco_ip_TBPdv) LIKE '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Gaveta Integrada" Then
          strSql = strSql & " WHERE convert(bit,DFGaveta_integrada_TBPdv) = '" & strGaveta & "'"
       ElseIf cbbCampos.Text = "Código da Impressora" Then
          strSql = strSql & " WHERE convert(nvarchar,FKCodigo_TBImpressoras_ecf) LIKE '%" & txtConsulta.Text & "%'"
       ElseIf cbbCampos.Text = "Impressões Utilizadas" Then
          strSql = strSql & " WHERE DFImpressoes_suportadas_TBPdv = '" & strImpressoes & "'"
       ElseIf cbbCampos.Text = "Leitor Serial" Then
          strSql = strSql & " WHERE DFLeitor_Serial_integrado = '" & strLeitor & "'"
       ElseIf cbbCampos.Text = "Porta COM" Then
          strSql = strSql & " WHERE DFPorta_com_leitor_serial = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Tipo Impressora" Then
          strSql = strSql & " WHERE DFTipo_impressora_orcamento_balcao_TBpdv = '" & strTipo_Impressora & "'"
       ElseIf cbbCampos.Text = "Caminho Impressora" Then
          strSql = strSql & " WHERE convert(nvarchar,DFCaminho_Impressora_Comum) LIKE '%" & txtConsulta.Text & "%'"
       
       End If
    End If
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgPdv, "1000,1000,1200,2000,1600,1700,1700,1500,1500,1500,2500,1200,2200", "Código,Nº ECF,Impressora,Nome,Endereço IP,Gaveta Integrada,Imp. Utilizadas,Leitor Serial,Porta COM,Tipo Impressora,Caminho Impressora,Empresa,Nome", "BDRetaguarda", "OTICA", Me
     
    hfgPdv.Col = 1
    hfgPdv.Row = 1
    
    If hfgPdv.Text = Empty Then
          hfgPdv.Rows = 2
          Movimentacoes.Monta_HFlex_Grid hfgPdv, "1000,1000,1000,2000,1600,1700,1700,1500,1500,1500,2500,1200,2200", "Código,Nº ECF,Impressora,Nome,Endereço IP,Gaveta Integrada,Imp. Utilizadas,Leitor Serial,Porta COM,Tipo Impressora,Caminho Impressora,Empresa,Nome", 13, "OTICA", Me
       Else
          hfgPdv.Col = 7
          contBens = 1
          Do While contBens <= hfgPdv.Rows - 1
              hfgPdv.Row = contBens
              If hfgPdv.Text = "0" Then
                 hfgPdv.Text = "CF"
              ElseIf hfgPdv.Text = "1" Then
                 hfgPdv.Text = "CNF"
              Else
                 hfgPdv.Text = "Ambas"
              End If
              contBens = contBens + 1
          Loop
          
          hfgPdv.Col = 10
          contBens = 1
          Do While contBens <= hfgPdv.Rows - 1
              hfgPdv.Row = contBens
              If hfgPdv.Text = "0" Then
                 hfgPdv.Text = "Não Fiscal"
              ElseIf hfgPdv.Text = "1" Then
                 hfgPdv.Text = "Comum"
              End If
              contBens = contBens + 1
          Loop
    End If
     
    frmAguarde.Show
    DoEvents

    Unload frmAguarde
    hfgPdv.Col = 0
    hfgPdv.Row = 1
    hfgPdv.SetFocus
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código do PDV")
    cbbCampos.AddItem ("Número ECF")
    cbbCampos.AddItem ("Código da Impressora")
    cbbCampos.AddItem ("Nome da Impressora")
    cbbCampos.AddItem ("Endereço IP")
    cbbCampos.AddItem ("Gaveta Integrada")
    cbbCampos.AddItem ("Impressões Utilizadas")
    cbbCampos.AddItem ("Leitor Serial")
    cbbCampos.AddItem ("Porta COM")
    cbbCampos.AddItem ("Tipo Impressora")
    cbbCampos.AddItem ("Caminho Impressora")
    
    cbbGaveta_integrada.Clear
    cbbGaveta_integrada.AddItem ("Sim")
    cbbGaveta_integrada.AddItem ("Não")
    
    cbbImpressoes_utilizadas.Clear
    cbbImpressoes_utilizadas.AddItem ("CF")
    cbbImpressoes_utilizadas.AddItem ("CNF")
    cbbImpressoes_utilizadas.AddItem ("Ambas")
    
    cbbTipo_Impressora.Clear
    cbbTipo_Impressora.AddItem ("Comum")
    cbbTipo_Impressora.AddItem ("Não Fiscal")
    
    cbbConsulta.Clear
    cbbConsulta.AddItem ("Sim")
    cbbConsulta.AddItem ("Não")
    
    cbbLeitor.Clear
    cbbLeitor.AddItem ("Sim")
    cbbLeitor.AddItem ("Não")
End Function

Private Sub txtEndereco_ip_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtEndereco_ip_KeyPress(KeyAscii As Integer)
'    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> Asc(".") Then
'       KeyAscii = 0
'    End If
End Sub

Private Sub txtEndereco_ip_LostFocus()
    If txtEndereco_ip.Text <> Empty And txtCodigo.Text <> Empty Then
       strSql = "SELECT PKCodigo_TBPdv,PKCodigo_TBPdv,DFEndereco_ip_TBPdv FROM TBPdv " & _
                "WHERE DFEndereco_ip_TBPdv = '" & txtEndereco_ip.Text & "'"
       
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstAplicacao, "Otica", Me)
       
       If rstAplicacao.RecordCount <> 0 And Not IsNull(rstAplicacao.Fields("DFEndereco_ip_TBPdv")) Then
          If rstAplicacao.Fields("PKCodigo_TBPdv") <> txtCodigo.Text Then
             MsgBox "Endereço de IP pertencente a outro cadastro de PDV. Verifique.", vbInformation, "Only Tech"
             txtEndereco_ip.SetFocus
             Set rstAplicacao = Nothing
             Exit Sub
          End If
       End If
       Set rstAplicacao = Nothing
    End If
End Sub

Private Sub txtImpressora_Change()
    dtcImpressora.BoundText = txtImpressora.Text
    If IsNumeric(txtImpressora.Text) = False Then txtImpressora.Text = Empty: Exit Sub
End Sub

Private Sub txtImpressora_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtImpressora_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumero_ECF_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_ECF_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
Private Function Monta_DataCombo()
    strSql = Empty
    strSql = "SELECT PKCodigo_TBImpressoras_ecf,DFNome_TBImpressoras_ecf FROM TBImpressoras_ecf"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBImpressoras_ecf", "DFNome_TBImpressoras_ecf", dtcImpressora, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    'dtcCodigo_empresa.boundtext = ---- Inserir aqui informações da DLLIntercomunicador de EXE's
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
End Function

Private Sub txtPorta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPorta_LostFocus()
    txtPorta.Text = UCase(txtPorta.Text)
End Sub

Private Function Limpa_Combos()
    cbbGaveta_integrada.Text = Empty
    cbbLeitor.Text = Empty
    cbbTipo_Impressora.Text = Empty
    cbbImpressoes_utilizadas.Text = Empty
    txtImpressora.Text = Empty
End Function

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo_TBPdv", txtCodigo.Text, "DFIntegrado_filiais_TBPdv", "TBPdv", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBPdv", Me.Top, Me.Left, Me.width, Me.Height, "PDV")
    
End Function
