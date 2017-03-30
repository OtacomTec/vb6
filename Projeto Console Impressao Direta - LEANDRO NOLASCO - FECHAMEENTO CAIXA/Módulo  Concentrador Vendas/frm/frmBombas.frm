VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmBombas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bombas"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBombas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8385
   Begin TabDlg.SSTab sstBomba 
      Height          =   6315
      Left            =   0
      TabIndex        =   19
      Top             =   330
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   11139
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
      TabPicture(0)   =   "frmBombas.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtcEmpresa"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "hfgBomba_Bico"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCodigo_Bomba"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDescricao_Bomba"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtNumero_Bicos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmBombas.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "cbbCampos"
      Tab(1).Control(2)=   "hfgBomba"
      Tab(1).Control(3)=   "cmdConsulta"
      Tab(1).Control(4)=   "cmdRefresh"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtConsulta"
      Tab(1).Control(6)=   "cmdBomba_Consulta_Empresa"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdBomba_Consulta_Empresa 
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
         Left            =   -67980
         Picture         =   "frmBombas.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtNumero_Bicos 
         Height          =   360
         Left            =   6900
         TabIndex        =   7
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox txtDescricao_Bomba 
         Height          =   360
         Left            =   1500
         TabIndex        =   6
         Top             =   1440
         Width           =   5355
      End
      Begin VB.TextBox txtCodigo_Bomba 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Código da Bomba"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Bicos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   120
         TabIndex        =   23
         Top             =   1830
         Width           =   8085
         Begin VB.TextBox txtCapacidade 
            Enabled         =   0   'False
            Height          =   360
            Left            =   5520
            MaxLength       =   100
            TabIndex        =   33
            ToolTipText     =   "Capacidade do Tanque"
            Top             =   1290
            Width           =   2415
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
            Left            =   6750
            TabIndex        =   17
            ToolTipText     =   "Remover"
            Top             =   1950
            Width           =   1185
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
            Left            =   5520
            TabIndex        =   16
            ToolTipText     =   "Incluir"
            Top             =   1950
            Width           =   1185
         End
         Begin VB.TextBox txtUltimo_Encerrante 
            Height          =   360
            Left            =   4110
            MaxLength       =   100
            TabIndex        =   10
            Top             =   630
            Width           =   1845
         End
         Begin VB.TextBox txtProduto 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   120
            MaxLength       =   6
            TabIndex        =   14
            Top             =   1950
            Width           =   1245
         End
         Begin VB.TextBox txtNumero_Maximo_Encerrante 
            Height          =   360
            Left            =   6000
            MaxLength       =   100
            TabIndex        =   11
            Top             =   630
            Width           =   1935
         End
         Begin VB.TextBox txtTanque 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   120
            MaxLength       =   6
            TabIndex        =   12
            ToolTipText     =   "Código do Tanque"
            Top             =   1290
            Width           =   1245
         End
         Begin VB.TextBox txtCodigo_Bico 
            Height          =   360
            Left            =   120
            MaxLength       =   100
            TabIndex        =   8
            Top             =   630
            Width           =   1245
         End
         Begin MSDataListLib.DataCombo dtcProduto 
            Height          =   360
            Left            =   1410
            TabIndex        =   15
            Top             =   1950
            Width           =   4050
            _ExtentX        =   7144
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
         Begin MSDataListLib.DataCombo dtcTanque 
            Height          =   360
            Left            =   1410
            TabIndex        =   13
            Top             =   1290
            Width           =   4050
            _ExtentX        =   7144
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
         Begin AutoCompletar.CbCompleta cbbTipo_Preco 
            Height          =   360
            Left            =   1410
            TabIndex        =   9
            Top             =   630
            Width           =   2655
            _ExtentX        =   4683
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Capacidade"
            Height          =   240
            Left            =   5520
            TabIndex        =   34
            Top             =   1050
            Width           =   990
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Último Encerrante"
            Height          =   240
            Left            =   4110
            TabIndex        =   29
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Produto"
            Height          =   240
            Left            =   120
            TabIndex        =   28
            Top             =   1710
            Width           =   660
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Nº Max. Encerrante"
            Height          =   240
            Left            =   6000
            TabIndex        =   27
            Top             =   390
            Width           =   1665
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tanque"
            Height          =   240
            Left            =   120
            TabIndex        =   26
            Top             =   1050
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   240
            Left            =   120
            TabIndex        =   25
            Top             =   390
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Preço"
            Height          =   240
            Left            =   1410
            TabIndex        =   24
            Top             =   390
            Width           =   1185
         End
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -73020
         TabIndex        =   1
         Top             =   720
         Width           =   4965
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
         Left            =   -67170
         Picture         =   "frmBombas.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   -67590
         Picture         =   "frmBombas.frx":383E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Consultar"
         Top             =   720
         Width           =   405
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgBomba 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   3
         Top             =   1170
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   8811
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgBomba_Bico 
         Height          =   1785
         Left            =   120
         TabIndex        =   18
         Top             =   4380
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   3149
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   780
         Width           =   8085
         _ExtentX        =   14261
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
         TabIndex        =   36
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nº Bicos"
         Height          =   240
         Left            =   6900
         TabIndex        =   32
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   1500
         TabIndex        =   31
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   585
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8970
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
            Picture         =   "frmBombas.frx":5538
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBombas.frx":5852
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBombas.frx":5B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBombas.frx":5F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBombas.frx":62A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBombas.frx":65BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBombas.frx":68D4
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
      Width           =   8385
      _ExtentX        =   14790
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
Attribute VB_Name = "frmBombas"
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
Dim strTamanho As String
Dim strNomes As String
Dim strNomes_bico As String
Dim strTamanho_bico As String
Dim strCombo As String
Dim strConsulta As String
Dim strClique_Bico As Integer
Dim strCampo_consulta As String
Dim strId_remover As String
Dim booAlterar As Boolean
Dim strId_Bomba As String
Dim strCasas_Decimais As Integer
Dim rstVerifica_Titulo As New ADODB.Recordset
Dim rstAplicacao As New ADODB.Recordset
Dim strID As String
Public strSql As String
Public strCodigo_Empresa_Consulta As String
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim log As New DLLSystemManager.log
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean

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
    
    'Call frmConsole_Relatorio_Bomba.Show
    
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
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    Else
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cbbTipo_Preco_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub cmdBomba_Consulta_Empresa_Click()
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
    
    Unload frmBombas_Consulta_Empresa
    frmAguarde.Show
    DoEvents
    frmBombas_Consulta_Empresa.Show
    Unload frmAguarde
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdIncluir_Click()
    Dim strIndice As String
    Dim intContador As Integer

    If txtCodigo_Bomba.Text = Empty Then
       MsgBox "Bomba inválida. Verifique!", vbCritical, "Only Tech"
       txtCodigo_Bomba.SetFocus
       Exit Sub
    ElseIf txtCodigo_Bico.Text = Empty Then
       MsgBox "Defina um Bico antes de acrescentar ao cadastro.", vbInformation, "Only Tech"
       txtCodigo_Bico.SetFocus
       Exit Sub
    ElseIf txtProduto.Text = Empty Then
       MsgBox "Defina um Produto antes de acrescentar ao cadastro.", vbInformation, "Only Tech"
       txtProduto.SetFocus
       Exit Sub
    ElseIf txtTanque.Text = Empty Then
       MsgBox "Defina um Tanque antes de acrescentar ao cadastro.", vbInformation, "Only Tech"
       txtTanque.SetFocus
       Exit Sub
    End If
    
    If txtNumero_Bicos.Text <> Empty Then
        If txtNumero_Bicos.Text <> 1 Then
           If CDbl(txtNumero_Bicos.Text) = (hfgBomba_Bico.Rows - 1) And cmdIncluir.Caption = "Incluir" Then
              MsgBox "O número de Bicos não pode ser maior que o limite da Bomba. Verifique.", vbInformation, "Only Tech"
              txtCodigo_Bico.SetFocus
              Exit Sub
           End If
        Else
           If CDbl(txtNumero_Bicos.Text) = (hfgBomba_Bico.Rows - 1) And cmdIncluir.Caption = "Incluir" Then
              hfgBomba_Bico.Row = 1
              hfgBomba_Bico.Col = 1
              If hfgBomba_Bico.Text <> Empty Then
                 MsgBox "O número de Bicos não pode ser maior que o limite da Bomba. Verifique.", vbInformation, "Only Tech"
                 txtCodigo_Bico.SetFocus
                 Exit Sub
              End If
           End If
        End If
    End If
    'Verificar se o item está no grid de itens do pedido
    intContador = 1

    Do While intContador <= hfgBomba_Bico.Rows - 1
        hfgBomba_Bico.Row = intContador
        hfgBomba_Bico.Col = 1
        If cmdIncluir.Caption = "Alterar" Then
           If hfgBomba_Bico.Text = txtCodigo_Bico.Text And hfgBomba_Bico.Row <> strClique_Bico Then
              MsgBox "O Bico alterado já pertence a outro item neste cadastro. Verifique.", vbInformation, "Only Tech"
              txtCodigo_Bico.SetFocus
              Exit Sub
           End If
        Else
           If hfgBomba_Bico.Text = txtCodigo_Bico.Text Then
              MsgBox "Bico já incluído neste cadastro. Verifique.", vbInformation, "Only Tech"
              txtCodigo_Bico.SetFocus
              Exit Sub
           End If
        End If
        intContador = intContador + 1
    Loop
    
    hfgBomba_Bico.Row = hfgBomba_Bico.TopRow
    If cmdIncluir.Caption = "Incluir" Then
       If hfgBomba_Bico.Text <> Empty Then
          strIndice = intContador
          hfgBomba_Bico.Rows = hfgBomba_Bico.Rows + 1
       Else
          strIndice = intContador - 1
       End If
    Else
       strIndice = strClique_Bico
    End If
    
    hfgBomba_Bico.Row = strIndice
    
    hfgBomba_Bico.Col = 0
    hfgBomba_Bico.ColWidth(0) = 500
    hfgBomba_Bico.Font.Name = "Tahoma"
    hfgBomba_Bico.CellFontSize = 7
    hfgBomba_Bico.CellFontBold = False
    hfgBomba_Bico.CellBackColor = &H80FFFF
    hfgBomba_Bico.Text = strIndice
    
    hfgBomba_Bico.Col = 1
    hfgBomba_Bico.Text = txtCodigo_Bico.Text
    
    hfgBomba_Bico.Col = 2
    hfgBomba_Bico.Text = cbbTipo_Preco.Text
    
    hfgBomba_Bico.Col = 3
    hfgBomba_Bico.Text = txtUltimo_Encerrante.Text
    
    hfgBomba_Bico.Col = 4
    hfgBomba_Bico.Text = txtNumero_Maximo_Encerrante.Text
    
    hfgBomba_Bico.Col = 5
    hfgBomba_Bico.Text = txtProduto.Text
    
    hfgBomba_Bico.Col = 6
    hfgBomba_Bico.Text = dtcProduto.Text
    
    hfgBomba_Bico.Col = 7
    hfgBomba_Bico.Text = txtTanque.Text
    
    hfgBomba_Bico.Col = 8
    hfgBomba_Bico.Text = dtcTanque.Text
    
    
    
    hfgBomba_Bico.Col = 0
    
    Do While intContador <= hfgBomba_Bico.Rows - 1
       hfgBomba_Bico.Row = intContador
       hfgBomba_Bico.Text = intContador
       intContador = intContador + 1
    Loop
       
    hfgBomba_Bico.Refresh
    
    txtCodigo_Bico.Text = Empty
    cbbTipo_Preco.Text = Empty
    txtUltimo_Encerrante.Text = Empty
    txtProduto.Text = Empty
    txtTanque.Text = Empty
    txtCapacidade.Text = Empty
    txtNumero_Maximo_Encerrante.Text = Empty
    
    txtCodigo_Bomba.Enabled = False
    txtDescricao_Bomba.Enabled = False
    txtNumero_Bicos.Enabled = False

    cmdIncluir.Caption = "Incluir"
    
    txtCodigo_Bico.SetFocus

End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub cmdRemover_Click()
    Dim intContador As Integer
    
    hfgBomba_Bico.Col = 0
    If hfgBomba_Bico.Text = Empty Then
       MsgBox "Não há Bico selecionado para esta Bomba.", vbInformation, "Only Tech"
       hfgBomba_Bico.SetFocus
       Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''
      'Verifica se o bico está vinculado a um Encerrante
      hfgBomba_Bico.Col = 1
      strSql = Empty
      strSql = "SELECT FKId_TBBomba_bico " & _
               "FROM TBEncerrante_Bomba " & _
               "INNER JOIN TBBomba_bico " & _
               "ON TBEncerrante_Bomba.FKId_TBBomba_bico = TBBomba_bico.PKId_TBBomba_bico " & _
               "WHERE IXCodigo_TBBomba_bico = '" & hfgBomba_Bico.Text & "'"
    
      Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstVerifica_Titulo, "OTICA", Me)
    
      If rstVerifica_Titulo.RecordCount <> 0 Then
         MsgBox "Este Bico encontra-se vinculado a um Encerrante e não pode ser removido. Verifique.", vbInformation, "Only Tech"
         Set rstVerifica_Titulo = Nothing
         Exit Sub
      End If
      Set rstVerifica_Titulo = Nothing
    ''''''''''''''''''''''''''''''''''''''''''
    
    hfgBomba_Bico.Col = 9
    'Guardando os Ids removidos para serem deletados no evento gravar
    If hfgBomba_Bico.Text <> Empty Then
       If strId_remover = Empty Then
          strId_remover = hfgBomba_Bico.Text
       Else
          strId_remover = strId_remover + "," + hfgBomba_Bico.Text
       End If
    End If
    
    If hfgBomba_Bico.Rows <= 2 Then
       hfgBomba_Bico.Clear
       Movimentacoes.Monta_HFlex_Grid hfgBomba_Bico, strTamanho_bico, strNomes_bico, 9, "Otica", Me
    Else
       hfgBomba_Bico.RemoveItem (hfgBomba_Bico.Row)
       hfgBomba_Bico.Col = 0
       intContador = 1
       Do While intContador <= hfgBomba_Bico.Rows - 1
          hfgBomba_Bico.Row = intContador
          hfgBomba_Bico.Text = intContador
          intContador = intContador + 1
       Loop
    End If
    
    hfgBomba_Bico.Refresh
    
    hfgBomba_Bico.Col = 0
    hfgBomba_Bico.Row = 1
    If booAlterar = False And hfgBomba_Bico.Rows <= 2 And hfgBomba_Bico.Text = Empty Then
       txtCodigo_Bomba.Enabled = True
       txtDescricao_Bomba.Enabled = True
       txtNumero_Bicos.Enabled = True
    End If
    
    cmdIncluir.Caption = "Incluir"
    
    txtTanque.Text = Empty
    txtCapacidade.Text = Empty
    txtNumero_Maximo_Encerrante.Text = Empty
    txtUltimo_Encerrante.Text = Empty
    cbbTipo_Preco.Text = Empty
    txtCodigo_Bico.Text = Empty
    
    txtCodigo_Bico.SetFocus
    hfgBomba_Bico.Col = 0
    hfgBomba_Bico.Row = 0
End Sub

Private Sub dtcEmpresa_LostFocus()
    strSql = "SELECT IXCodigo_TBTanque,DFDescricao_TBTanque FROM TBTanque WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBTanque", "DFDescricao_TBTanque", dtcTanque, strSql, "BDRetaguarda", "Otica", Me

    strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me

    dtcEmpresa.Enabled = False
End Sub

Private Sub dtcProduto_GotFocus()
    If txtTanque.Text = Empty Or dtcTanque.Text = Empty Then
       MsgBox "Tanque não definido. Verifique.", vbInformation, "Only Tech"
       txtTanque.SetFocus
       Exit Sub
    Else
       If txtProduto.Text = Empty Then
          Call Movimentacoes.Verifica_DataCombo(dtcProduto.Text)
          Call txtProduto_LostFocus
       End If
    End If
End Sub

Private Sub dtcProduto_LostFocus()
    txtProduto.Text = dtcProduto.BoundText
    If IsNumeric(txtProduto.Text) = False Or dtcProduto.Text = Empty Then txtProduto.Text = Empty: Exit Sub
End Sub

Private Sub dtcTanque_LostFocus()
    txtTanque.Text = dtcTanque.BoundText
    If IsNumeric(txtTanque.Text) = False Or dtcTanque.Text = Empty Then
       txtTanque.Text = Empty
    Else
       Call txtTanque_LostFocus
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
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstBomba, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.ocxUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
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
    
    sstBomba.TabEnabled(0) = False
    sstBomba.Tab = 1
    
    'ABASTECENDO VARIÁVEL PARA CASAS DECIMAIS
    strSql = "SELECT DFNumero_decimais_TBParametros_ecf FROM TBParametros_ecf " & _
             "WHERE FKCodigo_TBEmpresa = " & MDIPrincipal.ocxUsuario.Empresa & ""
                
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "OTICA", Me
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.ocxUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.ocxUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.ocxUsuario.Empresa, Me)
    
    If Not IsNull(rstAplicacao.Fields("DFNumero_decimais_TBParametros_ecf")) And rstAplicacao.RecordCount <> 0 Then
        strCasas_Decimais = rstAplicacao.Fields("DFNumero_decimais_TBParametros_ecf")
    End If
    Set rstAplicacao = Nothing
    
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
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    strCombo = Empty
    strCodigo_Empresa_Consulta = Empty
    Exit Sub
erro:
    Call erro.erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgBomba_Bico_Click()
    If hfgBomba_Bico.Col = 0 And hfgBomba_Bico.Text <> Empty And hfgBomba_Bico.Row <> strClique_Bico Then
        txtCodigo_Bico.Text = Empty
        txtProduto.Text = Empty
        txtTanque.Text = Empty
        cbbTipo_Preco.Text = Empty
        txtUltimo_Encerrante.Text = Empty
        txtNumero_Maximo_Encerrante.Text = Empty
        cmdIncluir.Caption = "Incluir"
    End If
End Sub

Private Sub hfgBomba_Bico_DblClick()
   If hfgBomba_Bico.Col = 0 And hfgBomba_Bico.Text <> Empty Then
       strClique_Bico = hfgBomba_Bico.Row
       cmdIncluir.Caption = "Alterar"
       txtCodigo_Bico.Text = hfgBomba_Bico.TextArray((hfgBomba_Bico.Row * hfgBomba_Bico.Cols + hfgBomba_Bico.Col + 1))
       cbbTipo_Preco.Text = hfgBomba_Bico.TextArray((hfgBomba_Bico.Row * hfgBomba_Bico.Cols + hfgBomba_Bico.Col + 2))
       txtUltimo_Encerrante.Text = hfgBomba_Bico.TextArray((hfgBomba_Bico.Row * hfgBomba_Bico.Cols + hfgBomba_Bico.Col + 3))
       txtNumero_Maximo_Encerrante.Text = hfgBomba_Bico.TextArray((hfgBomba_Bico.Row * hfgBomba_Bico.Cols + hfgBomba_Bico.Col + 4))
       txtProduto.Text = hfgBomba_Bico.TextArray((hfgBomba_Bico.Row * hfgBomba_Bico.Cols + hfgBomba_Bico.Col + 5))
       txtTanque.Text = hfgBomba_Bico.TextArray((hfgBomba_Bico.Row * hfgBomba_Bico.Cols + hfgBomba_Bico.Col + 7))
       Call txtTanque_LostFocus
    End If
    hfgBomba_Bico.SetFocus
End Sub

Private Sub hfgBomba_Click()
    Dim intContador As Integer
    
    If hfgBomba.Col = 0 And hfgBomba.Text <> Empty Then
       cbbTipo_Preco.Text = Empty
       Call Objetos.Limpa_TXT(Me)
       
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
        
       txtCodigo_Bomba.Text = hfgBomba.TextArray((hfgBomba.Row * hfgBomba.Cols + hfgBomba.Col + 1))
       txtDescricao_Bomba.Text = hfgBomba.TextArray((hfgBomba.Row * hfgBomba.Cols + hfgBomba.Col + 2))
       txtNumero_Bicos.Text = hfgBomba.TextArray((hfgBomba.Row * hfgBomba.Cols + hfgBomba.Col + 3))
       dtcEmpresa.BoundText = hfgBomba.TextArray((hfgBomba.Row * hfgBomba.Cols + hfgBomba.Col + 4))
       
       strSql = "SELECT IXCodigo_TBTanque,DFDescricao_TBTanque FROM TBTanque WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBTanque", "DFDescricao_TBTanque", dtcTanque, strSql, "BDRetaguarda", "Otica", Me
    
       strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me

       
       strID = hfgBomba.TextArray((hfgBomba.Row * hfgBomba.Cols + hfgBomba.Col + 6))
       
       'ABASTECENDO BICOS
        strSql = "SELECT TBBomba_bico.IXCodigo_TBBomba_bico," & _
                 "TBBomba_bico.DFTipo_preco_TBBomba_bico," & _
                 "TBBomba_bico.DFUltimo_encerrante_TBBomba_bico, " & _
                 "TBBomba_bico.DFNumero_maximo_encerrante_TBBomba_bico, " & _
                 "TBProduto.IXCodigo_TBProduto, " & _
                 "TBProduto.DFDescricao_TBProduto," & _
                 "TBTanque.IXCodigo_TBTanque, " & _
                 "TBTanque.DFDescricao_TBTanque,TBBomba_bico.PKId_TBBomba_bico " & _
                 "FROM TBBomba_bico " & _
                 "INNER JOIN TBProduto " & _
                 "ON TBProduto.PKId_TBProduto = TBBomba_bico.FKId_TBProduto " & _
                 "INNER JOIN TBTanque " & _
                 "ON TBTanque.PKId_TBTanque = TBBomba_bico.FKId_TBTanque " & _
                 "INNER JOIN TBBomba " & _
                 "ON TBBomba.PKId_TBBomba = TBBomba_bico.FKId_TBBomba " & _
                 "WHERE FKId_TBBomba = " & strID & ""
        
       Movimentacoes.Movimenta_HFlex_Grid strSql, hfgBomba_Bico, strTamanho_bico, strNomes_bico, "BDRetaguarda", "Otica", Me
        
       hfgBomba_Bico.Col = 1
       hfgBomba_Bico.Row = 1
       If hfgBomba_Bico.Text = Empty Then
          hfgBomba_Bico.Rows = 2
          Movimentacoes.Monta_HFlex_Grid hfgBomba_Bico, strTamanho_bico, strNomes_bico, 9, "Otica", Me
       Else
          intContador = 1
          hfgBomba_Bico.Col = 2
           
          Dim rstTipo_Preco As New ADODB.Recordset
           
          strSql = "SELECT DFNome_Preco_avista_TBTipo_preco,DFNome_Preco_promocao_TBTipo_preco," & _
                   "DFNome_Preco_revenda_TBTipo_preco,DFNome_Preco_especial_TBTipo_preco," & _
                   "DFNome_Preco_varejo_TBTipo_preco FROM TBTipo_preco "
          
          Select_geral strSql, "BDRetaguarda", rstTipo_Preco, "Otica", Me
                 
           Do While intContador <= hfgBomba_Bico.Rows - 1
              hfgBomba_Bico.Row = intContador
              
              If hfgBomba_Bico.Text = 1 Then
              
                 hfgBomba_Bico.Text = "1 - " & rstTipo_Preco("DFNome_Preco_avista_TBTipo_preco") & ""
                 
              ElseIf hfgBomba_Bico.Text = 2 Then

                 hfgBomba_Bico.Text = "2 - " & rstTipo_Preco("DFNome_Preco_promocao_TBTipo_preco") & ""
                 
              ElseIf hfgBomba_Bico.Text = 3 Then

                 hfgBomba_Bico.Text = "3 - " & rstTipo_Preco("DFNome_Preco_revenda_TBTipo_preco") & ""
                 
              ElseIf hfgBomba_Bico.Text = 4 Then

                 hfgBomba_Bico.Text = "4 - " & rstTipo_Preco("DFNome_Preco_especial_TBTipo_preco") & ""
                 
              ElseIf hfgBomba_Bico.Text = 5 Then

                 hfgBomba_Bico.Text = "5 - " & rstTipo_Preco("DFNome_Preco_varejo_TBTipo_preco") & ""
                 
              End If
              intContador = intContador + 1
            Loop
            Set rstTipo_Preco = Nothing
       End If
       
       booAlterar = True
       txtConsulta.Text = Empty
       txtCodigo_Bomba.Enabled = False
       txtDescricao_Bomba.Enabled = False
       txtNumero_Bicos.Enabled = False

       sstBomba.TabEnabled(0) = True
       sstBomba.Tab = 0
       txtCodigo_Bico.SetFocus
   End If
   
   Unload frmAguarde
   
End Sub

Private Sub hfgBomba_DblClick()
    hfgBomba.Sort = 1
End Sub

Private Sub hfgBomba_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgBomba_Click
    End If
End Sub

Private Sub sstBomba_Click(PreviousTab As Integer)
    If sstBomba.Tab = 0 Then
       If txtCodigo_Bomba.Enabled = False Then
          txtCodigo_Bico.SetFocus
       Else
          txtCodigo_Bomba.SetFocus
       End If
    End If
    If sstBomba.Tab = 1 Then
        If frmIntegracao.Visible = True Then
           Unload frmIntegracao
        End If
        If strCombo <> Empty And strCombo <> "Todos" Then
           cbbCampos.Text = strCombo
           txtConsulta.SetFocus
        ElseIf strCombo = "Todos" Then
           hfgBomba.Row = 1
           hfgBomba.Col = 0
           hfgBomba.SetFocus
        End If
    End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstBomba.Tab <> 1: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
    End Select
End Sub

Function Gravar()
    On Error GoTo erro
    
    Dim strCampos As String
    Dim strValores As String
    Dim strUltimo As String
    Dim strTipo_Preco As Integer
    Dim strID_Produto As String
    Dim strID_Tanque As String
    Dim strNumero As String
    
    Call Objetos.Retira_Espaco_Lateral(Me)
    Call Objetos.Maiusculo_TXT(Me)
    
    If txtCodigo_Bomba.Text = Empty Then
       MsgBox "O Código da Bomba não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtCodigo_Bomba.SetFocus
       Exit Function
    End If
    
    hfgBomba_Bico.Col = 0
    hfgBomba_Bico.Row = 1
    If hfgBomba_Bico.Text = Empty Then
       MsgBox "Não há Bico a ser cadastrado. Verifique.", vbInformation, "Only Tech"
       Exit Function
    End If
    
    strCampos = "IXCodigo_Bomba,DFDescricao_TBBomba,DFNumero_bicos_TBBomba,IXCodigo_TBEmpresa," & _
                "DFData_alteracao_TBBomba,DFIntegrado_filiais_TBBomba"
                
    If booIntegra_Portal = True Then
       strCampos = strCampos & ",DFIntegrado_portal_TBBomba"
    End If
    
    strValores = "" & txtCodigo_Bomba.Text & ",'" & txtDescricao_Bomba.Text & "','" & txtNumero_Bicos.Text & "'," & _
                 "" & dtcEmpresa.BoundText & ",'" & Format(Date, "YYYYMMDD") & "', 0"
                 
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
    End If

    If booAlterar = True Then
     
     On Error GoTo Erro_alteracao

       'abrindo conexao
       Conexao.Abrir_conexao "Otica"
       Conexao.CNconexao.BeginTrans
       
       log.Evento = "Alterar"
       strSql = "UPDATE TBBomba " & _
                "SET DFDescricao_TBBomba = '" & txtDescricao_Bomba.Text & "'," & _
                "    DFNumero_bicos_TBBomba  = '" & txtNumero_Bicos.Text & "', " & _
                "    DFData_alteracao_TBBomba = '" & Format(Date, "YYYYMMDD") & "'," & _
                "    DFIntegrado_filiais_TBBomba = 0 "
                
       If booIntegra_Portal = True Then
          strSql = strSql & ",DFIntegrado_portal_TBBomba = 0"
       End If
       
       strSql = strSql & "    WHERE PKId_TBBomba = " & strID & ""
    
       Conexao.CNconexao.Execute strSql

       intContador = 1

       Do While intContador <= hfgBomba_Bico.Rows - 1
            
            hfgBomba_Bico.Row = intContador
            
            hfgBomba_Bico.Col = 1
            strCodigo = hfgBomba_Bico.Text
            
            'Capturando número do tipo de preço
            hfgBomba_Bico.Col = 2
            If hfgBomba_Bico.Text <> Empty Then
               strTipo_Preco = Left(hfgBomba_Bico.Text, 1)
            End If
            
            hfgBomba_Bico.Col = 3
            strUltimo = hfgBomba_Bico.Text
            
            hfgBomba_Bico.Col = 5
            strID_Produto = Funcoes_Gerais.Localiza_ID("PKId_TBProduto", "IXCodigo_TBProduto", hfgBomba_Bico.Text, "TBProduto", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcEmpresa.BoundText)
            
            hfgBomba_Bico.Col = 7
            strID_Tanque = Funcoes_Gerais.Localiza_ID("PKId_TBTanque", "IXCodigo_TBTanque", hfgBomba_Bico.Text, "TBTanque", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcEmpresa.BoundText)

            hfgBomba_Bico.Col = 4
            strNumero = hfgBomba_Bico.Text
            
            hfgBomba_Bico.Col = 9
            If hfgBomba_Bico.Text <> Empty Then

                strSql = "UPDATE TBBomba_bico " & _
                         "SET FKId_TBProduto = '" & strID_Produto & "'," & _
                         "FKId_TBTanque = '" & strID_Tanque & "'," & _
                         "DFUltimo_encerrante_TBBomba_bico = " & Funcoes_Gerais.Grava_Moeda(strUltimo) & "," & _
                         "DFNumero_maximo_encerrante_TBBomba_bico = " & Funcoes_Gerais.Grava_Moeda(strNumero) & "," & _
                         "DFTipo_preco_TBBomba_bico = " & strTipo_Preco & ", " & _
                         "DFData_alteracao_TBBomba_bico = '" & Format(Date, "YYYYMMDD") & "'," & _
                         "DFIntegrado_filiais_TBBomba_bico = 0"
                         
                If booIntegra_Portal = True Then
                   strSql = strSql & ",DFIntegrado_portal_TBBomba_bico = 0 "
                End If
                
                strSql = strSql & "WHERE PKId_TBBomba_bico = " & hfgBomba_Bico.Text & ""
                
                Conexao.CNconexao.Execute strSql
  
             ElseIf hfgBomba_Bico.Text = Empty Then
            
                  strSql = "INSERT INTO TBBomba_bico (FKId_TBProduto,FKId_TBTanque,IXCodigo_TBBomba_bico," & _
                           "DFUltimo_encerrante_TBBomba_bico,DFNumero_maximo_encerrante_TBBomba_bico," & _
                           "DFTipo_preco_TBBomba_bico,FKId_TBBomba,DFData_alteracao_TBBomba_bico," & _
                           "DFIntegrado_filiais_TBBomba_bico"
                           
                  If booIntegra_Portal = True Then
                     strSql = strSql & ",DFIntegrado_portal_TBBomba_bico) "
                  Else
                     strSql = strSql & ") "
                  End If
                           
                  strSql = strSql & "SELECT '" & strID_Produto & "','" & strID_Tanque & "'," & _
                                    "" & strCodigo & "," & Funcoes_Gerais.Grava_Moeda(strUltimo) & "," & Funcoes_Gerais.Grava_Moeda(strNumero) & "," & _
                                    "" & strTipo_Preco & "," & strID & ",'" & Format(Date, "YYYYMMDD") & "',0"
                                    
                  If booIntegra_Portal = True Then
                     strSql = strSql & ",0) "
                  Else
                     strSql = strSql & ") "
                  End If
                  
                  Conexao.CNconexao.Execute strSql
                  
              End If
              
            intContador = intContador + 1
         Loop
       
        'Deletando registros antes da nova gravacao
        If strId_remover <> Empty Then
           strSql = "DELETE FROM TBBomba_bico WHERE FKId_TBBomba = " & strID & " " & _
                    "AND PKId_TBBomba_bico IN (" & strId_remover & ")"
           Conexao.CNconexao.Execute strSql
           strId_remover = Empty
        End If
           
       'fechando conexao
        Conexao.CNconexao.CommitTrans
        Conexao.Fechar_conexao
        
        log.Descricao = "Alterando o registro: " + " & txtCodigo_Bomba.Text & "
        log.Tipo = 1
        log.Hora = Format(Now, "hh:mm:ss")
        'Gravando log
        log.Gravar_log "OTICA", Me
    Else
       On Error GoTo Erro_inclusao
       
       
       'abrindo conexao
       Conexao.Abrir_conexao "Otica"
       Conexao.CNconexao.BeginTrans
       
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBBomba", strCampos, strValores, "Otica", Me, "BDRetaguarda")
       
       'localizando a ID gravada
       Dim strId_novo As String
       
       strSql = "SELECT MAX(PKId_TBBomba) as IdBomba FROM TBBomba"
    
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstAplicacao, "OTICA", Me)
       strId_novo = rstAplicacao.Fields("IdBomba")
     
       Set rstAplicacao = Nothing
       
       intContador = 1

       Do While intContador <= hfgBomba_Bico.Rows - 1
            
            hfgBomba_Bico.Row = intContador
            
            hfgBomba_Bico.Col = 1
            strCodigo = hfgBomba_Bico.Text
            
            'Capturando número do tipo de preço
            hfgBomba_Bico.Col = 2
            If hfgBomba_Bico.Text <> Empty Then
               strTipo_Preco = Left(hfgBomba_Bico.Text, 1)
            End If
            
            hfgBomba_Bico.Col = 3
            strUltimo = hfgBomba_Bico.Text
            
            hfgBomba_Bico.Col = 5
            strID_Produto = Funcoes_Gerais.Localiza_ID("PKId_TBProduto", "IXCodigo_TBProduto", hfgBomba_Bico.Text, "TBProduto", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcEmpresa.BoundText)
            
            hfgBomba_Bico.Col = 7
            strID_Tanque = Funcoes_Gerais.Localiza_ID("PKId_TBTanque", "IXCodigo_TBTanque", hfgBomba_Bico.Text, "TBTanque", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcEmpresa.BoundText)

            hfgBomba_Bico.Col = 4
            
            strSql = Empty
            strSql = "INSERT INTO TBBomba_bico (FKId_TBProduto,FKId_TBTanque,IXCodigo_TBBomba_bico," & _
                     "DFUltimo_encerrante_TBBomba_bico,DFNumero_maximo_encerrante_TBBomba_bico," & _
                     "DFTipo_preco_TBBomba_bico,FKId_TBBomba,DFData_alteracao_TBBomba_bico," & _
                     "DFIntegrado_filiais_TBBomba_bico"
                     
            If booIntegra_Portal = True Then
               strSql = strSql & ",DFIntegrado_portal_TBBomba_bico) "
            Else
               strSql = strSql & ") "
            End If
                     
            strSql = strSql & "SELECT '" & strID_Produto & "','" & strID_Tanque & "'," & _
                              "" & strCodigo & "," & Funcoes_Gerais.Grava_Moeda(strUltimo) & "," & Funcoes_Gerais.Grava_Moeda(hfgBomba_Bico.Text) & "," & _
                              "" & strTipo_Preco & "," & strId_novo & ",'" & Format(Date, "YYYYMMDD") & "',0"
                              
            If booIntegra_Portal = True Then
               strSql = strSql & ",0 "
            End If
            
            Conexao.CNconexao.Execute strSql

            intContador = intContador + 1
         Loop
        
       'fechando conexao
        Conexao.CNconexao.CommitTrans
        Conexao.Fechar_conexao
         
       log.Descricao = "Gravando o registro: " + " & txtCodigo_Bomba.Text & "
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
       hfgBomba.Visible = False
    End If
    
    sstBomba.TabEnabled(0) = False
    sstBomba.Tab = 1
    
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
    
    Call funcoes_banco.Excluir("TBBomba", "PKId_TBBomba", strId_novo, "Otica", Me, "BDRetaguarda")
    Call erro.erro(Me, "Otica", "Gravar")
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    Dim intContador As Integer
    
    On Error GoTo erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtCodigo_Bomba.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    
    'VERIFICA SE HÁ BICOS ATRELADOS A ENCERRANTES
    ''''''''''''''''''''''''''''''''''''''''''''''
    intContador = 1
    Do While intContador <= hfgBomba_Bico.Rows - 1
          'Verifica se o Bico está vinculado a um Encerrante
          hfgBomba_Bico.Row = intContador
          hfgBomba_Bico.Col = 1
          strSql = Empty
          strSql = "SELECT FKId_TBBomba_bico " & _
                   "FROM TBEncerrante_Bomba " & _
                   "INNER JOIN TBBomba_bico " & _
                   "ON TBEncerrante_Bomba.FKId_TBBomba_bico = TBBomba_bico.PKId_TBBomba_bico " & _
                   "WHERE IXCodigo_TBBomba_bico = '" & hfgBomba_Bico.Text & "'"
        
          Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstVerifica_Titulo, "OTICA", Me)
        
          If rstVerifica_Titulo.RecordCount <> 0 Then
             MsgBox "Há Bicos dessa bomba que encontram-se vinculados a um Encerrante e não podem ser excluídos. Verifique.", vbInformation, "Only Tech"
             Set rstVerifica_Titulo = Nothing
             Exit Function
          End If
          Set rstVerifica_Titulo = Nothing
          intContador = intContador + 1
     Loop
      
    'Iniciando conexao
    Conexao.Initial_Catalog = "BDRetaguarda"
    Conexao.Abrir_conexao ("Otica")
    
    Conexao.CNconexao.BeginTrans
    
    strSql = "DELETE FROM TBBomba_bico WHERE FKId_TBBomba = " & strID & ""
    
    Conexao.CNconexao.Execute strSql
    
    strSql = "DELETE FROM TBBomba WHERE PKId_TBBomba = " & strID & ""
    
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
       hfgBomba.Visible = False
    End If
        
    sstBomba.TabEnabled(0) = False
    sstBomba.Tab = 1
    
    Exit Function
    
erro:
    Conexao.CNconexao.RollbackTrans
    Conexao.Fechar_conexao
    
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
       hfgBomba.Visible = False
    End If
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstBomba.TabEnabled(0) = False
    sstBomba.Tab = 1
    txtCodigo_Bomba.Enabled = False
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo erro
    
    Dim rstBusca_Parametro As New ADODB.Recordset
    
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
    
    cbbTipo_Preco.Text = Empty
    txtCodigo_Bomba.Enabled = True
    txtDescricao_Bomba.Enabled = True
    txtNumero_Bicos.Enabled = True
    
    hfgBomba_Bico.Rows = 2
    Movimentacoes.Monta_HFlex_Grid hfgBomba_Bico, strTamanho_bico, strNomes_bico, 9, "Otica", Me
    cmdIncluir.Caption = "Incluir"
    
    txtCodigo_Bomba.SetFocus
    
    dtcEmpresa.BoundText = MDIPrincipal.ocxUsuario.Empresa
    
    strSql = "SELECT IXCodigo_TBTanque,DFDescricao_TBTanque FROM TBTanque WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBTanque", "DFDescricao_TBTanque", dtcTanque, strSql, "BDRetaguarda", "Otica", Me

    strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me

    booAlterar = False
    
    sstBomba.TabEnabled(0) = True
    sstBomba.Tab = 0
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    On Error GoTo erro
    
    strNomes = "Código,Descrição,Número de Bicos,Empresa,Nome,ID_Bomba"
    strTamanho = "1600,3500,1900,1500,2500,0"
    
    Movimentacoes.Monta_HFlex_Grid hfgBomba, strTamanho, strNomes, 6, "Otica", Me

    strNomes_bico = "Código,Tipo de Preço,Último Encerrante,Nº Max. Encerrante,Cod. Produto,Produto,Cod. Tanque,Tanque,ID"
    strTamanho_bico = "1000,2200,1800,1800,1500,2900,1500,2200,0"
    
    Movimentacoes.Monta_HFlex_Grid hfgBomba_Bico, strTamanho_bico, strNomes_bico, 9, "Otica", Me
    
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    dtcEmpresa.BoundText = MDIPrincipal.ocxUsuario.Empresa
    
    strSql = "SELECT IXCodigo_TBTanque,DFDescricao_TBTanque FROM TBTanque WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBTanque", "DFDescricao_TBTanque", dtcTanque, strSql, "BDRetaguarda", "Otica", Me

    strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me

    Call Monta_Combo
          
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Reposicao")
    Resume Next
End Function

Private Sub txtCodigo_Bico_LostFocus()
    If txtCodigo_Bico.Text <> Empty Then
        'VERIFICANDO EXISTENCIA DO REGISTRO NO BANCO PARA EMPRESA
           strSql = "SELECT IXCodigo_TBBomba_bico,IXCodigo_Bomba " & _
                    "FROM TBBomba_bico " & _
                    "INNER JOIN TBBomba ON TBBomba_bico.FKId_TBBomba = TBBomba.PKId_TBBomba " & _
                    "WHERE IXCodigo_TBBomba_bico = '" & txtCodigo_Bico.Text & "' " & _
                    "AND TBBomba.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
           
           Movimentacoes.Select_geral strSql, "BDRetaguarda", rstVerifica_Titulo, "Otica", Me
           
           If rstVerifica_Titulo.RecordCount <> 0 And IXCodigo_Bomba Then
             
                 MsgBox "Este registro já existe no banco para esta Empresa. Verifique.", vbInformation, "Only Tech"
                 Set rstVerifica_Titulo = Nothing
                 txtCodigo_Bico.SetFocus
                 Exit Sub

           End If
           Set rstVerifica_Titulo = Nothing
    End If
End Sub

Private Sub txtCodigo_Bomba_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Bomba_LostFocus()
    If txtCodigo_Bomba.Text <> Empty Then
        'VERIFICANDO EXISTENCIA DO REGISTRO NO BANCO PARA EMPRESA
        If booAlterar = False Then
           strSql = "SELECT IXCodigo_Bomba " & _
                    "FROM TBBomba " & _
                    "WHERE IXCodigo_Bomba = '" & txtCodigo_Bomba.Text & "' " & _
                    "AND TBBomba.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
           
           Movimentacoes.Select_geral strSql, "BDRetaguarda", rstVerifica_Titulo, "Otica", Me
           
           If rstVerifica_Titulo.RecordCount <> 0 Then
              MsgBox "Este registro já existe no banco para esta Empresa. Verifique.", vbInformation, "Only Tech"
              Set rstVerifica_Titulo = Nothing
              txtCodigo_Bomba.SetFocus
              Exit Sub
           End If
           Set rstVerifica_Titulo = Nothing
        End If
    End If
End Sub

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Sub txtCodigo_Bomba_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_Bico_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_Bico_GotFocus()
    If txtCodigo_Bomba.Text = Empty And sstBomba.Tab = 0 Then
       MsgBox "Bomba não definida. Verifique.", vbInformation, "Only Tech"
       txtCodigo_Bomba.SetFocus
    End If
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDescricao_Bomba_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_Bicos_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_Bicos_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtProduto_Change()
    dtcProduto.BoundText = txtProduto.Text
    If IsNumeric(txtProduto.Text) = False Then txtProduto.Text = Empty: Exit Sub
End Sub

Private Sub txtProduto_GotFocus()
    If txtTanque.Text = Empty Or dtcTanque.Text = Empty Then
       MsgBox "Tanque não definido. Verifique.", vbInformation, "Only Tech"
       txtTanque.SetFocus
       Exit Sub
    End If
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtProduto_LostFocus()
    If txtProduto.Text <> Empty Then
       Dim rstItem_verifica As New ADODB.Recordset
       Dim intID_Produto As String
        
       intID_Produto = Funcoes_Gerais.Localiza_ID("PKID_TBProduto", "IXCodigo_TBProduto", txtProduto.Text, "TBProduto", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcEmpresa.BoundText)
            
       strSql = Empty
       strSql = "SELECT FKId_TBProduto FROM TBItens_tabela_preco WHERE FKID_TBProduto = " & intID_Produto & ""
    
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstItem_verifica, "Otica", Me
        
       If rstItem_verifica.RecordCount = 0 Then
          MsgBox "Produto não cadastrado na tabela vigente. Verifique.", vbInformation, "Only Tech"
          txtProduto.SetFocus
       End If
    End If
End Sub

Private Sub txtTanque_Change()
    dtcTanque.BoundText = txtTanque.Text
    If IsNumeric(txtTanque.Text) = False Then
       txtTanque.Text = Empty
       Exit Sub
    End If
End Sub

Private Sub txtTanque_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtTanque_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTanque_LostFocus()
    If dtcTanque.Text = Empty Then
       txtTanque.Text = Empty
       txtCapacidade.Text = Empty
    Else
       strSql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto " & _
                "INNER JOIN TBEmpresa ON TBProduto.IXCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
                "INNER JOIN TBTanque ON TBTanque.IXCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
                "WHERE TBTanque.IXCodigo_TBTanque = " & txtTanque.Text & ""
                
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me
       
       strSql = "SELECT DFCapacidade_TBTanque FROM TBTanque " & _
                "WHERE TBTanque.IXCodigo_TBTanque = " & txtTanque.Text & ""
       Select_geral strSql, "BDRetaguarda", rstVerifica_Titulo, "Otica", Me
       
       If Not IsNull(rstVerifica_Titulo("DFCapacidade_TBTanque")) And rstVerifica_Titulo.RecordCount <> 0 Then
          Me.txtCapacidade.Text = rstVerifica_Titulo("DFCapacidade_TBTanque")
       End If
       Set rstVerifica_Titulo = Nothing
    End If
End Sub

Private Sub txtUltimo_Encerrante_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtUltimo_Encerrante_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtNumero_Maximo_Encerrante_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_Maximo_Encerrante_KeyPress(KeyAscii As Integer)
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
    cbbCampos.AddItem ("Nº de Bicos")
    
    Dim rstTipo_Preco As New ADODB.Recordset
    
    strSql = "SELECT * FROM TBTipo_preco"
    Select_geral strSql, "BDRetaguarda", rstTipo_Preco, "Otica", Me
    
    rstTipo_Preco.MoveFirst
    
    cbbTipo_Preco.Clear

    If rstTipo_Preco.RecordCount <> 0 Then
       cbbTipo_Preco.AddItem ("1 - " & rstTipo_Preco.Fields("DFNome_Preco_avista_TBTipo_preco") & "")
       cbbTipo_Preco.AddItem ("2 - " & rstTipo_Preco.Fields("DFNome_Preco_promocao_TBTipo_preco") & "")
       cbbTipo_Preco.AddItem ("3 - " & rstTipo_Preco.Fields("DFNome_Preco_revenda_TBTipo_preco") & "")
       cbbTipo_Preco.AddItem ("4 - " & rstTipo_Preco.Fields("DFNome_Preco_especial_TBTipo_preco") & "")
       cbbTipo_Preco.AddItem ("5 - " & rstTipo_Preco.Fields("DFNome_Preco_varejo_TBTipo_preco") & "")
    End If
    
End Function

Private Function Consulta()
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbInformation, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
    
    strSql = "SELECT IXCodigo_Bomba," & _
             "DFDescricao_TBBomba," & _
             "DFNumero_bicos_TBBomba,IXCodigo_TBEmpresa,DFRazao_Social_TBEmpresa,PKId_TBBomba " & _
             "FROM TBBomba " & _
             "INNER JOIN TBEmpresa ON TBBomba.IXCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa"

                     
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código" Then
          strSql = strSql & " WHERE IXCodigo_Bomba = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Descrição" Then
          strSql = strSql & " WHERE convert(nvarchar,DFDescricao_TBBomba) LIKE '%" & txtConsulta.Text & "%'"
       ElseIf cbbCampos.Text = "Nº de Bicos" Then
          strSql = strSql & " WHERE DFNumero_bicos_TBBomba = '" & txtConsulta.Text & "'"
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
    
    strSql = strSql & " ORDER BY IXCodigo_Bomba"
       
    frmAguarde.Show
    DoEvents

    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgBomba, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    hfgBomba.Row = 1
    hfgBomba.Col = 0
    If hfgBomba.Text = Empty Then
       hfgBomba.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgBomba, strTamanho, strNomes, 6, "Otica", Me
    End If
    
    hfgBomba.Col = 0
    hfgBomba.Row = 1
    
    Unload frmAguarde
    hfgBomba.SetFocus
    
End Function

Private Sub txtUltimo_Encerrante_LostFocus()
    If strCasas_Decimais = 2 Then
       txtUltimo_Encerrante.Text = Format(txtUltimo_Encerrante.Text, "#,###0.00")
    ElseIf strCasas_Decimais = 3 Then
       txtUltimo_Encerrante.Text = Format(txtUltimo_Encerrante.Text, "#,###0.000")
    Else
       txtUltimo_Encerrante.Text = Format(txtUltimo_Encerrante.Text, "#,##0.00")
    End If
End Sub

Private Sub txtNumero_Maximo_Encerrante_LostFocus()
    If strCasas_Decimais = 2 Then
       txtNumero_Maximo_Encerrante.Text = Format(txtNumero_Maximo_Encerrante.Text, "#,###0.00")
    ElseIf strCasas_Decimais = 3 Then
       txtNumero_Maximo_Encerrante.Text = Format(txtNumero_Maximo_Encerrante.Text, "#,###0.000")
    Else
       txtNumero_Maximo_Encerrante.Text = Format(txtNumero_Maximo_Encerrante.Text, "#,##0.00")
    End If
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("IXCodigo_Bomba", txtCodigo_Bomba.Text, "DFIntegrado_filiais_TBBomba", "TBBomba", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBBomba", Me.Top, Me.Left, Me.width, Me.Height, "Bomba")
    
End Function
