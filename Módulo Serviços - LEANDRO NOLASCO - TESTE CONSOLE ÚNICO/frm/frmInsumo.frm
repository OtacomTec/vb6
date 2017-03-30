VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmInsumo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insumo"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsumo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8460
   Begin TabDlg.SSTab sstInsumo 
      Height          =   6525
      Left            =   0
      TabIndex        =   21
      Top             =   330
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   11509
      _Version        =   393216
      Tab             =   1
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
      TabPicture(0)   =   "frmInsumo.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtTipo_Marcha"
      Tab(0).Control(1)=   "txtTecnica_Aplicada"
      Tab(0).Control(2)=   "txtFuncao"
      Tab(0).Control(3)=   "txtObservacao"
      Tab(0).Control(4)=   "txtConservacao"
      Tab(0).Control(5)=   "txtNome_cientifico"
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(7)=   "txtDescricao"
      Tab(0).Control(8)=   "txtCodigo"
      Tab(0).Control(9)=   "dtcFuncao"
      Tab(0).Control(10)=   "dtcTipo_Marcha"
      Tab(0).Control(11)=   "Label15"
      Tab(0).Control(12)=   "shpIntegrado"
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(14)=   "Label11"
      Tab(0).Control(15)=   "Label3"
      Tab(0).Control(16)=   "Label10"
      Tab(0).Control(17)=   "Label9"
      Tab(0).Control(18)=   "Label2"
      Tab(0).Control(19)=   "Label1"
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "&Análises"
      TabPicture(1)   =   "frmInsumo.frx":179E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgAnalise"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtAnalise"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtEspecificacao"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdSalvar"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdRemover"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdAlterar"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdConsulta_Insumo"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtNumero"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtPotencia"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "&Listagem"
      TabPicture(2)   =   "frmInsumo.frx":17BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "cbbIntergra_portal"
      Tab(2).Control(2)=   "cbbCampos"
      Tab(2).Control(3)=   "hfgInsumo"
      Tab(2).Control(4)=   "cmdConsulta"
      Tab(2).Control(5)=   "cmdRefresh"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtConsulta"
      Tab(2).ControlCount=   7
      Begin VB.CommandButton Command1 
         Caption         =   "Aplicar"
         Height          =   285
         Left            =   4950
         TabIndex        =   49
         Top             =   5910
         Width           =   735
      End
      Begin VB.TextBox txtPotencia 
         Height          =   375
         Left            =   3030
         TabIndex        =   48
         Top             =   5790
         Width           =   1785
      End
      Begin VB.TextBox txtNumero 
         Height          =   360
         Left            =   900
         TabIndex        =   47
         Top             =   5790
         Width           =   1485
      End
      Begin VB.TextBox txtTipo_Marcha 
         Height          =   360
         Left            =   -74880
         TabIndex        =   4
         ToolTipText     =   "Código da Função"
         Top             =   2070
         Width           =   1360
      End
      Begin VB.CommandButton cmdConsulta_Insumo 
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
         Left            =   7920
         Picture         =   "frmInsumo.frx":17D6
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Consulta Detalhada de Insumo"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtTecnica_Aplicada 
         Height          =   375
         Left            =   -70830
         MaxLength       =   800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         ToolTipText     =   "Técnica Aplicada"
         Top             =   5880
         Width           =   4095
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "Alterar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7170
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Alterar"
         Top             =   1800
         Width           =   1155
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
         Height          =   350
         Left            =   7170
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Remover"
         Top             =   2160
         Width           =   1155
      End
      Begin VB.CommandButton cmdSalvar 
         Cancel          =   -1  'True
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
         Height          =   350
         Left            =   7170
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Incluir"
         Top             =   1440
         Width           =   1155
      End
      Begin VB.TextBox txtEspecificacao 
         Height          =   1065
         Left            =   120
         MaxLength       =   800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1440
         Width           =   6975
      End
      Begin VB.TextBox txtAnalise 
         Height          =   360
         Left            =   120
         MaxLength       =   80
         TabIndex        =   15
         ToolTipText     =   "Descrição Insumo"
         Top             =   780
         Width           =   7755
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72510
         TabIndex        =   37
         Top             =   780
         Width           =   5055
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   -67020
         Picture         =   "frmInsumo.frx":1B60
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   -67410
         Picture         =   "frmInsumo.frx":2BA2
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtFuncao 
         Height          =   360
         Left            =   -74880
         TabIndex        =   2
         ToolTipText     =   "Código da Função"
         Top             =   1440
         Width           =   1360
      End
      Begin VB.TextBox txtObservacao 
         Height          =   375
         Left            =   -74880
         MaxLength       =   300
         TabIndex        =   13
         ToolTipText     =   "Observação"
         Top             =   5880
         Width           =   4005
      End
      Begin VB.TextBox txtConservacao 
         Height          =   375
         Left            =   -70830
         MaxLength       =   100
         TabIndex        =   7
         ToolTipText     =   "Conservação Insumo"
         Top             =   2730
         Width           =   4095
      End
      Begin VB.TextBox txtNome_cientifico 
         Height          =   375
         Left            =   -74880
         MaxLength       =   100
         TabIndex        =   6
         ToolTipText     =   "Nome Científico Insumo"
         Top             =   2730
         Width           =   4005
      End
      Begin VB.Frame Frame1 
         Caption         =   "Referências Bibliográficas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74850
         TabIndex        =   25
         Top             =   3180
         Width           =   8115
         Begin VB.TextBox txtReferencia5 
            Height          =   375
            Left            =   120
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            ToolTipText     =   "Referência 5 Insumo"
            Top             =   1890
            Width           =   7905
         End
         Begin VB.TextBox txtReferencia4 
            Height          =   375
            Left            =   4020
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            ToolTipText     =   "Referência 4 Insumo"
            Top             =   1230
            Width           =   4005
         End
         Begin VB.TextBox txtReferencia2 
            Height          =   375
            Left            =   4020
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            ToolTipText     =   "Referência 2 Insumo"
            Top             =   570
            Width           =   4005
         End
         Begin VB.TextBox txtReferencia3 
            Height          =   375
            Left            =   120
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            ToolTipText     =   "Referência 3 Insumo"
            Top             =   1230
            Width           =   3855
         End
         Begin VB.TextBox txtReferencia1 
            Height          =   375
            Left            =   120
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            ToolTipText     =   "Referência 1 Insumo"
            Top             =   570
            Width           =   3855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Referência 5"
            Height          =   240
            Left            =   120
            TabIndex        =   30
            Top             =   1650
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Referência 4"
            Height          =   240
            Left            =   4020
            TabIndex        =   29
            Top             =   990
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Referência 2"
            Height          =   240
            Left            =   4020
            TabIndex        =   28
            Top             =   330
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Referência 3"
            Height          =   240
            Left            =   120
            TabIndex        =   27
            Top             =   990
            Width           =   1080
         End
         Begin VB.Label lblObservacao 
            AutoSize        =   -1  'True
            Caption         =   "Referência 1"
            Height          =   240
            Left            =   120
            TabIndex        =   26
            Top             =   330
            Width           =   1080
         End
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   -73470
         MaxLength       =   100
         TabIndex        =   1
         ToolTipText     =   "Descrição Insumo"
         Top             =   780
         Width           =   6735
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   0
         ToolTipText     =   "Código Insumo"
         Top             =   780
         Width           =   1360
      End
      Begin MSDataListLib.DataCombo dtcFuncao 
         Height          =   360
         Left            =   -73470
         TabIndex        =   3
         ToolTipText     =   "Descrição da Função"
         Top             =   1440
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   8388608
         Text            =   ""
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgInsumo 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   38
         Top             =   1230
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   9128
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   -74880
         TabIndex        =   39
         Top             =   780
         Width           =   2325
         _ExtentX        =   4101
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgAnalise 
         Height          =   2985
         Left            =   120
         TabIndex        =   20
         Top             =   2610
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   5265
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin AutoCompletar.CbCompleta cbbIntergra_portal 
         Height          =   360
         Left            =   -72510
         TabIndex        =   44
         Top             =   780
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
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
      Begin MSDataListLib.DataCombo dtcTipo_Marcha 
         Height          =   360
         Left            =   -73470
         TabIndex        =   5
         ToolTipText     =   "Descrição da Função"
         Top             =   2070
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   8388608
         Text            =   ""
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Marcha"
         Height          =   240
         Left            =   -74880
         TabIndex        =   46
         Top             =   1830
         Width           =   1065
      End
      Begin VB.Shape shpIntegrado 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   165
         Left            =   -66780
         Shape           =   3  'Circle
         Top             =   6300
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Técnica Aplicada"
         Height          =   240
         Left            =   -70830
         TabIndex        =   43
         Top             =   5640
         Width           =   1440
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Especificação"
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Análise"
         Height          =   240
         Left            =   120
         TabIndex        =   41
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   40
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Função"
         Height          =   240
         Left            =   -74880
         TabIndex        =   34
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   240
         Left            =   -74880
         TabIndex        =   33
         Top             =   5640
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Conservação"
         Height          =   240
         Left            =   -70830
         TabIndex        =   32
         Top             =   2490
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nome Científico"
         Height          =   240
         Left            =   -74880
         TabIndex        =   31
         Top             =   2490
         Width           =   1335
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
         Left            =   -73470
         TabIndex        =   23
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   -74880
         TabIndex        =   22
         Top             =   540
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
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
      Left            =   10440
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
            Picture         =   "frmInsumo.frx":489C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsumo.frx":4BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsumo.frx":4ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsumo.frx":526A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsumo.frx":5604
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsumo.frx":591E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsumo.frx":5C38
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro Insumo                                                '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Peixoto                                                  '
' Data de Criação........: 04/03/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Public strSql As String
Dim strId As String
Dim booAlterar As Boolean
Dim conexao As New DLLConexao_Sistema.conexao
Dim log As New DLLSystemManager.log
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim strTamanho As String
Dim strNomes As String
Dim intClique_Analise As Integer
Dim intClique_Especificacao As Integer
Dim booIntegra_Portal As Boolean
Dim booIntegracao As Boolean
Dim intContador As Integer
Dim strAnalise_Antiga As String
    
Function Imprimir()
    On Error GoTo Erro
    'Tratamento de Erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Insumo.Show
    
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
       Me.cbbIntergra_portal.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    Else
       txtConsulta.Visible = True
       Me.cbbIntergra_portal.Visible = False
       txtConsulta.SetFocus
    End If
    
    If Me.cbbCampos.Text = "Integrado Portal" Then
       txtConsulta.Visible = False
       Me.cbbIntergra_portal.Visible = True
       Me.cbbIntergra_portal.AddItem ("Sim")
       Me.cbbIntergra_portal.AddItem ("Não")
    End If
    
End Sub

Private Sub cmdAlterar_Click()
    Dim strVerifica As String
    
    hfgAnalise.Col = 1
    hfgAnalise.Row = 1
    If hfgAnalise.Text = Empty Then Exit Sub
    
    If intClique_Analise = Empty And intClique_Especificacao = Empty Then
       MsgBox "Dados para alteração não selecionados. Verifique.", vbInformation, "Only Tech"
       Exit Sub
    ElseIf intClique_Analise <> Empty And intClique_Especificacao <> Empty Then
       MsgBox "Dados para alteração inválidos. Verifique.", vbInformation, "Only Tech"
       Exit Sub
    End If
    
    If txtAnalise.Text = Empty And intClique_Analise <> Empty Then
       MsgBox "Análise inválida. Verifique.", vbInformation, "Only Tech"
       txtAnalise.SetFocus
       Exit Sub
    ElseIf txtEspecificacao.Text = Empty And intClique_Especificacao <> Empty Then
       MsgBox "Especificação inválida. Verifique.", vbInformation, "Only Tech"
       txtEspecificacao.SetFocus
       Exit Sub
    End If

    intContador = 1
    Do While intContador <= hfgAnalise.Rows - 1
       hfgAnalise.Row = intContador
       
       If intClique_Analise <> Empty Then
          hfgAnalise.Col = 1
          If hfgAnalise.Text = txtAnalise.Text And hfgAnalise.Row <> intClique_Analise And strAnalise_Antiga <> hfgAnalise.Text Then
             MsgBox "A Análise alterada pertence a outro item neste cadastro. Verifique.", vbInformation, "Only Tech"
             txtAnalise.SetFocus
             Exit Sub
          End If
       ElseIf intClique_Especificacao <> Empty Then
          hfgAnalise.Col = 1
          strVerifica = hfgAnalise.Text
          hfgAnalise.Col = 2
          If strVerifica = strAnalise_Antiga And hfgAnalise.Text = txtEspecificacao.Text And hfgAnalise.Row <> intClique_Especificacao Then
             MsgBox "A Especificação alterada pertence a outro item desta Análise. Verifique.", vbInformation, "Only Tech"
             txtEspecificacao.SetFocus
             Exit Sub
          End If
       End If
       
       intContador = intContador + 1
    Loop
    
    If intClique_Especificacao <> Empty Then
       intIndice = intClique_Especificacao
       'Os campos sao medidos e uma constante de proporcionalidade é usada para o tamanho das linhas
       If Len(txtEspecificacao.Text) > 60 Then
          hfgAnalise.RowHeight(intIndice) = 285 * CDbl((Len(txtEspecificacao.Text)) / 55)
          hfgAnalise.WordWrap = True
       End If
    ElseIf intClique_Analise <> Empty Then
       intIndice = intClique_Analise
       'Os campos sao medidos e uma constante de proporcionalidade é usada para o tamanho das linhas
       If Len(txtAnalise.Text) > 20 Then
          hfgAnalise.RowHeight(intIndice) = 285 * CDbl((Len(txtAnalise.Text)) / 10)
          hfgAnalise.WordWrap = True
       End If
    End If
    
    hfgAnalise.Row = intIndice
    
    hfgAnalise.Col = 0
    hfgAnalise.ColWidth(0) = 500
    hfgAnalise.Font.Name = "Tahoma"
    hfgAnalise.CellFontSize = 7
    hfgAnalise.CellFontBold = False
    hfgAnalise.CellBackColor = &H80FFFF
    hfgAnalise.Text = intIndice
    
    'Alterando o grupo de análises com o mesmo nome
    If intClique_Analise <> Empty Then
       intContador = 1
       hfgAnalise.Col = 1
       Do While intContador <= hfgAnalise.Rows - 1
          hfgAnalise.Row = intContador
          If hfgAnalise.Text = strAnalise_Antiga Then
             hfgAnalise.Text = txtAnalise.Text
          End If
          intContador = intContador + 1
       Loop
    End If
    
    If intClique_Especificacao <> Empty Then
       hfgAnalise.Col = 2
       hfgAnalise.Text = txtEspecificacao.Text
    End If
    
    txtEspecificacao.Text = Empty
    txtAnalise.Text = Empty
    intClique_Especificacao = Empty
    intClique_Analise = Empty
    txtAnalise.SetFocus
    
End Sub

Private Sub cmdConsulta_Insumo_Click()
    frmAguarde.Show
    DoEvents
    Unload frmInsumo_Consulta_Detalhada_Insumo
    frmInsumo_Consulta_Detalhada_Insumo.Show
    Unload frmAguarde
End Sub

Private Sub cmdRemover_Click()
    Dim strExclui_Analise As String
    Dim booEspecificacao_Analise As Boolean
    
    If (hfgAnalise.Col <> 1 And hfgAnalise.Col <> 2) Or hfgAnalise.Row = 0 Or hfgAnalise.Text = Empty Then
       MsgBox "Não há dados selecionados para exclusão. Verifique.", vbInformation, "Only Tech"
       Exit Sub
    Else
       If hfgAnalise.Col = 1 Then
          intResp = MsgBox("Deseja excluir esta análise e todas as especificações vinculadas a ela?", vbYesNo, "Only Tech")
          If intResp = 7 Then Exit Sub
          booAnalise = True
          booEspecificacao_Analise = True
          strExclui_Analise = hfgAnalise.Text
       ElseIf hfgAnalise.Col = 2 Then
          booEspecificacao_Analise = False
       End If
    End If

    If hfgAnalise.Rows <= 2 Then
       hfgAnalise.Clear
       Movimentacoes.Monta_HFlex_Grid hfgAnalise, "2050,5500", "Análises,Especificações", 2, "Otica", Me
       hfgAnalise.ColAlignmentFixed(1) = 4
       hfgAnalise.ColAlignmentFixed(2) = 4
       hfgAnalise.RowHeight(1) = 285
    Else
       If booEspecificacao_Analise = True Then
          '''ROTINA PARA POSICIONAR A LINHA NA PRIMEIRA LINHA MESCLADA REFERENTE A ANALISE SELECIONADA'''
          intContador = 1
          Do While intContador <= hfgAnalise.Rows - 1
             hfgAnalise.Row = intContador
             If strExclui_Analise = hfgAnalise.Text Then
                Exit Do
             End If
             intContador = intContador + 1
          Loop
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Do While strExclui_Analise = hfgAnalise.Text
             If hfgAnalise.Rows = 2 Then
                Movimentacoes.Monta_HFlex_Grid hfgAnalise, "2050,5500", "Análises,Especificações", 2, "Otica", Me
                hfgAnalise.ColAlignmentFixed(1) = 4
                hfgAnalise.ColAlignmentFixed(2) = 4
                hfgAnalise.RowHeight(1) = 285
                txtAnalise.Text = Empty
                txtEspecificacao.Text = Empty
                hfgAnalise.Col = 0
                hfgAnalise.Row = 1
                Exit Sub
             End If
             hfgAnalise.RemoveItem (hfgAnalise.Row)
          Loop
       Else
          hfgAnalise.RemoveItem (hfgAnalise.Row)
       End If
       
       intContador = 1
       hfgAnalise.Col = 0
       Do While intContador <= hfgAnalise.Rows - 1
          hfgAnalise.Row = intContador
          hfgAnalise.Text = intContador
          intContador = intContador + 1
       Loop
    End If

    txtAnalise.Text = Empty
    txtEspecificacao.Text = Empty
    intClique_Especificacao = Empty
    intClique_Analise = Empty
    
    txtAnalise.SetFocus

    hfgAnalise.Col = 0
    hfgAnalise.Row = 1
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub cmdSalvar_Click()
    Dim intIndice As Integer
    
    If txtAnalise.Text = Empty Then
       MsgBox "Análise inválida. Verifique.", vbInformation, "Only Tech"
       txtAnalise.SetFocus
       Exit Sub
    ElseIf txtEspecificacao.Text = Empty Then
       MsgBox "Especificação inválida. Verifique.", vbInformation, "Only Tech"
       txtEspecificacao.SetFocus
       Exit Sub
    End If
    
    intContador = 1
    Do While intContador <= hfgAnalise.Rows - 1
       hfgAnalise.Row = intContador
       hfgAnalise.Col = 2
       strEspecificacao = hfgAnalise.Text
       hfgAnalise.Col = 1
       If UCase(hfgAnalise.Text) = UCase(txtAnalise.Text) And UCase(strEspecificacao) = UCase(txtEspecificacao.Text) Then
          MsgBox "Esta Análise e Especificação já foram definidas. Verifique.", vbInformation, "Only Tech"
          txtAnalise.SetFocus
          Exit Sub
       End If
       intContador = intContador + 1
    Loop
    
    hfgAnalise.Row = 1
    hfgAnalise.Col = 1
    
    If hfgAnalise.Text <> Empty Then
       intIndice = intContador
       hfgAnalise.Rows = hfgAnalise.Rows + 1
    Else
       intIndice = intContador - 1
    End If

    hfgAnalise.Row = intIndice
    
    hfgAnalise.Col = 0
    hfgAnalise.ColWidth(0) = 500
    hfgAnalise.Font.Name = "Tahoma"
    hfgAnalise.CellFontSize = 7
    hfgAnalise.CellFontBold = False
    hfgAnalise.CellBackColor = &H80FFFF
    hfgAnalise.Text = intIndice
    
    'Os campos sao medidos e uma constante de proporcionalidade é usada para o tamanho das linhas
    If Len(txtEspecificacao.Text) > Len(txtAnalise.Text) Then
       If Len(txtEspecificacao.Text) > 60 Then
          hfgAnalise.RowHeight(intIndice) = 285 * CDbl((Len(txtEspecificacao.Text)) / 48)
          hfgAnalise.WordWrap = True
       End If
    Else
       If Len(txtAnalise.Text) > 20 Then
          hfgAnalise.RowHeight(intIndice) = 285 * CDbl((Len(txtAnalise.Text)) / 10)
          hfgAnalise.WordWrap = True
       End If
    End If

    hfgAnalise.Col = 1
    hfgAnalise.Text = txtAnalise.Text
    hfgAnalise.Col = 2
    hfgAnalise.Text = txtEspecificacao.Text
    
    Call Ajusta_Analise
    
    hfgAnalise.MergeCol(1) = True
    hfgAnalise.MergeCol(0) = True
    hfgAnalise.MergeCells = flexMergeRestrictColumns
    hfgAnalise.ColAlignment(1) = 4
    
    hfgAnalise.Refresh
        
    txtAnalise.Text = Empty
    txtEspecificacao.Text = Empty
    intClique_Especificacao = Empty
    intClique_Analise = Empty
       
    hfgAnalise.Col = 0: hfgAnalise.Row = 1
    txtAnalise.SetFocus
End Sub

Private Sub Command1_Click()
    txtEspecificacao.Text = txtNumero.Text ^ txtPotencia.Text
End Sub

Private Sub dtcFuncao_GotFocus()
   If txtFuncao.Text = Empty Then
      Call Movimentacoes.Verifica_DataCombo(dtcFuncao.Text)
   End If
End Sub

Private Sub dtcFuncao_LostFocus()
   txtFuncao.Text = dtcFuncao.BoundText
   If IsNumeric(txtFuncao.Text) = False Or dtcFuncao.Text = Empty Then txtFuncao.Text = Empty: Exit Sub
End Sub

Private Sub dtcTipo_Marcha_LostFocus()
    txtTipo_Marcha.Text = dtcTipo_Marcha.BoundText
   If IsNumeric(txtTipo_Marcha.Text) = False Or dtcTipo_Marcha.Text = Empty Then txtTipo_Marcha.Text = Empty: Exit Sub
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
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
  
Private Sub Form_Load()
    On Error GoTo Erro
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Insumo"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstInsumo, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'PORTAL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifica se existe integração com o portal
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    If booIntegra_Portal = True Then
       Me.shpIntegrado.Visible = True
    Else
       Me.shpIntegrado.Visible = False
    End If
    If booIntegra_Portal = True Then
       On Error GoTo Erro_Portal
       Dim rstReg_nao_integrados As New ADODB.Recordset
       Dim strMensagem As String
       If Funcoes_Gerais.Verifica_registros_nao_integrados("TBinsumo", "PKCodigo_TBInsumo,DFDescricao_TBInsumo,DFIntegrado_TBInsumo", "Otica", "DFIntegrado_TBInsumo", rstReg_nao_integrados, Me) = True Then
          If rstReg_nao_integrados.RecordCount > 0 Then
             strMensagem = "Deseja atualizar as informações para o portal? Existem " & rstReg_nao_integrados.RecordCount & " registro(s) desatualizados."
             intRetorno = MsgBox(strMensagem, vbYesNo, "Only Tech")
             If intRetorno = 6 Then
                frmAguarde.Show
                Funcoes_Gerais.Atualiza_registros_nao_integrados rstReg_nao_integrados, "TBinsumo_portal", "PKCodigo_TBInsumo_portal", "PKCodigo_TBInsumo", "Portal", "PKCodigo_TBInsumo_portal,DFDescricao_TBInsumo_portal", Me, "ortofarma1", "Otica", "BDRetaguarda", "TBinsumo", "DFIntegrado_TBinsumo", "ortofarma1", "ortofarma7410"
                MsgBox "Dados no portal atualizados com sucesso!", vbInformation, "Only Tech"
                Unload frmAguarde
             End If
          End If
       End If
    End If
    
Fim_atu_portal:

    On Error GoTo Erro

    
    
    log.Descricao = "Inicializando o cadastro de Insumo"
    
    'Gravando o log
    log.Gravar_log "Otica", Me
    

    sstInsumo.TabEnabled(0) = False
    sstInsumo.TabEnabled(1) = False
    sstInsumo.Tab = 2
    
    Call Reposicao
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
        
    Exit Sub
    
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
        
Erro_Portal:
    MsgBox "Erro ao atualizar registros no Portal.Verifique", vbCritical, "Onlytech"
    Err.Clear
    Unload frmAguarde
    GoTo Fim_atu_portal
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    Unload frmInsumo_Consulta_Detalhada_Insumo
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando o cadastro de Insumo"
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Set log = Nothing
    
    strCombo = Empty
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgAnalise_Click()
    If hfgAnalise.Text <> Empty Then
       intClique_Analise = Empty
       intClique_Especificacao = Empty
       txtAnalise.Text = Empty
       txtEspecificacao.Text = Empty
    End If
End Sub

Private Sub hfgAnalise_DblClick()

    If hfgAnalise.Col = 1 And hfgAnalise.Text <> Empty Then
       intClique_Analise = hfgAnalise.Row
       intClique_Especificacao = Empty
       txtAnalise.Text = hfgAnalise.TextArray((hfgAnalise.Row * hfgAnalise.Cols + CDbl(hfgAnalise.Col - hfgAnalise.Col) + 1))
       strAnalise_Antiga = txtAnalise.Text
    ElseIf hfgAnalise.Col = 2 And hfgAnalise.Text <> Empty Then
       intClique_Especificacao = hfgAnalise.Row
       intClique_Analise = Empty
       'enviando o texto da analise referente a esta especificacao para posterior verificacao na alteração
       txtEspecificacao.Text = hfgAnalise.TextArray((hfgAnalise.Row * hfgAnalise.Cols + CDbl(hfgAnalise.Col - hfgAnalise.Col) + 2))
       strAnalise_Antiga = hfgAnalise.TextArray((hfgAnalise.Row * hfgAnalise.Cols + CDbl(hfgAnalise.Col - hfgAnalise.Col) + 1))
    End If
    
End Sub

Private Sub hfgInsumo_Click()

   If hfgInsumo.Col = 0 And hfgInsumo.Text <> Empty Then
     
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
        
        txtCodigo.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 1))
        txtDescricao.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 2))
        txtFuncao.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 3))
        txtNome_Cientifico.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 5))
        txtConservacao.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 6))
        txtReferencia1.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 7))
        txtReferencia2.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 8))
        txtReferencia3.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 9))
        txtReferencia4.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 10))
        txtReferencia5.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 11))
        txtObservacao.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 12))
        txtTecnica_Aplicada.Text = hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 13))
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'PORTAL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If booIntegra_Portal = True Then
            If hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 14)) = "Sim" Then
               Me.shpIntegrado.BackColor = &H8000&
            End If
            If hfgInsumo.TextArray((hfgInsumo.Row * hfgInsumo.Cols + hfgInsumo.Col + 14)) = "Não" Then
               Me.shpIntegrado.BackColor = vbRed
            End If
        End If
        
        'ABASTECENDO ANÁLISES
        strSql = "SELECT DFDescricao_TBAnalise_insumo, " & _
                 "DFDescricao_TBEspecificacao_analise_insumo " & _
                 "FROM TBAnalise_insumo " & _
                 "INNER JOIN TBEspecificacao_analise_insumo " & _
                 "ON TBAnalise_insumo.PKId_TBAnalise_Insumo = TBEspecificacao_analise_insumo.FKId_TBAnalise_Insumo " & _
                 "WHERE FKCodigo_TBInsumo = " & txtCodigo.Text & " " & _
                 "ORDER BY PKId_TBAnalise_Insumo,PKId_TBEspecificacao_analise_insumo"
        
        Movimenta_HFlex_Grid strSql, hfgAnalise, "2050,5500", "Análises,Especificações", "BDRetaguarda", "Otica", Me
        
        'centralizando o texto do cabeçalho
        hfgAnalise.ColAlignmentFixed(1) = 4
        hfgAnalise.ColAlignmentFixed(2) = 4
           
        hfgAnalise.Col = 1
        hfgAnalise.Row = 1
        
        If hfgAnalise.Text <> Empty Then
           Dim strAnalise As String
           'Acertando o tamanho das linhas - vide botao incluir e alterar
           intContador = 1
           Do While intContador <= hfgAnalise.Rows - 1
              hfgAnalise.Row = intContador
              
              hfgAnalise.Col = 1
              strAnalise = hfgAnalise.Text
              
              hfgAnalise.Col = 2
              'verificacao para montagem da proporcionalidade de acordo com o maior dos campos
              If Len(hfgAnalise.Text) > Len(strAnalise) Then
                 If Len(hfgAnalise.Text) > 60 Then
                    hfgAnalise.RowHeight(intContador) = 285 * CDbl((Len(hfgAnalise.Text)) / 49)
                    hfgAnalise.WordWrap = True
                 End If
              Else
                 If Len(strAnalise) > 20 Then
                    hfgAnalise.RowHeight(intContador) = 285 * CDbl((Len(strAnalise)) / 10)
                    hfgAnalise.WordWrap = True
                 End If
              End If
              intContador = intContador + 1
           Loop
           'rotina para ajuste da mesclagem
           Call Ajusta_Analise
           'habilitando a mesclagem
           hfgAnalise.MergeCol(1) = True
           hfgAnalise.MergeCells = flexMergeRestrictColumns
           hfgAnalise.ColAlignment(0) = 7
           hfgAnalise.ColAlignment(1) = 4
        Else
           hfgAnalise.Rows = 2
           Movimentacoes.Monta_HFlex_Grid hfgAnalise, "2050,5500", "Análises,Especificações", 2, "Otica", Me
           hfgAnalise.ColAlignmentFixed(1) = 4
           hfgAnalise.ColAlignmentFixed(2) = 4
        End If

        booAlterar = True
        txtConsulta.Text = Empty
        sstInsumo.TabEnabled(0) = True
        sstInsumo.TabEnabled(1) = True
        sstInsumo.Tab = 0
        
        txtCodigo.Enabled = False
                
   End If
   
   Unload frmAguarde
   
End Sub

Private Sub hfgInsumo_DblClick()
    hfgInsumo.Sort = 1
End Sub

Private Sub hfgInsumo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgInsumo_Click
    End If
End Sub

Private Sub sstInsumo_Click(PreviousTab As Integer)
    If sstInsumo.Tab = 0 Then
       txtDescricao.SetFocus
    ElseIf sstInsumo.Tab = 1 Then
       txtAnalise.SetFocus
    ElseIf sstInsumo.Tab = 2 Then
      If frmIntegracao.Visible = True Then
         Unload frmIntegracao
      End If
      If strCombo <> Empty And strCombo <> "Todos" Then
         cbbCampos.Text = strCombo
         If MDIPrincipal.OCXUsuario.PrivilégioConsultar = True Then: txtConsulta.SetFocus
      ElseIf strCombo = "Todos" Then
         hfgInsumo.Row = 1
         hfgInsumo.Col = 0
         hfgInsumo.SetFocus
      End If
   End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstInsumo.Tab <> 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
    End Select
End Sub

Function Gravar()
    On Error GoTo Erro
    
    'Verifica se os campos necessarios para gravar não estão nulos
    If txtDescricao.Text = Empty Then
       MsgBox "O campo descrição não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtDescricao.SetFocus
       Exit Function
    End If
    If txtFuncao.Text = Empty Then
       MsgBox "O campo código da função não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtFuncao.SetFocus
       Exit Function
    End If
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim intEmpresa As Integer
    Dim strCodigo_Insumo As String
    Dim strProx_Cod_Insumo As String
    Dim rstVerifica As New ADODB.Recordset
    
    intEmpresa = MDIPrincipal.OCXUsuario.Empresa
    
    If booAlterar = False Then
       strProx_Cod_Insumo = Funcoes_Gerais.Localiza_Proximo_Codigo("DFProximo_insumo_TBParametros_servicos", "FKCodigo_TBEmpresa", intEmpresa, "TBParametros_Servicos", "Otica", Me, "BDRetaguarda")
       txtCodigo.Text = strProx_Cod_Insumo
    End If
    
    strCampo = "PKCodigo_TBInsumo,DFDescricao_TBInsumo,FKCodigo_TBFuncao_insumo," & _
               "DFNome_cientifico_TBInsumo,DFObservacao_TBInsumo," & _
               "DFConservacao_TBInsumo,DFReferencia_biografica1_TBInsumo,DFReferencia_biografica2_TBInsumo," & _
               "DFReferencia_biografica3_TBInsumo,DFReferencia_biografica4_TBInsumo," & _
               "DFReferencia_biografica5_TBInsumo,DFTecnica_aplicada_TBInsumo,DFIntegrado_TBInsumo," & _
               "DFData_alteracao_TBInsumo,DFIntegrado_filiais_TBInsumo "
       
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBInsumo "
    End If
    
    strValores = "" & txtCodigo.Text & ",'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & txtFuncao.Text & "," & _
                 "'" & Funcoes_Gerais.Grava_String(txtNome_Cientifico.Text) & "','" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtConservacao.Text) & "','" & Funcoes_Gerais.Grava_String(txtReferencia1.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtReferencia2.Text) & "','" & Funcoes_Gerais.Grava_String(txtReferencia3.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtReferencia4.Text) & "','" & Funcoes_Gerais.Grava_String(txtReferencia5.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtTecnica_Aplicada.Text) & "',0," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0 "
                 
    If booIntegra_Portal = True Then
        strValores = strValores & ",0 "
    End If
              
    'Abrindo conexao
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    On Error GoTo Erro_transacao
    
    If booAlterar = True Then
       log.Evento = "Alterar"

       strSet = "UPDATE TBInsumo SET DFDescricao_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                "FKCodigo_TBFuncao_insumo = " & txtFuncao.Text & "," & _
                "DFNome_cientifico_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtNome_Cientifico.Text) & "'," & _
                "DFObservacao_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "'," & _
                "DFConservacao_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtConservacao.Text) & "'," & _
                "DFReferencia_biografica1_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtReferencia1.Text) & "'," & _
                "DFReferencia_biografica2_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtReferencia2.Text) & "'," & _
                "DFReferencia_biografica3_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtReferencia3.Text) & "'," & _
                "DFReferencia_biografica4_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtReferencia4.Text) & "'," & _
                "DFReferencia_biografica5_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtReferencia5.Text) & "'," & _
                "DFTecnica_aplicada_TBInsumo = '" & Funcoes_Gerais.Grava_String(txtTecnica_Aplicada.Text) & "', " & _
                "DFIntegrado_TBInsumo = 0," & _
                "DFData_alteracao_TBInsumo = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBInsumo = 0 "
                
      If booIntegra_Portal = True Then
         strSet = strSet & ",DFIntegrado_portal_TBInsumo = 0 "
      End If
      
      strSet = strSet & "WHERE PKCodigo_TBInsumo = " & txtCodigo.Text & ""
              
       conexao.CNConexao.Execute strSet
       
       Call Grava_Analise
       
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"

       strSql = "INSERT INTO TBInsumo(" & strCampo & ") VALUES(" & strValores & ")"
       
       conexao.CNConexao.Execute strSql
       
       Call Grava_Analise
       
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
       
       ''''' aqui começa a ATUALIZAÇÃO DA TABELA TBParametros_Servicos '''''
        
       'Somente para mostrar ao usuario o código que o cliente foi incluido
       strCodigo_Insumo = strProx_Cod_Insumo
       
       If strCodigo_Insumo <> Empty Then
          MsgBox "** O código desse Insumo é: " & strCodigo_Insumo & "", vbOKOnly, "Only Tech"
       End If
        
       strProx_Cod_Insumo = strProx_Cod_Insumo + 1
       
       strSet = "SET DFProximo_Insumo_TBParametros_Servicos = " & strProx_Cod_Insumo & "," & _
                "DFData_alteracao_TBParametros_servicos = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBParametros_servicos = 0 "
       
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBParametros_servicos = 0 "
       End If

       Call funcoes_banco.Alterar("TBParametros_Servicos", strSet, "FKCodigo_TBEmpresa", MDIPrincipal.OCXUsuario.Empresa, "Otica", Me, "BDRetaguarda")
    End If

    'Fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Portal
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If booIntegra_Portal = True Then
        intRetorno = MsgBox("Deseja atualizar as informações para o portal?", vbYesNo, "Only Tech")
        If intRetorno = 6 Then
            On Error GoTo Erro_Portal
            
            strCampo = "PKCodigo_TBInsumo_portal,DFDescricao_TBInsumo_portal"
            strValores = "" & txtCodigo.Text & ",'" & Funcoes_Gerais.Grava_String(Me.txtDescricao.Text) & "'"
                        
            Dim rstPortal_gravacao As New ADODB.Recordset
            
            strSql = Empty
            strSql = "SELECT COUNT(*) FROM TBInsumo_portal WHERE PKCodigo_TBInsumo_portal = " & Me.txtCodigo.Text & ""
            Movimentacoes.Select_geral strSql, "ortofarma1", rstPortal_gravacao, "Portal", Me
            
            If rstPortal_gravacao.BOF = True And rstPortal_gravacao.EOF = True Then
               log.Evento = "Alterar"
               strSet = "SET DFDescricao_TBInsumo_portal = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'"
               Call funcoes_banco.Alterar_Portal("ortofarma1", "TBInsumo_portal", strSet, "PKCodigo_TBInsumo_portal", txtCodigo.Text, "PKCodigo_TBInsumo", txtCodigo.Text, "Otica", Me, "BDRetaguarda", "TBInsumo", "DFIntegrado_TBInsumo")
               log.Descricao = "Alterando o registro no Portal: " + txtCodigo.Text
               log.Tipo = 1
               log.Hora = Format(Now, "hh:mm:ss")
               'Gravando log
               log.Gravar_log "OTICA", Me
            Else
               log.Evento = "Incluir Novo"
               Call funcoes_banco.Gravar_Portal("ortofarma1", "TBInsumo_portal", strCampo, strValores, Me, "Otica", "BDRetaguarda", "TBInsumo", "DFIntegrado_TBInsumo", "PKCodigo_TBInsumo", Me.txtCodigo.Text)
               log.Descricao = "Gravando o registro no Portal: " + txtCodigo.Text
               log.Tipo = 1
               log.Hora = Format(Now, "hh:mm:ss")
               'Gravando log
               log.Gravar_log "OTICA", Me
            End If
            
            Set rstPortal_gravacao = Nothing
            
        End If
        On Error GoTo Erro
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
       Me.hfgInsumo.Visible = False
    End If
    
    sstInsumo.TabEnabled(0) = False
    sstInsumo.TabEnabled(1) = False
    sstInsumo.Tab = 2
    
    Exit Function
    
Erro_transacao:
    
    conexao.CNConexao.RollbackTrans
    conexao.Fechar_conexao
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
       
Erro_Portal:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    MsgBox "Ocorreram erros na integração com o Portal!Contacte Only Tech.", vbCritical, "Only Tech"

    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    'Abrindo conexao
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'deletando registros filhos
    strSql = "DELETE TBEspecificacao_analise_insumo FROM TBEspecificacao_analise_insumo " & _
             "INNER JOIN TBAnalise_Insumo " & _
             "ON TBEspecificacao_analise_insumo.FKId_TBAnalise_Insumo = TBAnalise_Insumo.PKId_TBAnalise_Insumo " & _
             "WHERE TBAnalise_Insumo.FKCodigo_TBInsumo = " & txtCodigo.Text & ""
             
    conexao.CNConexao.Execute strSql
    
    strSql = "DELETE FROM TBAnalise_Insumo WHERE FKCodigo_TBInsumo = " & txtCodigo.Text & ""
    
    conexao.CNConexao.Execute strSql

    'Excluindo Registro
    strSql = "DELETE FROM TBInsumo WHERE PKCodigo_TBInsumo = " & txtCodigo.Text & ""
    
    conexao.CNConexao.Execute strSql
    
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'PORTAL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If booIntegra_Portal = True Then
        intRetorno = MsgBox("Deseja atualizar as informações para o portal?", vbYesNo, "Only Tech")
        If intRetorno = 6 Then
            Call funcoes_banco.Excluir("TBInsumo_portal", "PKCodigo_TBInsumo_portal", txtCodigo.Text, "Portal", Me, "ortofarma1")
        End If
    End If
    
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
       hfgInsumo.Visible = False
    End If
           
    sstInsumo.TabEnabled(0) = False
    sstInsumo.TabEnabled(1) = False
    sstInsumo.Tab = 2
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Excluir")
    Exit Function
    
    conexao.CNConexao.RollbackTrans
    conexao.Fechar_conexao
    
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
       hfgInsumo.Visible = False
    End If
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstInsumo.TabEnabled(0) = False
    sstInsumo.TabEnabled(1) = False
    sstInsumo.Tab = 2
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
    Dim rstBusca_Parametro As New ADODB.Recordset
    Dim strCodigo_Insumo As String
    
    Call Objetos.Limpa_TXT(Me)
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
        
    strSql = Empty
    strSql = "SELECT * FROM TBParametros_Servicos " & _
             "WHERE TBParametros_Servicos.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ""
             
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Parametro, "Otica", Me)
    
    If rstBusca_Parametro.RecordCount = 0 Then
       MsgBox "Parâmetros de Serviços não definidos. Verifique.", vbInformation, "Only Tech"
       Set rstBusca_Parametro = Nothing
       Exit Function
    End If
    
    strCodigo_Insumo = rstBusca_Parametro.Fields("DFProximo_Insumo_TBParametros_Servicos")
    Set rstBusca_Parametro = Nothing
        
    strSql = Empty
    strSql = "SELECT * FROM TBInsumo WHERE TBInsumo.PKCodigo_TBInsumo = " & strCodigo_Insumo & ""
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Parametro, "Otica", Me)
    
    If rstBusca_Parametro.RecordCount <> 0 Then
       MsgBox "O Código " & strCodigo_Insumo & " já existe, por favor, verifique o cadastro Parâmetros de Serviços e atualize o código do próximo Insumo.", vbInformation, "Only Tech"
       Set rstBusca_Parametro = Nothing
       Call Objetos.Limpa_TXT(Me)
       sstInsumo.Tab = 2
       Exit Function
    End If
    Set rstBusca_Parametro = Nothing
    
    If booIntegra_Portal = True Then
       Me.shpIntegrado.BackColor = vbRed
    End If
    
    hfgAnalise.Rows = 2
    Movimentacoes.Monta_HFlex_Grid hfgAnalise, "2050,5500", "Análises,Especificações", 2, "Otica", Me
    hfgAnalise.ColAlignmentFixed(1) = 4
    hfgAnalise.ColAlignmentFixed(2) = 4
    hfgAnalise.RowHeight(1) = 285
    
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
    
    sstInsumo.TabEnabled(0) = True
    sstInsumo.TabEnabled(1) = True
    sstInsumo.Tab = 0
    
    txtCodigo.Enabled = False
    txtDescricao.SetFocus
       
    booAlterar = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Sub txtAnalise_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

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
        Movimentacoes.Verifica_Numero "PKCodigo_TBInsumo", "TBInsumo", txtCodigo, "Otica", Me
    End If
End Sub

Private Function Reposicao()
    On Error GoTo Erro
    
    strTamanho = "900,3500,1200,3000,3000,3000," & _
                 "3000,3000," & _
                 "3000,3000,2000,3500,6000"
                 
    strNomes = "Código,Descrição,Cod. Função,Função,Nome Científico,Conservação," & _
               "Referência 1,Referência 2," & _
               "Referência 5,Referência 4,Referência 5,Observação,Técnica Aplicada"
    
    Movimentacoes.Monta_HFlex_Grid hfgInsumo, strTamanho, strNomes, 13, "Otica", Me

    Movimentacoes.Monta_HFlex_Grid hfgAnalise, "2050,5500", "Análises,Especificações", 2, "Otica", Me
    hfgAnalise.ColAlignmentFixed(1) = 4
    hfgAnalise.ColAlignmentFixed(2) = 4
    
    Call Monta_Combo

    strSql = "SELECT PKCodigo_TBFuncao_insumo,DFDescricao_TBFuncao_insumo FROM TBFuncao_insumo "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBFuncao_insumo", "DFDescricao_TBFuncao_insumo", dtcFuncao, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBTipo_marcha,DFDescricao_TBTipo_marcha FROM TBTipo_marcha"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBTipo_marcha", "DFDescricao_TBTipo_marcha", dtcTipo_Marcha, strSql, "BDRetaguarda", "Otica", Me
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Reposicao")
    Resume Next
End Function

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Sub txtDescricao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Consulta()
    
    If cbbCampos.Text <> "Todos" And cbbCampos.Text <> "Integrado Portal" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
    
    strSql = "SELECT TBInsumo.PKCodigo_TBInsumo,TBInsumo.DFDescricao_TBInsumo," & _
             "TBInsumo.FKCodigo_TBFuncao_insumo,TBFuncao_insumo.DFDescricao_TBFuncao_insumo," & _
             "DFNome_cientifico_TBInsumo,DFConservacao_TBInsumo," & _
             "DFReferencia_biografica1_TBInsumo,DFReferencia_biografica2_TBInsumo," & _
             "DFReferencia_biografica3_TBInsumo,DFReferencia_biografica4_TBInsumo," & _
             "DFReferencia_biografica5_TBInsumo,DFObservacao_TBInsumo,DFTecnica_aplicada_TBInsumo,DFIntegrado_TBInsumo " & _
             "FROM TBInsumo " & _
             "INNER JOIN TBFuncao_insumo ON TBInsumo.FKCodigo_TBFuncao_insumo = TBFuncao_insumo.PKCodigo_TBFuncao_insumo "
            
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    Funcoes_Gerais.Grava_String (txtConsulta.Text)
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE TBInsumo.PKCodigo_TBInsumo = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Descrição" Then
          strSql = strSql & " WHERE TBInsumo.DFDescricao_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Cod. Função" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE TBInsumo.FKCodigo_TBFuncao_insumo = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Função" Then
          strSql = strSql & " WHERE TBFuncao_insumo.DFDescricao_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Nome Científico" Then
          strSql = strSql & " WHERE TBInsumo.DFNome_cientifico_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Conservação" Then
          strSql = strSql & " WHERE TBInsumo.DFConservacao_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Referência 1" Then
          strSql = strSql & " WHERE TBInsumo.DFReferencia_biografica1_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Referência 2" Then
          strSql = strSql & " WHERE TBInsumo.DFReferencia_biografica2_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Referência 3" Then
          strSql = strSql & " WHERE TBInsumo.DFReferencia_biografica3_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Referência 4" Then
          strSql = strSql & " WHERE TBInsumo.DFReferencia_biografica4_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Referência 5" Then
          strSql = strSql & " WHERE TBInsumo.DFReferencia_biografica5_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Observação" Then
          strSql = strSql & " WHERE TBInsumo.DFObservacao_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Técnica Aplicada" Then
          strSql = strSql & " WHERE TBInsumo.DFTecnica_aplicada_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       End If
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'PORTAL
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If cbbCampos.Text = "Integrado Portal" Then
          Dim intIntegra As Integer
          If Me.cbbIntergra_portal.Text = "Sim" Then
             intIntegra = 1
          End If
          If Me.cbbIntergra_portal.Text = "Não" Then
             intIntegra = 0
          End If
          strSql = strSql & " WHERE TBInsumo.DFIntegrado_TBInsumo = " & intIntegra & ""
       End If
    End If
    
    frmAguarde.Show
    DoEvents
            
    strSql = strSql & " ORDER BY TBInsumo.PKCodigo_TBInsumo"
       
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgInsumo, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    hfgInsumo.Col = 0
    hfgInsumo.Row = 1
    If hfgInsumo.Text = Empty Then
       hfgInsumo.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgInsumo, strTamanho, strNomes, 14, "Otica", Me
    End If
    
    Unload frmAguarde
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código")
    cbbCampos.AddItem ("Descrição")
    cbbCampos.AddItem ("Cod. Função")
    cbbCampos.AddItem ("Função")
    cbbCampos.AddItem ("Nome Científico")
    cbbCampos.AddItem ("Conservação")
    cbbCampos.AddItem ("Referência 1")
    cbbCampos.AddItem ("Referência 2")
    cbbCampos.AddItem ("Referência 3")
    cbbCampos.AddItem ("Referência 4")
    cbbCampos.AddItem ("Referência 5")
    cbbCampos.AddItem ("Observação")
    cbbCampos.AddItem ("Técnica Aplicada")
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'PORTAL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If booIntegra_Portal = True Then
        cbbCampos.AddItem ("Integrado Portal")
    End If
End Function

Private Sub txtEspecificacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFuncao_Change()
    dtcFuncao.BoundText = txtFuncao.Text
    If IsNumeric(txtFuncao.Text) = False Then txtFuncao.Text = Empty: Exit Sub
End Sub

Private Sub txtFuncao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFuncao_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFuncao_LostFocus()
    If dtcFuncao.Text = Empty Then txtFuncao.Text = Empty
End Sub

Private Sub txtNome_cientifico_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConservacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtObservacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtReferencia1_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtReferencia2_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtReferencia3_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtReferencia4_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtReferencia5_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Public Function Ajusta_Analise()
    Dim strAnalise As String
    Dim intContador2 As Integer
    Dim intLinha1 As Integer
    Dim intLinha2 As Integer
    Dim strColuna1_1 As String
    Dim strColuna2_1 As String
    Dim strColuna1_2 As String
    Dim strColuna2_2 As String

    'Desabilitando a mescla de colunas antes da modificacao para corrigir um bug do vb
    hfgAnalise.MergeCol(1) = False

    intContador = 1
    Do While intContador <= hfgAnalise.Rows - 1
       hfgAnalise.Col = 1
       hfgAnalise.Row = intContador
       strAnalise = hfgAnalise.Text
       'Varre até que a linha tenha um texto diferente.
       'O contador continua o incremento para trabalhar a partir das linhas com o mesmo texto.
       Do While strAnalise = hfgAnalise.Text And intContador <= hfgAnalise.Rows - 1
          hfgAnalise.Row = intContador
          intLinha1 = hfgAnalise.Row
          If strAnalise <> hfgAnalise.Text Then
             Exit Do
          Else
             intContador = intContador + 1
          End If
       Loop
       'Um novo contador é implementado para varrer o restante do grid em busca de análises perdidas
       intContador2 = intContador
       'O contador recebe um decréscimo de 1 para continuar do ponto anterior ao loop de verificacao de repeticao
       intContador = intContador - 1
       Do While intContador2 <= hfgAnalise.Rows - 1
          hfgAnalise.Col = 1
          hfgAnalise.Row = intContador2
          If hfgAnalise.Text = strAnalise Then
             'Trocando as informacoes entre linhas
             intLinha2 = hfgAnalise.Row
             strColuna1_2 = hfgAnalise.Text
             hfgAnalise.Col = 2
             strColuna2_2 = hfgAnalise.Text

             hfgAnalise.Row = intLinha1
             hfgAnalise.Col = 1
             strColuna1_1 = hfgAnalise.Text
             hfgAnalise.Col = 2
             strColuna2_1 = hfgAnalise.Text
             
             'Sobrepondo pelas informacoes da linha logo abaixo da ultima igual
             hfgAnalise.Row = intLinha2
             hfgAnalise.Col = 1
             hfgAnalise.Text = strColuna1_1
             hfgAnalise.Col = 2
             hfgAnalise.Text = strColuna2_1
             
             hfgAnalise.Row = intLinha1
             hfgAnalise.Col = 1
             hfgAnalise.Text = strColuna1_2
             hfgAnalise.Col = 2
             hfgAnalise.Text = strColuna2_2
          End If
          intContador2 = intContador2 + 1
       Loop
       intContador = intContador + 1
    Loop
End Function

Private Function Grava_Analise()
    Dim strAnalise As String

    If booAlterar = True Then
       'deletando registros antes da inclusao
       strSql = "DELETE TBEspecificacao_analise_insumo FROM TBEspecificacao_analise_insumo " & _
                "INNER JOIN TBAnalise_Insumo " & _
                "ON TBEspecificacao_analise_insumo.FKId_TBAnalise_Insumo = TBAnalise_Insumo.PKId_TBAnalise_Insumo " & _
                "WHERE TBAnalise_Insumo.FKCodigo_TBInsumo = " & txtCodigo.Text & ""
                
       conexao.CNConexao.Execute strSql
       
       strSql = "DELETE FROM TBAnalise_Insumo WHERE FKCodigo_TBInsumo = " & txtCodigo.Text & ""
       
       conexao.CNConexao.Execute strSql
    End If
    
    hfgAnalise.Col = 1
    hfgAnalise.Row = 1
    If hfgAnalise.Text <> Empty Then
       intContador = 1
       Do While intContador <= hfgAnalise.Rows - 1
          hfgAnalise.Row = intContador
          hfgAnalise.Col = 1
          
          strSql = "INSERT INTO TBAnalise_Insumo(FKCodigo_TBInsumo,DFDescricao_TBAnalise_insumo) " & _
          "VALUES (" & txtCodigo.Text & ",'" & Funcoes_Gerais.Grava_String(hfgAnalise.Text) & "')"

          conexao.CNConexao.Execute strSql

          'Gravando as especificacoes até que a analise seja a mesma
          strAnalise = hfgAnalise.Text
          Do While strAnalise = hfgAnalise.Text And intContador <= hfgAnalise.Rows - 1

             hfgAnalise.Col = 2
             
             strSql = "INSERT INTO TBEspecificacao_analise_insumo(FKId_TBAnalise_Insumo," & _
                      "DFDescricao_TBEspecificacao_analise_insumo) " & _
                      "SELECT MAX(PKId_TBAnalise_Insumo),'" & Funcoes_Gerais.Grava_String(hfgAnalise.Text) & "' " & _
                      "FROM TBAnalise_Insumo"
             
             conexao.CNConexao.Execute strSql
             
             intContador = intContador + 1
             
             If intContador <= hfgAnalise.Rows - 1 Then
                hfgAnalise.Row = intContador
                hfgAnalise.Col = 1
             End If
          Loop
       Loop
    End If
    
End Function

Private Sub txtTecnica_Aplicada_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtTipo_Marcha_Change()
    dtcTipo_Marcha.BoundText = txtTipo_Marcha.Text
    If IsNumeric(txtTipo_Marcha.Text) = False Then txtTipo_Marcha.Text = Empty: Exit Sub
End Sub

Private Sub txtTipo_Marcha_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtTipo_Marcha_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub dtcTipo_Marcha_GotFocus()
    If Me.txtTipo_Marcha.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcTipo_Marcha.Text)
    End If
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo_TBInsumo", txtCodigo.Text, "DFIntegrado_filiais_TBInsumo", "TBInsumo", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBInsumo", Me.Top, Me.Left, Me.Width, Me.Height, "Insumo")
    
End Function






