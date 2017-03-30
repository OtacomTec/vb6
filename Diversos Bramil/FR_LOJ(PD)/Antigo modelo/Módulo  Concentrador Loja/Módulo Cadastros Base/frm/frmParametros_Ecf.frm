VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmParametros_Ecf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parâmetros ECF"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmParametros_Ecf.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5745
   Begin TabDlg.SSTab sstParametros_Ecf 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   10504
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
      TabPicture(0)   =   "frmParametros_Ecf.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtProxima_Serie_Orcamento_Balcao"
      Tab(0).Control(1)=   "txtProximo_Orcamento_Balcao"
      Tab(0).Control(2)=   "txtFinalizadora_Sangria"
      Tab(0).Control(3)=   "txtFinalizadora_Abertura"
      Tab(0).Control(4)=   "txtProduto_Associado_Desconto"
      Tab(0).Control(5)=   "txtPercentual_Taxa"
      Tab(0).Control(6)=   "txtProduto_Associado_Taxa"
      Tab(0).Control(7)=   "txtCodigo"
      Tab(0).Control(8)=   "cbbTipo"
      Tab(0).Control(9)=   "cbbNumeros_Decimais"
      Tab(0).Control(10)=   "cbbDesconto"
      Tab(0).Control(11)=   "cbbCodigo_Inicial"
      Tab(0).Control(12)=   "cbbPreco_Peso"
      Tab(0).Control(13)=   "dtcProduto_Associado_Taxa"
      Tab(0).Control(14)=   "dtcProduto_Associado_Desconto"
      Tab(0).Control(15)=   "cbbControla_Vendedor"
      Tab(0).Control(16)=   "dtcFinalizadora_Abertura"
      Tab(0).Control(17)=   "dtcFinalizadora_Sangria"
      Tab(0).Control(18)=   "dtcEmpresa"
      Tab(0).Control(19)=   "Label16"
      Tab(0).Control(20)=   "Label15"
      Tab(0).Control(21)=   "Label14"
      Tab(0).Control(22)=   "Label13"
      Tab(0).Control(23)=   "Label12"
      Tab(0).Control(24)=   "Label11"
      Tab(0).Control(25)=   "Label10"
      Tab(0).Control(26)=   "Label9"
      Tab(0).Control(27)=   "Label8"
      Tab(0).Control(28)=   "Label7"
      Tab(0).Control(29)=   "Label4"
      Tab(0).Control(30)=   "Label5"
      Tab(0).Control(31)=   "Label3"
      Tab(0).Control(32)=   "Label1"
      Tab(0).Control(33)=   "Label2"
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmParametros_Ecf.frx":179E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cbbCampos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "HfgParametros_Ecf"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdConsulta"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdRefresh"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtConsulta"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtProxima_Serie_Orcamento_Balcao 
         Height          =   375
         Left            =   -71430
         TabIndex        =   18
         Top             =   5400
         Width           =   2025
      End
      Begin VB.TextBox txtProximo_Orcamento_Balcao 
         Height          =   360
         Left            =   -73050
         TabIndex        =   17
         Top             =   5400
         Width           =   1545
      End
      Begin VB.TextBox txtFinalizadora_Sangria 
         Height          =   360
         Left            =   -74880
         TabIndex        =   14
         Top             =   4740
         Width           =   1300
      End
      Begin VB.TextBox txtFinalizadora_Abertura 
         Height          =   360
         Left            =   -74880
         TabIndex        =   12
         Top             =   4080
         Width           =   1300
      End
      Begin VB.TextBox txtProduto_Associado_Desconto 
         Height          =   360
         Left            =   -74880
         TabIndex        =   10
         Top             =   3420
         Width           =   1300
      End
      Begin VB.TextBox txtPercentual_Taxa 
         Height          =   360
         Left            =   -70230
         TabIndex        =   9
         Top             =   2760
         Width           =   825
      End
      Begin VB.TextBox txtProduto_Associado_Taxa 
         Height          =   360
         Left            =   -74880
         TabIndex        =   7
         Top             =   2760
         Width           =   1300
      End
      Begin VB.TextBox txtCodigo 
         Height          =   360
         Left            =   -74880
         TabIndex        =   1
         Top             =   1440
         Width           =   1815
      End
      Begin AutoCompletar.CbCompleta cbbTipo 
         Height          =   360
         Left            =   -73020
         TabIndex        =   2
         Top             =   1440
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
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   2520
         TabIndex        =   20
         Top             =   780
         Width           =   2235
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   5220
         Picture         =   "frmParametros_Ecf.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   4830
         Picture         =   "frmParametros_Ecf.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgParametros_Ecf 
         Height          =   4575
         Left            =   120
         TabIndex        =   22
         Top             =   1230
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   8070
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   120
         TabIndex        =   19
         Top             =   780
         Width           =   2385
         _ExtentX        =   4207
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
      Begin AutoCompletar.CbCompleta cbbNumeros_Decimais 
         Height          =   360
         Left            =   -71190
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
      Begin AutoCompletar.CbCompleta cbbDesconto 
         Height          =   360
         Left            =   -74880
         TabIndex        =   4
         Top             =   2100
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
      Begin AutoCompletar.CbCompleta cbbCodigo_Inicial 
         Height          =   360
         Left            =   -73020
         TabIndex        =   5
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
      Begin AutoCompletar.CbCompleta cbbPreco_Peso 
         Height          =   360
         Left            =   -71190
         TabIndex        =   6
         Top             =   2100
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
      Begin MSDataListLib.DataCombo dtcProduto_Associado_Taxa 
         Height          =   360
         Left            =   -73530
         TabIndex        =   8
         Top             =   2760
         Width           =   3255
         _ExtentX        =   5741
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
      Begin MSDataListLib.DataCombo dtcProduto_Associado_Desconto 
         Height          =   360
         Left            =   -73530
         TabIndex        =   11
         Top             =   3420
         Width           =   4155
         _ExtentX        =   7329
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
      Begin AutoCompletar.CbCompleta cbbControla_Vendedor 
         Height          =   360
         Left            =   -74880
         TabIndex        =   16
         Top             =   5400
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
      Begin MSDataListLib.DataCombo dtcFinalizadora_Abertura 
         Height          =   360
         Left            =   -73530
         TabIndex        =   13
         Top             =   4080
         Width           =   4155
         _ExtentX        =   7329
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
      Begin MSDataListLib.DataCombo dtcFinalizadora_Sangria 
         Height          =   360
         Left            =   -73530
         TabIndex        =   15
         Top             =   4740
         Width           =   4155
         _ExtentX        =   7329
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
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   360
         Left            =   -74880
         TabIndex        =   38
         Top             =   780
         Width           =   5505
         _ExtentX        =   9710
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
      Begin VB.Label Label16 
         Caption         =   "Próx. Série Orç. Balcão"
         Height          =   240
         Left            =   -71400
         TabIndex        =   41
         Top             =   5160
         Width           =   2025
      End
      Begin VB.Label Label15 
         Caption         =   "Próx. Orç. Balcão"
         Height          =   240
         Left            =   -73050
         TabIndex        =   40
         Top             =   5160
         Width           =   1755
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   240
         Left            =   -74880
         TabIndex        =   39
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Finalizadora Sangria"
         Height          =   240
         Left            =   -74880
         TabIndex        =   37
         Top             =   4500
         Width           =   1755
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Finalizadora Abertura"
         Height          =   240
         Left            =   -74880
         TabIndex        =   36
         Top             =   3840
         Width           =   1845
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Controla Vendedor"
         Height          =   240
         Left            =   -74880
         TabIndex        =   35
         Top             =   5160
         Width           =   1605
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Produto Associado Desc."
         Height          =   240
         Left            =   -74880
         TabIndex        =   34
         Top             =   3180
         Width           =   2100
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "% Taxa"
         Height          =   240
         Left            =   -70230
         TabIndex        =   33
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Produto Associado Taxa"
         Height          =   240
         Left            =   -74880
         TabIndex        =   32
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cód. Parâmetro ECF"
         Height          =   240
         Left            =   -74880
         TabIndex        =   31
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código Inicial"
         Height          =   240
         Left            =   -73020
         TabIndex        =   30
         Top             =   1860
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Preço Peso Balança"
         Height          =   240
         Left            =   -71190
         TabIndex        =   29
         Top             =   1860
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Desconto"
         Height          =   240
         Left            =   -74880
         TabIndex        =   28
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Decimais"
         Height          =   240
         Left            =   -71190
         TabIndex        =   27
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Quantidade"
         Height          =   240
         Left            =   -73020
         TabIndex        =   24
         Top             =   1200
         Width           =   1410
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   5745
      _ExtentX        =   10134
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5820
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
            Picture         =   "frmParametros_Ecf.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":5578
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParametros_Ecf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Cadastro Base                                                  '
' Objetivo...............: Cadastro de Parâmetros Ecf                                     '
' Equipe Responsável.....: Jones, Giordano,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Desenvolvedor..........: Rafael Gomes                                                   '
' Data de Criação........: 14/01/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strTamanho As String
Dim strNomes As String
Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Dim booAlterar As Boolean
Public strSql As String
Dim conexao As New DLLConexao_Sistema.conexao
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim log As New DLLSystemManager.log

Function Imprimir()
    On Error GoTo Erro
    'Tratamento de erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Relatorio_Parametros_Ecf.Show
    
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

Private Sub dtcFinalizadora_Abertura_GotFocus()
    If txtFinalizadora_Abertura.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora_Abertura)
    End If
End Sub

Private Sub dtcFinalizadora_Abertura_LostFocus()
    txtFinalizadora_Abertura.Text = dtcFinalizadora_Abertura.BoundText
End Sub

Private Sub dtcFinalizadora_Sangria_GotFocus()
    If txtFinalizadora_Sangria.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora_Sangria)
    End If
End Sub

Private Sub dtcFinalizadora_Sangria_LostFocus()
    txtFinalizadora_Sangria.Text = dtcFinalizadora_Sangria.BoundText
End Sub

Private Sub dtcProduto_Associado_Desconto_GotFocus()
    If txtProduto_Associado_Desconto.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcProduto_Associado_Desconto)
    End If
End Sub

Private Sub dtcProduto_Associado_Desconto_LostFocus()
    txtProduto_Associado_Desconto.Text = dtcProduto_Associado_Desconto.BoundText
End Sub

Private Sub dtcProduto_Associado_Taxa_GotFocus()
    If txtProduto_Associado_Taxa.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcProduto_Associado_Taxa)
    End If
End Sub

Private Sub dtcProduto_Associado_Taxa_LostFocus()
    txtProduto_Associado_Taxa.Text = dtcProduto_Associado_Taxa.BoundText
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
    log.Programa = "Cadastro de Parâmetros EOF"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstParametros_Ecf, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o cadastro de Parâmetros EOF"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    Call Reposicao
    
    sstParametros_Ecf.TabEnabled(0) = False
    sstParametros_Ecf.Tab = 1
      
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Set log = Nothing
    
    strCombo = Empty
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgParametros_ecf_Click()
    If HfgParametros_Ecf.Col = 0 And HfgParametros_Ecf.Text <> Empty Then
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
       
      txtCodigo.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 1))
      cbbTipo.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 2))
      cbbNumeros_Decimais.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 3))
      cbbDesconto.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 4))
      cbbCodigo_Inicial.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 5))
      cbbPreco_Peso.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 6))
      txtProduto_Associado_Taxa.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 7))
      txtPercentual_Taxa.Text = Format(HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 8)), "#,###0.00")
      txtProduto_Associado_Desconto.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 9))
      txtFinalizadora_Abertura.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 10))
      txtFinalizadora_Sangria.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 11))
      cbbControla_Vendedor.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 12))
      dtcEmpresa.BoundText = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 13))
      txtProximo_Orcamento_Balcao = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 15))
      txtProxima_Serie_Orcamento_Balcao = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 16))
             
      txtCodigo.Enabled = False
      booAlterar = True
      txtConsulta.Text = Empty
      sstParametros_Ecf.TabEnabled(0) = True
      sstParametros_Ecf.Tab = 0
      Me.cbbTipo.SetFocus
   End If
   
   Unload frmAguarde
   
End Sub

Private Sub hfgParametros_ecf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgParametros_ecf_Click
    End If
End Sub

Private Sub sstParametros_ecf_Click(PreviousTab As Integer)
    If sstParametros_Ecf.Tab = 0 Then
       Me.cbbTipo.SetFocus
    ElseIf sstParametros_Ecf.Tab = 1 Then
        If strCombo <> Empty And strCombo <> "Todos" Then
           cbbCampos.Text = strCombo
           txtConsulta.SetFocus
        ElseIf strCombo = "Todos" Then
           HfgParametros_Ecf.Row = 1
           HfgParametros_Ecf.Col = 0
           HfgParametros_Ecf.SetFocus
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
    Dim strTipo As String
    Dim strDesconto As String
    Dim strPreco_Peso As String
    Dim strControla_Vendedor As String
    
    If cbbTipo.Text = "Inteiro" Then: strTipo = "I": Else: strTipo = "F"
    If cbbDesconto.Text = "$ - Dinheiro" Then: strDesconto = "$": Else: strDesconto = "%"
    If cbbPreco_Peso.Text = "0 - Preço" Then: strPreco_Peso = 0: Else: strPreco_Peso = 1
    If cbbControla_Vendedor.Text = "Sim" Then: strControla_Vendedor = 1: Else: strControla_Vendedor = 0
    
    strCampo = "PKCodigo_TBParametros_ecf," & _
               "DFTipo_quantidade_TBParametros_ecf," & _
               "DFNumero_decimais_TBParametros_ecf," & _
               "DFTipo_desconto_TBParametros_ecf," & _
               "DFCodigo_inicial_peso_variavel_TBParametros_ecf," & _
               "DFPreco_peso_balanca_TBParametros_ecf," & _
               "DFProduto_associado_taxa_TBParametros_ecf," & _
               "DFPercentual_taxa_TBParametros_ecf," & _
               "DFProduto_Desconto_associado_TBParametros_ecf," & _
               "DFControla_vendedor_TBParametros_ecf," & _
               "DFFinalizadora_abertura_TBParametros_ecf," & _
               "DFFinalizadora_sangria_TBParametros_ecf," & _
               "FKCodigo_TBEmpresa," & _
               "DFProximo_orcamento_balcao_TBParametros_ecf," & _
               "DFProximo_serie_orcamento_balcao_TBParametros_ecf"
                
    strValores = "'" & txtCodigo.Text & "', " & _
                 "'" & strTipo & "', " & _
                 "'" & cbbNumeros_Decimais.Text & "', " & _
                 "'" & strDesconto & "', " & _
                 "'" & cbbCodigo_Inicial.Text & "', " & _
                 "'" & strPreco_Peso & "', " & _
                 " " & txtProduto_Associado_Taxa.Text & ", " & _
                 " " & Funcoes_Gerais.Grava_Moeda(txtPercentual_Taxa.Text) & ", " & _
                 " " & txtProduto_Associado_Desconto.Text & ", " & _
                 "'" & strControla_Vendedor & "', " & _
                 " " & txtFinalizadora_Abertura.Text & ", " & _
                 " " & txtFinalizadora_Sangria.Text & ", " & _
                 " " & dtcEmpresa.BoundText & ", " & _
                 " " & txtProximo_Orcamento_Balcao.Text & ", " & _
                 "'" & txtProxima_Serie_Orcamento_Balcao.Text & "'"
                     
    If booAlterar = True Then
       
       log.Evento = "Alterar"
       strSet = "SET DFTipo_quantidade_TBParametros_ecf = '" & strTipo & "'," & _
                "    DFNumero_decimais_TBParametros_ecf = '" & cbbNumeros_Decimais.Text & "'," & _
                "    DFTipo_desconto_TBParametros_ecf = '" & strDesconto & "'," & _
                "    DFCodigo_inicial_peso_variavel_TBParametros_ecf = '" & cbbCodigo_Inicial.Text & "'," & _
                "    DFPreco_peso_balanca_TBParametros_ecf = '" & strPreco_Peso & "'," & _
                "    DFProduto_associado_taxa_TBParametros_ecf = " & txtProduto_Associado_Taxa.Text & "," & _
                "    DFPercentual_taxa_TBParametros_ecf = " & Funcoes_Gerais.Grava_Moeda(txtPercentual_Taxa.Text) & "," & _
                "    DFProduto_Desconto_associado_TBParametros_ecf = " & txtProduto_Associado_Desconto.Text & "," & _
                "    DFControla_vendedor_TBParametros_ecf = '" & strControla_Vendedor & "'," & _
                "    DFFinalizadora_abertura_TBParametros_ecf = " & txtFinalizadora_Abertura.Text & "," & _
                "    DFFinalizadora_sangria_TBParametros_ecf = " & txtFinalizadora_Sangria.Text & "," & _
                "    FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & "," & _
                "    DFProximo_orcamento_balcao_TBParametros_ecf = " & txtProximo_Orcamento_Balcao.Text & "," & _
                "    DFProximo_serie_orcamento_balcao_TBParametros_ecf = '" & txtProxima_Serie_Orcamento_Balcao.Text & "' "
 
       Call funcoes_banco.Alterar("TBParametros_ecf", strSet, "PKCodigo_TBParametros_ecf", txtCodigo.Text, "Otica", Me, "BDRetaguarda")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBParametros_ecf", strCampo, strValores, "Otica", Me, "BDRetaguarda")
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
    
    If booPrivilegio_Consultar = False Then
       HfgParametros_Ecf.Visible = False
    End If
    
    sstParametros_Ecf.TabEnabled(0) = False
    sstParametros_Ecf.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBParametros_ecf", "PKCodigo_TBParametros_ecf", txtCodigo.Text, "Otica", Me, "BDRetaguarda")
    
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
       HfgParametros_Ecf.Visible = False
    End If
        
    sstParametros_Ecf.TabEnabled(0) = False
    sstParametros_Ecf.Tab = 1
    
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
       HfgParametros_Ecf.Visible = False
    End If
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Call Monta_Combo
    
    sstParametros_Ecf.TabEnabled(0) = False
    sstParametros_Ecf.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
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
    
    sstParametros_Ecf.TabEnabled(0) = True
    sstParametros_Ecf.Tab = 0
    
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    booAlterar = False
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    On Error GoTo Erro
    
    strNomes = "Código,Tipo Quantidade,Nº Decimais,Tipo Desconto,Código Inicial,Preço Peso,Prod.Associado Taxa," & _
               "% Taxa,Prod.Associado Desc.,Finalizadora Abertura,Finalizadora Sangria,Controla Vendedor," & _
               "Empresa,Nome,Prox. Orç. Balcão,Prox. Série Orç. Balcão"
    
    strTamanho = "1000,1500,1500,1500,1500,1500,1600,1000,1600,1600,1600,1300,1000,0,1000,1500"
    
    Movimentacoes.Monta_HFlex_Grid HfgParametros_Ecf, strTamanho, strNomes, "16", "Otica", Me
    
    Call Monta_Combo
    Call Monta_DataCombo
              
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Reposicao")
    Resume Next
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
    If txtCodigo <> Empty Then
       Movimentacoes.Verifica_Numero "PKCodigo_TBParametros_ecf", "TBParametros_ecf", txtCodigo, "Otica", Me
    End If
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub


Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código")
    cbbCampos.AddItem ("Tipo Quantidade")
    cbbCampos.AddItem ("Nº Decimais")
    cbbCampos.AddItem ("Tipo Desconto")
    cbbCampos.AddItem ("Código Inicial")
    cbbCampos.AddItem ("Preço Peso")
    cbbCampos.AddItem ("Produto Associado Taxa")
    cbbCampos.AddItem ("% Taxa")
    cbbCampos.AddItem ("Produto Associado Desc.")
    cbbCampos.AddItem ("Finalizadora Abertura")
    cbbCampos.AddItem ("Finalizadora Sangria")
    cbbCampos.AddItem ("Controla Vendedor")
    cbbCampos.AddItem ("Código da Empresa")
    cbbCampos.AddItem ("Nome Empresa")
    cbbCampos.AddItem ("Prox. Orç. Balcão")
    cbbCampos.AddItem ("Prox. Série Orç. Balcão")
    
    cbbTipo.Clear
    cbbTipo.AddItem ("Inteiro")
    cbbTipo.AddItem ("Fração")
    
    cbbNumeros_Decimais.Clear
    cbbNumeros_Decimais.AddItem ("2")
    cbbNumeros_Decimais.AddItem ("3")
    
    cbbControla_Vendedor.Clear
    cbbControla_Vendedor.AddItem ("Sim")
    cbbControla_Vendedor.AddItem ("Não")
    
    cbbDesconto.Clear
    cbbDesconto.AddItem ("$ - Dinheiro")
    cbbDesconto.AddItem ("% - Percentagem")
    
    cbbCodigo_Inicial.Clear
    cbbCodigo_Inicial.AddItem ("0")
    cbbCodigo_Inicial.AddItem ("1")
    cbbCodigo_Inicial.AddItem ("2")
    cbbCodigo_Inicial.AddItem ("3")
    cbbCodigo_Inicial.AddItem ("4")
    cbbCodigo_Inicial.AddItem ("5")
    cbbCodigo_Inicial.AddItem ("6")
    cbbCodigo_Inicial.AddItem ("7")
    cbbCodigo_Inicial.AddItem ("8")
    cbbCodigo_Inicial.AddItem ("9")
    
    cbbPreco_Peso.Clear
    cbbPreco_Peso.AddItem ("0 - Preço")
    cbbPreco_Peso.AddItem ("1 - Peso")
    
End Function

Private Function Consulta()
    Dim strPercentual As String
    Dim strControla_Vendedor As String
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
    
    If cbbCampos.Text = "% Taxa" Then
       strPercentual = Format(txtConsulta.Text, "#,###0.00")
    End If
    
    If cbbCampos.Text = "Controla Vendedor" Then
       If txtConsulta.Text = "SIM" Then
          strControla_Vendedor = 1
       Else
          strControla_Vendedor = 0
       End If
    End If
    
    strSql = "SELECT TBParametros_ecf.PKCodigo_TBParametros_ecf," & _
             "TBParametros_ecf.DFTipo_quantidade_TBParametros_ecf," & _
             "TBParametros_ecf.DFNumero_decimais_TBParametros_ecf," & _
             "TBParametros_ecf.DFTipo_desconto_TBParametros_ecf," & _
             "TBParametros_ecf.DFCodigo_inicial_peso_variavel_TBParametros_ecf," & _
             "TBParametros_ecf.DFPreco_peso_balanca_TBParametros_ecf," & _
             "TBParametros_ecf.DFProduto_associado_taxa_TBParametros_ecf," & _
             "TBParametros_ecf.DFPercentual_taxa_TBParametros_ecf," & _
             "TBParametros_ecf.DFProduto_Desconto_associado_TBParametros_ecf," & _
             "TBParametros_ecf.DFFinalizadora_abertura_TBParametros_ecf," & _
             "TBParametros_ecf.DFFinalizadora_sangria_TBParametros_ecf," & _
             "TBParametros_ecf.DFControla_vendedor_TBParametros_ecf," & _
             "TBParametros_ecf.FKCodigo_TBEmpresa," & _
             "TBEmpresa.DFRazao_Social_TBEmpresa," & _
             "TBParametros_ecf.DFProximo_orcamento_balcao_TBParametros_ecf," & _
             "TBParametros_ecf.DFProximo_serie_orcamento_balcao_TBParametros_ecf " & _
             "FROM TBParametros_ecf " & _
             "INNER JOIN TBEmpresa ON  TBParametros_ecf.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa"
                     
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código" Then
          strSql = strSql & " WHERE convert(nvarchar,PKCodigo_TBParametros_ecf) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Tipo Quantidade" Then
          strSql = strSql & " WHERE convert(nvarchar,DFTipo_quantidade_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Nº Decimais" Then
          strSql = strSql & " WHERE convert(nvarchar,DFNumero_decimais_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Tipo Desconto" Then
          strSql = strSql & " WHERE convert(nvarchar,DFTipo_desconto_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Código Inicial" Then
          strSql = strSql & " WHERE convert(nvarchar,DFCodigo_inicial_peso_variavel_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Preço Peso" Then
          strSql = strSql & " WHERE convert(nvarchar,DFPreco_peso_balanca_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Produto Associado Taxa" Then
          strSql = strSql & " WHERE convert(nvarchar,DFProduto_associado_taxa_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "% Taxa" Then
          strSql = strSql & " WHERE convert(money,DFProduto_associado_taxa_TBParametros_ecf) = " & strPercentual & ""
       ElseIf cbbCampos.Text = "Produto Associado Desc." Then
          strSql = strSql & " WHERE convert(nvarchar,DFProduto_Desconto_associado_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Finalizadora Abertura" Then
          strSql = strSql & " WHERE convert(nvarchar,DFFinalizadora_abertura_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Finalizadora Sangria" Then
          strSql = strSql & " WHERE convert(nvarchar,DFFinalizadora_sangria_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Controla Vendedor" Then
          strSql = strSql & " WHERE convert(nvarchar,DFControla_vendedor_TBParametros_ecf) = '" & strControla_Vendedor & "'"
       ElseIf cbbCampos.Text = "Código da Empresa" Then
          strSql = strSql & " WHERE convert(nvarchar,FKCodigo_TBEmpresa) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Nome Empresa" Then
          strSql = strSql & " WHERE convert(nvarchar,DFNome_TBEmpresa) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Prox. Orç. Balcão" Then
          strSql = strSql & " WHERE convert(nvarchar,DFProximo_orcamento_balcao_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Prox. Série Orç. Balcão" Then
          strSql = strSql & " WHERE convert(nvarchar,DFProximo_serie_orcamento_balcao_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       End If
    End If
                             
    frmAguarde.Show
    DoEvents
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, HfgParametros_Ecf, strTamanho, strNomes, "BDRetaguarda", "Otica", Me, "S"
    
    If HfgParametros_Ecf.Rows > 1 And HfgParametros_Ecf.Text <> Empty Then
       For I = 1 To HfgParametros_Ecf.Rows - 1
           HfgParametros_Ecf.Row = I
           HfgParametros_Ecf.Col = 2
           If HfgParametros_Ecf.Text = "I" Then
              HfgParametros_Ecf.Text = "Inteiro"
           Else
              HfgParametros_Ecf.Text = "Fração"
           End If
           HfgParametros_Ecf.Col = 4
           If HfgParametros_Ecf.Text = "$" Then
              HfgParametros_Ecf.Text = "$ - Dinheiro"
           Else
              HfgParametros_Ecf.Text = "% - Percentagem"
           End If
           HfgParametros_Ecf.Col = 6
           If HfgParametros_Ecf.Text = "Não" Then
              HfgParametros_Ecf.Text = "0 - Preço"
           Else
              HfgParametros_Ecf.Text = "1 - Peso"
           End If
       Next I
    Else
       'Removendo linhas do grid, evitando assim que fiquem linhas em branco.
       HfgParametros_Ecf.ClearStructure
       Do While HfgParametros_Ecf.Rows <= HfgParametros_Ecf.Rows + 1
          HfgParametros_Ecf.Col = 1
          If HfgParametros_Ecf.Text = "" And HfgParametros_Ecf.Rows = 2 Then
             Exit Do
          End If
          HfgParametros_Ecf.Row = HfgParametros_Ecf.Rows - 1
          HfgParametros_Ecf.RemoveItem HfgParametros_Ecf.Rows - 1
       Loop
        
       Movimentacoes.Monta_HFlex_Grid HfgParametros_Ecf, strTamanho, strNomes, "16", "Otica", Me
    End If
    
    Unload frmAguarde
    
    HfgParametros_Ecf.Refresh
    HfgParametros_Ecf.SetFocus
End Function

Private Function Monta_DataCombo()
    
    strSql = Empty
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = Empty
    strSql = "SELECT * FROM TBProduto"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto_Associado_Desconto, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = Empty
    strSql = "SELECT * FROM TBProduto"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto_Associado_Desconto, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = Empty
    strSql = "SELECT * FROM TBProduto"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto_Associado_Taxa, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = Empty
    strSql = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_Abertura, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = Empty
    strSql = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_Sangria, strSql, "BDRetaguarda", "Otica", Me
    
End Function

Private Sub txtFinalizadora_Abertura_Change()
    dtcFinalizadora_Abertura.BoundText = txtFinalizadora_Abertura.Text
End Sub

Private Sub txtFinalizadora_Abertura_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFinalizadora_Sangria_Change()
    dtcFinalizadora_Sangria.BoundText = txtFinalizadora_Sangria.Text
End Sub

Private Sub txtFinalizadora_Sangria_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPercentual_Taxa_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPercentual_Taxa_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPercentual_Taxa_LostFocus()
    txtPercentual_Taxa.Text = Format(txtPercentual_Taxa, "#,###0.00")
End Sub

Private Sub txtProduto_Associado_Desconto_Change()
    dtcProduto_Associado_Desconto.BoundText = txtProduto_Associado_Desconto.Text
End Sub

Private Sub txtProduto_Associado_Desconto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtProduto_Associado_Taxa_Change()
    dtcProduto_Associado_Taxa.BoundText = txtProduto_Associado_Taxa.Text
End Sub

Private Sub txtProduto_Associado_Taxa_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub
