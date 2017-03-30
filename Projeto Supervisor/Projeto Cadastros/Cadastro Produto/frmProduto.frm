VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmProduto 
   Caption         =   "Produto"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstProduto 
      Height          =   5325
      Left            =   0
      TabIndex        =   18
      Top             =   330
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   9393
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
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "frmProduto.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(14)=   "Label16"
      Tab(0).Control(15)=   "Label15"
      Tab(0).Control(16)=   "txtId_Tributacao"
      Tab(0).Control(17)=   "txtImagem"
      Tab(0).Control(18)=   "Picture1"
      Tab(0).Control(19)=   "dtcTributacao"
      Tab(0).Control(20)=   "dtpInicio_Promocao"
      Tab(0).Control(21)=   "dtpFim_Promocao"
      Tab(0).Control(22)=   "txtCodigo"
      Tab(0).Control(23)=   "txtDescricao"
      Tab(0).Control(24)=   "txtDescricao_Resumida"
      Tab(0).Control(25)=   "cbbCST2"
      Tab(0).Control(26)=   "cbbUnidade"
      Tab(0).Control(27)=   "cbbTipo_Preco"
      Tab(0).Control(28)=   "txtPreco_Promocao"
      Tab(0).Control(29)=   "txtPreco_Venda"
      Tab(0).Control(30)=   "txtCusto_Real"
      Tab(0).Control(31)=   "txtCusto_Contabil"
      Tab(0).Control(32)=   "txtCusto_Medio"
      Tab(0).Control(33)=   "cbbCST1"
      Tab(0).Control(34)=   "txtCodigo_Categoria"
      Tab(0).Control(35)=   "dtcCategoria"
      Tab(0).Control(36)=   "txtCodigo_Tributacao"
      Tab(0).Control(37)=   "CommonDialog1"
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "Código de Barras"
      TabPicture(1)   =   "frmProduto.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label19"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtCodigo_barras"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtCodigo2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdIncluir"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdRemover"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtDescricao2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "hfgCodigo_barra"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Listagem"
      TabPicture(2)   =   "frmProduto.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "adgProduto"
      Tab(2).Control(2)=   "txtConsulta"
      Tab(2).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgCodigo_barra 
         Height          =   3165
         Left            =   120
         TabIndex        =   49
         Top             =   2010
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BandDisplay     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox txtDescricao2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1290
         MaxLength       =   60
         TabIndex        =   48
         Top             =   780
         Width           =   5175
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
         Height          =   375
         Left            =   5310
         TabIndex        =   47
         Top             =   1470
         Width           =   1185
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3930
         TabIndex        =   46
         Top             =   1470
         Width           =   1185
      End
      Begin VB.TextBox txtCodigo2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   4
         TabIndex        =   43
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo_barras 
         Height          =   375
         Left            =   120
         MaxLength       =   60
         TabIndex        =   42
         Top             =   1470
         Width           =   3645
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -70560
         Top             =   2820
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtConsulta 
         Height          =   375
         Left            =   -74880
         TabIndex        =   39
         Top             =   720
         Width           =   6405
      End
      Begin VB.TextBox txtCodigo_Tributacao 
         Height          =   375
         Left            =   -74880
         MaxLength       =   40
         TabIndex        =   3
         Top             =   2100
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dtcCategoria 
         Height          =   360
         Left            =   -73710
         TabIndex        =   2
         Top             =   2730
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   635
         _Version        =   393216
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
      Begin VB.TextBox txtCodigo_Categoria 
         Height          =   375
         Left            =   -74880
         MaxLength       =   40
         TabIndex        =   1
         Top             =   2730
         Width           =   1095
      End
      Begin VB.ComboBox cbbCST1 
         Height          =   360
         Left            =   -74880
         TabIndex        =   7
         Top             =   3390
         Width           =   1725
      End
      Begin VB.TextBox txtCusto_Medio 
         Height          =   360
         Left            =   -72540
         MaxLength       =   40
         TabIndex        =   17
         Top             =   4080
         Width           =   1065
      End
      Begin VB.TextBox txtCusto_Contabil 
         Height          =   360
         Left            =   -73860
         MaxLength       =   40
         TabIndex        =   16
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtCusto_Real 
         Height          =   360
         Left            =   -74880
         MaxLength       =   40
         TabIndex        =   15
         Top             =   4080
         Width           =   945
      End
      Begin VB.TextBox txtPreco_Venda 
         Height          =   360
         Left            =   -74880
         MaxLength       =   40
         TabIndex        =   11
         Top             =   4770
         Width           =   1785
      End
      Begin VB.TextBox txtPreco_Promocao 
         Height          =   360
         Left            =   -73020
         MaxLength       =   40
         TabIndex        =   12
         Top             =   4770
         Width           =   1665
      End
      Begin VB.ComboBox cbbTipo_Preco 
         Height          =   360
         Left            =   -69780
         TabIndex        =   10
         Top             =   4080
         Width           =   1305
      End
      Begin VB.ComboBox cbbUnidade 
         Height          =   360
         Left            =   -71400
         TabIndex        =   9
         Top             =   4080
         Width           =   1545
      End
      Begin VB.ComboBox cbbCST2 
         Height          =   360
         Left            =   -73080
         TabIndex        =   8
         Top             =   3390
         Width           =   4635
      End
      Begin VB.TextBox txtDescricao_Resumida 
         Height          =   375
         Left            =   -74880
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1440
         Width           =   4125
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   5
         Top             =   780
         Width           =   5205
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   0
         Top             =   780
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFim_Promocao 
         Height          =   375
         Left            =   -69810
         TabIndex        =   14
         Top             =   4770
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19529729
         CurrentDate     =   37777
      End
      Begin MSComCtl2.DTPicker dtpInicio_Promocao 
         Height          =   375
         Left            =   -71280
         TabIndex        =   13
         Top             =   4770
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19529729
         CurrentDate     =   37777
      End
      Begin MSDataListLib.DataCombo dtcTributacao 
         Height          =   360
         Left            =   -73710
         TabIndex        =   4
         Top             =   2100
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   635
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid adgProduto 
         Height          =   4005
         Left            =   -74880
         TabIndex        =   40
         Top             =   1200
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   7064
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Enabled         =   0   'False
         Height          =   2085
         Left            =   -70650
         ScaleHeight     =   2025
         ScaleWidth      =   2115
         TabIndex        =   36
         Top             =   1230
         Width           =   2175
      End
      Begin VB.TextBox txtImagem 
         Height          =   375
         Left            =   -70440
         MaxLength       =   60
         TabIndex        =   37
         Top             =   1740
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtId_Tributacao 
         Height          =   375
         Left            =   -69840
         MaxLength       =   40
         TabIndex        =   35
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Barra"
         Height          =   240
         Left            =   120
         TabIndex        =   45
         Top             =   1230
         Width           =   1380
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   44
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tributação"
         Height          =   240
         Left            =   -74880
         TabIndex        =   34
         Top             =   1860
         Width           =   915
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria"
         Height          =   240
         Left            =   -74880
         TabIndex        =   33
         Top             =   2490
         Width           =   825
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Custo Médio"
         Height          =   240
         Left            =   -72510
         TabIndex        =   32
         Top             =   3810
         Width           =   1050
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Custo Real"
         Height          =   240
         Left            =   -74880
         TabIndex        =   31
         Top             =   3810
         Width           =   915
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Custo Contábil"
         Height          =   240
         Left            =   -73860
         TabIndex        =   30
         Top             =   3810
         Width           =   1230
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fim Promoção"
         Height          =   240
         Left            =   -69780
         TabIndex        =   29
         Top             =   4500
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Início Promoção"
         Height          =   240
         Left            =   -71250
         TabIndex        =   28
         Top             =   4500
         Width           =   1365
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Preço Promoção"
         Height          =   240
         Left            =   -72990
         TabIndex        =   27
         Top             =   4500
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Preço"
         Height          =   240
         Left            =   -69780
         TabIndex        =   26
         Top             =   3810
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Preço Venda"
         Height          =   240
         Left            =   -74880
         TabIndex        =   25
         Top             =   4500
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unidade de Venda"
         Height          =   240
         Left            =   -71400
         TabIndex        =   24
         Top             =   3810
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Situaçao Tributária"
         Height          =   240
         Left            =   -74880
         TabIndex        =   23
         Top             =   3150
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição Resumida"
         Height          =   240
         Left            =   -74880
         TabIndex        =   22
         Top             =   1200
         Width           =   2205
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   -73680
         TabIndex        =   20
         Top             =   540
         Width           =   825
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
         TabIndex        =   19
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Path imagem/Invisivel"
         Height          =   240
         Left            =   -70440
         TabIndex        =   38
         Top             =   1560
         Visible         =   0   'False
         Width           =   1875
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6750
      Top             =   360
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
            Picture         =   "frmProduto.frx":0054
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":036E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":0688
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":0A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":0DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduto.frx":10D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
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
Attribute VB_Name = "frmProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Cadastrar de Produtos                                          '
' Data de Criação........: 30/04/2003                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião                        '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim conexao As New DLLConexao_Sistema.conexao
Public log As New DLLSystemManager.log
Dim strSQL As String
Dim strCod_barra_ant(400) As String
Dim I As Integer


Private Sub adgProduto_DblClick()
   
    On Error Resume Next
    
    Dim strCampos_Grid As String
    Dim strTamanhos_Campos_Grid As String
    
    tlbBotoes.Buttons.Item(1).Enabled = False
    tlbBotoes.Buttons.Item(2).Enabled = True
    tlbBotoes.Buttons.Item(3).Enabled = True
    tlbBotoes.Buttons.Item(4).Enabled = True
    tlbBotoes.Buttons.Item(5).Enabled = False
    
    cmdIncluir.Enabled = True
    Picture1.Enabled = True
    
    txtCodigo.Text = adgProduto.Columns(0).Value
    txtDescricao.Text = adgProduto.Columns(1).Value
    txtCodigo2.Text = txtCodigo.Text
    txtDescricao2.Text = txtDescricao.Text
    txtCodigo_Categoria.Text = adgProduto.Columns(2).Value
    dtcCategoria.Text = adgProduto.Columns(3).Value
    
    txtId_Tributacao.Text = adgProduto.Columns(5).Value
    
    txtCodigo_Tributacao.Text = adgProduto.Columns(6).Value
    dtcTributacao.Text = adgProduto.Columns(7).Value
    txtDescricao_Resumida.Text = adgProduto.Columns(8).Value
    
    cbbCST1.Text = adgProduto.Columns(9).Value
    cbbCST2.Text = adgProduto.Columns(10).Value
        
    If adgProduto.Columns(11).Value = "UND" Then
       cbbUnidade.Text = "Unidade"
    Else
      If adgProduto.Columns(11).Value = "PT" Then
         cbbUnidade.Text = "Pacote"
      Else
        If adgProduto.Columns(11).Value = "FD" Then
           cbbUnidade.Text = "Fardo"
        Else
          If adgProduto.Columns(11).Value = "KG" Then
             cbbUnidade.Text = "Kilo"
          End If
        End If
      End If
    End If
    
    txtPreco_Venda.Text = adgProduto.Columns(12).Value
    cbbTipo_Preco.Text = adgProduto.Columns(13).Value
    txtPreco_Promocao.Text = adgProduto.Columns(14).Value
    dtpInicio_Promocao.Value = adgProduto.Columns(15).Value
    dtpFim_Promocao.Value = adgProduto.Columns(16).Value
    txtCusto_Real.Text = adgProduto.Columns(17).Value
    txtCusto_Contabil.Text = adgProduto.Columns(18).Value
    txtCusto_Medio.Text = adgProduto.Columns(19).Value
    Picture1.Picture = LoadPicture(adgProduto.Columns(20).Value)
    txtImagem.Text = adgProduto.Columns(20).Value
    
    strSQL = "SELECT TBCodigo_Barras.DFCodigo_TBCodigo_barras FROM TBCodigo_Barras " & _
             "WHERE TBCodigo_Barras.FKCodigo_TBProduto = " & txtCodigo.Text & " "
             
    strCampos_Grid = "Código de Barras"
    strTamanhos_Campos_Grid = "6000"
    
    Call Movimentacoes.Movimenta_HFlex_Grid(strSQL, hfgCodigo_barra, strTamanhos_Campos_Grid, strCampos_Grid, "BDSupervisor", "PDV", Me)
    
    'Abastece a matriz de codigo de barras para verificação de
    'futuras alterações
    
    hfgCodigo_barra.Row = hfgCodigo_barra.TopRow
    For I = 1 To hfgCodigo_barra.Rows - 1
        hfgCodigo_barra.Row = I
        strCod_barra_ant(I) = hfgCodigo_barra.Text
    Next I
                
    booAlterar = True
    sstProduto.TabEnabled(1) = True
    txtConsulta.Text = Empty
    sstProduto.Tab = 0
    Me.txtDescricao.SetFocus
    
End Sub

Private Sub adgProduto_HeadClick(ByVal ColIndex As Integer)
    strCampo_consulta = adgProduto.Columns(ColIndex).DataField
    txtConsulta.SetFocus
End Sub

Private Sub cmdIncluir_Click()

    Dim strCodigo_Barras As String
        
    'Verifica a digitação de códigos repetidos
    
    For I = 1 To 400
        If strCod_barra_ant(I) = txtCodigo_barras.Text Then
           MsgBox "Este código de barras já foi digitado.", vbInformation, "Produto"
           txtCodigo_barras.Text = Empty
           txtCodigo_barras.SetFocus
           Exit Sub
        End If
    Next I
    
    strCodigo_Barras = txtCodigo_barras.Text
    
    For I = 1 To 400
        If strCod_barra_ant(I) = Empty And txtCodigo_barras.Text <> Empty Then
           strCod_barra_ant(I) = txtCodigo_barras.Text
           txtCodigo_barras.Text = Empty
           Exit For
        End If
    Next I
          
    conexao.Initial_Catalog = "BDSupervisor"
    conexao.Abrir_conexao ("PDV")
    
    strSQL = "INSERT INTO TBCodigo_barras (FKCodigo_TBProduto,DFCodigo_TBCodigo_Barras) " & _
             "VALUES (" & txtCodigo.Text & " , " & strCodigo_Barras & " ) "
                                                                                                                             
    conexao.CNConexao.Execute strSQL
    conexao.Fechar_conexao
    
    hfgCodigo_barra.AddItem strCodigo_Barras, 1
    txtCodigo_barras.SetFocus
       
End Sub

Private Sub cmdRemover_Click()

    If hfgCodigo_barra.Text = Empty Then
       MsgBox "Não a código de barras selecionado.", vbInformation, "Produto"
       Exit Sub
    End If
    
    conexao.Initial_Catalog = "BDSupervisor"
    conexao.Abrir_conexao ("PDV")
    
    strSQL = "DELETE FROM TBCodigo_Barras WHERE DFCodigo_TBCodigo_Barras = " & hfgCodigo_barra.Text & " "
    conexao.CNConexao.Execute strSQL
    
    conexao.Fechar_conexao
    
    For I = 1 To 400
        If strCod_barra_ant(I) = hfgCodigo_barra.Text Then
           strCod_barra_ant(I) = Empty
           Exit For
        End If
    Next I
    
    If hfgCodigo_barra.Rows <= 2 Then
       txtCodigo_barras.Text = Empty
       hfgCodigo_barra.Clear
       hfgCodigo_barra.ColWidth(0) = 6000
       hfgCodigo_barra.Row = 0
       hfgCodigo_barra.Col = 0
       hfgCodigo_barra.Text = "Código de Barras"
    Else
       hfgCodigo_barra.RemoveItem (hfgCodigo_barra.Row)
    End If
                      
    txtCodigo_barras.SetFocus
    
End Sub

Private Sub dtcCategoria_Click(Area As Integer)

    txtCodigo_Categoria.Text = dtcCategoria.BoundText
    
End Sub

Private Sub dtcTributacao_Change()
    
    txtCodigo_Tributacao.Text = dtcTributacao.BoundText
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 78: Call Novo     'CTRL+N
                       Case 71: Call Gravar   'CTRL+G
                       Case 67: Call Cancelar 'CTRL+C
                       Case 69: Call Excluir  'CTRL+E
                       Case 83: Unload Me     'CTRL+S
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
       
    sstProduto.TabEnabled(1) = False
    
    'Informações constantes para o log
    
    'Ver
    log.Data = Date
    
    'Ver
'   strEstacao_log = MDIPrincipal_Cadastro_Base.strEstação
'   strUsuario_log = MDIPrincipal_Cadastro_Base.UsuárioOCX.NomeReduzido
    log.Estacao = "INFO-888"
    log.Usuario = "Adão"
    log.Programa = "Cadastro de Produtos"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando o cadastro de Produtos"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando o log
    log.Gravar_log "PDV", Me
    
    'Define a aparência do grid de Código de Barras
    
    hfgCodigo_barra.ColWidth(0) = 6000
    hfgCodigo_barra.Row = 0
    hfgCodigo_barra.Col = 0
    hfgCodigo_barra.Text = "Código de Barras"
    
    sstProduto.Tab = 2
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = True
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    Call Reposicao
    
    Exit Sub
    
Erro:
    Call Erro.Erro(Me, "PDV", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

     On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    Exit Sub
Erro:

    Call Erro.Erro(Me, "PDV", "Unload")
    Exit Sub
    
End Sub

Private Sub Picture1_DblClick()
    
    CommonDialog1.Filter = "Arquivos de bitmap|*.bmp|GIF|*.gif|Formato de intercâmbio de Arquivo JPEG|*.jpg;*.jpeg"
    CommonDialog1.ShowOpen
    txtImagem.Text = CommonDialog1.FileName
    Picture1.Picture = LoadPicture(txtImagem.Text)
    
End Sub

Private Sub tlbbotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           'Case 5: Call Imprimir
           Case 7: Unload Me
    End Select
End Sub

Function Gravar()

    On Error GoTo Erro
    
    Dim strSet As String
    Dim strCampo As String
    Dim strvalores As String
    Dim strUnidade As String
    Dim strTipo As String
    Dim strData_Ini As String
    Dim strData_Fin As String
    Dim strIDtributacao As String
    
    sstProduto.TabEnabled(1) = False
    Picture1.Enabled = False
    
    If cbbCST1.Text = "0 - Nacional" Then
       cbbCST1.Text = 0
    Else
      If cbbCST1.Text = "1 - Estrangeira - Importação direta" Then
         cbbCST1.Text = 1
      Else
         cbbCST1.Text = 2
      End If
    End If
    
    If cbbCST2.Text = "00 - Tributada integralmente" Then
       cbbCST2.Text = "00"
    Else
      If cbbCST2.Text = "10 - Tributada e com cobrança do ICMS por substituição tributária" Then
         cbbCST2.Text = 10
      Else
        If cbbCST2.Text = "20 - Com redução de base de cálculo" Then
           cbbCST2.Text = 20
        Else
          If cbbCST2.Text = "30 - Isenta ou não tributada e com cobrança do ICMS por substituição tributária" Then
             cbbCST2.Text = 30
          Else
            If cbbCST2.Text = "40 - Isenta" Then
               cbbCST2.Text = 40
            Else
              If cbbCST2.Text = "41 - Não tributada" Then
                 cbbCST2.Text = 41
              Else
                If cbbCST2.Text = "50 - Suspensão" Then
                   cbbCST2.Text = 50
                Else
                  If cbbCST2.Text = "51 - Diferimento" Then
                     cbbCST2.Text = 51
                  Else
                    If cbbCST2.Text = "60 - ICMS cobrado anteriormente por substituição tributária" Then
                       cbbCST2.Text = 60
                    Else
                      If cbbCST2.Text = "70 - Com redução de base de cálculo e cobrança do ICMS por substituição tributária" Then
                         cbbCST2.Text = 70
                      Else
                         cbbCST2.Text = 90
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
        
    If cbbUnidade.Text = "Unidade" Then
       strUnidade = "UND"
    Else
      If cbbUnidade.Text = "Pacote" Then
         strUnidade = "PT"
      Else
        If cbbUnidade.Text = "Fardo" Then
           strUnidade = "FD"
        Else
          If cbbUnidade.Text = "Kilo" Then
             strUnidade = "KG"
          End If
        End If
      End If
    End If
    
    If cbbTipo_Preco.Text = "Normal" Then
       strTipo = 1
    Else
      If cbbTipo_Preco.Text = "Promoção Relâmpago" Then
         strTipo = 2
      Else
        If cbbTipo_Preco.Text = "Panfleto" Then
           strTipo = 3
        Else
          If cbbTipo_Preco.Text = "Televisão" Then
             strTipo = 4
          Else
            strTipo = 5
          End If
        End If
      End If
    End If
    
    'esta função localiza o ID da tributacao.
    'Ela se faz necessaria pois é o ID da tributacao que é gravado no produto,
    'e a função que monta o datacombo nao pega o ID, pega somente o codigo e a descricao.
    
    strIDtributacao = Funcoes_Gerais.Localiza_ID("PKId_TBTributacao", "DFCodigo_Fiscal_TbTributacao", txtCodigo_Tributacao.Text, "TBTributacao", "PDV", Me, "BDSupervisor")
    
    strData_Ini = Format(dtpInicio_Promocao.Value, "YYYYMMDD")
    strData_Fin = Format(dtpFim_Promocao.Value, "YYYYMMDD")
    
    strCampo = "PKCodigo_TBProduto,FKCodigo_TBCategoria,FKId_TBTributacao," & _
               "DFDescricao_TBProduto,DFDescricao_Resumida_TBProduto,DFCst1_TBProduto," & _
               "DFCst2_TBProduto,DFUnidade_Venda_TBProduto,DFPreco_Venda_TBProduto," & _
               "DFTipo_Preco_TBProduto,DFPreco_Promocao_TBProduto," & _
               "DFData_Inicio_Promocao_TBProduto,DFData_Fim_Promocao_TBProduto,DFCusto_Real_TBProduto," & _
               "DFCusto_Contabil_TBProduto,DFCusto_Medio_TBProduto,DFPath_Imagem_TBProduto"
               
    strvalores = " " & txtCodigo.Text & " , " & txtCodigo_Categoria.Text & " , " & strIDtributacao & " , " & _
                 " '" & txtDescricao.Text & "' , '" & txtDescricao_Resumida.Text & "' , " & _
                 " '" & cbbCST1.Text & "' , '" & cbbCST2.Text & "' , '" & strUnidade & "' , " & txtPreco_Venda.Text & " , " & _
                 " '" & strTipo & "' , " & txtPreco_Promocao.Text & " , '" & strData_Ini & "' , " & _
                 " '" & strData_Fin & "' , " & txtCusto_Real.Text & " , " & txtCusto_Contabil.Text & " , " & _
                 " " & txtCusto_Medio.Text & " , '" & txtImagem.Text & "' "
                 
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET FKCodigo_TBCategoria = " & txtCodigo_Categoria.Text & " , " & _
                "    FKId_TBTributacao = " & txtId_Tributacao.Text & " , " & _
                "    DFDescricao_TBProduto = '" & txtDescricao.Text & "' , " & _
                "    DFDescricao_Resumida_TBProduto = '" & txtDescricao_Resumida.Text & "' , " & _
                "    DFCst1_TBProduto = '" & cbbCST1.Text & "' , " & _
                "    DFCst2_TBProduto = '" & cbbCST2.Text & "' , " & _
                "    DFUnidade_Venda_TBProduto = '" & strUnidade & "' , " & _
                "    DFPreco_Venda_TBProduto = " & txtPreco_Venda.Text & " , " & _
                "    DFTipo_Preco_TBProduto = '" & strTipo & "' , " & _
                "    DFPreco_Promocao_TBProduto = " & txtPreco_Promocao.Text & " , " & _
                "    DFData_Inicio_Promocao_TBProduto = '" & strData_Ini & "' , " & _
                "    DFData_Fim_Promocao_TBProduto = '" & strData_Fin & "' , " & _
                "    DFCusto_Real_TBProduto = " & txtCusto_Real.Text & " , " & _
                "    DFCusto_Contabil_TBProduto = " & txtCusto_Contabil.Text & " , " & _
                "    DFCusto_Medio_TBProduto  = " & txtCusto_Medio.Text & " , " & _
                "    DFPath_Imagem_TBProduto = '" & txtImagem.Text & "' "
       Call funcoes_banco.Alterar("TBProduto", strSet, "PKCodigo_TBProduto", txtCodigo.Text, "PDV", Me)
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "PDV", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBProduto", strCampo, strvalores, "PDV", Me)
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "PDV", Me
    End If
    
    Call Objetos.Limpa_TXT(Me)
    cbbCST1.Text = Empty
    cbbCST2.Text = Empty
    cbbTipo_Preco.Text = Empty
    cbbUnidade.Text = Empty
    dtcCategoria.Text = Empty
    dtcTributacao.Text = Empty
    'limpa o HFlexGrid
    hfgCodigo_barra.Clear
    Picture1.Picture = LoadPicture(txtImagem.Text)
    Picture1.Refresh
    
    Call Reposicao
    
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = True
    tlbBotoes.Buttons.Item(5).Enabled = True
        
    Exit Function
    
Erro:
    Call Erro.Erro(Me, "PDV", "Gravar")
    Exit Function
    
End Function

Private Function Excluir()

    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
           
    If hfgCodigo_barra.Rows >= 2 Then
       conexao.Initial_Catalog = "BDSupervisor"
       conexao.Abrir_conexao ("PDV")
       strSQL = "DELETE FROM TBCodigo_Barras WHERE FKCodigo_TBProduto = " & txtCodigo.Text & " "
       conexao.CNConexao.Execute strSQL
       conexao.Fechar_conexao
    End If
     
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBProduto", "PKCodigo_TBProduto", adgProduto.Columns(0).Value, "PDV", Me, "BDSupervisor")
    
    'Gravando log
    log.Gravar_log "PDV", Me
    
    For I = 1 To 400
        strCod_barra_ant(I) = Empty
    Next I
           
    Call Objetos.Limpa_TXT(Me)
    hfgCodigo_barra.Clear
    cbbCST1.Text = Empty
    cbbCST2.Text = Empty
    cbbTipo_Preco.Text = Empty
    cbbUnidade.Text = Empty
    dtcCategoria.Text = Empty
    dtcTributacao.Text = Empty
    Picture1.Picture = LoadPicture(txtImagem.Text)
    Picture1.Refresh
    
    Call Reposicao
    
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = True
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    Call Reposicao
    
    Exit Function
    
Erro:

    Call Erro.Erro(Me, "PDV", "Excluir")
    Exit Function

End Function
Private Function Cancelar()

    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    cbbCST1.Text = Empty
    cbbCST2.Text = Empty
    cbbTipo_Preco.Text = Empty
    cbbUnidade.Text = Empty
    dtcCategoria.Text = Empty
    dtcTributacao.Text = Empty
    Picture1.Enabled = False
    sstProduto.TabEnabled(1) = False
    
    'limpa o Hflexgrid
    hfgCodigo_barra.Clear
    
    Picture1.Picture = LoadPicture(txtImagem.Text)
    Picture1.Refresh
       
    'Inserir log
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = True
    tlbBotoes.Buttons.Item(5).Enabled = True
    txtCodigo.Enabled = False
    txtCodigo_Categoria.SetFocus
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "PDV", "Cancelar")
    Exit Function

End Function

Private Function Novo()

    On Error GoTo Erro
    
    sstProduto.Tab = 0
    Call Objetos.Limpa_TXT(Me)
    Picture1.Enabled = True
    cmdIncluir.Enabled = True
    cbbCST1.Text = Empty
    cbbCST2.Text = Empty
    cbbTipo_Preco.Text = Empty
    cbbUnidade.Text = Empty
    dtcCategoria.Text = Empty
    dtcTributacao.Text = Empty
    'txtImagem.Text = "C:\Projetos\Sistemas\PDV\Figuras\Imagem_Padrao_Produto.JPG"
    Picture1.Picture = LoadPicture(txtImagem.Text)
    Picture1.Refresh
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    tlbBotoes.Buttons.Item(1).Enabled = False
    tlbBotoes.Buttons.Item(2).Enabled = True
    tlbBotoes.Buttons.Item(3).Enabled = True
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = False
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    booAlterar = False
    
    Exit Function
    
Erro:

    Call Erro.Erro(Me, "PDV", "Novo")
    Exit Function

End Function



Private Sub txtCodigo_Categoria_LostFocus()

    dtcCategoria.BoundText = txtCodigo_Categoria.Text
    
End Sub

Private Sub txtCodigo_LostFocus()

    Movimentacoes.Verifica_Numero "PKCodigo_TBProduto", "TBProduto", txtCodigo, "PDV", Me
    
End Sub

Private Function Reposicao()

    On Error GoTo Erro

    Dim strCampos_Grid As String
    Dim strTamanhos_Campos_Grid As String
            
    strSQL = "SELECT TBProduto.PKCodigo_TBProduto,TBProduto.DFDescricao_TBProduto,TBProduto.FKCodigo_TBCategoria," & _
             "TBCategoria.DFDescricao_TBCategoria,TBProduto.FKId_TBTributacao," & _
             "TBTributacao.PKId_TBTributacao,TBTributacao.DFCodigo_Fiscal_TBTributacao," & _
             "TBTributacao.DFDescricao_TBTributacao," & _
             "TBProduto.DFDescricao_Resumida_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto," & _
             "TBProduto.DFUnidade_Venda_TBProduto,TBProduto.DFTipo_Preco_TBProduto," & _
             "TBProduto.DFPreco_Venda_TBProduto,TBProduto.DFPreco_Promocao_TBProduto," & _
             "TBProduto.DFData_Inicio_Promocao_TBProduto," & _
             "TBProduto.DFData_Fim_Promocao_TBProduto,TBProduto.DFCusto_Real_TBProduto," & _
             "TBProduto.DFCusto_Contabil_TBProduto,TBProduto.DFCusto_Medio_TBProduto," & _
             "TBProduto.DFPath_Imagem_TBProduto FROM TBProduto " & _
             "INNER JOIN TBCategoria ON TBProduto.FKCodigo_TBCategoria = TBCategoria.PKCodigo_TBCategoria " & _
             "INNER JOIN TBTributacao ON TBProduto.FKId_TBTributacao = TBTributacao.PKId_TBTributacao"

    If txtConsulta.Text <> Empty Then
       strSQL = strSQL & " WHERE " & strCampo_consulta & " LIKE '" & txtConsulta.Text & "%' "
    End If
        
    strCampos_Grid = "Código,Produto,Código,Categoria,ID,ID,Código,Tributação,Descrição Resumida,CST 1,CST 2,Unidade Venda," & _
                     "Tipo do Preço,Preço de Venda,Preço de Promoçao,Data Início,Data Final," & _
                     "Custo Real,Custo Contábil,Custo Médio,Imagem"
                     
    strTamanhos_Campos_Grid = "1200,2500,1200,2500,0,0,1200,2500,2500,1200,1200,1600,1450,1500,2000,1500,1500,1200,1500,1300,3200"
    
    Movimentacoes.Movimenta_Data_Grid strSQL, adgProduto, strTamanhos_Campos_Grid, strCampos_Grid, "BDSupervisor", "PDV", Me
    
    strSQL = "SELECT TBCategoria.PKCodigo_TBCategoria,TBCategoria.DFDescricao_TBCategoria FROM TBCategoria"
    Call Movimentacoes.Movimenta_DataCombo("PKCodigo_TBCategoria", "DFDescricao_TBCategoria", dtcCategoria, strSQL, "BDSupervisor", "PDV", Me)
    
    strSQL = "SELECT TBTributacao.PKId_TBTributacao,TBTributacao.DFCodigo_Fiscal_TBTributacao,TBTributacao.DFDescricao_TBTributacao FROM TBTributacao"
    Call Movimentacoes.Movimenta_DataCombo("DFCodigo_Fiscal_TBTributacao", "DFDescricao_TBTributacao", dtcTributacao, strSQL, "BDSupervisor", "PDV", Me)
    
    Call Abastece_Combos
    
    Exit Function

Erro:
    Call Erro.Erro(Me, "PDV", "Reposicao")
    Resume Next

End Function

Private Sub txtCodigo_Tributacao_Change()
    
    dtcTributacao.BoundText = txtCodigo_Tributacao.Text
       
End Sub

Private Sub txtConsulta_Change()

   Call Reposicao
   
End Sub

Private Sub txtDescricao_LostFocus()
    
    txtDescricao.Text = UCase(txtDescricao.Text)
    
End Sub

Private Sub txtDescricao_Resumida_LostFocus()
    
    txtDescricao_Resumida.Text = UCase(txtDescricao_Resumida.Text)
    
End Sub
Private Function Abastece_Combos()

    cbbCST1.Clear
    cbbCST1.AddItem ("0 - Nacional")
    cbbCST1.AddItem ("1 - Estrangeira - Importação direta")
    cbbCST1.AddItem ("2 - Estrangeira - Adquirida no mercado interno")
    
    cbbCST2.Clear
    cbbCST2.AddItem ("00 - Tributada integralmente")
    cbbCST2.AddItem ("10 - Tributada e com cobrança do ICMS por substituição tributária")
    cbbCST2.AddItem ("30 - Isenta ou não tributada e com cobrança do ICMS por substituição tributária")
    cbbCST2.AddItem ("40 - Isenta")
    cbbCST2.AddItem ("30 - Isenta ou não tributada e com cobrança do ICMS por substituição tributária")
    cbbCST2.AddItem ("41 - Não tributada")
    cbbCST2.AddItem ("50 - Suspensão")
    cbbCST2.AddItem ("51 - Diferimento")
    cbbCST2.AddItem ("60 - ICMS cobrado anteriormente por substituição tributária")
    cbbCST2.AddItem ("70 - Com redução de base de cálculo e cobrança do ICMS por substituição tributária")
    cbbCST2.AddItem ("90 – Outras")
        
    cbbTipo_Preco.Clear
    cbbTipo_Preco.AddItem ("Normal")
    cbbTipo_Preco.AddItem ("Promoção Relampago")
    cbbTipo_Preco.AddItem ("Panfleto")
    cbbTipo_Preco.AddItem ("Televisão")
    cbbTipo_Preco.AddItem ("Promoção")
    
    cbbUnidade.Clear
    cbbUnidade.AddItem ("Unidade")
    cbbUnidade.AddItem ("Pacote")
    cbbUnidade.AddItem ("Fardo")
    cbbUnidade.AddItem ("Kilo")
        
End Function

