VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmComanda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comanda"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmComanda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7830
   Begin TabDlg.SSTab sstComanda 
      Height          =   6015
      Left            =   0
      TabIndex        =   11
      Top             =   330
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
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
      TabPicture(0)   =   "frmComanda.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTotal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblCupom"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label12"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtpHora_Lancamento"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtcVendedor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtpLancamento"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtNumero_Comanda"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCodigo_Vendedor"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdImprimir"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNumero_Pessoas"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmComanda.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRefresh"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdConsulta"
      Tab(1).Control(2)=   "txtConsulta"
      Tab(1).Control(3)=   "hfgComanda"
      Tab(1).Control(4)=   "cbbcampos"
      Tab(1).Control(5)=   "dtpIni"
      Tab(1).Control(6)=   "dtpFin"
      Tab(1).Control(7)=   "Label29"
      Tab(1).Control(8)=   "lblA"
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtNumero_Pessoas 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   120
         MaxLength       =   40
         TabIndex        =   40
         Text            =   "1"
         Top             =   5550
         Width           =   1425
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
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
         Left            =   1650
         TabIndex        =   39
         Top             =   5550
         Width           =   1155
      End
      Begin VB.TextBox txtCodigo_Vendedor 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   1
         Top             =   780
         Width           =   975
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
         Left            =   -67680
         Picture         =   "frmComanda.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   810
         Width           =   375
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
         Left            =   -68070
         Picture         =   "frmComanda.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Consultar"
         Top             =   810
         Width           =   375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72960
         TabIndex        =   18
         Top             =   810
         Width           =   4845
      End
      Begin VB.TextBox txtNumero_Comanda 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   780
         Width           =   1245
      End
      Begin VB.Frame Frame1 
         Caption         =   "Informações para troca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         Left            =   120
         TabIndex        =   12
         Top             =   1770
         Width           =   7545
         Begin VB.CommandButton cmdConsulta_Cliente 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7050
            Picture         =   "frmComanda.frx":44F6
            Style           =   1  'Graphical
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Consulta detalhada do produto "
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtDescricao_Produto 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   2340
            TabIndex        =   42
            Top             =   570
            Width           =   4665
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
            Left            =   6390
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1035
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
            Left            =   5190
            TabIndex        =   7
            Top             =   1170
            Width           =   1035
         End
         Begin VB.TextBox txtTotal_Item 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   360
            Left            =   2790
            MaxLength       =   40
            TabIndex        =   6
            Top             =   1170
            Width           =   1425
         End
         Begin VB.TextBox txtPreco 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   360
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   5
            Top             =   1170
            Width           =   1155
         End
         Begin VB.TextBox txtCodigo_Produto 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   120
            MaxLength       =   13
            TabIndex        =   3
            Top             =   570
            Width           =   2175
         End
         Begin VB.TextBox txtQuantidade 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   120
            MaxLength       =   40
            TabIndex        =   4
            Top             =   1170
            Width           =   975
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgProduto 
            Height          =   1755
            Left            =   120
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1590
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   3096
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            SelectionMode   =   1
            Appearance      =   0
            RowSizingMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
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
            Left            =   2580
            TabIndex        =   38
            Top             =   1240
            Width           =   165
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Item"
            Height          =   240
            Left            =   2790
            TabIndex        =   33
            Top             =   930
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "="
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
            Left            =   2580
            TabIndex        =   32
            Top             =   1230
            Width           =   165
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
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
            Left            =   1170
            TabIndex        =   31
            Top             =   1230
            Width           =   135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preço"
            Height          =   240
            Left            =   1380
            TabIndex        =   30
            Top             =   930
            Width           =   480
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Produto (F2 Consulta)"
            Height          =   240
            Left            =   120
            TabIndex        =   26
            Top             =   330
            Width           =   1875
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade"
            Height          =   240
            Left            =   120
            TabIndex        =   13
            Top             =   930
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker dtpLancamento 
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   1410
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   50921473
         CurrentDate     =   37949
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgComanda 
         Height          =   4635
         Left            =   -74880
         TabIndex        =   14
         Top             =   1200
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   8176
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin AutoCompletar.CbCompleta cbbcampos 
         Height          =   360
         Left            =   -74880
         TabIndex        =   15
         Top             =   810
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
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   360
         Left            =   -72960
         TabIndex        =   16
         Top             =   810
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
         Format          =   50921473
         CurrentDate     =   37949
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   360
         Left            =   -71280
         TabIndex        =   17
         Top             =   810
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
         Format          =   50921473
         CurrentDate     =   37949
      End
      Begin MSDataListLib.DataCombo dtcVendedor 
         Height          =   360
         Left            =   2430
         TabIndex        =   2
         Top             =   780
         Width           =   5280
         _ExtentX        =   9313
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
      Begin MSComCtl2.DTPicker dtpHora_Lancamento 
         Height          =   360
         Left            =   1500
         TabIndex        =   10
         Top             =   1410
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50921474
         CurrentDate     =   37858
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Pessoas"
         Height          =   240
         Left            =   120
         TabIndex        =   41
         Top             =   5310
         Width           =   960
      End
      Begin VB.Label lblCupom 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblCupom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3990
         TabIndex        =   37
         Top             =   1410
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cupom:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2970
         TabIndex        =   36
         Top             =   1410
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   6090
         TabIndex        =   35
         Top             =   5670
         Width           =   1560
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   5520
         TabIndex        =   34
         Top             =   5670
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hora Abertura"
         Height          =   240
         Left            =   1500
         TabIndex        =   28
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         Height          =   240
         Left            =   1410
         TabIndex        =   27
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Comanda"
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   540
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.lançamento"
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   1170
         Width           =   1230
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   20
         Top             =   540
         Width           =   435
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
         Left            =   -71490
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7830
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
            Picture         =   "frmComanda.frx":4880
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComanda.frx":4B9A
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComanda.frx":4EB4
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComanda.frx":524E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComanda.frx":55E8
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComanda.frx":5902
            Key             =   "IMG6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7830
      _ExtentX        =   13811
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
            ImageKey        =   "IMG4"
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
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageKey        =   "IMG2"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Excluir"
            Object.ToolTipText     =   "Excluir registro - CTRL+E"
            ImageKey        =   "IMG6"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Imprimir"
            Object.ToolTipText     =   "Imprimir - CTRL+I"
            ImageKey        =   "IMG3"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sair"
            Object.ToolTipText     =   "Sair - CTRL+S"
            ImageKey        =   "IMG5"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvFiltrar 
      Height          =   480
      Left            =   30
      TabIndex        =   43
      Top             =   6420
      Width           =   3285
      lastProp        =   500
      _cx             =   5794
      _cy             =   847
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Cadastro Base                                                  '
' Objetivo...............: Cadastrar Comanda                                              '
' Data de Criação........: 16/01/2005                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo As String
Dim strTamanho As String
Public strID_Troca As String
Dim strCodigo_Produto_ant(5000) As Integer
Dim strNomes As String
Dim strCombo As String
Dim strConsulta As String
Dim booAlterar As Boolean
Dim I As Integer
Public strSql As String
Dim CNconexao As New DLLConexao_Sistema.conexao
Dim conexao As New DLLConexao_Sistema.conexao
Dim Relatorio As New CRAXDRT.Report
Dim Aplicacao As New CRAXDRT.Application
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
'------------------------------------------------------------
'Declaração da variavel do intercomunicador de mensagens
Private Cliente_mensagem_exe As VetorDeMensagens.ClienteDeMensagens
Dim log As New DLLSystemManager.log

Function Imprimir()
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Relatorio_Comanda.Show
    
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
    ElseIf cbbCampos.Text = "Data Lançamento" Then
       dtpIni.Visible = True
       dtpFin.Visible = True
       lblA.Visible = True
       txtConsulta.Visible = False
       dtpIni.SetFocus
    Else
       dtpIni.Visible = False
       dtpFin.Visible = False
       lblA.Visible = False
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub


Private Sub cmdConsulta_Cliente_Click()
    frmConsulta_Produto_Comanda.Show
End Sub

Private Sub cmdImprimir_Click()
        
    Dim strValor_Pessoa As String
    Dim strNumero_Pessoas As String
    Dim rstImprime As New ADODB.Recordset
    Dim strData As String
          
    If Len(txtNumero_Pessoas.Text) = 0 Or txtNumero_Pessoas.Text = "" Then
       MsgBox "Digite o nº de Pessoas antes de imprimir.", vbInformation, "Only Tech"
       txtNumero_Pessoas.SetFocus
       Exit Sub
    End If
    
    'Abrindo uma conexão nova
    CNconexao.Banco = "BDRetaguarda"
    CNconexao.Abrir_conexao "Otica"
        
    'Deletando Itens para gravação
    Call funcoes_banco.Excluir("TBItens_comanda", "FKCodigo_TBComanda", txtNumero_Comanda.Text, "Otica", Me, "BDRetaguarda")
    
    Dim strQuantidade_Incluir As String
    Dim strPreco_Incluir As String
    Dim strTotal_Item_Incluir As String
                                    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    'Gravando os itens que estão no GRID no momento'
    ''''''''''''''''''''''''''''''''''''''''''''''''
    
    If hfgProduto.Rows > 1 Then
       For I = 1 To hfgProduto.Rows - 1
           hfgProduto.Row = I
           hfgProduto.Col = 2
           If hfgProduto.Text = Empty Then
              Exit For
           End If
           strID_Produto = Funcoes_Gerais.Localiza_ID("PKID_TBProduto", "IXCodigo_TBProduto", hfgProduto.Text, "TBproduto", "Otica", Me, "BDRetaguarda")
           hfgProduto.Col = 4
           strQuantidade_Incluir = hfgProduto.Text
           hfgProduto.Col = 5
           strPreco_Incluir = hfgProduto.Text
           hfgProduto.Col = 6
           strTotal_Item_Incluir = hfgProduto.Text
           conexao.Initial_Catalog = "BDRetaguarda"
           conexao.Abrir_conexao ("Otica")
           strSql = "INSERT INTO TBItens_comanda (FKCodigo_TBComanda,FKId_TBProduto,DFQuantidade_TBItens_comanda," & _
                    "DFPreco_TBItens_comanda,DFValor_total_TBItens_comanda ) " & _
                    "VALUES (" & txtNumero_Comanda.Text & "," & strID_Produto & "," & _
                    " " & Funcoes_Gerais.Grava_Moeda(strQuantidade_Incluir) & "," & _
                    " " & Funcoes_Gerais.Grava_Moeda(strPreco_Incluir) & "," & _
                    " " & Funcoes_Gerais.Grava_Moeda(strTotal_Item_Incluir) & ") "
           conexao.CNconexao.Execute strSql
           conexao.Fechar_conexao
       Next I
    End If
        
    CNconexao.CNconexao.BeginTrans
    CNconexao.CNconexao.CommitTrans
    
    '''''''''''''''''''''''''''''''
    'Calculando o valor por pessoa'
    '''''''''''''''''''''''''''''''
    strValor_Pessoa = CDbl(lblTotal.Caption) / CDbl(txtNumero_Pessoas.Text)
    strValor_Pessoa = Format(strValor_Pessoa, "#,###0.00")
    strNumero_Pessoas = txtNumero_Pessoas.Text
    
        
    'Impressão no Crystal-------------------------------------------------
        
'    Dim intTamanho_string As Integer
'    Dim inttamanho_From As Integer
'    Dim strCaminho As String
'    Dim strSql_antes_from As String
'    Dim strSql_pos_from As String
'    Dim strRemontada_sql As String
'    Dim strNome_cliente As String
'    Dim adrImprime As New ADODB.Recordset
'
'    'On Error GoTo Erro
'
'    'Inserindo a hora no nome da tabela
'    Tabela = "TBTEMP_RELATORIO" & time
'    Tabela = Replace(Tabela, ":", "_")
'
'    'Montando a nova string  de SQL com o INTO para criação da tabela temporária
'    intTamanho_string = Len(strSql)
'    inttamanho_From = InStr(1, strSql, "FROM")
'    strSql_antes_from = Mid(strSql, 1, inttamanho_From - 1)
'    strSql_pos_from = Mid(strSql, inttamanho_From, intTamanho_string)
'    strRemontada_sql = strSql_antes_from + "INTO " & Tabela & " " + strSql_pos_from
'
'    'On Error GoTo Erro
'
'    CNconexao.CNconexao.Execute strRemontada_sql
'
'    'Abrindo a recordset com as informações da tabela temporaria
'    adrImprime.Open "SELECT * FROM " & Tabela & "", CNconexao.CNconexao, adOpenKeyset, adLockOptimistic
'
'    strCaminho = Funcoes_Gerais.Abrir_relatorio_registro("Otica", Me) & "\rptEmissao_comanda.rpt"
'
'    DoEvents
'
'    Set Relatorio = Aplicacao.OpenReport(strCaminho)
'
'    Relatorio.Database.Tables.Item(1).SetDataSource adrImprime, 3
'    Relatorio.FormulaFields.GetItemByName("Divide_Conta").Text = "'" + strValor_Pessoa & "'"
'    Relatorio.FormulaFields.GetItemByName("Calculo_Divide").Text = "'" + strNumero_Pessoas & "'"
'
'    Relatorio.DiscardSavedData
'
'    'Indica que a impresão é direta para a impressora
'    Relatorio.PrintOut False
'
'    crvFiltrar.ReportSource = Relatorio
'    crvFiltrar.Refresh
'    crvFiltrar.ViewReport
'
'    Set adrImprime = Nothing
'    Set Aplicacao = Nothing
'
'    CNconexao.CNconexao.Execute "DROP TABLE " & Tabela & " "
'    Set Relatorio = Nothing
'

    'Impressão ---------------------------------------------------------------------------------------------
    
    strSql = Empty
    strSql = "SELECT TBItens_comanda.PKId_TBItens_comanda, " & _
             "TBProduto.IXCodigo_TBProduto, " & _
             "TBProduto.DFDescricao_resumida_TBProduto, " & _
             "TBItens_comanda.DFQuantidade_TBItens_comanda, " & _
             "TBItens_comanda.DFPreco_TBItens_comanda, " & _
             "TBItens_comanda.DFValor_total_TBItens_comanda," & _
             "TBItens_comanda.FKCodigo_TBComanda," & _
             "TBVendedor.IXCodigo_TBVendedor," & _
             "TBVendedor.DFNome_TBVendedor," & _
             "TBVendedor.IXCodigo_TBEmpresa," & _
             "TBEmpresa.DFRazao_Social_TBEmpresa " & _
             "FROM TBItens_comanda " & _
             "INNER JOIN TBComanda ON TBComanda.PKCodigo_TBComanda = TBItens_comanda.FKCodigo_TBComanda " & _
             "INNER JOIN TBProduto ON TBProduto.PKId_TBProduto = TBItens_comanda.FKId_TBProduto " & _
             "INNER JOIN TBVendedor ON TBVendedor.PKId_TBVendedor = TBComanda.FKId_TBVendedor " & _
             "INNER JOIN TBEmpresa ON TBVendedor.IXCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
             "WHERE TBItens_comanda.FKCodigo_TBComanda = " & txtNumero_Comanda.Text & " " & _
             "ORDER BY TBProduto.DFDescricao_TBProduto"
    
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstImprime, "Otica", Me)
    
    strData = Date
    
    'Cabeçalho
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    'Empresa
    strLinha_Impressao = rstImprime.Fields("DFRazao_Social_TBEmpresa")
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
    
    'N ° Comanda
    strLinha_Impressao = "COMANDA: " & rstImprime.Fields("FKCodigo_TBComanda") & "     " & strData
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
    
    'Cabeçalho
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    'Cabeçalho 1
    strLinha_Impressao = "PRODUTO (COD.INTERNO) QUANTIDADE   X  VLR.UNT.    TOTAL"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    Dim dblTotal_Comanda As Double
    dblTotal_Comanda = Empty
    
    Do While rstImprime.EOF = False And rstImprime.BOF = False
       Dim strDescricao_Produto As String * 20
       Dim strCodigo_Produto As String * 6
       Dim strQuantidade As String * 6
       Dim strPreco_Unitario As String * 8
       Dim strPreco_Total_Item As String * 10
       
       strCodigo_Produto = rstImprime.Fields("IXCodigo_TBProduto")
       strDescricao_Produto = rstImprime.Fields("DFDescricao_resumida_TBProduto")
       strQuantidade = Format(rstImprime.Fields("DFQuantidade_TBItens_comanda"), "#,###0.00")
       strPreco_Unitario = Format(rstImprime.Fields("DFPreco_TBItens_comanda"), "#,###0.00")
       strPreco_Total_Item = Format(rstImprime.Fields("DFValor_total_TBItens_comanda"), "#,###0.00")
                    
       'Linha 1
       strLinha_Impressao = strCodigo_Produto & " " & strDescricao_Produto & " " & strQuantidade & " X  " & strPreco_Unitario & " =  " & strPreco_Total_Item
       sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
       iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
       
       dblTotal_Comanda = dblTotal_Comanda + rstImprime.Fields("DFValor_total_TBItens_comanda")
       
       rstImprime.MoveNext
    Loop
    
    Set rstImprime = Nothing
    
    'Salto
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    'Rodapé Total
    strLinha_Impressao = "            Total.: " & Format(dblTotal_Comanda, "#,###0.00")
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 1, 1)
    
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    'Rodapé - Mensagem
    strLinha_Impressao = "Obrigado pela preferência.Volte Sempre!"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 1)
    
    'Salto
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    'Rodapé - Mensagem 2
    strLinha_Impressao = "Este documento não tem validade fiscal"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 1)
    
    'Salto
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    strLinha_Impressao = "  "
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 1, 1)
    
    CNconexao.Fechar_conexao
    
End Sub

Private Sub cmdIncluir_Click()
    Dim strCodigo_Produto As String
    Dim strDescricao_Produto As String
    Dim strIndice As String
    
    If txtQuantidade.Text = Empty Or txtQuantidade.Text = "0,00" Or txtQuantidade.Text = "0" Then
       MsgBox "Informe uma quantidade.", vbInformation, "Only Tech"
       txtQuantidade.Text = "0,00"
       txtQuantidade.SetFocus
       Exit Sub
    End If
                       
    If booAlterar = False Then
       strIndice = hfgProduto.Rows - 1
    Else
       strIndice = hfgProduto.Rows
    End If
    I = strIndice
    
    hfgProduto.AddItem strIndice + Chr$(9) + Chr$(9) + txtCodigo_Produto.Text + Chr$(9) + txtDescricao_Produto.Text + Chr$(9) + txtQuantidade.Text + Chr$(9) + txtPreco.Text + Chr$(9) + txtTotal_Item.Text, I
    
    lblTotal.Caption = CDbl(lblTotal.Caption) + CDbl(txtTotal_Item.Text)
    lblTotal.Caption = Format(lblTotal.Caption, "#,###0.00")
    
    Call Colori_Grid
    
    txtCodigo_Produto.Text = Empty
    txtDescricao_Produto.Text = Empty
    txtQuantidade.Text = "0,00"
    txtPreco.Text = "0,00"
    txtTotal_Item.Text = "0,00"
                    
    hfgProduto.Refresh
    
    hfgProduto.TopRow = hfgProduto.Rows - 1
       
    txtCodigo_Produto.SetFocus
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub cmdRemover_Click()
    Dim strID_Produto As String
    Dim strCodigo_Produto As String
    Dim strCodigo_Tabela As String
    Dim strPreco_Item As String
    
    If hfgProduto.Text = Empty Then
       MsgBox "Não a Produto selecionado.", vbInformation, "Only Tech"
       Exit Sub
    End If
    
    strPreco_Item = hfgProduto.TextArray((hfgProduto.Row * hfgProduto.Cols + hfgProduto.Col + 6))
    
    lblTotal.Caption = CDbl(lblTotal.Caption) - CDbl(strPreco_Item)
    lblTotal.Caption = Format(lblTotal.Caption, "#,###0.00")
    
    If hfgProduto.Rows <= 2 Then
       txtCodigo_Produto.Text = Empty
       txtQuantidade.Text = Empty
       txtPreco.Text = Empty
       dtcProduto.Text = Empty
       txtTotal_Item.Text = Empty
       hfgProduto.Clear
       strCampo = "ID,Produto,Descrição,Quant.,Preço,Total Item,NºComanda"
       strTamanho = "0,800,2580,900,1000,1200,0"
       Call Movimentacoes.Monta_HFlex_Grid(hfgProduto, strTamanho, strCampo, 7, "Otica", Me)
    Else
       hfgProduto.RemoveItem (hfgProduto.Row)
    End If
    
    hfgProduto.Refresh
    
    txtCodigo_Produto.Text = Empty
    txtQuantidade.Text = Empty
    txtPreco.Text = Empty
    'dtcProduto.Text = Empty
    txtTotal_Item.Text = Empty

    txtCodigo_Produto.SetFocus
End Sub

'Private Sub dtcProduto_LostFocus()
'    If dtcProduto.Text <> Empty Then
'       txtCodigo_Produto.Text = dtcProduto.BoundText
'       Dim rstBusca_Preco As New ADODB.Recordset
'       Dim rstBusca_Paramentro As New ADODB.Recordset
'       Dim strTabela_Vigente As String
'       Dim strID_Produto As String
'       strSql = Empty
'       strSql = "SELECT * FROM TBParametros_Venda WHERE IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ""
'       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Paramentro, "Otica", Me)
'       strTabela_Vigente = rstBusca_Paramentro.Fields("DFNumero_tabela_vigente_TBParametros_venda")
'       Set rstBusca_Paramentro = Nothing
'       strID_Produto = Funcoes_Gerais.Localiza_ID("PKID_TBProduto", "IXCodigo_TBProduto", txtCodigo_Produto.Text, "TBproduto", "Otica", Me, "BDRetaguarda")
'       strSql = Empty
'       strSql = "SELECT * FROM TBItens_tabela_preco WHERE FKCodigo_TBTabela_preco = " & strTabela_Vigente & "AND " & _
'                "FKId_TBProduto = " & strID_Produto & ""
'       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Preco, "Otica", Me)
'       If rstBusca_Preco.RecordCount = 0 Then
'          MsgBox "Este Produto não possui preço de varejo cadastrado. Verifique.", vbInformation, "Only Tech"
'          txtCodigo_Produto.Text = Empty
'          dtcProduto.Text = Empty
'          txtPreco.Text = Empty
'          txtTotal_Item.Text = Empty
'          txtQuantidade.Text = Empty
'          txtCodigo_Produto.SetFocus
'          Exit Sub
'       End If
'       txtPreco.Text = Format(rstBusca_Preco.Fields("DFPreco_varejo_TBItens_tabela_preco"), "#,###0.00")
'       txtQuantidade.SetFocus
'    End If
'End Sub

Private Sub dtpLancamento_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtcVendedor_GotFocus()
    If txtCodigo_Vendedor.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcVendedor)
    End If
End Sub

Private Sub dtcVendedor_LostFocus()
    If dtcVendedor.Text <> Empty Then
       txtCodigo_Vendedor.Text = dtcVendedor.BoundText
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
    If KeyCode = 113 Then
       frmConsulta_Produto_Comanda.Show
    End If
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
    log.Programa = "Cadastro de Comanda"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, sstComanda, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando Cadastro de Comanda"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    dtpLancamento.Value = Date
    dtpHora_Lancamento.Value = Format(Now, "hh:mm:ss")
    lblTotal.Caption = "0,00"
    txtQuantidade.Text = "0,00"
    txtPreco.Text = "0,00"
    txtTotal_Item.Text = "0,00"
    
    strCampo = "ID,Produto,Descrição,Quant.,Preço,Total Item,NºComanda"
    strTamanho = "0,800,2580,900,1000,1200,0"
    Call Movimentacoes.Monta_HFlex_Grid(hfgProduto, strTamanho, strCampo, 7, "Otica", Me)
       
    sstComanda.TabEnabled(0) = False
    sstComanda.Tab = 1
            
    Call Reposicao
    
    '-------------------------------------------------------------------------------------------------------
    'Abrindo Impressora não fiscal
    '-------------------------------------------------------------------------------------------------------
    Dim intPorta As Integer
    Dim strComunica As String
    
    ' Fecha a porta que está aberta
    intPorta = FechaPorta()
    If intPorta <= 0 Then
       MsgBox "Problemas ao Fechar a Porta de Comunicação com a imp. não fiscal.Reinicie a aplicação", vbCritical, "Only Tech"
    End If

    ' Abre a porta de comunicacao com imp. não fiscal
    intPorta = IniciaPorta("LPT1")
    If intPorta <= 0 Then
       MsgBox "Problemas ao Abrir a Porta de Comunicação com a imp. não fiscal.Reinicie a aplicação", vbCritical, "Only Tech"
    End If
           
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
              
    strEvento_log = "Unload"
    
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    
    strCombo = Empty
        
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Unload")
    Exit Sub
End Sub

Private Sub hfgComanda_Click()
    If hfgComanda.Col = 0 Then
        
       On Error Resume Next
              
       Dim rstBaixa As New ADODB.Recordset
       
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
        
       txtNumero_Comanda.Text = hfgComanda.TextArray((hfgComanda.Row * hfgComanda.Cols + hfgComanda.Col + 1))
       txtCodigo_Vendedor.Text = hfgComanda.TextArray((hfgComanda.Row * hfgComanda.Cols + hfgComanda.Col + 2))
       dtpLancamento.Value = hfgComanda.TextArray((hfgComanda.Row * hfgComanda.Cols + hfgComanda.Col + 4))
       dtpHora_Lancamento.Value = hfgComanda.TextArray((hfgComanda.Row * hfgComanda.Cols + hfgComanda.Col + 5))
                          
       Label10.Visible = True
       lblCupom.Visible = True
       If Trim(hfgComanda.TextArray((hfgComanda.Row * hfgComanda.Cols + hfgComanda.Col + 6))) <> Empty Then
          lblCupom.ForeColor = &HC0&
          lblCupom.Caption = "FECHADO"
          cmdIncluir.Enabled = False
          cmdRemover.Enabled = False
          cmdImprimir.Enabled = False
       Else
          lblCupom.ForeColor = &H800000
          lblCupom.Caption = "EM ANDAMENTO"
          cmdIncluir.Enabled = True
          cmdRemover.Enabled = True
          cmdImprimir.Enabled = True
       End If
          
       hfgProduto.Clear
       Call Limpa_Grid
       
       Dim rstBusca_Itens As New ADODB.Recordset
       
       strSql = "SELECT TBItens_comanda.PKId_TBItens_comanda, " & _
                "TBProduto.IXCodigo_TBProduto, " & _
                "TBProduto.DFDescricao_TBProduto, " & _
                "TBItens_comanda.DFQuantidade_TBItens_comanda, " & _
                "TBItens_comanda.DFPreco_TBItens_comanda, " & _
                "TBItens_comanda.DFValor_total_TBItens_comanda," & _
                "TBItens_comanda.FKCodigo_TBComanda " & _
                "FROM TBItens_comanda " & _
                "INNER JOIN TBProduto ON TBProduto.PKId_TBProduto = TBItens_comanda.FKId_TBProduto " & _
                "WHERE TBItens_comanda.FKCodigo_TBComanda = " & txtNumero_Comanda.Text & " " & _
                "ORDER BY TBProduto.DFDescricao_TBProduto"
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Itens, "Otica", Me)
       
       strCampo = "ID,Produto,Descrição,Quant.,Preço,Total Item,NºComanda"
       strTamanho = "0,800,2580,900,1000,1200,0"
       
       frmAguarde.Show
       DoEvents
       If rstBusca_Itens.RecordCount <> 0 Then
          Call Movimentacoes.Movimenta_HFlex_Grid(strSql, hfgProduto, strTamanho, strCampo, "BDRetaguarda", "Otica", Me)
          lblTotal.Caption = rstBusca_Itens.Fields("Total")
       Else
          Call Movimentacoes.Monta_HFlex_Grid(hfgProduto, strTamanho, strCampo, 7, "Otica", Me)
          lblTotal.Caption = "0,00"
       End If
       Set rstBusca_Itens = Nothing
       
       lblTotal.Caption = "0,00"
       If hfgProduto.Rows > 1 Then
          For I = 1 To hfgProduto.Rows - 1
              hfgProduto.Row = I
              hfgProduto.Col = 6
              lblTotal.Caption = CDbl(lblTotal.Caption) + CDbl(hfgProduto.Text)
          Next I
       End If
       lblTotal.Caption = Format(lblTotal.Caption, "#,###0.00")
                         
       Unload frmAguarde
       
       hfgProduto.Refresh
                                                                     
       txtCodigo_Produto.Enabled = True
       dtcProduto.Enabled = True
       cbbUnidade.Enabled = True
       txtQuantidade.Enabled = True
       dtpCadastro.Enabled = False
       
       txtQuantidade.Text = "0,00"
       txtPreco.Text = "0,00"
       txtTotal_Item.Text = "0,00"
                    
       booAlterar = True
       txtConsulta.Text = Empty
       sstComanda.TabEnabled(0) = True
       sstComanda.Tab = 0
       
       txtNumero_Comanda.Enabled = False
       txtCodigo_Vendedor.Enabled = False
       dtcVendedor.Enabled = False
       'cmdImprimir.Enabled = True
                           
    End If
    Unload frmAguarde
End Sub

Private Sub hfgComanda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgComanda_Click
    End If
End Sub


Private Sub sstComanda_Click(PreviousTab As Integer)
    If sstComanda.Tab = 0 Then
       If lblCupom.Visible = True Then
          If lblCupom.Caption = "EM ANDAMENTO" Then
             txtCodigo_Produto.SetFocus
          End If
       End If
    ElseIf sstComanda.Tab = 1 Then
       If strCombo <> Empty And strCombo <> "Todos" Then
          cbbCampos.Text = strCombo
          txtConsulta.SetFocus
       ElseIf strCombo = "Todos" Then
          hfgComanda.Row = 1
          hfgComanda.Col = 0
          hfgComanda.SetFocus
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
    Dim strValores As String
    Dim booVerifica As Boolean
    Dim strID_Vendedor As String
    Dim strID_Produto As String
    Dim strData As String
    Dim strHora As String
    
    Call Objetos.Maiusculo_TXT(Me)
          
    strData = Format(dtpLancamento.Value, "YYYYMMDD")
    strHora = Format(dtpHora_Lancamento.Value, "hh:mm:ss")
    strID_Vendedor = Funcoes_Gerais.Localiza_ID("PKID_TBVendedor", "IXCodigo_TBVendedor", txtCodigo_Vendedor.Text, "TBVendedor", "Otica", Me, "BDRetaguarda")
               
    If booAlterar = False Then
       strCampo = "PKCodigo_TBComanda,FKId_TBVendedor,DFData_lancamento_TBComanda,DFHora_abertura_TBComanda"
                    
       strValores = " " & txtNumero_Comanda.Text & "," & strID_Vendedor & ",'" & strData & "'," & _
                    "'" & strHora & "'"
       Call funcoes_banco.Gravar("TBComanda", strCampo, strValores, "Otica", Me, "BDRetaguarda")
       log.Descricao = "Gravando o registro: " + txtNumero_Comanda.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "Otica", Me
    End If
                    
    'Deletando Itens para gravação
    Call funcoes_banco.Excluir("TBItens_comanda", "FKCodigo_TBComanda", txtNumero_Comanda.Text, "Otica", Me, "BDRetaguarda")
    
    Dim strQuantidade_Incluir As String
    Dim strPreco_Incluir As String
    Dim strTotal_Item_Incluir As String
                                
    If hfgProduto.Rows > 1 Then
       For I = 1 To hfgProduto.Rows - 1
           hfgProduto.Row = I
           hfgProduto.Col = 2
           If hfgProduto.Text = Empty Then
              Exit For
           End If
           strID_Produto = Funcoes_Gerais.Localiza_ID("PKID_TBProduto", "IXCodigo_TBProduto", hfgProduto.Text, "TBproduto", "Otica", Me, "BDRetaguarda")
           hfgProduto.Col = 4
           strQuantidade_Incluir = hfgProduto.Text
           hfgProduto.Col = 5
           strPreco_Incluir = hfgProduto.Text
           hfgProduto.Col = 6
           strTotal_Item_Incluir = hfgProduto.Text
           conexao.Initial_Catalog = "BDRetaguarda"
           conexao.Abrir_conexao ("Otica")
           strSql = "INSERT INTO TBItens_comanda (FKCodigo_TBComanda,FKId_TBProduto,DFQuantidade_TBItens_comanda," & _
                    "DFPreco_TBItens_comanda,DFValor_total_TBItens_comanda ) " & _
                    "VALUES (" & txtNumero_Comanda.Text & "," & strID_Produto & "," & _
                    " " & Funcoes_Gerais.Grava_Moeda(strQuantidade_Incluir) & "," & _
                    " " & Funcoes_Gerais.Grava_Moeda(strPreco_Incluir) & "," & _
                    " " & Funcoes_Gerais.Grava_Moeda(strTotal_Item_Incluir) & ") "
           conexao.CNconexao.Execute strSql
           conexao.Fechar_conexao
       Next I
    End If
            
    Call Objetos.Limpa_TXT(Me)
                 
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    
    If booPrivilegio_Consultar = False Then
       Me.hfgComanda.Visible = False
    End If
    
    sstComanda.TabEnabled(0) = False
    sstComanda.Tab = 1
           
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
              
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtNumero_Comanda.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "Otica", Me
           
    Call funcoes_banco.Excluir("TBItens_comanda", "FKCodigo_TBComanda", txtNumero_Comanda.Text, "Otica", Me, "BDRetaguarda")
    Call funcoes_banco.Excluir("TBComanda", "PKCodigo_TBComanda", txtNumero_Comanda.Text, "Otica", Me, "BDRetaguarda")
       
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
       hfgComanda.Visible = False
    End If
     
    sstComanda.TabEnabled(0) = False
    sstComanda.Tab = 1
        
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Excluir")
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Grid
                          
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
       hfgComanda.Visible = False
    End If
    
    sstComanda.TabEnabled(0) = False
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    
    sstComanda.Tab = 1
    cmdImprimir.Enabled = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
    sstComanda.TabEnabled(0) = True
    sstComanda.Tab = 0
           
    hfgProduto.Clear
    Call Limpa_Grid
    
    Call Objetos.Limpa_TXT(Me)
    
    strCampo = "ID,Produto,Descrição,Quant.,Preço,Total Item,NºComanda"
    strTamanho = "0,800,2580,900,1000,1200,0"
    Call Movimentacoes.Monta_HFlex_Grid(hfgProduto, strTamanho, strCampo, 7, "Otica", Me)
    
    Call Monta_DataCombo
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
        
    'Gravando Log
    log.Gravar_log "Otica", Me
            
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
    
    lblTotal.Caption = "0,00"
    lblCupom.Caption = Empty
    Label10.Visible = False
    txtQuantidade.Text = "0,00"
    txtPreco.Text = "0,00"
    txtTotal_Item.Text = "0,00"
    
    txtNumero_Comanda.Enabled = True
    txtCodigo_Vendedor.Enabled = True
    dtcVendedor.Enabled = True
    cmdImprimir.Enabled = False
                     
    txtNumero_Comanda.SetFocus
                       
    booAlterar = False
           
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    strTamanho = "1300,1200,2200,1600,1600,1300"
    strNomes = "NºComanda,Vendedor,Nome,Data Lançamento,Hora Abertura,Nº Cupom"
    
    On Error GoTo Erro
   
    Movimentacoes.Monta_HFlex_Grid hfgComanda, strTamanho, strNomes, 6, "Otica", Me
            
    Call Monta_DataCombo
    Call Monta_Combo
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Reposicao")
    Resume Next
End Function

'Private Sub txtCodigo_Produto_Change()
    'dtcProduto.BoundText = txtCodigo_Produto.Text
'End Sub

Private Sub txtCodigo_Produto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Produto_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_Produto_LostFocus()
        
    If txtCodigo_Produto.Text <> Empty Then
       Dim strDigito_Peso_Variavel As String
       Dim strDigito_Produto_Digitado As String
       Dim strCodigo_Produto_Etiqueta As String
       Dim strID_Produto As String
       Dim strPreco_Peso_Parametro As String
       Dim strTabela_Vigente As String
       Dim rstBusca_Preco As New ADODB.Recordset
       Dim rstBusca_Paramentros As New ADODB.Recordset
       Dim rstBusca_Produto As New ADODB.Recordset
       Dim strPreco_Tabela As String
       Dim strTotal As String
       Dim strPreco_Peso As String
       Dim strDecimal As String
       Dim strQuantidade As String
           
       strSql = Empty
       strSql = "SELECT * FROM TBParametros_ecf WHERE FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ""
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Paramentros, "Otica", Me)
        
       strDigito_Peso_Variavel = rstBusca_Paramentros.Fields("DFCodigo_inicial_peso_variavel_TBParametros_ecf")
       If rstBusca_Paramentros.Fields("DFPreco_peso_balanca_TBParametros_ecf") = False Then
          strPreco_Peso_Parametro = 0
       Else
          strPreco_Peso_Parametro = 1
       End If
       Set rstBusca_Paramentros = Nothing
        
       strSql = Empty
       strSql = "SELECT * FROM TBParametros_venda WHERE IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ""
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Paramentros, "Otica", Me)
       
       strTabela_Vigente = rstBusca_Paramentros.Fields("DFNumero_tabela_vigente_TBParametros_venda")
       Set rstBusca_Paramentros = Nothing
           
       If Len(txtCodigo_Produto.Text) > 6 Then
          strDigito_Produto_Digitado = Left(txtCodigo_Produto.Text, 1)
          If strDigito_Peso_Variavel = strDigito_Produto_Digitado Then
             strCodigo_Produto_Etiqueta = Mid(txtCodigo_Produto.Text, 2, 4)
             strPreco_Peso = Mid(txtCodigo_Produto.Text, 6, 7)
             strSql = Empty
             strSql = "SELECT * FROM TBProduto WHERE IXCodigo_TBproduto = " & strCodigo_Produto_Etiqueta & " "
             Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Produto, "Otica", Me)
             If rstBusca_Produto.RecordCount = 0 Then
                MsgBox "Produto não Cadastrado, Verifique.", vbCritical, "Only Tech"
                txtCodigo_Produto.SetFocus
                Exit Sub
             End If
             txtDescricao_Produto.Text = rstBusca_Produto.Fields("DFDescricao_resumida_TBProduto")
             strID_Produto = Funcoes_Gerais.Localiza_ID("PKId_TBProduto", "IXCodigo_TBProduto", strCodigo_Produto_Etiqueta, "TBProduto", "Otica", Me, "BDRetaguarda")
             strSql = Empty
             strSql = "SELECT TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco " & _
                      "FROM TBItens_tabela_preco " & _
                      "WHERE FKCodigo_TBTabela_preco = " & strTabela_Vigente & " AND " & _
                      "FKId_TBProduto = " & strID_Produto & ""
             Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Preco, "Otica", Me)
             If rstBusca_Preco.RecordCount = 0 Then
                MsgBox "Produto não cadastrado na tabela de preço vigente.Verifique.", vbCritical, "Only Tech"
                txtCodigo_Produto.SetFocus
                Exit Sub
             End If
             strPreco_Tabela = Format(rstBusca_Preco.Fields("DFPreco_varejo_TBItens_tabela_preco"), "#,###0.00")
             Set rstBusca_Preco = Nothing
             If strPreco_Peso_Parametro = 0 Then
                strPreco_Peso = Mid(txtCodigo_Produto.Text, 6, 5)
                strDecimal = Mid(txtCodigo_Produto.Text, 11, 2)
                strPreco_Peso = strPreco_Peso & "," & strDecimal
                strPreco_Peso = Format(strPreco_Peso, "#,###0.00")
                strQuantidade = CDbl(strPreco_Peso) / CDbl(strPreco_Tabela)
                strQuantidade = Format(strQuantidade, "#,###0.000")
                txtQuantidade.Text = strQuantidade
                txtPreco.Text = strPreco_Tabela
                strTotal = CDbl(strPreco_Tabela) * CDbl(strQuantidade)
                strTotal = Format(strTotal, "#,###0.00")
                txtTotal_Item.Text = strTotal
             Else
                strPreco_Peso = Format(strPreco_Peso, "#,###0.000")
                strTotal = strPreco_Peso * strPreco_Tabela
                txtQuantidade.Text = strPreco_Peso
                txtPreco.Text = Format(strPreco_Tabela, "#,###0.00")
                txtTotal_Item.Text = Format(strTotal, "#,###0.00")
             End If
             txtCodigo_Produto.Text = rstBusca_Produto.Fields("IXCodigo_TBProduto")
             Set rstBusca_Produto = Nothing
          Else
             strID_Produto = Funcoes_Gerais.Localiza_ID("FKId_TBProduto", "IXCodigo_TBCodigo_barras", txtCodigo_Produto.Text, "TBCodigo_barras", "Otica", Me, "BDRetaguarda")
             strSql = Empty
             strSql = "SELECT * FROM TBProduto WHERE PKId_TBproduto = " & strID_Produto & " "
             Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Produto, "Otica", Me)
             If rstBusca_Produto.RecordCount = 0 Then
                MsgBox "Produto não Cadastrado, Verifique.", vbCritical, "Only Tech"
                txtCodigo_Produto.SetFocus
                Exit Sub
             End If
             txtCodigo_Produto.Text = rstBusca_Produto.Fields("IXCodigo_TBProduto")
             Set rstBusca_Produto = Nothing
             strSql = Empty
             strSql = "SELECT TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco, " & _
                      "TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                      "FROM TBItens_tabela_preco " & _
                      "INNER JOIN TBProduto ON TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                      "WHERE FKCodigo_TBTabela_preco = " & strTabela_Vigente & " AND " & _
                      "FKId_TBProduto = " & strID_Produto & ""
             Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Preco, "Otica", Me)
             If rstBusca_Preco.RecordCount = 0 Then
                MsgBox "Produto não cadastrado na tabela de preço vigente.Verifique.", vbCritical, "Only Tech"
                txtCodigo_Produto.SetFocus
                Exit Sub
             End If
             strPreco_Tabela = Format(rstBusca_Preco.Fields("DFPreco_varejo_TBItens_tabela_preco"), "#,###0.00")
             txtQuantidade.Text = 1
             txtPreco.Text = strPreco_Tabela
             txtDescricao_Produto.Text = rstBusca_Preco.Fields("DFDescricao_resumida_TBProduto")
             'imgProduto.Picture = LoadPicture(rstBusca_Preco.Fields("DFPath_imagem_TBProduto"))
             strTotal = CDbl(txtQuantidade.Text) * CDbl(txtPreco.Text)
             strTotal = Format(strTotal, "#,###0.00")
             txtPreco.Text = strTotal
             Set rstBusca_Preco = Nothing
          End If
       Else
          strSql = Empty
          strSql = "SELECT * FROM TBProduto WHERE IXCodigo_TBProduto = " & txtCodigo_Produto.Text & " "
          Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Produto, "Otica", Me)
          If rstBusca_Produto.RecordCount = 0 Then
             MsgBox "Produto não Cadastrado, Verifique.", vbCritical, "Only Tech"
             txtCodigo_Produto.SetFocus
             Exit Sub
          End If
          Set rstBusca_Produto = Nothing
          strID_Produto = Funcoes_Gerais.Localiza_ID("PKId_TBProduto", "IXCodigo_TBProduto", txtCodigo_Produto.Text, "TBProduto", "Otica", Me, "BDRetaguarda")
          strSql = Empty
          strSql = "SELECT * FROM TBProduto WHERE PKId_TBproduto = " & strID_Produto & " "
          Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Produto, "Otica", Me)
          If rstBusca_Produto.RecordCount = 0 Then
             MsgBox "Produto não Cadastrado, Verifique.", vbCritical, "Only Tech"
             txtCodigo_Produto.SetFocus
             Exit Sub
          End If
          Set rstBusca_Produto = Nothing
          strSql = Empty
          strSql = "SELECT TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco, " & _
                   "TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                   "FROM TBItens_tabela_preco " & _
                   "INNER JOIN TBProduto ON TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                   "WHERE FKCodigo_TBTabela_preco = " & strTabela_Vigente & " AND " & _
                   "FKId_TBProduto = " & strID_Produto & ""
          Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Preco, "Otica", Me)
          If rstBusca_Preco.RecordCount = 0 Then
             MsgBox "Produto não cadastrado na tabela de preço vigente.Verifique.", vbCritical, "Only Tech"
             txtCodigo_Produto.SetFocus
             Exit Sub
          End If
          strPreco_Tabela = Format(rstBusca_Preco.Fields("DFPreco_varejo_TBItens_tabela_preco"), "#,###0.00")
          txtPreco.Text = strPreco_Tabela
          txtQuantidade.Text = 1
          txtDescricao_Produto.Text = rstBusca_Preco.Fields("DFDescricao_resumida_TBProduto")
          'imgProduto.Picture = LoadPicture(rstBusca_Preco.Fields("DFPath_imagem_TBProduto"))
          strTotal = CDbl(txtQuantidade.Text) * CDbl(txtPreco.Text)
          strTotal = Format(strTotal, "#,###0.00")
          txtPreco.Text = strTotal
          Set rstBusca_Preco = Nothing
       End If
       txtQuantidade.SetFocus
    End If
End Sub

Private Sub txtCodigo_Vendedor_Change()
    dtcVendedor.BoundText = txtCodigo_Vendedor.Text
End Sub

Private Sub txtCodigo_Vendedor_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Vendedor_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Function Consulta()
    Dim strDataInicial As String
    Dim strDataFinal  As String
                       
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
     
    strDataInicial = Format(dtpIni.Value, "YYYYMMDD")
    strDataFinal = Format(dtpFin.Value, "YYYYMMDD")
          
    strSql = "SELECT TBComanda.PKCodigo_TBComanda," & _
             "TBVendedor.IXCodigo_TBVendedor," & _
             "TBVendedor.DFNome_TBVendedor," & _
             "TBComanda.DFData_lancamento_TBComanda," & _
             "TBComanda.DFHora_abertura_TBComanda," & _
             "TBComanda.DFNumero_cupom_TBComanda " & _
             "FROM TBComanda " & _
             "INNER JOIN TBVendedor ON TBComanda.FKId_TBVendedor = TBVendedor.PKId_TBVendedor"
                      
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Número da Comanda" Then
           strSql = strSql & " WHERE convert(nvarchar,PKCodigo_TBComanda) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Código do Vendedor" Then
           strSql = strSql & " WHERE convert(nvarchar,IXCodigo_TBVendedor) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Nome do Vendedor" Then
           strSql = strSql & " WHERE convert(nvarchar,DFNome_TBVendedor) = '%" & txtConsulta.Text & "%'"
       ElseIf cbbCampos.Text = "Data Lançamento" Then
           strSql = strSql & " WHERE TBComanda.DFData_lancamento_TBComanda >= '" & strDataInicial & "' AND" & _
                             " TBComanda.DFData_lancamento_TBComanda <= '" & strDataFinal & "'"
       ElseIf cbbCampos.Text = "Hora Atendimento" Then
           strSql = strSql & " WHERE TBComanda.DFData_lancamento_TBComanda >= '" & strDataInicial & "' AND" & _
                             " TBComanda.DFData_lancamento_TBComanda <= '" & strDataFinal & "'"
       Else
           strSql = strSql & " WHERE convert(nvarchar,DFNumero_cupom_TBComanda) = '" & txtConsulta.Text & "'"
       End If
    End If
    
    frmAguarde.Show
    DoEvents
    
    strTamanho = "1300,1200,2200,1600,1600,1300"
    strNomes = "NºComanda,Vendedor,Nome,Data Lançamento,Hora Abertura,Nº Cupom"
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgComanda, strTamanho, strNomes, "BDRetaguarda", "Otica", Me, "S"
                       
    Unload frmAguarde
        
    hfgComanda.Refresh
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Número da Comanda")
    cbbCampos.AddItem ("Código do Vendedor")
    cbbCampos.AddItem ("Nome do Vendedor")
    cbbCampos.AddItem ("Data Lançamento")
    cbbCampos.AddItem ("Hora Atendimento")
    cbbCampos.AddItem ("Número Cupom")
End Function

Private Function Colori_Grid()
    hfgProduto.Col = 0
    hfgProduto.Row = I
    hfgProduto.ColWidth(0) = 400
    hfgProduto.Font.Name = "Tahoma"
    hfgProduto.FixedAlignment(0) = 2
    hfgProduto.CellFontSize = 7
    hfgProduto.CellBackColor = &H80FFFF
    hfgProduto.CellFontBold = False
    hfgProduto.Text = I
End Function

Private Sub txtNumero_Comanda_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_Comanda_LostFocus()
    If booAlterar = False Then
       Call Movimentacoes.Verifica_Numero("PKCodigo_TBComanda", "TBComanda", txtNumero_Comanda, "Otica", Me)
    End If
End Sub

Private Sub txtNumero_Pessoas_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_Pessoas_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_LostFocus()
    txtPreco.Text = Format(txtPreco.Text, "#,###0.00")
End Sub

Private Sub txtQuantidade_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtQuantidade_LostFocus()
    txtQuantidade.Text = Format(txtQuantidade.Text, "#,###0.00")
    If txtQuantidade.Text <> Empty Then
       txtTotal_Item.Text = CDbl(txtQuantidade.Text) * CDbl(txtPreco.Text)
       txtTotal_Item.Text = Format(txtTotal_Item.Text, "#,###0.00")
    End If
End Sub

Private Function Monta_DataCombo()
    'strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto FROM TBProduto"
    'Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT TBVendedor.IXCodigo_TBVendedor,TBVendedor.DFNome_TBVendedor FROM TBVendedor"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBVendedor", "DFNome_TBVendedor", dtcVendedor, strSql, "BDRetaguarda", "Otica", Me
End Function

Private Function Limpa_Grid()
    If hfgProduto.Rows >= 2 Then
       I = hfgProduto.Rows - 1
       Do While I <= hfgProduto.Rows - 1
          hfgProduto.Row = I
          hfgProduto.Col = 1
          If hfgProduto.Row > 1 Then
             hfgProduto.RemoveItem (hfgProduto.Row)
          Else
             hfgProduto.Clear
             strCampo = "ID,Produto,Descrição,Quant.,Preço,Total Item,NºComanda"
             strTamanho = "0,800,2580,900,1000,1200,0"
             Call Movimentacoes.Monta_HFlex_Grid(hfgProduto, strTamanho, strCampo, 7, "Otica", Me)
             Exit Do
          End If
          I = I - 1
       Loop
    End If
End Function
