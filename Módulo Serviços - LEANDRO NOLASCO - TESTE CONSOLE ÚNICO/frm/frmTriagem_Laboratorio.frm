VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTriagem_Laboratorio 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Triagem Laboratório"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "frmTriagem_Laboratorio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7875
   Begin TabDlg.SSTab sstTriagem_Laboratorio 
      Height          =   4545
      Left            =   0
      TabIndex        =   13
      Top             =   330
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   8017
      _Version        =   393216
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
      TabPicture(0)   =   "frmTriagem_Laboratorio.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblStatus"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "shpIntegrado"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtcFabricante"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtcInsumo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtcCliente"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCodigo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtLote"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtFabricante"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtInsumo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCodigo_Cliente"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtID_Portal"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "&Resultado"
      TabPicture(1)   =   "frmTriagem_Laboratorio.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtObservacao"
      Tab(1).Control(1)=   "dtpData_Resultado"
      Tab(1).Control(2)=   "cbbConforme"
      Tab(1).Control(3)=   "Label15"
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(5)=   "Label4"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&Listagem"
      TabPicture(2)   =   "frmTriagem_Laboratorio.frx":17BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdConsulta"
      Tab(2).Control(1)=   "cmdRefresh"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtConsulta"
      Tab(2).Control(3)=   "hfgTriagem_Laboratorio"
      Tab(2).Control(4)=   "cbbCampos"
      Tab(2).Control(5)=   "dtpConsulta_Data_Fim"
      Tab(2).Control(6)=   "dtpConsulta_Data_Inicio"
      Tab(2).Control(7)=   "lblAte"
      Tab(2).Control(8)=   "Label6"
      Tab(2).ControlCount=   9
      Begin VB.TextBox txtID_Portal 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Código Lote"
         Top             =   780
         Width           =   1005
      End
      Begin VB.Frame Frame2 
         Caption         =   "Competência"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   4620
         TabIndex        =   37
         Top             =   3180
         Width           =   3135
         Begin AutoCompletar.CbCompleta cbbCompetencia_Mes 
            Height          =   360
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Competência (Mês)"
            Top             =   570
            Width           =   1515
            _ExtentX        =   2672
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
         Begin MSComCtl2.DTPicker dtpCompetencia_Ano 
            Height          =   360
            Left            =   1680
            TabIndex        =   12
            ToolTipText     =   "Competência (Ano)"
            Top             =   570
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   20185091
            CurrentDate     =   38797
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Ano"
            Height          =   240
            Left            =   1680
            TabIndex        =   39
            Top             =   330
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mês"
            Height          =   240
            Left            =   120
            TabIndex        =   38
            Top             =   330
            Width           =   345
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   120
         TabIndex        =   32
         Top             =   3180
         Width           =   4485
         Begin MSComCtl2.DTPicker dtpData_Lancamento 
            Height          =   375
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Data Lançamento"
            Top             =   570
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   38800
         End
         Begin MSComCtl2.DTPicker dtpData_Fabricacao 
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            ToolTipText     =   "Data Fabricação"
            Top             =   570
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   38800
         End
         Begin MSComCtl2.DTPicker dtpData_Validade 
            Height          =   375
            Left            =   3000
            TabIndex        =   10
            ToolTipText     =   "Data Validade"
            Top             =   570
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   38800
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Validade"
            Height          =   240
            Left            =   3000
            TabIndex        =   35
            Top             =   330
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fabricação"
            Height          =   240
            Left            =   1560
            TabIndex        =   34
            Top             =   330
            Width           =   930
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Lançamento"
            Height          =   240
            Left            =   120
            TabIndex        =   33
            Top             =   330
            Width           =   1035
         End
      End
      Begin VB.TextBox txtCodigo_Cliente 
         Height          =   360
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Código Cliente"
         Top             =   1440
         Width           =   1250
      End
      Begin VB.TextBox txtInsumo 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Código Insumo"
         Top             =   2100
         Width           =   1250
      End
      Begin VB.TextBox txtFabricante 
         Height          =   360
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Código Fabricante"
         Top             =   2760
         Width           =   1250
      End
      Begin VB.TextBox txtLote 
         Height          =   375
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Código Lote"
         Top             =   780
         Width           =   1845
      End
      Begin VB.TextBox txtObservacao 
         Height          =   2745
         Left            =   -74880
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         ToolTipText     =   "Descrição Resultado"
         Top             =   1440
         Width           =   7605
      End
      Begin MSComCtl2.DTPicker dtpData_Resultado 
         Height          =   375
         Left            =   -74880
         TabIndex        =   24
         ToolTipText     =   "Data Resultado"
         Top             =   780
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   38800
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Código Triagem"
         Top             =   780
         Width           =   1250
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
         Left            =   -67950
         Picture         =   "frmTriagem_Laboratorio.frx":17D6
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   -67560
         Picture         =   "frmTriagem_Laboratorio.frx":34D0
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   405
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72990
         TabIndex        =   18
         Top             =   780
         Width           =   4965
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgTriagem_Laboratorio 
         Height          =   3165
         Left            =   -74880
         TabIndex        =   15
         Top             =   1230
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5583
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
         TabIndex        =   16
         Top             =   780
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
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   360
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Descrição Cliente"
         Top             =   1440
         Width           =   6315
         _ExtentX        =   11139
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
      Begin MSDataListLib.DataCombo dtcInsumo 
         Height          =   360
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Descrição Insumo"
         Top             =   2100
         Width           =   6315
         _ExtentX        =   11139
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
      Begin MSDataListLib.DataCombo dtcFabricante 
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         ToolTipText     =   "Descrição Fabricante"
         Top             =   2760
         Width           =   6315
         _ExtentX        =   11139
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
      Begin AutoCompletar.CbCompleta cbbConforme 
         Height          =   360
         Left            =   -73320
         TabIndex        =   40
         ToolTipText     =   "Status: Conforme / Não Conforme"
         Top             =   780
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSComCtl2.DTPicker dtpConsulta_Data_Fim 
         Height          =   360
         Left            =   -69540
         TabIndex        =   20
         Top             =   780
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20185091
         CurrentDate     =   37923
      End
      Begin MSComCtl2.DTPicker dtpConsulta_Data_Inicio 
         Height          =   360
         Left            =   -72990
         TabIndex        =   19
         Top             =   780
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20185091
         CurrentDate     =   37923
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "ID Portal"
         Height          =   240
         Left            =   3360
         TabIndex        =   44
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lblAte 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "até"
         Height          =   240
         Left            =   -70680
         TabIndex        =   42
         Top             =   930
         Width           =   270
      End
      Begin VB.Shape shpIntegrado 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   165
         Left            =   7560
         Shape           =   3  'Circle
         Top             =   4290
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         Height          =   240
         Left            =   -73320
         TabIndex        =   41
         Top             =   540
         Width           =   840
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Left            =   4440
         TabIndex        =   36
         Top             =   780
         Width           =   3345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Insumo"
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   1860
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fabricante"
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   240
         Left            =   1440
         TabIndex        =   28
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observação"
         Height          =   240
         Left            =   -74880
         TabIndex        =   27
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Resultado"
         Height          =   240
         Left            =   -74880
         TabIndex        =   26
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   540
         Width           =   585
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
         TabIndex        =   17
         Top             =   540
         Width           =   435
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
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
      Left            =   8040
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
            Picture         =   "frmTriagem_Laboratorio.frx":4512
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriagem_Laboratorio.frx":482C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriagem_Laboratorio.frx":4B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriagem_Laboratorio.frx":4EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriagem_Laboratorio.frx":527A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriagem_Laboratorio.frx":5594
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriagem_Laboratorio.frx":58AE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTriagem_Laboratorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro Triagem Laboratório Integrado ao Portal               '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rodrigo Santos                                                 '
' Data de Criação........: 24/03/2006                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strCombo As String
Dim strConsulta As String
Dim strNomes As String
Dim strTamanho As String
Dim strCampo_consulta As String
Dim intContador As Integer
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
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
Dim intRetorno As Integer
Dim rstTriagem As New ADODB.Recordset

Function Imprimir()
    On Error GoTo Erro
    'Tratamento de Erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique.", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Triagem_Laboratorio.Show
    
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
       dtpConsulta_Data_Inicio.Visible = False
       dtpConsulta_Data_Fim.Visible = False
       lblAte.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Data Lançamento" Or cbbCampos.Text = "Data Fabricação" Or cbbCampos.Text = "Data Validade" Or cbbCampos.Text = "Data Resultado" Then
       txtConsulta.Visible = False
       dtpConsulta_Data_Inicio.Visible = True
       dtpConsulta_Data_Fim.Visible = True
       dtpConsulta_Data_Fim.Value = Date
       dtpConsulta_Data_Inicio.Value = Date - 15
       lblAte.Visible = True
    Else
       txtConsulta.Visible = True
       txtConsulta.SetFocus
       dtpConsulta_Data_Inicio.Visible = False
       dtpConsulta_Data_Fim.Visible = False
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

Private Sub dtpConsulta_Data_Fim_LostFocus()
    If dtpConsulta_Data_Fim.Value < dtpConsulta_Data_Inicio.Value Then
       MsgBox "Data Final menor que Data Início. Verifique!", vbInformation, "OnlyTech"
       dtpConsulta_Data_Inicio.SetFocus
    End If
End Sub

Private Sub dtpConsulta_Data_Inicio_LostFocus()
    If dtpConsulta_Data_Fim.Value < dtpConsulta_Data_Inicio.Value Then
       MsgBox "Data Final menor que Data Início. Verifique!", vbInformation, "OnlyTech"
       dtpConsulta_Data_Inicio.SetFocus
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
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0
       SendKeys "{TAB}"
    End If

End Sub
  
Private Sub Form_Load()
    On Error GoTo Erro
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Tipo Marcha"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstTriagem_Laboratorio, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
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
       
       If Funcoes_Gerais.Verifica_registros_nao_integrados("TBTriagem", "FKCodigo_TBInsumo,FKCodigo_TBFabricante,FKId_TBCliente,DFData_lancamento_TBTriagem,DFData_fabricacao_TBTriagem,DFLote_TBTriagem,DFMes_ano_competencia_TBTriagem,DFAno_competencia_TBTriagem,DFData_validade_TBTriagem,PKId_TBTriagem", "Otica", "DFIntegrado_TBTriagem", rstReg_nao_integrados, Me) = True Then
          If rstReg_nao_integrados.RecordCount > 0 Then
             strMensagem = "Deseja atualizar as informações para o portal? Existem " & rstReg_nao_integrados.RecordCount & " registro(s) desatualizados."
             intRetorno = MsgBox(strMensagem, vbYesNo, "Only Tech")
             If intRetorno = 6 Then
                frmAguarde.Show
                Funcoes_Gerais.Atualiza_registros_nao_integrados rstReg_nao_integrados, "TBTriagem_portal", "PKId_TBTriagem_portal ", "PKId_TBTriagem", "Portal", "FKCodigo_TBInsumo_portal,FKCodigo_TBFabricante_portal,FKId_TBCliente_portal,DFData_lancamento_TBTriagem_portal,DFData_fabricacao_TBTriagem_portal,DFLote_TBTriagem_portal,DFMes_ano_competencia_TBTriagem_portal,DFAno_competencia_TBTriagem_portal,DFData_validade_TBTriagem_portal,DFID_Int_Retaguarda_TBTriagem_portal", Me, "ortofarma1", "Otica", "BDRetaguarda", "TBTriagem", "DFIntegrado_TBTriagem", "ortofarma1", "ortofarma7410", 10
                MsgBox "Dados no portal atualizados com sucesso!", vbInformation, "Only Tech"
                Unload frmAguarde
             End If
          End If
       End If
       
    End If
    
Fim_atu_portal:

    On Error GoTo Erro
    
    log.Descricao = "Inicializando o cadastro de Triagem Laboratório"
    
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    sstTriagem_Laboratorio.TabEnabled(0) = False
    sstTriagem_Laboratorio.TabEnabled(1) = False
    sstTriagem_Laboratorio.Tab = 2
    
    dtpData_Resultado.Value = Date
    dtpData_Fabricacao.Value = Date
    dtpData_Validade.Value = Date
    dtpData_Lancamento.Value = Date
    dtpData_Lancamento.Enabled = False
    cbbConforme.Text = "Em Andamento"
    
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
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando o cadastro de Triagem Laboratório"
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

Private Sub hfgTriagem_Laboratorio_Click()
   Dim strIntegrado As String
   If hfgTriagem_Laboratorio.Col = 0 And hfgTriagem_Laboratorio.Text <> Empty Then
           
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
       
      txtCodigo.Text = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 1))
      txtFabricante.Text = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 2))
      txtInsumo.Text = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 4))
      txtCodigo_Cliente.Text = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 6))
      dtpData_Lancamento.Value = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 8))
      dtpData_Fabricacao.Value = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 9))
      dtpData_Validade.Value = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 10))
      txtLote.Text = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 11))
      cbbCompetencia_Mes.Text = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 12))
      dtpCompetencia_Ano.Value = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 13))
      dtpData_Resultado.Value = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 14))
      cbbConforme.Text = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 15))
      txtObservacao.Text = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 16))
      strIntegrado = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 17))
      txtID_Portal = hfgTriagem_Laboratorio.TextArray((hfgTriagem_Laboratorio.Row * hfgTriagem_Laboratorio.Cols + hfgTriagem_Laboratorio.Col + 18))
      
      'PREENCHENDO A LABEL
      lblStatus.Caption = UCase(cbbConforme.Text)
       
      If strIntegrado = "True" Then
         shpIntegrado.BackColor = &H8000&
      Else
         shpIntegrado.BackColor = vbRed
      End If
       
      booAlterar = True
      txtConsulta.Text = Empty
      sstTriagem_Laboratorio.TabEnabled(0) = True
      sstTriagem_Laboratorio.TabEnabled(1) = True
      sstTriagem_Laboratorio.Tab = 0
                              
   End If
   
   Unload frmAguarde
   
End Sub

Private Sub hfgTriagem_Laboratorio_DblClick()
    hfgTriagem_Laboratorio.Sort = 1
End Sub

Private Sub hfgTriagem_Laboratorio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgTriagem_Laboratorio_Click
    End If
End Sub

Private Sub lblStatus_Change()
    If lblStatus.Caption = "EM ANDAMENTO" Then
       lblStatus.ForeColor = &H80FFFF
    ElseIf lblStatus.Caption = "CONFORME" Then
       lblStatus.ForeColor = &HC00000
    ElseIf lblStatus.Caption = "NÃO CONFORME" Then
       lblStatus.ForeColor = &HFF&
    End If
End Sub

Private Sub sstTriagem_Laboratorio_Click(PreviousTab As Integer)
   If sstTriagem_Laboratorio.Tab = 0 Then
      txtLote.SetFocus
   ElseIf sstTriagem_Laboratorio.Tab = 1 Then
      dtpData_Resultado.SetFocus
   ElseIf sstTriagem_Laboratorio.Tab = 2 Then
      If frmIntegracao.Visible = True Then
         Unload frmIntegracao
      End If
      If strCombo <> Empty And strCombo <> "Todos" Then
         cbbCampos.Text = strCombo
         txtConsulta.SetFocus
      ElseIf strCombo = "Todos" Then
         hfgTriagem_Laboratorio.Row = 1
         hfgTriagem_Laboratorio.Col = 0
         hfgTriagem_Laboratorio.SetFocus
      End If
   End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstTriagem_Laboratorio.Tab <> 2: Call Gravar
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
    If txtCodigo_Cliente.Text = Empty Then
       MsgBox "O campo Código Cliente não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtCodigo_Cliente.SetFocus
       Exit Function
    ElseIf txtFabricante.Text = Empty Then
       MsgBox "O campo Código do Fabricante não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtFabricante.SetFocus
       Exit Function
    ElseIf txtInsumo.Text = Empty Then
       MsgBox "O campo Número Sequencial não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtInsumo.SetFocus
       Exit Function
    End If
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim intID_Cliente As Integer
    Dim rstVerifica_Resultado As New ADODB.Recordset
    Dim rstVerifica_Triplicidade As New ADODB.Recordset
    Dim rstReg_nao_integrados As New ADODB.Recordset
    Dim rstLocaliza_Id As New ADODB.Recordset
    Dim rstBusca_ID As New ADODB.Recordset
    Dim intConforme As Integer
    Dim strID_Portal As String
    Dim strID_Resultado_Portal As String
    Dim strProxima_ID As String
    
    If cbbConforme.Text = "Conforme" Then
       intConforme = 0
    ElseIf cbbConforme.Text = "Não Conforme" Then
       intConforme = 1
    Else
       intConforme = 2
    End If
    
    intID_Cliente = Funcoes_Gerais.Localiza_ID("PKId_TBCliente", "IXCodigo_TBCliente", dtcCliente.BoundText, "TBCliente", "Otica", Me, "BDRetaguarda")
    
    Call Objetos.Maiusculo_TXT(Me)
           
    strCampo = "FKCodigo_TBFabricante,FKCodigo_TBInsumo," & _
               "FKId_TBCliente,DFData_lancamento_TBTriagem," & _
               "DFData_fabricacao_TBTriagem,DFLote_TBTriagem," & _
               "DFMes_ano_competencia_TBTriagem,DFIntegrado_TBTriagem," & _
               "DFData_validade_TBTriagem,DFAno_competencia_TBTriagem," & _
               "DFData_alteracao_TBTriagem,DFIntegrado_filiais_TBTriagem "
                              
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBTriagem "
    End If
    
    strValores = "" & dtcFabricante.BoundText & "," & dtcInsumo.BoundText & "," & _
                 "'" & intID_Cliente & "','" & Format(dtpData_Lancamento.Value, "YYYYMMDD") & "'," & _
                 "'" & Format(dtpData_Fabricacao.Value, "YYYYMMDD") & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtLote.Text) & "','" & cbbCompetencia_Mes.Text & "',0," & _
                 "'" & Format(dtpData_Validade.Value, "YYYYMMDD") & "','" & Format(dtpCompetencia_Ano.Value, "YYYY") & "'," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0 "
    
    If booIntegra_Portal = True Then
        strValores = strValores & ",0 "
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                    VERIFICANDO EXISTÊNCIA DE TRÊS TRIAGENS PARA O MESMO INSUMO
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    strSql = "SELECT FKCodigo_TBInsumo,FKId_TBCliente,FKCodigo_TBFabricante,DFLote_TBTriagem," & _
             "FKCodigo_TBRamo_atividade,DFMes_ano_competencia_TBTriagem,DFAno_competencia_TBTriagem " & _
             "FROM TBTriagem " & _
             "INNER JOIN TBCliente ON TBTriagem.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
             "WHERE FKCodigo_TBInsumo = '" & dtcInsumo.BoundText & "' AND FKCodigo_TBRamo_atividade = " & _
             "(SELECT FKCodigo_TBRamo_atividade FROM TBCliente " & _
             "WHERE PKId_TBCliente = '" & intID_Cliente & "' ) AND " & _
             "FKCodigo_TBFabricante = '" & dtcFabricante.BoundText & "' AND DFLote_TBTriagem = '" & Funcoes_Gerais.Grava_String(txtLote.Text) & "' " & _
             "AND DFMes_ano_competencia_TBTriagem = '" & cbbCompetencia_Mes.Text & "' AND " & _
             "DFAno_competencia_TBTriagem = '" & Format(dtpCompetencia_Ano.Value, "YYYY") & "'"
    
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstVerifica_Triplicidade, "Otica", Me)
    
    If rstVerifica_Triplicidade.RecordCount <> 0 Then
       If rstVerifica_Triplicidade.RecordCount >= 3 Then
          MsgBox "Já existem três triagens lançadas para esse insumo. Verifique.", vbInformation, "OnlyTech"
          Call Cancelar
          Exit Function
       End If
    End If
    
    'INDICANDO O BANCO A CONECTAR-SE
    conexao.Initial_Catalog = "BDRetaguarda"

    'ESTABELECENDO CONEXÃO COM O BANCO
    conexao.Abrir_conexao ("Otica")

    'INDICA O INICIO DA TRANSAÇÃO JUNTO O BANCO
    conexao.CNConexao.BeginTrans
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       
       strSet = "UPDATE TBTriagem " & _
                "SET FKCodigo_TBFabricante = '" & dtcFabricante.BoundText & "', " & _
                "FKCodigo_TBInsumo = '" & dtcInsumo.BoundText & "', " & _
                "FKId_TBCliente = '" & intID_Cliente & "', " & _
                "DFData_lancamento_TBTriagem = '" & Format(dtpData_Lancamento.Value, "YYYYMMDD") & "', " & _
                "DFData_fabricacao_TBTriagem = '" & Format(dtpData_Fabricacao.Value, "YYYYMMDD") & "', " & _
                "DFLote_TBTriagem = '" & Funcoes_Gerais.Grava_String(txtLote.Text) & "', " & _
                "DFMes_ano_competencia_TBTriagem = '" & cbbCompetencia_Mes.Text & "', " & _
                "DFData_validade_TBTriagem = '" & Format(dtpData_Validade.Value, "YYYYMMDD") & "', " & _
                "DFAno_competencia_TBTriagem = '" & Format(dtpCompetencia_Ano.Value, "YYYY") & "'," & _
                "DFData_alteracao_TBTriagem = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBTriagem = 0 "
        
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBTriagem = 0 "
       End If
                
       strSet = strSet & "WHERE PKId_TBTriagem = '" & txtCodigo.Text & "'"
                     
       conexao.CNConexao.Execute strSet
       
       'VERIFICANDO A EXISTENCIA DE DADOS NA TBTRIAGEM_RESULTADO REFERENTE A TRIAGEM
       strSql = Empty
       strSql = "SELECT PKId_TBTriagem FROM TBTriagem_resultado WHERE PKId_TBTriagem = '" & txtCodigo.Text & "' "
       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstVerifica_Resultado, "Otica", Me)
       
       'MONTANDO A QUERY DA TABELA FILHA
       If rstVerifica_Resultado.EOF = True And rstVerifica_Resultado.BOF = True Then
          strSql = Empty
          strSql = "INSERT INTO TBTriagem_resultado (PKId_TBTriagem,DFData_resultado_TBTriagem_resultado," & _
                   "DFResultado_DFData_resultado_TBTriagem,DFObservacao_TBTriagem_resultado,DFIntegrado_TBTriagem_resultado," & _
                   "DFData_alteracao_TBTriagem_resultado,DFIntegrado_filiais_TBTriagem_resultado"
                   
          If booIntegra_Portal = True Then
             strSql = strSql & ",DFIntegrado_portal_TBTriagem_resultado) "
          Else
             strSql = strSql & ") "
          End If
                 


          strSql = strSql & "VALUES ('" & txtCodigo.Text & "','" & Format(dtpData_Resultado.Value, "YYYYMMDD") & "'," & _
                   "'" & intConforme & "','" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "','0'," & _
                   "'" & Format(Date, "YYYYMMDD") & "',0"
                   
          If booIntegra_Portal = True Then
             strSql = strSql & ",0) "
          Else
             strSql = strSql & ") "
          End If
       
       Else
          strSql = Empty
          strSql = "UPDATE TBTriagem_resultado " & _
                   "SET DFData_resultado_TBTriagem_resultado = '" & Format(dtpData_Resultado.Value, "YYYYMMDD") & "', " & _
                   "DFResultado_DFData_resultado_TBTriagem = '" & intConforme & "', " & _
                   "DFObservacao_TBTriagem_resultado = '" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "'," & _
                   "DFData_alteracao_TBTriagem_resultado = '" & Format(Date, "YYYYMMDD") & "'," & _
                   "DFIntegrado_filiais_TBTriagem_resultado = 0 "
                   
          strSql = strSql & ",DFIntegrado_portal_TBTriagem_resultado = 0 "
          
          strSql = strSql & "WHERE PKId_TBTriagem = '" & txtCodigo.Text & "'"
       End If
                   
       'GRAVANDO OS DADOS DA TABELA FILHA
       conexao.CNConexao.Execute strSql
       
    Else
       log.Evento = "Incluir Novo"

       strSql = "INSERT INTO TBTriagem (" & strCampo & ") " & _
                "VALUES (" & strValores & ")"
       
       conexao.CNConexao.Execute strSql
    End If
     
    'COMITANDO A TRANSACAO
    conexao.CNConexao.CommitTrans
    'FECHANDO A CONEXÃO
    conexao.CNConexao.Close
    
    'LOCALIZANDO ID DA TRIAGEM GRAVADA
    strSql = Empty
    strSql = "SELECT MAX(PKId_TBTriagem)AS IDTRIAGEM FROM TBTriagem "
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_ID, "Otica", Me)
        
    'INCREMENTANDO A ID
    If rstBusca_ID!IDTRIAGEM <> Empty Then
       strProxima_ID = rstBusca_ID!IDTRIAGEM
    Else
       strProxima_ID = 1
    End If
    
    'GRAVANDO LOG
    If booAlterar = True Then
       log.Descricao = "Alterando o registro: " + strProxima_ID
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
     Else
       log.Descricao = "Incluindo o registro: " + strProxima_ID
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
     End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Portal Tabela Triagem
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If booIntegra_Portal = True Then
        intRetorno = MsgBox("Deseja atualizar as informações para o portal?", vbYesNo, "Only Tech")
        If intRetorno = 6 Then
            On Error GoTo Erro_Portal
            
            strCampo = "FKCodigo_TBInsumo_portal,FKCodigo_TBFabricante_portal,FKId_TBCliente_portal,DFData_lancamento_TBTriagem_portal,DFData_fabricacao_TBTriagem_portal,DFLote_TBTriagem_portal,DFMes_ano_competencia_TBTriagem_portal,DFAno_competencia_TBTriagem_portal,DFData_validade_TBTriagem_portal,DFID_Int_Retaguarda_TBTriagem_portal"
            strValores = "" & dtcInsumo.BoundText & "," & dtcFabricante.BoundText & "," & intID_Cliente & ",'" & Format(dtpData_Lancamento.Value, "YYYYMMDD") & "','" & Format(dtpData_Fabricacao.Value, "YYYYMMDD") & "','" & Funcoes_Gerais.Grava_String(txtLote.Text) & "','" & cbbCompetencia_Mes.Text & "','" & Format(dtpCompetencia_Ano.Value, "YYYY") & "','" & Format(dtpData_Validade.Value, "YYYYMMDD") & "','" & strProxima_ID & "' "

            If booAlterar = True Then
               log.Evento = "Alterar"
               strSet = "SET FKCodigo_TBFabricante_portal = '" & dtcFabricante.BoundText & "', " & _
                        "FKCodigo_TBInsumo_portal = '" & dtcInsumo.BoundText & "', " & _
                        "FKId_TBCliente_portal = '" & intID_Cliente & "', " & _
                        "DFData_lancamento_TBTriagem_portal = '" & Format(dtpData_Lancamento.Value, "YYYYMMDD") & "', " & _
                        "DFData_fabricacao_TBTriagem_portal = '" & Format(dtpData_Fabricacao.Value, "YYYYMMDD") & "', " & _
                        "DFLote_TBTriagem_portal = '" & Funcoes_Gerais.Grava_String(txtLote.Text) & "', " & _
                        "DFMes_ano_competencia_TBTriagem_portal = '" & cbbCompetencia_Mes.Text & "', " & _
                        "DFAno_competencia_TBTriagem_portal = '" & dtpCompetencia_Ano.Value & "', " & _
                        "DFData_validade_TBTriagem_portal = '" & Format(dtpData_Validade.Value, "YYYYMMDD") & "', " & _
                        "DFID_Int_Retaguarda_TBTriagem_portal = '" & strProxima_ID & "' "
                        
               Call funcoes_banco.Alterar_Portal("ortofarma1", "TBTriagem_portal ", strSet, "PKId_TBTriagem_portal", Me.txtID_Portal.Text, "PKId_TBTriagem", strProxima_ID, "Otica", Me, "BDRetaguarda", "TBTriagem", "DFIntegrado_TBTriagem")
               log.Descricao = "Alterando o registro no Portal: " + strProxima_ID
               log.Tipo = 1
               log.Hora = Format(Now, "hh:mm:ss")
               'Gravando log
               log.Gravar_log "OTICA", Me
            Else
               log.Evento = "Incluir Novo"
               Call funcoes_banco.Gravar_Portal("ortofarma1", "TBTriagem_portal", strCampo, strValores, Me, "Otica", "BDRetaguarda", "TBTriagem", "DFIntegrado_TBTriagem", "PKId_TBTriagem", strProxima_ID)
               log.Descricao = "Gravando o registro no Portal: " + strProxima_ID
               log.Tipo = 1
               log.Hora = Format(Now, "hh:mm:ss")
               'Gravando log
               log.Gravar_log "OTICA", Me
            End If
                       
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'ATUALIZANDO TABELA FILHA TBTRIAGEM_RESULTADO
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If booAlterar = True Then
            
               strID_Portal = Funcoes_Gerais.Localiza_ID("PKId_TBTriagem_portal", "DFID_Int_Retaguarda_TBTriagem_portal", txtCodigo.Text, "TBTriagem_portal", "Portal", Me, "ortofarma1")
               strID_Resultado_Portal = Funcoes_Gerais.Localiza_ID("PKId_TBTriagem_resultado_portal", "PKId_TBTriagem_portal", strID_Portal, "TBTriagem_resultado_portal", "Portal", Me, "ortofarma1")
               
               If strID_Resultado_Portal <> 0 Then
                  
                  strSet = "SET PKId_TBTriagem_portal = '" & strID_Portal & "', " & _
                           "DFData_resultado_TBTriagem_resultado_portal = '" & Format(dtpData_Resultado.Value, "YYYYMMDD") & "', " & _
                           "DFResultado_DFData_resultado_TBTriagem_portal = '" & intConforme & "', " & _
                           "DFObservacao_TBTriagem_resultado_portal = '" & txtObservacao.Text & "' "
                        
                  Call funcoes_banco.Alterar_Portal("ortofarma1", "TBTriagem_resultado_portal ", strSet, "PKId_TBTriagem_resultado_portal", strID_Resultado_Portal, "PKId_TBTriagem_resultado", strProxima_ID, "Otica", Me, "BDRetaguarda", "TBTriagem_resultado", "DFIntegrado_TBTriagem_resultado")
               
               Else
                  
                  strCampo = "PKId_TBTriagem_portal,DFData_resultado_TBTriagem_resultado_portal,DFResultado_DFData_resultado_TBTriagem_portal,DFObservacao_TBTriagem_resultado_portal"
                  strValores = "" & Me.txtID_Portal.Text & ",'" & Format(dtpData_Resultado.Value, "YYYYMMDD") & "','" & intConforme & "','" & txtObservacao.Text & "'"
                  
                  Call funcoes_banco.Gravar_Portal("ortofarma1", "TBTriagem_resultado_portal", strCampo, strValores, Me, "Otica", "BDRetaguarda", "TBTriagem_resultado", "DFIntegrado_TBTriagem_resultado", "PKId_TBTriagem_resultado", strProxima_ID)
                  
               End If
            End If
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
       Me.hfgTriagem_Laboratorio.Visible = False
    End If
    
    sstTriagem_Laboratorio.TabEnabled(0) = False
    sstTriagem_Laboratorio.TabEnabled(1) = False
    sstTriagem_Laboratorio.Tab = 2
    
    Exit Function

Erro:
    If conexao.CNConexao.State <> adStateClosed Then
       conexao.CNConexao.RollbackTrans
       conexao.Fechar_conexao
    End If
    
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
       
Erro_Portal:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    MsgBox "Ocorreram erros na integração com o Portal!Contacte Only Tech.", vbCritical, "Only Tech"

    Exit Function
End Function

Private Function Excluir()
    Dim strID_Portal_Excluir As String
    On Error GoTo Erro
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'abrindo conexao
    conexao.Initial_Catalog = "BDRetaguarda"
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    'Excluindo Registro filho
    strSql = "DELETE FROM TBTriagem_resultado WHERE FKId_TBTriagem =  '" & txtCodigo.Text & "'"
    
    conexao.CNConexao.Execute strSql
    
    'Excluindo Registro Principal
    strSql = "DELETE FROM TBTriagem WHERE PKId_TBTriagem =  '" & txtCodigo.Text & "'"
    
    conexao.CNConexao.Execute strSql
    
    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'PORTAL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If booIntegra_Portal = True Then
        intRetorno = MsgBox("Deseja atualizar as informações para o portal?", vbYesNo, "Only Tech")
        If intRetorno = 6 Then
           
           strID_Portal_Excluir = Funcoes_Gerais.Localiza_ID("PKId_TBTriagem_portal", "DFID_Int_Retaguarda_TBTriagem_portal", txtCodigo.Text, "TBTriagem_portal", "Portal", Me, "ortofarma1")
           
           If strID_Portal_Excluir <> 0 Then
              Call funcoes_banco.Excluir("TBTriagem_resultado_portal", "PKId_TBTriagem_portal", strID_Portal_Excluir, "Portal", Me, "ortofarma1")
              Call funcoes_banco.Excluir("TBTriagem_portal", "PKId_TBTriagem_portal", strID_Portal_Excluir, "Portal", Me, "ortofarma1")
           Else
              MsgBox "Esse registro não está cadastrado no portal para ser excluido. Verifique.", vbInformation, "OnlyTech"
           End If
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
       hfgTriagem_Laboratorio.Visible = False
    End If
           
    sstTriagem_Laboratorio.TabEnabled(0) = False
    sstTriagem_Laboratorio.TabEnabled(1) = False
    sstTriagem_Laboratorio.Tab = 2
    
    Exit Function
Erro:
    If conexao.CNConexao.State <> adStateClosed Then
       conexao.CNConexao.RollbackTrans
       conexao.Fechar_conexao
    End If
    
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
       hfgTriagem_Laboratorio.Visible = False
    End If
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstTriagem_Laboratorio.TabEnabled(0) = False
    sstTriagem_Laboratorio.TabEnabled(1) = False
    sstTriagem_Laboratorio.Tab = 2
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
    Dim rstBusca_ID As New ADODB.Recordset
    Dim strCodigo_Tipo_Marcha As String

    Call Objetos.Limpa_TXT(Me)
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Set rstBusca_ID = Nothing
     
    sstTriagem_Laboratorio.TabEnabled(0) = True
    sstTriagem_Laboratorio.TabEnabled(1) = False
    sstTriagem_Laboratorio.Tab = 0
                    
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
        
    txtLote.SetFocus
    lblStatus.Caption = "EM ANDAMENTO"
    cbbCompetencia_Mes.Text = Empty
    dtpData_Fabricacao.Value = Date
    dtpData_Validade.Value = Date
    
    If booIntegra_Portal = True Then
       Me.shpIntegrado.BackColor = vbRed
    End If
       
    booAlterar = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    On Error GoTo Erro
    
    strNomes = "Cod. Triagem,Cod. Fabricante,Desc. Fabricante,Cod. Insumo, Desc. Insumo," & _
               "Cod. Cliente,Nome Cliente,Data Lançamento,Data Fabricação,Data Validade," & _
               "Lote,Competência Mês,Competência Ano,Data Resultado,Resultado," & _
               "Observação,Integrado,Id Portal"
               
    strTamanho = "1300,1450,4000,1300,4000," & _
                 "1300,4000,1600,1450,1450," & _
                 "1500,1600,1600,1450,1450," & _
                 "6500,0,1500"

    Movimentacoes.Monta_HFlex_Grid hfgTriagem_Laboratorio, strTamanho, strNomes, 18, "Otica", Me
        
    Call Monta_Combo
    Call Monta_DataCombo
            
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Reposicao")
    Resume Next
End Function

Private Function Consulta()
    Dim intPos As Integer
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text <> "Data Lançamento" And cbbCampos.Text <> "Data Fabricação" And cbbCampos.Text <> "Data Validade" And cbbCampos.Text <> "Data Resultado" Then
          If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
             MsgBox "Selecione um campo e digite os dados para consulta.", vbInformation, "Only Tech"
             cbbCampos.SetFocus
             Exit Function
           End If
       End If
    End If
    
    strSql = "SELECT TBTriagem.PKId_TBTriagem,FKCodigo_TBFabricante,DFNome_TBFabricante," & _
             "FKCodigo_TBInsumo,DFDescricao_TBInsumo,IXCodigo_TBCliente,DFNome_TBCliente," & _
             "DFData_lancamento_TBTriagem,DFData_fabricacao_TBTriagem,DFData_validade_TBTriagem," & _
             "DFLote_TBTriagem,DFMes_ano_competencia_TBTriagem,DFAno_competencia_TBTriagem," & _
             "TBTriagem_resultado.DFData_resultado_TBTriagem_resultado," & _
             "TBTriagem_resultado.DFResultado_DFData_resultado_TBTriagem," & _
             "TBTriagem_resultado.DFObservacao_TBTriagem_resultado,DFIntegrado_TBTriagem,DFCodigo_Identificador_TBTriagem " & _
             "FROM TBTriagem " & _
             "INNER JOIN TBFabricante ON TBTriagem.FKCodigo_TBFabricante = TBFabricante.PKCodigo_TBFabricante " & _
             "INNER JOIN TBInsumo ON TBTriagem.FKCodigo_TBInsumo = TBInsumo.PKCodigo_TBInsumo " & _
             "INNER JOIN TBCliente ON TBTriagem.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
             "LEFT JOIN TBTriagem_resultado ON TBTriagem.PKId_TBTriagem = TBTriagem_resultado.PKId_TBTriagem "

    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    Funcoes_Gerais.Grava_String (txtConsulta.Text)
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Cod. Triagem" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE TBTriagem.PKId_TBTriagem = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Cod. Fabricante" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE TBTriagem.FKCodigo_TBFabricante = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Desc. Fabricante" Then
          strSql = strSql & " WHERE TBFabricante.DFNome_TBFabricante LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Cod. Insumo" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE TBTriagem.FKCodigo_TBInsumo = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Desc. Insumo" Then
          strSql = strSql & " WHERE TBInsumo.DFDescricao_TBInsumo LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Cod. Cliente" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE TBCliente.IXCodigo_TBCliente = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Nome Cliente" Then
          strSql = strSql & " WHERE TBCliente.DFNome_TBCliente LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Data Lançamento" Then
          strSql = strSql & " WHERE TBTriagem.DFData_lancamento_TBTriagem BETWEEN '" & Format(dtpConsulta_Data_Inicio.Value, "YYYYMMDD") & "' AND '" & Format(dtpConsulta_Data_Fim.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Data Fabricação" Then
          strSql = strSql & " WHERE TBTriagem.DFData_fabricacao_TBTriagem BETWEEN '" & Format(dtpConsulta_Data_Inicio.Value, "YYYYMMDD") & "' AND '" & Format(dtpConsulta_Data_Fim.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Data Validade" Then
          strSql = strSql & " WHERE TBTriagem.DFData_validade_TBTriagem BETWEEN '" & Format(dtpConsulta_Data_Inicio.Value, "YYYYMMDD") & "' AND '" & Format(dtpConsulta_Data_Fim.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Data Resultado" Then
          strSql = strSql & " WHERE TBTriagem_resultado.DFData_resultado_TBTriagem_resultado BETWEEN '" & Format(dtpConsulta_Data_Inicio.Value, "YYYYMMDD") & "' AND '" & Format(dtpConsulta_Data_Fim.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Lote" Then
          strSql = strSql & " WHERE TBTriagem.DFLote_TBTriagem LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Competência Mês" Then
          strSql = strSql & " WHERE TBTriagem.DFMes_ano_competencia_TBTriagem LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Competência Ano" Then
          strSql = strSql & " WHERE TBTriagem.DFAno_competencia_TBTriagem LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Resultado" Then
          strSql = strSql & " WHERE TBTriagem_resultado.DFResultado_DFData_resultado_TBTriagem LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Observação" Then
          strSql = strSql & " WHERE TBTriagem_resultado.DFObservacao_TBTriagem_resultado LIKE '%" & txtConsulta.Text & "%' "
       End If
    End If

    frmAguarde.Show
    DoEvents
            
    strSql = strSql & " ORDER BY TBTriagem.PKId_TBTriagem"
        
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstTriagem, "Otica", Me
    
    If rstTriagem.EOF = True And rstTriagem.BOF = True Then
       GoTo FIM
    End If
    
    Dim contador_colunas As Long
    Dim Linhas As Long
    
    hfgTriagem_Laboratorio.Cols = 19
    hfgTriagem_Laboratorio.Rows = rstTriagem.RecordCount + 1
    
    contador_colunas = 2
    Linhas = 1
    
    rstTriagem.MoveFirst
    
    Do While Linhas <= rstTriagem.RecordCount
       DoEvents
       hfgTriagem_Laboratorio.Row = Linhas
       hfgTriagem_Laboratorio.Col = 0
       Me.hfgTriagem_Laboratorio.ColWidth(0) = 300
       hfgTriagem_Laboratorio.CellBackColor = &H80FFFF
       hfgTriagem_Laboratorio.CellFontBold = False
       hfgTriagem_Laboratorio.CellFontSize = 7
       hfgTriagem_Laboratorio.Text = Linhas
       contador_colunas = 1
       Do While contador_colunas <= rstTriagem.Fields.Count
          
          hfgTriagem_Laboratorio.Col = 0
          hfgTriagem_Laboratorio.Text = rstTriagem.AbsolutePosition
          
          hfgTriagem_Laboratorio.Col = 1
          hfgTriagem_Laboratorio.Text = rstTriagem!PKId_TBTriagem
          
          hfgTriagem_Laboratorio.Col = 2
          hfgTriagem_Laboratorio.Text = rstTriagem!FKCodigo_TBFabricante
          
          hfgTriagem_Laboratorio.Col = 3
          If IsNull(rstTriagem!DFNome_TBFabricante) = True Then
             hfgTriagem_Laboratorio.Text = ""
          Else
             hfgTriagem_Laboratorio.Text = rstTriagem!DFNome_TBFabricante
          End If
          hfgTriagem_Laboratorio.Col = 4
          If IsNull(rstTriagem!FKCodigo_TBInsumo) = True Then
             hfgTriagem_Laboratorio.Text = ""
          Else
             hfgTriagem_Laboratorio.Text = rstTriagem!FKCodigo_TBInsumo
          End If
          hfgTriagem_Laboratorio.Col = 5
          If IsNull(rstTriagem!DFDescricao_TBInsumo) = True Then
             hfgTriagem_Laboratorio.Text = ""
          Else
             hfgTriagem_Laboratorio.Text = rstTriagem!DFDescricao_TBInsumo
          End If
          hfgTriagem_Laboratorio.Col = 6
          If IsNull(rstTriagem!IXCodigo_TBCliente) = True Then
             hfgTriagem_Laboratorio.Text = ""
          Else
             hfgTriagem_Laboratorio.Text = rstTriagem!IXCodigo_TBCliente
          End If
          hfgTriagem_Laboratorio.Col = 7
          If IsNull(rstTriagem!DFNome_TBCliente) = True Then
             hfgTriagem_Laboratorio.Text = ""
          Else
             hfgTriagem_Laboratorio.Text = rstTriagem!DFNome_TBCliente
          End If
          hfgTriagem_Laboratorio.Col = 8
          hfgTriagem_Laboratorio.Text = rstTriagem!DFData_lancamento_TBTriagem
          hfgTriagem_Laboratorio.Col = 9
          hfgTriagem_Laboratorio.Text = rstTriagem!DFData_fabricacao_TBTriagem
          hfgTriagem_Laboratorio.Col = 10
          hfgTriagem_Laboratorio.Text = rstTriagem!DFData_validade_TBTriagem
          hfgTriagem_Laboratorio.Col = 11
          hfgTriagem_Laboratorio.Text = rstTriagem!DFLote_TBTriagem
          hfgTriagem_Laboratorio.Col = 12
          hfgTriagem_Laboratorio.Text = rstTriagem!DFMes_ano_competencia_TBTriagem
          hfgTriagem_Laboratorio.Col = 13
          hfgTriagem_Laboratorio.Text = rstTriagem!DFAno_competencia_TBTriagem
          
          hfgTriagem_Laboratorio.Col = 14
          If IsNull(rstTriagem!DFResultado_DFData_resultado_TBTriagem) = True Then
             hfgTriagem_Laboratorio.Text = ""
          Else
             hfgTriagem_Laboratorio.Text = rstTriagem!DFResultado_DFData_resultado_TBTriagem
          End If
          
          hfgTriagem_Laboratorio.Col = 15
          If IsNull(rstTriagem!DFResultado_DFData_resultado_TBTriagem) = True Then
             hfgTriagem_Laboratorio.Text = "Em Andamento"
          Else
             If rstTriagem!DFResultado_DFData_resultado_TBTriagem = 0 Then
                hfgTriagem_Laboratorio.Text = "Conforme"
             End If
             If rstTriagem!DFResultado_DFData_resultado_TBTriagem = 1 Then
                hfgTriagem_Laboratorio.Text = "Não Conforme"
             End If
             If rstTriagem!DFResultado_DFData_resultado_TBTriagem = 2 Then
                hfgTriagem_Laboratorio.Text = "Em Andamento"
             End If
          End If
          
          hfgTriagem_Laboratorio.Col = 16
          If IsNull(rstTriagem!DFObservacao_TBTriagem_resultado) = True Then
             hfgTriagem_Laboratorio.Text = ""
          Else
             hfgTriagem_Laboratorio.Text = rstTriagem!DFObservacao_TBTriagem_resultado
          End If
          
          hfgTriagem_Laboratorio.Col = 17
          hfgTriagem_Laboratorio.Text = rstTriagem!DFIntegrado_TBTriagem
          
          hfgTriagem_Laboratorio.Col = 18
          If IsNull(rstTriagem!DFCodigo_Identificador_TBTriagem) = True Then
             hfgTriagem_Laboratorio.Text = ""
          Else
             hfgTriagem_Laboratorio.Text = rstTriagem!DFCodigo_Identificador_TBTriagem
          End If
          
          contador_colunas = contador_colunas + 1
       Loop
       rstTriagem.MoveNext
       contador_colunas = 1
       Linhas = Linhas + 1
    Loop
FIM:

    Set rstTriagem = Nothing
    
    'Alinhando Fabricante a esquerda
    hfgTriagem_Laboratorio.ColAlignment(3) = 0
       
    Unload frmAguarde
    
End Function

Private Function Monta_Combo()

    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Cod. Triagem")
    cbbCampos.AddItem ("Cod. Fabricante")
    cbbCampos.AddItem ("Desc. Fabricante")
    cbbCampos.AddItem ("Cod. Insumo")
    cbbCampos.AddItem ("Desc. Insumo")
    cbbCampos.AddItem ("Cod. Cliente")
    cbbCampos.AddItem ("Nome Cliente")
    cbbCampos.AddItem ("Data Lançamento")
    cbbCampos.AddItem ("Data Fabricação")
    cbbCampos.AddItem ("Data Validade")
    cbbCampos.AddItem ("Lote")
    cbbCampos.AddItem ("Competência Mês")
    cbbCampos.AddItem ("Competência Ano")
    cbbCampos.AddItem ("Data Resultado")
    cbbCampos.AddItem ("Resultado")
    cbbCampos.AddItem ("Observação")
    
    cbbCompetencia_Mes.Clear
    cbbCompetencia_Mes.AddItem ("Janeiro")
    cbbCompetencia_Mes.AddItem ("Fevereiro")
    cbbCompetencia_Mes.AddItem ("Março")
    cbbCompetencia_Mes.AddItem ("Abril")
    cbbCompetencia_Mes.AddItem ("Maio")
    cbbCompetencia_Mes.AddItem ("Junho")
    cbbCompetencia_Mes.AddItem ("Julho")
    cbbCompetencia_Mes.AddItem ("Agosto")
    cbbCompetencia_Mes.AddItem ("Setembro")
    cbbCompetencia_Mes.AddItem ("Outubro")
    cbbCompetencia_Mes.AddItem ("Novembro")
    cbbCompetencia_Mes.AddItem ("Dezembro")
    
    cbbConforme.Clear
    cbbConforme.AddItem ("Conforme")
    cbbConforme.AddItem ("Não Conforme")
    cbbConforme.AddItem ("Em Andamento")
    
End Function

Private Function Monta_DataCombo()
           
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBInsumo,DFDescricao_TBInsumo FROM TBInsumo"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBInsumo", "DFDescricao_TBInsumo", dtcInsumo, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBFabricante,DFNome_TBFabricante FROM TBFabricante"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBFabricante", "DFNome_TBFabricante", dtcFabricante, strSql, "BDRetaguarda", "Otica", Me
    
End Function

Private Sub txtCodigo_Cliente_Change()
    dtcCliente.BoundText = txtCodigo_Cliente.Text
    If IsNumeric(txtCodigo_Cliente.Text) = False Then txtCodigo_Cliente.Text = Empty: Exit Sub
End Sub

Private Sub txtCodigo_Cliente_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_Cliente_LostFocus()
    If dtcCliente.Text = Empty Then txtCodigo_Cliente.Text = Empty
End Sub

Private Sub txtFabricante_Change()
    dtcFabricante.BoundText = txtFabricante.Text
    If IsNumeric(txtFabricante.Text) = False Then txtFabricante.Text = Empty: Exit Sub
End Sub

Private Sub txtFabricante_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFabricante_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtInsumo_Change()
    dtcInsumo.BoundText = txtInsumo.Text
    If IsNumeric(txtInsumo.Text) = False Then txtInsumo.Text = Empty: Exit Sub
End Sub

Private Sub txtInsumo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtInsumo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtLote_LostFocus()
    If txtLote.Text <> Empty Then txtLote.Text = UCase(txtLote.Text)
End Sub

Private Sub dtcFabricante_GotFocus()
    If txtFabricante.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFabricante)
    End If
End Sub

Private Sub dtcFabricante_LostFocus()
    txtFabricante.Text = dtcFabricante.BoundText
    If IsNumeric(txtFabricante.Text) = False Or dtcFabricante.Text = Empty Then txtFabricante.Text = Empty: Exit Sub
End Sub

Private Sub dtcInsumo_GotFocus()
    If txtInsumo.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcInsumo)
    End If
End Sub

Private Sub dtcInsumo_LostFocus()
    txtInsumo.Text = dtcInsumo.BoundText
    If IsNumeric(txtInsumo.Text) = False Or dtcInsumo.Text = Empty Then txtInsumo.Text = Empty: Exit Sub
End Sub

Private Sub dtcCliente_GotFocus()
    If txtCodigo_Cliente.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCliente)
    End If
End Sub

Private Sub dtcCliente_LostFocus()
    txtCodigo_Cliente.Text = dtcCliente.BoundText
    If IsNumeric(txtCodigo_Cliente.Text) = False Or dtcCliente.Text = Empty Then txtCodigo_Cliente.Text = Empty: Exit Sub
End Sub

Private Sub txtObservacao_LostFocus()
    If txtObservacao.Text <> Empty Then txtObservacao.Text = UCase(txtObservacao.Text)
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKId_TBTriagem", txtCodigo.Text, "DFIntegrado_filiais_TBTriagem", "TBTriagem", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBTriagem", Me.Top, Me.Left, Me.Width, Me.Height, "Triagem")
    
End Function
