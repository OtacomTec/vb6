VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmExplorer 
   Caption         =   "Project1"
   ClientHeight    =   6165
   ClientLeft      =   1980
   ClientTop       =   2355
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   12015
   Tag             =   "Explorer"
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2610
      TabIndex        =   11
      Top             =   1950
      Width           =   765
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   6
      Top             =   645
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2055
      TabIndex        =   4
      Top             =   975
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   8467
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item2"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item3"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item4"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   12015
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   12015
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " ListView:"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   3
         Tag             =   " ListView:"
         Top             =   12
         Width           =   3216
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   " TreeView:"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Tag             =   " TreeView:"
         Top             =   12
         Width           =   2016
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5895
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15558
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "06/09/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "09:13"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2685
      Top             =   3510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4800
      Left            =   0
      TabIndex        =   5
      Top             =   975
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   8467
      _Version        =   393217
      Indentation     =   441
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageListGeralGrande 
      Left            =   6360
      Top             =   1890
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   30
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":0000
            Key             =   "ico_Laboratório"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":0454
            Key             =   "ico_Usuário"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":08A8
            Key             =   "ico_Computador"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":305C
            Key             =   "ico_Software"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":3378
            Key             =   "ico_SistemaOperacional"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":5B2C
            Key             =   "ico_Programa"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":5E48
            Key             =   "ico_Departamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":6164
            Key             =   "ico_Drive3.5"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":8918
            Key             =   "ico_Drive5.25"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":B0CC
            Key             =   "ico_DriveCD"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":B520
            Key             =   "ico_DriveJazz250"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":C374
            Key             =   "ico_Empresa"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":C690
            Key             =   "ico_Hardware"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":C9AC
            Key             =   "ico_HD"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":F160
            Key             =   "ico_ÁreaDeTrabalho"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":F5B4
            Key             =   "ico_Memória"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":FA08
            Key             =   "ico_Modem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":FE5C
            Key             =   "ico_Mouse"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":10178
            Key             =   "ico_Placa"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":10494
            Key             =   "ico_PontoDeRede"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":107B0
            Key             =   "ico_Processador"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":10C04
            Key             =   "ico_SalaFechada"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":10F20
            Key             =   "ico_SalaAberta"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1123C
            Key             =   "ico_SCSI"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":11690
            Key             =   "ico_Servidor"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":11AE4
            Key             =   "ico_Teclado"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":143D8
            Key             =   "ico_USB"
            Object.Tag             =   "ico_Trânsito"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1482C
            Key             =   "ico_Monitor"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":15108
            Key             =   "ico_Sucata"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":15424
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListGeral 
      Left            =   5790
      Top             =   1890
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":15740
            Key             =   "ico_Laboratório"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":15B94
            Key             =   "ico_Computador"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":18348
            Key             =   "ico_Software"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":18664
            Key             =   "ico_SistemaOperacional"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1AE18
            Key             =   "ico_Programa"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1B134
            Key             =   "ico_Departamento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1B450
            Key             =   "ico_Drive3.5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":1DC04
            Key             =   "ico_Drive5.25"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":203B8
            Key             =   "ico_DriveCD"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2080C
            Key             =   "ico_DriveJazz250"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":21658
            Key             =   "ico_Empresa"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":21974
            Key             =   "ico_Hardware"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":21C90
            Key             =   "ico_HD"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":24444
            Key             =   "ico_ÁreaDeTrabalho"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":24898
            Key             =   "ico_Memória"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":24CEC
            Key             =   "ico_Modem"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":25140
            Key             =   "ico_Mouse"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2545C
            Key             =   "ico_Placa"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":25778
            Key             =   "ico_PontoDeRede"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":25A94
            Key             =   "ico_Processador"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":25EE8
            Key             =   "ico_SalaFechada"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":26204
            Key             =   "ico_SalaAberta"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":26520
            Key             =   "ico_SCSI"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":26974
            Key             =   "ico_Servidor"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":26DC8
            Key             =   "ico_Teclado"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":296BC
            Key             =   "ico_USB"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":29B10
            Key             =   "ico_Trânsito"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":29F64
            Key             =   "ico_Monitor"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2A840
            Key             =   "ico_Sucata"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2AB5C
            Key             =   "ico_Explorer"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2AE78
            Key             =   "ico_Usuário"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListToolBar 
      Left            =   6930
      Top             =   1890
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2B2CC
            Key             =   "tool_Ícones Grandes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2B3E0
            Key             =   "tool_Ícones Pequenos"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2B4F4
            Key             =   "tool_Ícones Lista"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2B608
            Key             =   "tool_Ícones Detalhes"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2B71C
            Key             =   "tool_Novo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2B830
            Key             =   "tool_Abrir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2B944
            Key             =   "tool_Salvar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2BA58
            Key             =   "tool_Imprimir"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2BB6C
            Key             =   "tool_Recortar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2BC80
            Key             =   "tool_Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2BD94
            Key             =   "tool_Colar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2BEA8
            Key             =   "tool_Explorer"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2C1C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2C2D6
            Key             =   "tool_Explorer2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2C3E8
            Key             =   "tool_SetaDireita"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorer.frx":2C73A
            Key             =   "tool_SetaEsquerda"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   741
      BandCount       =   4
      ImageList       =   "ImageListToolBar"
      _CBWidth        =   12015
      _CBHeight       =   420
      _Version        =   "6.7.8988"
      Child1          =   "ToolbarPadrão"
      MinHeight1      =   330
      Width1          =   3270
      NewRow1         =   0   'False
      Child2          =   "Toolbar"
      MinHeight2      =   330
      Width2          =   3360
      NewRow2         =   0   'False
      Child3          =   "ToolbarFerramentasGerais"
      MinHeight3      =   330
      Width3          =   3360
      NewRow3         =   0   'False
      MinHeight4      =   360
      NewRow4         =   0   'False
      Begin MSComctlLib.Toolbar ToolbarPadrão 
         Height          =   330
         Left            =   165
         TabIndex        =   10
         Top             =   45
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageListToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Novo"
               ImageKey        =   "tool_Novo"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Empresa"
                     Text            =   "Empresa"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Departamento"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Setor"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Abrir"
               ImageKey        =   "tool_Abrir"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Empresa"
                     Text            =   "Empresa"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Outros"
                     Text            =   "Outros"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Salvar"
               ImageKey        =   "tool_Salvar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageKey        =   "tool_Imprimir"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageKey        =   "tool_Recortar"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageKey        =   "tool_Copiar"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageKey        =   "tool_Colar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar 
         Height          =   330
         Left            =   3465
         TabIndex        =   9
         Top             =   45
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageListToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Explorer"
               ImageKey        =   "tool_Explorer2"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Adicionar Empresa"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarFerramentasGerais 
         Height          =   330
         Left            =   6855
         TabIndex        =   8
         Top             =   45
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageListToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Ícones Grandes"
               ImageKey        =   "tool_Ícones Grandes"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Ícones Pequenos"
               ImageKey        =   "tool_Ícones Pequenos"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Ícones Lista"
               ImageKey        =   "tool_Ícones Lista"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Ícones Detalhes"
               ImageKey        =   "tool_Ícones Detalhes"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image ImageDrag 
      Height          =   480
      Left            =   6210
      Picture         =   "frmExplorer.frx":2CA8C
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   255
      Width           =   150
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim mbMoving As Boolean
Const sglSplitLimit = 500

Const cns_TPNÓ_PARDRÃO = "pad"
Const cns_TPNÓ_EMPRESA = "emp"
Const cns_TPNÓ_DEPARTAMENTO = "dep"
Const cns_TPNÓ_SETOR = "set"
Const cns_TPNÓ_SALA = "sal"
Const cns_TPNÓ_EQUIPAMENTO = "equ"
Const cns_TPNÓ_HARDWARE = "hdw"
Const cns_TPNÓ_SOFTWARE = "sfw"
Const cns_TPNÓ_PEÇA = "peç"
Const cns_TPNÓ_PROGRAMA = "prg"


Private Enum tp_TipodeNó
    tp_TNÓ_Nenhum = 0
    tp_TNÓ_Empresa = 1
    tp_TNÓ_Departamento = 2
    tp_TNÓ_Setor = 3
    tp_TNÓ_Sala = 4
    tp_TNÓ_Equipamento = 5
    tp_TNÓ_Hardware = 6
    tp_TNÓ_Peça = 7
    tp_TNÓ_software = 12
    tp_TNÓ_Programa = 13
        
End Enum

Private Type NóSelecionado
    CódigoDoNó As String
    RótuloDoNó As String
    NaturezaDoNó As String
    Atual As Node
End Type

'Private SourceNode As Object
Private SourceType As tp_TipodeNó
Private TargetNode As Object
Private Nó As NóSelecionado

Private MedidaTopo_s As Single


Private Sub Command1_Click()
    '' Nó.Atual.
End Sub

Private Sub CoolBar1_Resize()
    If CoolBar1.Visible Then
        tvTreeView.Top = CoolBar1.Height + picTitles.Height
    Else
        tvTreeView.Top = picTitles.Height
    End If
    lvListView.Top = tvTreeView.Top
    
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - 10 - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - 10 - (picTitles.Top + picTitles.Height)
    End If
    lvListView.Height = tvTreeView.Height
End Sub

Private Sub Form_Activate()
    'TipoDeJanalaAtiva_pi = 1
    
    'JanelaAtiva_tp.TipoDeJanela_int = enHD_TDJ_Explorer
    'JanelaAtiva_tp.ArgumentoChave_str = Empty
    'Set JanelaAtiva_tp.JanelaAtiva_frm = Me
    
End Sub

Private Sub Form_Deactivate()
'    JanelaAtiva_tp.TipoDeJanela_int = enHD_TDJ_Nenhuma
'    JanelaAtiva_tp.ArgumentoChave_str = Empty
'    Set JanelaAtiva_tp.JanelaAtiva_frm = Nothing

    'TipoDeJanalaAtiva_pi = 0
End Sub

Private Sub Form_Load()
    Dim nodX As Node
    'Set SourceNode = Nothing
    
    MedidaTopo_s = picTitles.Top + picTitles.Height + 10
    
    tvTreeView.Top = 1
    
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    tvTreeView.ImageList = ImageListGeral
    lvListView.Icons = ImageListGeralGrande
    lvListView.SmallIcons = ImageListGeral
    'lvListView.ColumnHeaderIcons = ImageListGeral
    
    Set nodX = tvTreeView.Nodes.Add(, , "pad_Laboratório", "Laboratório", "ico_Laboratório", "ico_Laboratório")
    Set nodX = tvTreeView.Nodes.Add(, , "pad_Trânsito", "Trânsito", "ico_Trânsito", "ico_Trânsito")
    Set nodX = tvTreeView.Nodes.Add(, , "pad_Sucata", "Sucata", "ico_Sucata", "ico_Sucata")
    
    
    
    Set nodX = tvTreeView.Nodes.Add(, , "emp_Matriz", "Matriz", "ico_Empresa", "ico_Empresa")
    Set nodX = tvTreeView.Nodes.Add("emp_Matriz", tvwChild, "dep_Recursos Humanos", "Recursos Humanos", "ico_Departamento")
    Set nodX = tvTreeView.Nodes.Add("emp_Matriz", tvwChild, "dep_Contabilidade", "Contabilidade", "ico_Departamento")
    
    Set nodX = tvTreeView.Nodes.Add("emp_Matriz", tvwChild, "dep_Informática", "Informática", "ico_Departamento")
    
     Set nodX = tvTreeView.Nodes.Add("dep_Recursos Humanos", tvwChild, "set_Recrutamento", "Recrutamento", "ico_SalaFechada", "ico_SalaAberta")
    Set nodX = tvTreeView.Nodes.Add("dep_Recursos Humanos", tvwChild, "set_Admissão", "Admissão", "ico_SalaFechada", "ico_SalaAberta")
    Set nodX = tvTreeView.Nodes.Add("dep_Informática", tvwChild, "set_Suporte", "Suporte", "ico_SalaFechada", "ico_SalaAberta")
    Set nodX = tvTreeView.Nodes.Add("dep_Informática", tvwChild, "set_Desenvolvimento", "Sala 13 - Desenvolvimento", "ico_SalaFechada", "ico_SalaAberta")
    
    'Computador 211
    Set nodX = tvTreeView.Nodes.Add("set_Desenvolvimento", tvwChild, "equ_Computador211", "Computador 211", "ico_Computador")
        'Hardware
        Set nodX = tvTreeView.Nodes.Add("equ_Computador211", tvwChild, "hdw_Computador211", "Hardware", "ico_Hardware")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador211", tvwChild, "peç_Monitor", "Monitor Philips", "ico_Monitor")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador211", tvwChild, "peç_HD", "HD Maxtor 1958000FFS8", "ico_HD")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador211", tvwChild, "peç_Drive", "Drive 1.44", "ico_Drive3.5")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador211", tvwChild, "peç_Modem", "Modem", "ico_Modem")
             Set nodX = tvTreeView.Nodes.Add("hdw_Computador211", tvwChild, "peç_Teclado", "Teclado Genérico 101 EUA", "ico_Teclado")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador211", tvwChild, "peç_Mouse", "Mouse 2 Teclas", "ico_Mouse")
    
    'Software
        Set nodX = tvTreeView.Nodes.Add("equ_Computador211", tvwChild, "sfw_Computador211", "Software", "ico_Software")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador211", tvwChild, "prg_Programa001", "Windows 2000", "ico_SistemaOperacional")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador211", tvwChild, "prg_Programa002", "Windows 98SE", "ico_SistemaOperacional")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador211", tvwChild, "prg_Programa003", "Visual Basic 6", "ico_Programa")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador211", tvwChild, "prg_Programa004", "Explorer 6.0", "ico_Programa")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador211", tvwChild, "prg_Programa005", "MSDE 2000", "ico_Programa")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador211", tvwChild, "prg_Programa006", "Access 97", "ico_Programa")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador211", tvwChild, "prg_Programa007", "Explorer 6.0", "ico_Programa")
            
    'Computador 028
    Set nodX = tvTreeView.Nodes.Add("set_Desenvolvimento", tvwChild, "equ_Computador028", "Computador 028", "ico_Computador")
        'Hardware
        Set nodX = tvTreeView.Nodes.Add("equ_Computador028", tvwChild, "hdw_Computador028", "Hardware", "ico_Hardware")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador028", tvwChild, "peç_028_Monitor", "Monitor Philips", "ico_Monitor")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador028", tvwChild, "peç_028_HD", "HD Maxtor 1958000FFS8", "ico_HD")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador028", tvwChild, "peç_028_Drive", "Drive 1.44", "ico_Drive3.5")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador028", tvwChild, "peç_028_Modem", "Modem", "ico_Modem")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador028", tvwChild, "peç_028_Teclado", "Teclado Genérico 101 EUA", "ico_Teclado")
            Set nodX = tvTreeView.Nodes.Add("hdw_Computador028", tvwChild, "peç_028_Mouse", "Mouse 2 Teclas", "ico_Mouse")
    
    'Software
        Set nodX = tvTreeView.Nodes.Add("equ_Computador028", tvwChild, "sfw_Computador028", "Software", "ico_Software")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador028", tvwChild, "prg_028_Programa001", "Windows 2000", "ico_SistemaOperacional")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador028", tvwChild, "prg_028_Programa002", "Windows 98SE", "ico_SistemaOperacional")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador028", tvwChild, "prg_028_Programa003", "Visual Basic 6", "ico_Programa")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador028", tvwChild, "prg_028_Programa004", "Explorer 6.0", "ico_Programa")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador028", tvwChild, "prg_028_Programa005", "MSDE 2000", "ico_Programa")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador028", tvwChild, "prg_028_Programa006", "Access 97", "ico_Programa")
            Set nodX = tvTreeView.Nodes.Add("sfw_Computador028", tvwChild, "prg_028_Programa007", "Explorer 6.0", "ico_Programa")
    
    Set nodX = tvTreeView.Nodes.Add(, , "emp_PV01", "PV01 - Areal", "ico_Empresa", "ico_Empresa")
    Set nodX = tvTreeView.Nodes.Add("emp_PV01", tvwChild, "dep_Geral", "Geral", "ico_Departamento")
    Set nodX = tvTreeView.Nodes.Add("dep_Geral", tvwChild, "set_CPD", "CPD", "ico_SalaFechada", "ico_SalaAberta")
        
    nodX.EnsureVisible
    tvTreeView.Refresh
    
End Sub





Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Unload Me

End Sub



Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
    If Source = imgSplitter Then
        SizeControls x
    End If
End Sub


Sub SizeControls(x As Single)
    On Error Resume Next
    
    If x < 1500 Then x = 1500
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    
    
    tvTreeView.Width = x
    imgSplitter.Left = x
    lvListView.Left = x + 40
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40
    'If tbToolBar.Visible Then
    If CoolBar1.Visible Then
        'tvTreeView.Top = tbToolBar.Height + picTitles.Height
        'tvTreeView.Top = MedidaTopo_s
        tvTreeView.Top = CoolBar1.Height + picTitles.Height
    Else
        tvTreeView.Top = picTitles.Height
    End If
    lvListView.Top = tvTreeView.Top

    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - 10 - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - 10 - (picTitles.Top + picTitles.Height)
    End If

    lvListView.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Dim nodX As Node
    Dim i As Integer
        
    Select Case Button.Key
        Case "Computador"
            Set nodX = tvTreeView.Nodes.Add(, , "Computador", "Computador", 1, 1)
            Set nodX = tvTreeView.Nodes.Add("Computador", tvwChild, "Hardware", "Hardware", 2)
            Set nodX = tvTreeView.Nodes.Add("Computador", tvwChild, "Software", "Software", 3)
            
        Case "Hardware"
            If tvTreeView.SelectedItem.Text = "" Then Exit Sub
            Set nodX = tvTreeView.Nodes.Add(tvTreeView.SelectedItem.Text, tvwChild, "Hardware", "Hardware", 2)
            
        Case "SoftWare"
            If tvTreeView.SelectedItem.Text = "" Then Exit Sub
            Set nodX = tvTreeView.Nodes.Add(tvTreeView.SelectedItem.Text, tvwChild, "Software", "SoftWare", 3)
            
    End Select
    nodX.EnsureVisible
    tvTreeView.Refresh
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    'tbToolBar.Visible = mnuViewToolbar.Checked
    CoolBar1.Visible = mnuViewToolbar.Checked
    SizeControls imgSplitter.Left
End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    tvAdicionarItem Me.tvTreeView, (Nó.Atual.Key), "novo_", "Novo Item", "ico_Computador", "ico_Computador"
    
End Sub

Private Sub ToolbarFerramentasGerais_ButtonClick(ByVal Button As MSComctlLib.Button)
            Select Case Button.Key
                Case "Ícones Grandes"
                    lvListView.View = lvwIcon
                Case "Ícones Pequenos"
                    lvListView.View = lvwSmallIcon
                Case "Ícones Lista"
                    lvListView.View = lvwList
                Case "Ícones Detalhes"
                    Dim clr As ColumnHeader
                    lvListView.View = lvwReport
                    'Set clr = frmJanelaAtiva.lvListView.ColumnHeader.Add(, , "Item")
                
                    'Set clr = frmJanelaAtiva.lvListView.ColumnHeader.Add(, , "Item", , , 2)
                    'Set clr = frmJanelaAtiva.lvListView.ColumnHeader.Add(, , "Item", , , 3)
            End Select
            lvListView.Sorted = True
End Sub

Private Sub tvTreeView_BeforeLabelEdit(Cancel As Integer)
   ' MsgBox Nó.CódigoDoNó
    
    
End Sub

Private Sub tvTreeView_DragDrop(Source As Control, x As Single, y As Single)
    If Not (tvTreeView.DropHighlight Is Nothing) Then
        Set SourceNode.Parent = tvTreeView.DropHighlight
        'SourceNode.Key
        Set SourceNode = tvTreeView.DropHighlight
        'tvTreeView.Nodes.Item
        Set tvTreeView.DropHighlight = Nothing
    End If
    Set SourceNode = Nothing
    SourceType = tp_TNÓ_Nenhum
End Sub

Private Sub tvTreeView_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim target As Node
    Dim highlight As Boolean

    ' See what node we're above.
    Set target = tvTreeView.HitTest(x, y)
    
    ' If it's the same as last time, do nothing.
    If target Is TargetNode Then Exit Sub
    Set TargetNode = target
    
    highlight = False
    If Not (TargetNode Is Nothing) Then
        ' See what kind of node were above.
        'Select Case NodeType(TargetNode)
        Select Case SourceType
            Case tp_TNÓ_software
                If NodeType(TargetNode) + 7 = SourceType Then highlight = True
                Debug.Print NodeType(TargetNode) + 7 & " " & SourceType
            Case tp_TNÓ_Equipamento
                If NodeType(TargetNode) + 2 = SourceType Then highlight = True
            Case Else
                If NodeType(TargetNode) + 1 = SourceType Then highlight = True
        End Select
        'If NodeType(TargetNode) + 1 = SourceType Then _
            highlight = True
    End If
    
    If highlight Then
        Set tvTreeView.DropHighlight = TargetNode
    Else
        Set tvTreeView.DropHighlight = Nothing
    End If
End Sub

Private Sub tvTreeView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set SourceNode = tvTreeView.HitTest(x, y)
'    tvTreeView_Click
    
    
End Sub

Private Sub tvTreeView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        'If SourceNode = Nothing Then Exit Sub
        SourceType = NodeType(SourceNode)
        Set tvTreeView.SelectedItem = SourceNode
        tvTreeView.DragIcon = ImageDrag.Picture
        tvTreeView.Drag vbBeginDrag
    End If
End Sub

Private Sub tvTreeView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'MsgBox Button
    'MsgBox Nó.NaturezaDoNó
    Select Case Nó.NaturezaDoNó
        
    End Select
    If Button = 2 Then PopupMenu MDIForm1.mnuJanelaExplorerTreeView ' .mnuFerramentas 'mnuPropriedades
    
    
End Sub

Private Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    'MsgBox Node.Child
'tvTreeView_Click
    Dim Node_Filho As Node
    Dim list_item As ListItem
    
    Nó.CódigoDoNó = Mid(Node.Key, 5, Len(Node.Key) - 4)
    Nó.RótuloDoNó = Node.Text
    Nó.NaturezaDoNó = Mid(Node.Key, 1, 3)
    Set Nó.Atual = Node

    If Node.Children <> 0 Then
        Set Node_Filho = Node.Child
        lblTitle(1).Caption = Node.FullPath
        lvListView.ListItems.Clear
        incounter = Node_Filho.FirstSibling.Index
        imagem = tvTreeView.Nodes.Item(incounter).Image
        Set list_item = lvListView.ListItems.Add(, Node_Filho.Key & incounter, tvTreeView.Nodes.Item(incounter), imagem, imagem)
        While incounter <> Node_Filho.LastSibling.Index
            imagem = tvTreeView.Nodes.Item(incounter).Next.Image
            'Set list_item = lvListView.ListItems.Add(, Node_Filho.Key & incounter, tvTreeView.Nodes.Item(incounter), imagem, imagem)
            Set list_item = lvListView.ListItems.Add(, Node_Filho.Next.Key & incounter, tvTreeView.Nodes.Item(incounter).Next.Text, imagem, imagem)
            incounter = tvTreeView.Nodes(incounter).Next.Index
        Wend
    Else
        lvListView.ListItems.Clear
    End If
    
End Sub
Private Function NodeType(test_node As Node) As tp_TipodeNó
    'If test_node = Nothing Then Exit Sub
    On Error GoTo erro1
    Select Case Left$(test_node.Key, 3)
        Case cns_TPNÓ_EMPRESA
            NodeType = tp_TNÓ_Empresa
        Case cns_TPNÓ_DEPARTAMENTO
            NodeType = tp_TNÓ_Departamento
        Case cns_TPNÓ_SETOR
            NodeType = tp_TNÓ_Setor
        Case cns_TPNÓ_SALA
            NodeType = tp_TNÓ_Sala
        Case cns_TPNÓ_EQUIPAMENTO
            NodeType = tp_TNÓ_Equipamento
        Case cns_TPNÓ_HARDWARE
            NodeType = tp_TNÓ_Hardware
        Case cns_TPNÓ_SOFTWARE
            NodeType = tp_TNÓ_software
        Case cns_TPNÓ_PEÇA
            NodeType = tp_TNÓ_Peça
        Case cns_TPNÓ_SOFTWARE
            NodeType = tp_TNÓ_software
        Case cns_TPNÓ_PROGRAMA
            NodeType = tp_TNÓ_Programa
        Case Else
            NodeType = tp_TNÓ_Nenhum
    End Select
    
    Exit Function
erro1:
    If Err = 91 Then Resume Next
    
    
End Function


