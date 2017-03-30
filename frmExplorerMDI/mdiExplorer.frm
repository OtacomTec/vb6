VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8295
   ClientLeft      =   1710
   ClientTop       =   1755
   ClientWidth     =   12855
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin VB.PictureBox imgSplitter 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7905
      Left            =   2685
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7905
      ScaleWidth      =   30
      TabIndex        =   6
      Top             =   390
      Width           =   30
      Begin VB.PictureBox picSplitter 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillColor       =   &H00808080&
         Height          =   4800
         Left            =   300
         ScaleHeight     =   2090.126
         ScaleMode       =   0  'User
         ScaleWidth      =   780
         TabIndex        =   7
         Top             =   30
         Visible         =   0   'False
         Width           =   72
      End
   End
   Begin VB.PictureBox picPainelVertical 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7905
      Left            =   0
      ScaleHeight     =   7905
      ScaleWidth      =   2685
      TabIndex        =   4
      Top             =   390
      Width           =   2685
      Begin MSComctlLib.TreeView tvTreeView 
         Height          =   4800
         Left            =   30
         TabIndex        =   5
         Top             =   60
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   8467
         _Version        =   393217
         Indentation     =   441
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
      End
      Begin VB.Image ImageDrag 
         Height          =   480
         Left            =   2190
         Picture         =   "mdiExplorer.frx":0000
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList ImageListGeralGrande 
      Left            =   630
      Top             =   450
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
            Picture         =   "mdiExplorer.frx":0152
            Key             =   "ico_Laboratório"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":05A6
            Key             =   "ico_Usuário"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":09FA
            Key             =   "ico_Computador"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":31AE
            Key             =   "ico_Software"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":34CA
            Key             =   "ico_SistemaOperacional"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":5C7E
            Key             =   "ico_Programa"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":5F9A
            Key             =   "ico_Departamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":62B6
            Key             =   "ico_Drive3.5"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":8A6A
            Key             =   "ico_Drive5.25"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":B21E
            Key             =   "ico_DriveCD"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":B672
            Key             =   "ico_DriveJazz250"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":C4C6
            Key             =   "ico_Empresa"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":C7E2
            Key             =   "ico_Hardware"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":CAFE
            Key             =   "ico_HD"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":F2B2
            Key             =   "ico_ÁreaDeTrabalho"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":F706
            Key             =   "ico_Memória"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":FB5A
            Key             =   "ico_Modem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":FFAE
            Key             =   "ico_Mouse"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":102CA
            Key             =   "ico_Placa"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":105E6
            Key             =   "ico_PontoDeRede"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":10902
            Key             =   "ico_Processador"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":10D56
            Key             =   "ico_SalaFechada"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":11072
            Key             =   "ico_SalaAberta"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1138E
            Key             =   "ico_SCSI"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":117E2
            Key             =   "ico_Servidor"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":11C36
            Key             =   "ico_Teclado"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1452A
            Key             =   "ico_USB"
            Object.Tag             =   "ico_Trânsito"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1497E
            Key             =   "ico_Monitor"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1525A
            Key             =   "ico_Sucata"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":15576
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListToolBar 
      Left            =   1200
      Top             =   450
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
            Picture         =   "mdiExplorer.frx":15892
            Key             =   "tool_Ícones Grandes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":159A6
            Key             =   "tool_Ícones Pequenos"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":15ABA
            Key             =   "tool_Ícones Lista"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":15BCE
            Key             =   "tool_Ícones Detalhes"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":15CE2
            Key             =   "tool_Novo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":15DF6
            Key             =   "tool_Abrir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":15F0A
            Key             =   "tool_Salvar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1601E
            Key             =   "tool_Imprimir"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":16132
            Key             =   "tool_Recortar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":16246
            Key             =   "tool_Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1635A
            Key             =   "tool_Colar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1646E
            Key             =   "tool_Explorer"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1678A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1689C
            Key             =   "tool_Explorer2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":169AE
            Key             =   "tool_SetaDireita"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":16D00
            Key             =   "tool_SetaEsquerda"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   688
      ImageList       =   "ImageListToolBar"
      _CBWidth        =   12855
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Child1          =   "ToolbarPadrão"
      MinHeight1      =   330
      Width1          =   3270
      NewRow1         =   0   'False
      Child2          =   "Toolbar"
      MinHeight2      =   330
      Width2          =   6525
      NewRow2         =   0   'False
      Child3          =   "ToolbarFerramentasGerais"
      MinHeight3      =   330
      Width3          =   3360
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar ToolbarFerramentasGerais 
         Height          =   330
         Left            =   10020
         TabIndex        =   3
         Top             =   30
         Width           =   2745
         _ExtentX        =   4842
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
      Begin MSComctlLib.Toolbar Toolbar 
         Height          =   330
         Left            =   3465
         TabIndex        =   2
         Top             =   30
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageListToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
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
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarPadrão 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
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
   End
   Begin MSComctlLib.ImageList ImageListGeral 
      Left            =   60
      Top             =   450
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
            Picture         =   "mdiExplorer.frx":17052
            Key             =   "ico_Laboratório"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":174A6
            Key             =   "ico_Computador"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":19C5A
            Key             =   "ico_Software"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":19F76
            Key             =   "ico_SistemaOperacional"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1C72A
            Key             =   "ico_Programa"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1CA46
            Key             =   "ico_Departamento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1CD62
            Key             =   "ico_Drive3.5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":1F516
            Key             =   "ico_Drive5.25"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":21CCA
            Key             =   "ico_DriveCD"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":2211E
            Key             =   "ico_DriveJazz250"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":22F6A
            Key             =   "ico_Empresa"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":23286
            Key             =   "ico_Hardware"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":235A2
            Key             =   "ico_HD"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":25D56
            Key             =   "ico_ÁreaDeTrabalho"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":261AA
            Key             =   "ico_Memória"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":265FE
            Key             =   "ico_Modem"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":26A52
            Key             =   "ico_Mouse"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":26D6E
            Key             =   "ico_Placa"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":2708A
            Key             =   "ico_PontoDeRede"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":273A6
            Key             =   "ico_Processador"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":277FA
            Key             =   "ico_SalaFechada"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":27B16
            Key             =   "ico_SalaAberta"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":27E32
            Key             =   "ico_SCSI"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":28286
            Key             =   "ico_Servidor"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":286DA
            Key             =   "ico_Teclado"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":2AFCE
            Key             =   "ico_USB"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":2B422
            Key             =   "ico_Trânsito"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":2B876
            Key             =   "ico_Monitor"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":2C152
            Key             =   "ico_Sucata"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":2C46E
            Key             =   "ico_Explorer"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiExplorer.frx":2C78A
            Key             =   "ico_Usuário"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFerramentas 
      Caption         =   "Ferramentas"
      Begin VB.Menu itmBarra 
         Caption         =   "Barra1"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu itmBarra 
         Caption         =   "Barra2"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu itmBarra 
         Caption         =   "Barra3"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mnuBotãoDireito 
      Caption         =   "Botão Direito"
      Visible         =   0   'False
      Begin VB.Menu mnuJanelaExplorerTreeView 
         Caption         =   "JanelaExplorerTreeView"
         Begin VB.Menu itmJanelaExplorer 
            Caption         =   "&Enviar Para..."
            Index           =   1
         End
         Begin VB.Menu itmJanelaExplorer 
            Caption         =   "&Editar"
            Index           =   2
         End
         Begin VB.Menu itmJanelaExplorer 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu itmJanelaExplorer 
            Caption         =   "&Propriedades"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  Dim i As Integer

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

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    Me.tvTreeView.Height = Me.picPainelVertical.Height - 50
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    i = 1
End Sub

Private Sub itmBarra_Click(Index As Integer)
    itmBarra(Index).Checked = Not CoolBar1.Bands(Index).Visible
    CoolBar1.Bands(Index).Visible = itmBarra(Index).Checked
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Explorer"
            NovoExplorerLocalização "Novo Explorer"
    End Select
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    
End Sub

Private Sub itmJanelaExplorer_Click(Index As Integer)
    Select Case Index
        Case 2 'Editar
            'EmpresaEditar (SourceNode.Text)
            EmpresaEditar (SourceNode.Text)
            
        Case 4
            NovoPropriedades (SourceNode.Text)
    End Select
End Sub

Private Sub MDIForm_Load()
    Dim nodX As Node
    Me.tvTreeView.Width = Me.picPainelVertical.Width - 50
    Me.tvTreeView.Height = Me.picPainelVertical.Height - 100
    
    'Set SourceNode = Nothing
    tvTreeView.Top = 1
    
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    tvTreeView.ImageList = ImageListGeral
    'lvListView.Icons = ImageListGeralGrande
    'lvListView.SmallIcons = ImageListGeral
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

Private Sub MDIForm_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    'SizeControls imgSplitter.Left
    Me.tvTreeView.Height = Me.picPainelVertical.Height - 50
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Explorer"
            NovoExplorerLocalização "Computadores Explorer"
        Case "Adicionar Empresa"
        
        Case Else
            frm_HD_Novo.Show
    End Select
    frmExplorer.Show
    
End Sub

Private Sub ToolbarFerramentasGerais_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case JanelaAtiva_tp.TipoDeJanela_int
        Case enHD_TDJ_Explorer
            Select Case Button.Key
                Case "Ícones Grandes"
                    JanelaAtiva_tp.JanelaAtiva_frm.lvListView.View = lvwIcon
                Case "Ícones Pequenos"
                    JanelaAtiva_tp.JanelaAtiva_frm.lvListView.View = lvwSmallIcon
                Case "Ícones Lista"
                    JanelaAtiva_tp.JanelaAtiva_frm.lvListView.View = lvwList
                Case "Ícones Detalhes"
                    Dim clr As ColumnHeader
                    JanelaAtiva_tp.JanelaAtiva_frm.lvListView.View = lvwReport
                    'Set clr = frmJanelaAtiva.lvListView.ColumnHeader.Add(, , "Item")
                
                    'Set clr = frmJanelaAtiva.lvListView.ColumnHeader.Add(, , "Item", , , 2)
                    'Set clr = frmJanelaAtiva.lvListView.ColumnHeader.Add(, , "Item", , , 3)
            End Select
        Case Else
        
    End Select
End Sub

Private Sub ToolbarPadrão_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim rsCL As clBD
    Dim rs As adodb.Recordset
    Set rsCL = New clBD
    Dim lstrSQL As String
    
    
    

    Select Case Button.Key
        Case "Novo": frm_HD_Novo.Show 1
        Case "Abrir"
            frm_HD_Novo.Caption = "Abrir..."
            frm_HD_Novo.Show 1
        Case "Salvar"
            'Select Case JanelaAtiva_tp.TipoDeJanela_int
             '   Case enHD_TDJ_EmpresaEditar
                    Salvar (JanelaAtiva_tp.TipoDeJanela_int)
             '   Case Else
                
            'End Select
    End Select
End Sub

Private Sub ToolbarPadrão_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Parent.Key
        Case "Novo"
            Select Case ButtonMenu.Key
                Case "Empresa"
                    frm_HD_Novo.Show
            End Select
    End Select
End Sub
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    

    'If mbMoving Then
    '    sglPos = X + imgSplitter.Left
    '    If sglPos < sglSplitLimit Then
    '        picSplitter.Left = sglSplitLimit
    '    ElseIf sglPos > Me.Width - sglSplitLimit Then
    '        picSplitter.Left = Me.Width - sglSplitLimit
    '    Else
    '        picSplitter.Left = sglPos
    '    End If
    'End If
    
    'On Error Resume Next
    If i = 1 Then
       Me.picPainelVertical.Width = x + Me.picPainelVertical.Width
       Me.tvTreeView.Width = Me.picPainelVertical.Width - 50
        'Form1.Caption = "Side width=" & Picture1.Width
    End If

End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    i = 0
    'SizeControls picSplitter.Left
    'picSplitter.Visible = False
    'mbMoving = False
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
        'lblTitle(1).Caption = Node.FullPath
        'lvListView.ListItems.Clear
        incounter = Node_Filho.FirstSibling.Index
        imagem = tvTreeView.Nodes.Item(incounter).Image
        'Set list_item = lvListView.ListItems.Add(, Node_Filho.Key & incounter, tvTreeView.Nodes.Item(incounter), imagem, imagem)
        While incounter <> Node_Filho.LastSibling.Index
            imagem = tvTreeView.Nodes.Item(incounter).Next.Image
            'Set list_item = lvListView.ListItems.Add(, Node_Filho.Key & incounter, tvTreeView.Nodes.Item(incounter), imagem, imagem)
            'Set list_item = lvListView.ListItems.Add(, Node_Filho.Next.Key & incounter, tvTreeView.Nodes.Item(incounter).Next.Text, imagem, imagem)
            incounter = tvTreeView.Nodes(incounter).Next.Index
        Wend
    Else
        'lvListView.ListItems.Clear
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


