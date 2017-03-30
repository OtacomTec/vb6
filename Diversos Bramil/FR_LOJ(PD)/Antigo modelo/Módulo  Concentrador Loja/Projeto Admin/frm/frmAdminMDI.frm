VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{FBC84465-367A-4234-AF5D-140071197D57}#1.0#0"; "OCXMenu_Vertical.ocx"
Object = "{D0159C1D-A983-4698-8940-3BE45A260C35}#1.0#0"; "SegundoPlanoMDI.ocx"
Object = "{C5014412-BD55-402F-8335-07C273732964}#1.1#0"; "AplicativoUsuário.ocx"
Begin VB.MDIForm frmAdminMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Only  Tech - Concentrador de Loja"
   ClientHeight    =   8595
   ClientLeft      =   1155
   ClientTop       =   45
   ClientWidth     =   13890
   Icon            =   "frmAdminMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin AplicativoUsuárioOCX.AplicativoUsuário AplicativoUsuário 
      Index           =   0
      Left            =   5220
      Top             =   1110
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin OCXSegundoPlano.SegundoPlanoMDI SegundoPlano 
      Left            =   6480
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox imgSplitter 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8145
      Left            =   1635
      MousePointer    =   9  'Size W E
      ScaleHeight     =   8145
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   450
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
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
         Width           =   72
      End
   End
   Begin VB.PictureBox picPainelVertical 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8145
      Left            =   0
      ScaleHeight     =   8145
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   450
      Width           =   1635
      Begin OCXMenu_Vertical.MenuVertical mnvMDI 
         Height          =   11085
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   19553
         MenuCaption1    =   "Menu Principal"
         ItemMenuMax1    =   3
         ItemMenuÍcone11 =   "frmAdminMDI.frx":1782
         ItemMenuCaption11=   "Cadastros Base"
         ItemMenuÍcone12 =   "frmAdminMDI.frx":1A9C
         ItemMenuCaption12=   "Interfaces"
         ItemMenuÍcone13 =   "frmAdminMDI.frx":1DB6
         ItemMenuCaption13=   "Gerencial"
      End
      Begin VB.Image ImageDrag 
         Height          =   480
         Left            =   690
         Picture         =   "frmAdminMDI.frx":20D0
         Top             =   1170
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList ImageListGeralGrande 
      Left            =   5790
      Top             =   1080
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
            Picture         =   "frmAdminMDI.frx":2222
            Key             =   "ico_Laboratório"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2676
            Key             =   "ico_Usuário"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2ACA
            Key             =   "ico_Computador"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":527E
            Key             =   "ico_Software"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":559A
            Key             =   "ico_SistemaOperacional"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":7D4E
            Key             =   "ico_Programa"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":806A
            Key             =   "ico_Departamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":8386
            Key             =   "ico_Drive3.5"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":AB3A
            Key             =   "ico_Drive5.25"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":D2EE
            Key             =   "ico_DriveCD"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":D742
            Key             =   "ico_DriveJazz250"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":E596
            Key             =   "ico_Empresa"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":E8B2
            Key             =   "ico_Hardware"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":EBCE
            Key             =   "ico_HD"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":11382
            Key             =   "ico_ÁreaDeTrabalho"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":117D6
            Key             =   "ico_Memória"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":11C2A
            Key             =   "ico_Modem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1207E
            Key             =   "ico_Mouse"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1239A
            Key             =   "ico_Placa"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":126B6
            Key             =   "ico_PontoDeRede"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":129D2
            Key             =   "ico_Processador"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":12E26
            Key             =   "ico_SalaFechada"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":13142
            Key             =   "ico_SalaAberta"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1345E
            Key             =   "ico_SCSI"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":138B2
            Key             =   "ico_Servidor"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":13D06
            Key             =   "ico_Teclado"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":165FA
            Key             =   "ico_USB"
            Object.Tag             =   "ico_Trânsito"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":16A4E
            Key             =   "ico_Monitor"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1732A
            Key             =   "ico_Sucata"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":17646
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListToolBar 
      Left            =   5520
      Top             =   1770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":17962
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1966C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":19F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1AC20
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1BA72
            Key             =   "tool_AreaDeTrabalho"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1BB84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":21E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":23B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":250BA
            Key             =   "tool_AreaDeTrabalho2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2564C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2596E
            Key             =   "ico_calculadora"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":25C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":27DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":28C14
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   794
      BandCount       =   1
      ImageList       =   "ImageListToolBar"
      _CBWidth        =   13890
      _CBHeight       =   450
      _Version        =   "6.7.8988"
      Child1          =   "ToolbarAreaDeTrabalho"
      MinHeight1      =   390
      Width1          =   2895
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar ToolbarAreaDeTrabalho 
         Height          =   390
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   13770
         _ExtentX        =   24289
         _ExtentY        =   688
         ButtonWidth     =   3704
         ButtonHeight    =   582
         TextAlignment   =   1
         ImageList       =   "ImageListToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Gerenciador de Tarefas"
               Key             =   "GerenciadorDeTarefas"
               Description     =   "Gerencie Tarefas da área de trabalho"
               Object.ToolTipText     =   "Gerencie Tarefas da área de trabalho"
               ImageIndex      =   3
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Gerenciador de Logins"
               Key             =   "Login"
               Description     =   "Gerencie Logins na área de trabalho"
               Object.ToolTipText     =   "Opções de Login - Gerencie Logins na área de trabalho"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Calculadora"
               Key             =   "calculadora"
               Description     =   "Calculadora para uso no sistema"
               Object.ToolTipText     =   "Calculadora para uso no sistema"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Bloco de Notas"
               Key             =   "bloco_notas"
               Description     =   "Bloco de Notas para uso no sistema"
               ImageIndex      =   14
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
            Picture         =   "frmAdminMDI.frx":28F2E
            Key             =   "ico_Laboratório"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":29382
            Key             =   "ico_Computador"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2BB36
            Key             =   "ico_Software"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2BE52
            Key             =   "ico_SistemaOperacional"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2E606
            Key             =   "ico_Programa"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2E922
            Key             =   "ico_Departamento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2EC3E
            Key             =   "ico_Drive3.5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":313F2
            Key             =   "ico_Drive5.25"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":33BA6
            Key             =   "ico_DriveCD"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":33FFA
            Key             =   "ico_DriveJazz250"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":34E46
            Key             =   "ico_Empresa"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":35162
            Key             =   "ico_Hardware"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3547E
            Key             =   "ico_HD"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":37C32
            Key             =   "ico_ÁreaDeTrabalho"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":38086
            Key             =   "ico_Memória"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":384DA
            Key             =   "ico_Modem"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3892E
            Key             =   "ico_Mouse"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":38C4A
            Key             =   "ico_Placa"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":38F66
            Key             =   "ico_PontoDeRede"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":39282
            Key             =   "ico_Processador"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":396D6
            Key             =   "ico_SalaFechada"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":399F2
            Key             =   "ico_SalaAberta"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":39D0E
            Key             =   "ico_SCSI"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3A162
            Key             =   "ico_Servidor"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3A5B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3CEAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3D2FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3D752
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3E02E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3E34A
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3E666
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListGeralPequeno 
      Left            =   4590
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3EABA
            Key             =   "ico_Aplicativo"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuAreaDeTrabalho 
      Caption         =   "Área de Trabalho"
      WindowList      =   -1  'True
      Begin VB.Menu mnuExibir 
         Caption         =   "Exi&bir"
         Begin VB.Menu itmBarra 
            Caption         =   "Exibir Barra Vertical"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu itmBarra 
            Caption         =   "Barra Padrão"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu itmBarra 
            Caption         =   "Barra de Login e Área de Trabalho"
            Checked         =   -1  'True
            Index           =   2
         End
      End
      Begin VB.Menu mnuAreaDeTrabalhoOrganizar 
         Caption         =   "Organizar"
         Begin VB.Menu itmAreaDeTrabalhoOrganizar 
            Caption         =   "Em Cascata"
            Index           =   0
         End
         Begin VB.Menu itmAreaDeTrabalhoOrganizar 
            Caption         =   "Lado a Lado Horizontal"
            Index           =   1
         End
         Begin VB.Menu itmAreaDeTrabalhoOrganizar 
            Caption         =   "Lado a Lado Vertical"
            Index           =   2
         End
         Begin VB.Menu itmAreaDeTrabalhoOrganizar 
            Caption         =   "Ícones"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "&Cadastros"
      Visible         =   0   'False
      Begin VB.Menu smnCadastro_empresa 
         Caption         =   "&Empresas"
      End
      Begin VB.Menu smnParametros_gerais 
         Caption         =   "&Parâmetros_gerais"
      End
      Begin VB.Menu smnCadastro_cidades 
         Caption         =   "&Cidades"
      End
      Begin VB.Menu smnCadastro_clientes 
         Caption         =   "C&lientes"
      End
      Begin VB.Menu smnCadastro_pacientes 
         Caption         =   "P&acientes"
      End
      Begin VB.Menu smnCadastro_oftalmo 
         Caption         =   "&Oftalmo"
      End
      Begin VB.Menu smnCadastro_consulta_produto 
         Caption         =   "Co&nsulta produto"
      End
   End
   Begin VB.Menu mnuMercadologico 
      Caption         =   "&Mercadológico"
      Visible         =   0   'False
      Begin VB.Menu smnMercadologico_secao 
         Caption         =   "&Seção"
      End
      Begin VB.Menu mnuMercadologico_Produto 
         Caption         =   "&Produto"
      End
      Begin VB.Menu mnuMercadologico_analise_posicao_estoque 
         Caption         =   "&Análise posição de estoque"
      End
      Begin VB.Menu mnuMercadologico_resumo_balanco 
         Caption         =   "&Resumo balanço"
      End
   End
   Begin VB.Menu mnuAdministracao 
      Caption         =   "&Administração"
      Visible         =   0   'False
      Begin VB.Menu smnAdministracao_supervisor 
         Caption         =   "&Supervisor"
      End
      Begin VB.Menu smnAdministracao_vendedor 
         Caption         =   "&Vendedor"
      End
      Begin VB.Menu smnAdministracao_plano_pagamentos 
         Caption         =   "&Plano de pagamentos"
      End
      Begin VB.Menu smnAdministracao_historico_padrao 
         Caption         =   "&Histórico padrão"
      End
      Begin VB.Menu smnAdministracao_faixa_comissao 
         Caption         =   "&Faixa de comissão"
         Begin VB.Menu smnAdministracao_faixa_comissao_Supervisor 
            Caption         =   "Supervisor"
         End
         Begin VB.Menu smnAdministracao_faixa_comissao_vendedor 
            Caption         =   "Vendedor"
         End
      End
      Begin VB.Menu smnAdministracao_tabela_precos 
         Caption         =   "&Tabela de preços"
      End
      Begin VB.Menu smnAdministracao_tabela_custos 
         Caption         =   "Tabela de &custos"
      End
      Begin VB.Menu mnuMala_direta 
         Caption         =   "&Mala direta"
         Begin VB.Menu smnMala_direta_etiquetas 
            Caption         =   "&Etiquetas"
            Begin VB.Menu smnMala_direta_etiquetas_Fornecedor 
               Caption         =   "Fornecedor"
            End
            Begin VB.Menu smnMala_direta_etiquetas_Cliente 
               Caption         =   "Cliente"
            End
         End
         Begin VB.Menu smnMala_direta 
            Caption         =   "&Carta de cobrança"
            Begin VB.Menu sisMala_direta_carta_cobranca_Modelo1 
               Caption         =   "Modelo 1"
            End
            Begin VB.Menu sisMala_direta_carta_cobranca_Modelo2 
               Caption         =   "Modelo 2"
            End
            Begin VB.Menu sisMala_direta_carta_cobranca_Modelo3 
               Caption         =   "Modelo 3"
            End
            Begin VB.Menu sisMala_direta_carta_cobranca_Modelo4 
               Caption         =   "Modelo 4"
            End
         End
         Begin VB.Menu smnMala_direta_marketing 
            Caption         =   "&Marketing"
            Begin VB.Menu sisMala_direta_marketing_Renovar_oculos 
               Caption         =   "Renovar Óculos"
            End
            Begin VB.Menu sisMala_direta_marketing_Aviso_oculos_prontos 
               Caption         =   "Aviso Óculos prontos"
            End
            Begin VB.Menu sisMala_direta_marketing_anuncio_lancamento 
               Caption         =   "Anúncio de lançamento"
            End
         End
      End
   End
   Begin VB.Menu mnuComercial 
      Caption         =   "Comercia&l"
      Visible         =   0   'False
      Begin VB.Menu mnuVendas 
         Caption         =   "&Vendas"
         Begin VB.Menu smnVendas_receituario 
            Caption         =   "&Receituário"
         End
         Begin VB.Menu smnVendas_orcamentos 
            Caption         =   "&Orçamentos"
         End
         Begin VB.Menu smnVendas_relacao_entregas 
            Caption         =   "R&elação de entregas"
         End
         Begin VB.Menu smnVendas_ajuste_entrega 
            Caption         =   "&Ajuste de entregas"
         End
         Begin VB.Menu smnVendas_resumo_diario_vendas 
            Caption         =   "Re&sumo diário de vendas"
         End
         Begin VB.Menu smnVendas_comissoes 
            Caption         =   "&Comissões"
         End
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "Cai&xa"
         Begin VB.Menu smnCaixa_abertura 
            Caption         =   "&Abertura"
         End
         Begin VB.Menu smnCaixa_baixa_crediario 
            Caption         =   "&Baixa crediário"
         End
         Begin VB.Menu smnCaixa_imprime_carne 
            Caption         =   "&Imprime carnê"
         End
         Begin VB.Menu smnCaixa_emissao_recibo 
            Caption         =   "&Emissão de recibo"
         End
         Begin VB.Menu smnCaixa_movimento_periodo 
            Caption         =   "&Movimento do período"
         End
         Begin VB.Menu smnCaixa_recebimento_periodo 
            Caption         =   "&Recebimento do período"
         End
         Begin VB.Menu smnCaixa_mapa_resumo 
            Caption         =   "Ma&pa resumo"
         End
         Begin VB.Menu smnCaixa_extrato 
            Caption         =   "E&xtrato (Crédito e Débito)"
         End
         Begin VB.Menu smnCaixa_encerramento_dia 
            Caption         =   "E&ncerramento do dia"
         End
      End
      Begin VB.Menu mnCompras 
         Caption         =   "C&ompras"
         Begin VB.Menu smnCompras_fornecedor 
            Caption         =   "&Fornecedor"
         End
         Begin VB.Menu smnCompras_ordem_compra 
            Caption         =   "&Ordem de compra"
         End
         Begin VB.Menu smnCompras_necessidade_compra 
            Caption         =   "&Necessidade de compra"
         End
         Begin VB.Menu smnCompras_nota_entrada 
            Caption         =   "Nota de &Entrada"
         End
         Begin VB.Menu smnCompras_comparativo_ne_oc 
            Caption         =   "Co&mparativo nota de entrada X ordem de compra"
         End
         Begin VB.Menu smnCompras_notas_recebidas 
            Caption         =   "Notas &recebidas"
         End
      End
   End
   Begin VB.Menu mnuFinanceiro 
      Caption         =   "Financeiro"
      Visible         =   0   'False
      Begin VB.Menu mnuContas_receber 
         Caption         =   "Contas a &Receber"
         Begin VB.Menu smnContas_receber_alineas 
            Caption         =   "&Alineas"
         End
         Begin VB.Menu smnContas_receber_bancos 
            Caption         =   "&Bancos"
         End
         Begin VB.Menu smnContas_receber_titulos 
            Caption         =   "&Títulos"
         End
         Begin VB.Menu smnContas_receber_titulos_recebidos 
            Caption         =   "Títulos &recebidos"
         End
         Begin VB.Menu smnContas_receber_titulos_aberto 
            Caption         =   "T&ítulos em aberto"
         End
         Begin VB.Menu smnContas_receber_extrato_cliente 
            Caption         =   "&Extrato do cliente"
         End
         Begin VB.Menu smnContas_receber_clientes_atraso 
            Caption         =   "&Clientes em atraso"
         End
         Begin VB.Menu smnContas_receber_cheques_devolvidos 
            Caption         =   "C&heques devolvidos"
         End
         Begin VB.Menu smnContas_receber_relacao_cheques_devolvidos 
            Caption         =   "Relação de cheques &devolvidos"
         End
      End
      Begin VB.Menu mnuContas_pagar 
         Caption         =   "Contas a &Pagar"
         Begin VB.Menu smnContas_pagar_titulos 
            Caption         =   "&Títulos"
         End
         Begin VB.Menu smnContas_pagar_relacao_titulos 
            Caption         =   "&Relação de títulos"
         End
         Begin VB.Menu smnContas_pagar_extrato_diario 
            Caption         =   "&Extrato diário"
         End
         Begin VB.Menu smnContas_pagar_extrato_fornecedor 
            Caption         =   "E&xtrato fornecedor"
         End
      End
   End
   Begin VB.Menu mnuEstatistica 
      Caption         =   "&Estatísticas"
      Visible         =   0   'False
      Begin VB.Menu smnEstatistica_estoque_analitico 
         Caption         =   "&Estoque analitíco"
      End
      Begin VB.Menu smnEstatistica_estoque_financeiro 
         Caption         =   "Estoque &financeiro"
      End
      Begin VB.Menu smnEstatistica_produtos_nao_vendidos 
         Caption         =   "&Produtos não vendidos"
      End
      Begin VB.Menu smnEstatistica_curva_abc 
         Caption         =   "&Curvas ABC"
         Begin VB.Menu sisEstatistica_curva_abc_Produto 
            Caption         =   "Produto"
         End
         Begin VB.Menu sisEstatistica_curva_abc_Vendedor 
            Caption         =   "Vendedor"
         End
         Begin VB.Menu sisEstatistica_curva_abc_Cliente 
            Caption         =   "Cliente"
         End
         Begin VB.Menu sisEstatistica_curva_abc_Cidade 
            Caption         =   "Cidade"
         End
         Begin VB.Menu sisEstatistica_curva_abc_Fornecedor 
            Caption         =   "Fornecedor"
         End
         Begin VB.Menu sisEstatistica_curva_abc_Grupo 
            Caption         =   "Grupo"
         End
         Begin VB.Menu sisEstatistica_curva_abc_Filial 
            Caption         =   "Filial"
         End
      End
   End
   Begin VB.Menu mnuRotinas 
      Caption         =   "R&otinas"
      Visible         =   0   'False
      Begin VB.Menu mnuRotina_mensal 
         Caption         =   "Rotina &mensal"
         Begin VB.Menu smnRotina_mensal_movimentos 
            Caption         =   "&Movimentos"
         End
         Begin VB.Menu smnRotina_mensal_apuracao_resultados 
            Caption         =   "&Apuração de resultados"
         End
      End
      Begin VB.Menu mnuRotinas_administrativas 
         Caption         =   "R&otinas administrativas"
         Begin VB.Menu smnRotinas_administrativas_atualiza_tabela_precos 
            Caption         =   "Atualiza &tabela de preços"
         End
         Begin VB.Menu smnRotinas_administrativas_ajuste_custos 
            Caption         =   "Ajuste &custos"
         End
         Begin VB.Menu smnRotinas_administrativas_ajuste_precos 
            Caption         =   "Ajuste &Preços"
         End
         Begin VB.Menu smnRotinas_administrativas_ajuste_estoque 
            Caption         =   "Ajuste &Estoque"
         End
         Begin VB.Menu smnRotinas_administrativas_relacao_descontos_acresimos_cencedidos 
            Caption         =   "&Relação de descontos e acrésimos concedidos"
            Begin VB.Menu sisRotinas_administrativas_relacao_descontos_acresimos_cencedidos_Vendedor 
               Caption         =   "Vendedor"
            End
            Begin VB.Menu sisRotinas_administrativas_relacao_descontos_acresimos_cencedidos_Orcamento 
               Caption         =   "Orçamento"
            End
         End
      End
   End
   Begin VB.Menu mnuFerramentas 
      Caption         =   "Ferramentas"
      Visible         =   0   'False
      Begin VB.Menu smnFerramentas_System_manager 
         Caption         =   "&System Manager"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      Visible         =   0   'False
      Begin VB.Menu smnAjuda 
         Caption         =   "&? Ajuda"
      End
      Begin VB.Menu smnIndice 
         Caption         =   "&Índice"
      End
      Begin VB.Menu smnSobre 
         Caption         =   "&Sobre"
      End
   End
End
Attribute VB_Name = "frmAdminMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim boMousePressionado As Boolean
Dim X As Long
Dim lngPID As Long 'ID do Processo do Aplicativo no Windows
Dim PID As Long
Dim lngIDhwnd As Long 'ID da handle da Janela no Windows
Public log As New DLLSystemManager.log
Dim acesso As New DLLSystemManager.Acessibilidade
Dim rstAplicacao As New ADODB.Recordset
Dim strID As String
Dim booDesign_time As Boolean
Dim strCaminho As String
Public intAtual_ID As Integer
   
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    boMousePressionado = True
End Sub

Private Sub itmAreaDeTrabalhoOrganizar_Click(Index As Integer)
    Me.Arrange Index
End Sub

Private Sub itmBarra_Click(Index As Integer)
    If Index = 0 Then
        picPainelVertical.Visible = Not itmBarra(Index).Checked
        itmBarra(Index).Checked = picPainelVertical.Visible
    End If
End Sub

Private Sub MDIForm_Activate()
    
    'Informações para gravar o LOG
    'Informações Constantes para o log
    log.Usuario = AplicativoUsuário(0).Nome
    log.Programa = "Admin do Sistema"
    log.Estacao = strEstação
    
    'Informações Variaveis para o log
    log.Evento = "Load Admin"
    log.Descricao = "Inicializando o Sistema"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Se true, projeto em Design Time
    booDesign_time = False
    
    'Gravando o log
    log.Gravar_log "Otica", Me
    
End Sub

Private Sub MDIForm_Load()

    Set AT = New AreaDeTrabalho
    
    'Setando e passando a estação local para a mensagem do intercomunicador
    Set FCRegistro = New DLLSystemManager.Registro
    strEstação = FCRegistro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName")
    
    Set SegundoPlano.FormulárioMDI = Me
    SegundoPlano.AutoAtualizar = True
    SegundoPlano.DistânciaDaBorda = 50
    SegundoPlano.Cor(CorEmCima_enPDC) = &HC00000
    SegundoPlano.Cor(CorEmBaixo_enPDC) = vbBlack
    SegundoPlano.EstiloDoFundo = FundoGradiente_enMDIF
    
    SegundoPlano.ArquivoDaImagem = Funcoes_Gerais.Abrir_figura_registro("Otica", Me) & "\fundo1024_800.jpg"  ' App.Path & "\img\Gestor Logo.wmf"
    
    SegundoPlano.PosiçãoDaFigura = FiguraCantoInferiorDireito_enPF ' FiguraNoCentro_enPF ' = SemFigura_enPF ' FiguraAjustada_enPF '  FiguraCantoInferiorDireito_enPF ' FiguraAjustada_enPF ' FiguraCantoInferiorDireito_enPF
    
    mnvMDI.Width = Me.picPainelVertical.Width - 50
    mnvMDI.Height = Me.picPainelVertical.Height - 100
    mnvMDI.Top = 1
    mnvMDI.Refresh
    
    frmAdminDeskTopCliente.Show
    frmAdminDeskTopCliente.Visible = False
    
    'Passando as dimensões do admin para os módulos do sistema
    frmAdminMDI.AplicativoUsuário(intAtual_ID).Height_Admin = Me.Height - 1140
    frmAdminMDI.AplicativoUsuário(intAtual_ID).Width_Admin = Me.Width - 1785
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set SegundoPlano.FormulárioMDI = Nothing
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    Me.mnvMDI.Height = Me.picPainelVertical.Height - 50
    Me.mnvMDI.Height = Me.ScaleHeight ' Me.Height - 500
    SegundoPlano.Atualizar
    Me.Arrange 3
    
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picPainelVertical.Visible = False Then Exit Sub
    'imgSplitter
    If boMousePressionado = True Then
        Me.picPainelVertical.Width = X + Me.picPainelVertical.Width
        Me.mnvMDI.Width = Me.picPainelVertical.Width - 50
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    boMousePressionado = False
    Set SegundoPlano.FormulárioMDI = Me
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    Set SegundoPlano.FormulárioMDI = Nothing
    
    'Informações Variaveis para o log
    log.Evento = "Unload Admin"
    log.Descricao = "Fechando o Sistema"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando o log
    log.Gravar_log "Otica", Me
    'Limpando inf de usuarios do reg. da maquina
    Call Movimentacoes.Limpa_Contingencia_Acessibilidade("Otica")
    
End Sub

Private Sub mnvMDI_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    Dim X As Long
    Dim lngPID As Long 'ID do Processo do Aplicativo no Windows
    Dim PID As Long
    Dim lngIDhwnd As Long 'ID da handle da Janela no Windows
    Dim MSG As String
   
    If frmAdminMDI.AplicativoUsuário(0).NomeReduzido = "" Then Exit Sub
    
    If booDesign_time = True Then
       strCaminho = "P:\Sistemas\CLIENT LOGICX\EXE"
    Else
       strCaminho = App.Path
    End If
    
    '--------------------------------------------------------------------------------------------------
    'IMPORTANTE:
    '
    'Muita atenção com o parâmetro da chamada da classe  - AdicionarAplicativo, caption_form, este tem que ser
    'exatamente o caption do exe a ser chamado;
    '
    '--------------------------------------------------------------------------------------------------
    Dim booAcesso As Boolean
    
    Select Case MenuNumber
        Case 1  'Menu Parâmetros
            Select Case MenuItem
                Case 1 'Cadastros de Base
                    'Acessibilidade
                    mnvMDI.ItemMenuAtual = 1
                    booAcesso = Movimentacoes.Acessibilidade_Menu(mnvMDI.ItemMenuCaption, "Otica", "BDRetaguarda", AplicativoUsuário(0).Codigo)
                    If booAcesso = False Then
                       MsgBox "Você não possui privilégios para acessar este módulo!Verifique com o administrador do sistema.", vbCritical, "Logicx"
                       Exit Sub
                    End If
                    'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
                    lngIDhwnd = AT.AdicionarAplicativo(strCaminho & "\Concentrador - Cadastros Base.exe", _
                                             frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                             lngPID, Me.hwnd, "Concentrador - Cadastros Base")
                                                
                    'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
                    frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Cadastros de  Base", lngIDhwnd, lngPID
                    
                Case 2 'Faturamento
                    'Acessibilidade
                    mnvMDI.ItemMenuAtual = 2
                    booAcesso = Movimentacoes.Acessibilidade_Menu(mnvMDI.ItemMenuCaption, "Otica", "BDRetaguarda", AplicativoUsuário(0).Codigo)
                    If booAcesso = False Then
                       MsgBox "Você não possui privilégios para acessar este módulo!Verifique com o administrador do sistema.", vbCritical, "Logicx"
                       Exit Sub
                    End If
                    'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
                    lngIDhwnd = AT.AdicionarAplicativo(strCaminho & "\Faturamento.exe", _
                                             frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                             lngPID, Me.hwnd, "Faturamento")
                    'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
                    frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Faturamento", lngIDhwnd, lngPID
                Case 3 'Compras
                    'Acessibilidade
                    mnvMDI.ItemMenuAtual = 3
                    booAcesso = Movimentacoes.Acessibilidade_Menu(mnvMDI.ItemMenuCaption, "Otica", "BDRetaguarda", AplicativoUsuário(0).Codigo)
                    If booAcesso = False Then
                       MsgBox "Você não possui privilégios para acessar este módulo!Verifique com o administrador do sistema.", vbCritical, "Logicx"
                       Exit Sub
                    End If
                    'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
                    lngIDhwnd = AT.AdicionarAplicativo(strCaminho & "\Compras.exe", _
                                             frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                             lngPID, Me.hwnd, "Compras")
                    'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
                    frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Compras", lngIDhwnd, lngPID
            End Select
    End Select
    
    Me.mnvMDI.Refresh
    
End Sub

Private Sub sisEstatistica_curva_abc_Grupo_Click()

   'Informações Variaveis para o log
   log.Evento = "Load Exes"
   log.Descricao = "Chamada de Programas - Estatisticas - Curva ABC - Grupos"
   log.Tipo = 1
   log.Data = Date
   log.Hora = Format(Now, "hh:mm:ss")
    
   'Gravando o log
   log.Gravar_log "Otica", Me
   
   'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
   lngIDhwnd = AT.AdicionarAplicativo(App.Path & "\exe\Curva_abc_grupo.exe", _
                                      frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                      lngPID, Me.hwnd, "Curva abc grupo")
   'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
   frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Curva ABC Grupo", lngIDhwnd, lngPID
   
   Me.mnvMDI.Refresh

End Sub

Private Sub smnAdministracao_faixa_comissao_Supervisor_Click()
    
   'Informações Variaveis para o log
   log.Evento = "Load Exes"
   log.Descricao = "Chamada de Programas - Faixa de Comissão do Supervisor"
   log.Tipo = 1
   log.Data = Date
   log.Hora = Format(Now, "hh:mm:ss")
    
   'Gravando o log
   log.Gravar_log "Otica", Me
   
   'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
   lngIDhwnd = AT.AdicionarAplicativo(App.Path & "\exe\Faixa_comissao_supervisor.exe", _
                                      frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                      lngPID, Me.hwnd, "Faixa Comissao do Supervisor")
   'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
   frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Cadastro de Faixa Comissao do Supervisor", lngIDhwnd, lngPID
   
End Sub

Private Sub smnAdministracao_supervisor_Click()
    'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
    lngIDhwnd = AT.AdicionarAplicativo(App.Path & "\exe\Faixa_comissao_supervisor.exe", _
                                       frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                       lngPID, Me.hwnd, "Supervisor")
    'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
    frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Cadastro de Bancos", lngIDhwnd, lngPID
    'Call Identifica_usuario
End Sub

Private Sub smnContas_receber_bancos_Click()

   'Informações Variaveis para o log
   log.Evento = "Load Exes"
   log.Descricao = "Chamada de Programas - Bancos"
   log.Tipo = 1
   log.Data = Date
   log.Hora = Format(Now, "hh:mm:ss")
    
   'Gravando o log
   log.Gravar_log "Otica", Me
   
   'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
   lngIDhwnd = AT.AdicionarAplicativo(App.Path & "\exe\teste1.exe", _
                                      frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                      lngPID, Me.hwnd, "Bancos")
   'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
   frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Cadastro de Bancos", lngIDhwnd, lngPID
   
End Sub

Private Sub smnContas_receber_titulos_Click()

   'Informações Variaveis para o log
   log.Evento = "Load Exes"
   log.Descricao = "Chamada de Programas - Bancos"
   log.Tipo = 1
   log.Data = Date
   log.Hora = Format(Now, "hh:mm:ss")
    
   'Gravando o log
   log.Gravar_log "Otica", Me
   
   'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
   lngIDhwnd = AT.AdicionarAplicativo(App.Path & "\exe\Cadastro_Base.exe", _
                                      frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                      lngPID, Me.hwnd, "Cadastros Base")
   'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
   frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Cadastros de Base", lngIDhwnd, lngPID
   
End Sub

Private Sub ToolbarAreaDeTrabalho_ButtonClick(ByVal Button As MSComctlLib.Button)
      
    If frmAdminMDI.AplicativoUsuário(0).NomeReduzido = "" Then Exit Sub
    
    If booDesign_time = True Then
       strCaminho = "P:\Sistemas\CLIENT LOGICX\EXE"
    Else
       strCaminho = App.Path
    End If
        
    Select Case Button.Key
        Case "GerenciadorDeTarefas"
            frmAdminDesktopPropriedades.Show 1
        Case "Login"
            ExibirLoginOpções
        Case "bloco_notas"
             'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
             lngIDhwnd = AT.AdicionarAplicativo(strCaminho & "\Logix_Bloco_Notas.exe", _
                                      frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                      lngPID, Me.hwnd, "Sem título - Bloco de Notas")
                                         
             'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
             frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Sem título - Bloco de Notas", lngIDhwnd, lngPID
        Case "calculadora"
             'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
             lngIDhwnd = AT.AdicionarAplicativo(strCaminho & "\Logicx_Calculadora.exe", _
                                      frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                      lngPID, Me.hwnd, "EuroCalc v2.1", True)
                                         
             'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
             frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "EuroCalc v2.1", lngIDhwnd, lngPID
    End Select
End Sub

Private Sub ToolbarAreaDeTrabalho_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "AT_DeixarInvisível"
            frmAdminMDI.ActiveForm.Visible = False
            ToolbarAreaDeTrabalho.Buttons(1).ButtonMenus(3).Enabled = True
            'ToolbarAreaDeTrabalho.Buttons(1).ButtonMenus(3).
        Case "AT_Proteger"
            frmAdminMDI.ActiveForm.WindowState = 1
            frmAdminMDI.ActiveForm.Enabled = False
        Case "AT_ProgramasAbertos"
            frmAdminDesktopPropriedades.Show 1
            'frmAdminMDI.ActiveForm.Enabled = False
    End Select
End Sub
