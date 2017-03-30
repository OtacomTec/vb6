VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{20756D14-EE4C-11D6-9F91-000102C349D1}#1.0#0"; "MenuVertical.ocx"
Object = "{06DDD466-EE4A-11D6-9F91-000102C349D1}#1.1#0"; "SegundoPlanoMDI.ocx"
Object = "{D3F9E3A8-F26B-11D6-9F91-000102C349D1}#2.2#0"; "AplicativoUsuário.ocx"
Begin VB.MDIForm frmAdminMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Logicx - Supervisor de Pdv's"
   ClientHeight    =   8190
   ClientLeft      =   1155
   ClientTop       =   2055
   ClientWidth     =   11880
   Icon            =   "frmAdminMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin AplicativoUsuárioOCX.AplicativoUsuário AplicativoUsuário 
      Index           =   0
      Left            =   2910
      Top             =   1050
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin SegundoPlano.SegundoPlanoMDI SegundoPlano 
      Left            =   2400
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox imgSplitter 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7800
      Left            =   1695
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7800
      ScaleWidth      =   30
      TabIndex        =   3
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
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   72
      End
   End
   Begin VB.PictureBox picPainelVertical 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7800
      Left            =   0
      ScaleHeight     =   7800
      ScaleWidth      =   1695
      TabIndex        =   2
      Top             =   390
      Width           =   1695
      Begin GMMenuVertical.MenuVertical mnvMDI 
         Height          =   7755
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "Barra de Atalhos"
         Top             =   0
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   13679
         MenuMax         =   8
         MenuAtual       =   2
         MenuCaption1    =   "Parâmetros"
         ItemMenuMax1    =   10
         ItemMenuÍcone11 =   "frmAdminMDI.frx":1782
         ItemMenuCaption11=   "Acessibilidade"
         ItemMenuÍcone12 =   "frmAdminMDI.frx":1A9C
         ItemMenuCaption12=   "Gerais"
         ItemMenuÍcone13 =   "frmAdminMDI.frx":1DB6
         ItemMenuCaption13=   "PDV"
         ItemMenuÍcone14 =   "frmAdminMDI.frx":20D0
         ItemMenuCaption14=   "TEF"
         ItemMenuÍcone15 =   "frmAdminMDI.frx":23EA
         ItemMenuCaption15=   "Fiscais"
         ItemMenuÍcone16 =   "frmAdminMDI.frx":2704
         ItemMenuCaption16=   "Balanças"
         ItemMenuÍcone17 =   "frmAdminMDI.frx":2A1E
         ItemMenuCaption17=   "Toten"
         ItemMenuÍcone18 =   "frmAdminMDI.frx":2D38
         ItemMenuCaption18=   "Etiquetas"
         ItemMenuÍcone19 =   "frmAdminMDI.frx":3052
         ItemMenuCaption19=   "Bancos"
         ItemMenuÍcone110=   "frmAdminMDI.frx":336C
         ItemMenuCaption110=   "Teclados"
         MenuCaption2    =   "Cadastros"
         ItemMenuMax2    =   10
         ItemMenuÍcone21 =   "frmAdminMDI.frx":3686
         ItemMenuCaption21=   "Mercadológico"
         ItemMenuÍcone22 =   "frmAdminMDI.frx":39A0
         ItemMenuCaption22=   "Produtos"
         ItemMenuÍcone23 =   "frmAdminMDI.frx":3CBA
         ItemMenuCaption23=   "Tributação"
         ItemMenuÍcone24 =   "frmAdminMDI.frx":3FD4
         ItemMenuCaption24=   "Finalizadora"
         ItemMenuÍcone25 =   "frmAdminMDI.frx":42EE
         ItemMenuCaption25=   "Lista Negra"
         ItemMenuÍcone26 =   "frmAdminMDI.frx":4608
         ItemMenuCaption26=   "Lista Branca"
         ItemMenuÍcone27 =   "frmAdminMDI.frx":4922
         ItemMenuCaption27=   "Fornecedores"
         ItemMenuÍcone28 =   "frmAdminMDI.frx":4C3C
         ItemMenuCaption28=   "Clientes"
         ItemMenuÍcone29 =   "frmAdminMDI.frx":4F56
         ItemMenuCaption29=   "Vendedores"
         ItemMenuÍcone210=   "frmAdminMDI.frx":5270
         ItemMenuCaption210=   "Alineas"
         MenuCaption3    =   "Flash Vendas"
         ItemMenuMax3    =   6
         ItemMenuÍcone31 =   "frmAdminMDI.frx":558A
         ItemMenuCaption31=   "Produtos"
         ItemMenuÍcone32 =   "frmAdminMDI.frx":58A4
         ItemMenuCaption32=   "Mercadológico"
         ItemMenuÍcone33 =   "frmAdminMDI.frx":5BBE
         ItemMenuCaption33=   "Faixa Horário"
         ItemMenuÍcone34 =   "frmAdminMDI.frx":5ED8
         ItemMenuCaption34=   "Operador"
         ItemMenuÍcone35 =   "frmAdminMDI.frx":61F2
         ItemMenuCaption35=   "Vendedor"
         ItemMenuÍcone36 =   "frmAdminMDI.frx":650C
         ItemMenuCaption36=   "Finalizadora"
         MenuCaption4    =   "Caixa Geral"
         ItemMenuMax4    =   6
         ItemMenuÍcone41 =   "frmAdminMDI.frx":6826
         ItemMenuCaption41=   "Sangria"
         ItemMenuÍcone42 =   "frmAdminMDI.frx":6A00
         ItemMenuCaption42=   "Mapa Resumo"
         ItemMenuÍcone43 =   "frmAdminMDI.frx":6D1A
         ItemMenuCaption43=   "Fechamento"
         ItemMenuÍcone44 =   "frmAdminMDI.frx":7534
         ItemMenuCaption44=   "Fundo de Caixa"
         ItemMenuÍcone45 =   "frmAdminMDI.frx":784E
         ItemMenuCaption45=   "Comissão"
         ItemMenuÍcone46 =   "frmAdminMDI.frx":7B68
         ItemMenuCaption46=   "Cancelamento"
         MenuCaption5    =   "Comunicação"
         ItemMenuMax5    =   3
         ItemMenuÍcone51 =   "frmAdminMDI.frx":7E82
         ItemMenuCaption51=   "Exportação"
         ItemMenuÍcone52 =   "frmAdminMDI.frx":819C
         ItemMenuCaption52=   "Importação"
         ItemMenuÍcone53 =   "frmAdminMDI.frx":84B6
         ItemMenuCaption53=   "Cargas"
         MenuCaption6    =   "Emissor N.F"
         ItemMenuÍcone61 =   "frmAdminMDI.frx":87D0
         MenuCaption7    =   "Etiquetas"
         ItemMenuÍcone71 =   "frmAdminMDI.frx":8AEA
         MenuCaption8    =   "Painel de Controle"
         ItemMenuMax8    =   3
         ItemMenuÍcone81 =   "frmAdminMDI.frx":8E04
         ItemMenuCaption81=   "System Manager"
         ItemMenuÍcone82 =   "frmAdminMDI.frx":911E
         ItemMenuCaption82=   "Conf. de Sistema"
         ItemMenuÍcone83 =   "frmAdminMDI.frx":9438
         ItemMenuCaption83=   "Monitor Replicação"
      End
      Begin VB.Image ImageDrag 
         Height          =   480
         Left            =   690
         Picture         =   "frmAdminMDI.frx":9752
         Top             =   1170
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList ImageListGeralGrande 
      Left            =   2880
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
            Picture         =   "frmAdminMDI.frx":98A4
            Key             =   "ico_Laboratório"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":9CF8
            Key             =   "ico_Usuário"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":A14C
            Key             =   "ico_Computador"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":C900
            Key             =   "ico_Software"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":CC1C
            Key             =   "ico_SistemaOperacional"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":F3D0
            Key             =   "ico_Programa"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":F6EC
            Key             =   "ico_Departamento"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":FA08
            Key             =   "ico_Drive3.5"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":121BC
            Key             =   "ico_Drive5.25"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":14970
            Key             =   "ico_DriveCD"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":14DC4
            Key             =   "ico_DriveJazz250"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":15C18
            Key             =   "ico_Empresa"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":15F34
            Key             =   "ico_Hardware"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":16250
            Key             =   "ico_HD"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":18A04
            Key             =   "ico_ÁreaDeTrabalho"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":18E58
            Key             =   "ico_Memória"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":192AC
            Key             =   "ico_Modem"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":19700
            Key             =   "ico_Mouse"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":19A1C
            Key             =   "ico_Placa"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":19D38
            Key             =   "ico_PontoDeRede"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1A054
            Key             =   "ico_Processador"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1A4A8
            Key             =   "ico_SalaFechada"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1A7C4
            Key             =   "ico_SalaAberta"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1AAE0
            Key             =   "ico_SCSI"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1AF34
            Key             =   "ico_Servidor"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1B388
            Key             =   "ico_Teclado"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1DC7C
            Key             =   "ico_USB"
            Object.Tag             =   "ico_Trânsito"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1E0D0
            Key             =   "ico_Monitor"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1E9AC
            Key             =   "ico_Sucata"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1ECC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListToolBar 
      Left            =   2310
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1EFE4
            Key             =   "tool_Ícones Grandes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1F0F8
            Key             =   "tool_Ícones Pequenos"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1F20C
            Key             =   "tool_Ícones Lista"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1F320
            Key             =   "tool_Ícones Detalhes"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1F434
            Key             =   "tool_Novo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1F548
            Key             =   "tool_Abrir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1F65C
            Key             =   "tool_Salvar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1F770
            Key             =   "tool_Imprimir"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1F884
            Key             =   "tool_Recortar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1F998
            Key             =   "tool_Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1FAAC
            Key             =   "tool_Colar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1FBC0
            Key             =   "tool_Explorer"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1FEDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":1FFEE
            Key             =   "tool_Explorer2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":20100
            Key             =   "tool_SetaDireita"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":20452
            Key             =   "tool_SetaEsquerda"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":207A4
            Key             =   "tool_AreaDeTrabalhoMask"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":208B6
            Key             =   "tool_AreaDeTrabalho"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":209C8
            Key             =   "tool_AreaDeTrabalho2"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":20F5A
            Key             =   "tool_Usuário"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":210AC
            Key             =   "tool_UsuárioLogin"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":211FE
            Key             =   "tool_GerenciadorDeTarefas"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   688
      BandCount       =   2
      ImageList       =   "ImageListToolBar"
      _CBWidth        =   11880
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Child1          =   "ToolbarPadrão"
      MinHeight1      =   330
      Width1          =   3270
      NewRow1         =   0   'False
      Child2          =   "ToolbarAreaDeTrabalho"
      MinHeight2      =   330
      Width2          =   3000
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar ToolbarAreaDeTrabalho 
         Height          =   330
         Left            =   3465
         TabIndex        =   5
         Top             =   30
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageListToolBar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "GerenciadorDeTarefas"
               Object.ToolTipText     =   "Gerenciador De Tarefas da Área de Trabalho"
               ImageKey        =   "tool_GerenciadorDeTarefas"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Login"
               Object.ToolTipText     =   "Opções de Login"
               ImageKey        =   "tool_Usuário"
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
            Picture         =   "frmAdminMDI.frx":21550
            Key             =   "ico_Laboratório"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":219A4
            Key             =   "ico_Computador"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":24158
            Key             =   "ico_Software"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":24474
            Key             =   "ico_SistemaOperacional"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":26C28
            Key             =   "ico_Programa"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":26F44
            Key             =   "ico_Departamento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":27260
            Key             =   "ico_Drive3.5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":29A14
            Key             =   "ico_Drive5.25"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2C1C8
            Key             =   "ico_DriveCD"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2C61C
            Key             =   "ico_DriveJazz250"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2D468
            Key             =   "ico_Empresa"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2D784
            Key             =   "ico_Hardware"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":2DAA0
            Key             =   "ico_HD"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":30254
            Key             =   "ico_ÁreaDeTrabalho"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":306A8
            Key             =   "ico_Memória"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":30AFC
            Key             =   "ico_Modem"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":30F50
            Key             =   "ico_Mouse"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3126C
            Key             =   "ico_Placa"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":31588
            Key             =   "ico_PontoDeRede"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":318A4
            Key             =   "ico_Processador"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":31CF8
            Key             =   "ico_SalaFechada"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":32014
            Key             =   "ico_SalaAberta"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":32330
            Key             =   "ico_SCSI"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":32784
            Key             =   "ico_Servidor"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":32BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":354CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":35920
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":35D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":36650
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":3696C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminMDI.frx":36C88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListGeralPequeno 
      Left            =   3450
      Top             =   450
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
            Picture         =   "frmAdminMDI.frx":370DC
            Key             =   "ico_Aplicativo"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFerramentas 
      Caption         =   "Ferramentas"
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
   Begin VB.Menu mnuAreaDeTrabalho 
      Caption         =   "Área de Trabalho"
      WindowList      =   -1  'True
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
End
Attribute VB_Name = "frmAdminMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim boMousePressionado As Boolean

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    'Me.tvTreeView.Height = Me.picPainelVertical.Height - 50
    'If Me.Width < 3000 Then Me.Width = 3000
    'Me.mnvMDI.Height = Me.picPainelVertical.Height - 50
    'Me.mnvMDI.Height = Me.ScaleHeight
End Sub

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
    Else
        itmBarra(Index).Checked = Not CoolBar1.Bands(Index).Visible
        CoolBar1.Bands(Index).Visible = itmBarra(Index).Checked
    End If
End Sub

Private Sub MDIForm_Load()
    Set SvMsg = New VetorDeMensagens.ServidorDeMensagens
    Set AT = New AreaDeTrabalho
    Set FCRegistro = New DLLFuncoesGerais.Registro
    strEstação = FCRegistro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName")
    
    Set SegundoPlano.FormulárioMDI = Me
    SegundoPlano.AutoAtualizar = True
    SegundoPlano.DistânciaDaBorda = 50
    SegundoPlano.Cor(CorEmCima_enPDC) = vbWhite
    SegundoPlano.Cor(CorEmBaixo_enPDC) = vbYellow
    SegundoPlano.ArquivoDaImagem = "C:\Projetos\Imagens Logicx\fundo.jpg" ' App.Path & "\img\Gestor Logo.wmf"
    
    SegundoPlano.EstiloDoFundo = FundoGradiente_enMDIF
    SegundoPlano.PosiçãoDaFigura = FiguraCantoInferiorDireito_enPF ' FiguraNoCentro_enPF ' = SemFigura_enPF ' FiguraAjustada_enPF '  FiguraCantoInferiorDireito_enPF ' FiguraAjustada_enPF ' FiguraCantoInferiorDireito_enPF
    
    
    mnvMDI.Width = Me.picPainelVertical.Width - 50
    mnvMDI.Height = Me.picPainelVertical.Height - 100
    mnvMDI.Top = 1
    mnvMDI.Refresh
    frmAdminDeskTopCliente.Show
    frmAdminDeskTopCliente.Visible = False
    
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
    
    'End If
    'imgSplitter
    If boMousePressionado = True Then
        Me.picPainelVertical.Width = X + Me.picPainelVertical.Width
        Me.mnvMDI.Width = Me.picPainelVertical.Width - 50
     '  MDIForm_Resize
    End If

End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    boMousePressionado = False
    'Set SegundoPlano.FormulárioMDI = Nothing
    Set SegundoPlano.FormulárioMDI = Me
    'SegundoPlano.AutoAtualizar = True
    'SegundoPlano.Cor(CorEmCima_enPDC) = vbBlue
    'SegundoPlano.Cor(CorEmBaixo_enPDC) = vbBlack
    'SegundoPlano.ArquivoDaImagem = "c:\windows\floresta.bmp"
    'SegundoPlano.EstiloDoFundo = FundoGradiente_enMDIF
    'SegundoPlano.PosiçãoDaFigura = FiguraNoCentro_enPF
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    Set SegundoPlano.FormulárioMDI = Nothing
End Sub

Private Sub mnvMDI_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    Dim X As Long
    Dim lngPID As Long 'ID do Processo do Aplicativo no Windows
    Dim PID As Long
    Dim lngIDhwnd As Long 'ID da handle da Janela no Windows
    Dim MSG As String
    If frmAdminMDI.AplicativoUsuário(0).NomeReduzido = "" Then Exit Sub
    
    Select Case MenuNumber
        Case 1  'Menu Parâmetros
            Select Case MenuItem
                Case 1 'Menu System Manager
                    'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
                    lngIDhwnd = AT.AdicionarAplicativo(App.Path & "\exe\SystemManager.exe", _
                                                       frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                                       lngPID)
                    'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
                    frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Bloco de Notas", lngIDhwnd, lngPID
                Case 2 ' Cadastros de Base
                    'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
                    lngIDhwnd = AT.AdicionarAplicativo(App.Path & "\exe\Configurador_provedor_dados.exe", _
                                                       frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                                       lngPID)
                    'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
                    frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Configurador de Sistemas ", lngIDhwnd, lngPID
                Case 3 'Mercadologia
                    lngIDhwnd = AT.AdicionarAplicativo(App.Path & "\exe\Mercadologia.exe", frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, lngPID)   ' (Handle))
                    frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Bloco de Notas", lngIDhwnd, lngPID
                    
            End Select
            
        Case 2  'Menu Cadastros
            Select Case MenuItem
                Case 10 'Alineas
                    'Executar um Aplicativo e Adicionar na Área de Trabalho do Usuário Atual (0)
                    lngIDhwnd = AT.AdicionarAplicativo(App.Path & "\exe\Alineas.exe", _
                                                       frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, _
                                                       lngPID)
                    'Adicionar o Aplicativo na lista de Programas Abertos pelo Usuáiro Atual
                    frmAdminMDI.AplicativoUsuário(0).Janela.AdicionarPrograma "Bloco de Notas", lngIDhwnd, lngPID
                
            End Select
    End Select
    MSG = strEstação & "¤" & _
          frmAdminMDI.AplicativoUsuário(0).NomeReduzido & "¤" & _
          frmAdminMDI.AplicativoUsuário(0).Senha & "¤" & _
          frmAdminMDI.AplicativoUsuário(0).Nome & "¤" & _
          frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho & "¤" & _
          frmAdminMDI.AplicativoUsuário(0).PrivilégioAcessar & "¤" & _
          frmAdminMDI.AplicativoUsuário(0).PrivilégioAlterar & "¤" & _
          frmAdminMDI.AplicativoUsuário(0).PrivilégioConsultar & "¤" & _
          frmAdminMDI.AplicativoUsuário(0).PrivilégioExcluir & "¤" & _
          frmAdminMDI.AplicativoUsuário(0).PrivilégioIncluir
    SvMsg.EnviarMensagemID Me.hwnd, MSG, lngIDhwnd

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Computer Manager"
        
    End Select
End Sub



Private Sub ToolbarAreaDeTrabalho_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "GerenciadorDeTarefas"
            frmAdminDesktopPropriedades.Show 1
                    
        Case "Login"
            ExibirLoginOpções
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
