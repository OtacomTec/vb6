VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrincipal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Aplicativo de teste usando a API de comunicação e o driver de spooler"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   8040
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame9 
      Caption         =   "Modelo da Impressora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   27
      Top             =   75
      Width           =   3135
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmPrincipal.frx":030A
         Left            =   240
         List            =   "frmPrincipal.frx":0329
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   300
         Width           =   2655
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Porta de Comunicação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   25
      Top             =   75
      Width           =   2175
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmPrincipal.frx":0393
         Left            =   480
         List            =   "frmPrincipal.frx":03A9
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   19
      Top             =   7440
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      CausesValidation=   0   'False
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Usando a API"
      TabPicture(0)   =   "frmPrincipal.frx":03D0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAcentos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdImprimeTextoSemFormatacao"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdImprimeTextoComFormatacao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdTesteTextoFormatado"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdImprimirCaracterGrafico"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdCortarPapel"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdAguardarImpressaoTexto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdVerificarPapelPresenter"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdCortarParcial"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Usando Código de Barras com a API"
      TabPicture(1)   =   "frmPrincipal.frx":03EC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame18"
      Tab(1).Control(1)=   "frmCodigo"
      Tab(1).Control(2)=   "frmFonte"
      Tab(1).Control(3)=   "frmPosicaoCaracter"
      Tab(1).Control(4)=   "frmLarguraBarras"
      Tab(1).Control(5)=   "frmCodigoBarras"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Usando o Driver de Spooler"
      TabPicture(2)   =   "frmPrincipal.frx":0408
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(1)=   "Frame11"
      Tab(2).Control(2)=   "Frame7"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Impressão de Bitmap"
      TabPicture(3)   =   "frmPrincipal.frx":0424
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "LabelBmpFile"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "ImageBmp"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label14"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label15"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "FileName"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Command1"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame13"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Frame14"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Command2"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Frame15"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "ComboBitola"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "ComboAlgorithm"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
      Begin VB.ComboBox ComboAlgorithm 
         Height          =   315
         ItemData        =   "frmPrincipal.frx":0440
         Left            =   6720
         List            =   "frmPrincipal.frx":044A
         TabIndex        =   102
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox ComboBitola 
         Height          =   315
         ItemData        =   "frmPrincipal.frx":045C
         Left            =   6720
         List            =   "frmPrincipal.frx":046F
         TabIndex        =   98
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame15 
         Caption         =   "Girar"
         Height          =   855
         Left            =   5640
         TabIndex        =   94
         Top             =   3960
         Width           =   1935
         Begin VB.TextBox Degrees 
            Height          =   285
            Left            =   840
            MaxLength       =   3
            TabIndex        =   95
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Graus"
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   6240
         TabIndex        =   92
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Frame Frame14 
         Caption         =   "Redimensionar"
         Height          =   3735
         Left            =   5280
         TabIndex        =   86
         Top             =   1920
         Width           =   2775
         Begin VB.CheckBox AjustaBtn 
            Caption         =   "Ajusta na largura do papel"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox BmpWidth 
            Height          =   285
            Left            =   840
            MaxLength       =   3
            TabIndex        =   90
            Text            =   "100"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox BmpHeight 
            Height          =   285
            Left            =   870
            MaxLength       =   3
            TabIndex        =   89
            Text            =   "100"
            Top             =   390
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   91
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label11 
            Caption         =   "Largura"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Altura"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Orientação do Papel"
         Height          =   1095
         Left            =   120
         TabIndex        =   84
         Top             =   960
         Width           =   3735
         Begin VB.CommandButton Command3 
            Caption         =   "Imprimir"
            Height          =   375
            Left            =   2040
            TabIndex        =   97
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton PaisagemBtn 
            Caption         =   "Paisagem"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton RetratoBtn 
            Caption         =   "Retrato"
            Height          =   375
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   4440
         TabIndex        =   83
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox FileName 
         Height          =   375
         Left            =   1680
         TabIndex        =   81
         Top             =   480
         Width           =   2535
      End
      Begin VB.Frame Frame18 
         Height          =   975
         Left            =   -74760
         TabIndex        =   78
         Top             =   4980
         Width           =   7695
         Begin VB.CommandButton cmdImprimirCodBarras 
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   79
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label lbImprimirCodigo 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Imprimir Código de Barras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   80
            Top             =   0
            Width           =   7710
         End
      End
      Begin VB.Frame frmCodigo 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74760
         TabIndex        =   66
         Top             =   4020
         Width           =   7575
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Left            =   1080
            TabIndex        =   77
            Text            =   "1234567"
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame frmFonte 
         Caption         =   "Fonte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -69240
         TabIndex        =   65
         Top             =   2220
         Width           =   2055
         Begin VB.OptionButton optCondensada 
            Caption         =   "Condensada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton optNormal 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   600
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame frmPosicaoCaracter 
         Caption         =   "Posição dos Caracteres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -72360
         TabIndex        =   64
         Top             =   2220
         Width           =   3015
         Begin VB.OptionButton optNaoImprime 
            Caption         =   "Não Imprime os Caracteres"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   1245
            Width           =   2655
         End
         Begin VB.OptionButton optAbaixo 
            Caption         =   "Abaixo do código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   645
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optAcima 
            Caption         =   "Acima do código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   360
            Width           =   2295
         End
         Begin VB.OptionButton optAcimaAbaixo 
            Caption         =   "Acima e Abaixo do Código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   945
            Width           =   2655
         End
      End
      Begin VB.Frame frmLarguraBarras 
         Caption         =   "Largura das Barras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74760
         TabIndex        =   63
         Top             =   2220
         Width           =   2295
         Begin VB.OptionButton optGrossas 
            Caption         =   "Grossas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   1200
            Width           =   1695
         End
         Begin VB.OptionButton optMedias 
            Caption         =   "Médias (Defaut)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   840
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optFinas 
            Caption         =   "Finas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame frmCodigoBarras 
         Caption         =   "Código de Barras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74760
         TabIndex        =   62
         Top             =   900
         Width           =   7575
         Begin VB.ComboBox cbBarras 
            Height          =   315
            ItemData        =   "frmPrincipal.frx":0497
            Left            =   1680
            List            =   "frmPrincipal.frx":04C2
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   480
            Width           =   3615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Informações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -74880
         TabIndex        =   59
         Top             =   780
         Width           =   7815
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   420
            Width           =   5295
         End
         Begin VB.Label Label7 
            Caption         =   "Impressora:"
            Height          =   255
            Left            =   780
            TabIndex        =   61
            Top             =   450
            Width           =   855
         End
      End
      Begin VB.Frame Frame11 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   53
         Top             =   1740
         Width           =   7815
         Begin VB.CommandButton cmdLigarSensorPoucoPapel 
            Caption         =   "Desligar Sensor de Pouco Papel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4395
            TabIndex        =   57
            Top             =   1140
            Width           =   3015
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4395
            TabIndex        =   56
            Top             =   1620
            Width           =   3015
         End
         Begin VB.CommandButton cmdModificarFonte 
            Caption         =   "Modificar Fonte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4395
            TabIndex        =   55
            Top             =   660
            Width           =   3015
         End
         Begin VB.TextBox Text5 
            Height          =   2175
            Left            =   195
            TabIndex        =   54
            Text            =   "Bematech Soluções"
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Entre com o seu texto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   58
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame12 
         Height          =   975
         Left            =   -74880
         TabIndex        =   50
         Top             =   4860
         Width           =   7695
         Begin VB.CommandButton cmdImprimirFigura 
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   51
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Imprimir Figura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   7710
         End
      End
      Begin VB.CommandButton cmdCortarParcial 
         Caption         =   "Corte Parcial do Papel"
         Height          =   375
         Left            =   -72240
         TabIndex        =   49
         Top             =   5580
         Width           =   2415
      End
      Begin VB.CommandButton cmdVerificarPapelPresenter 
         Caption         =   "Verificar papel no presenter"
         Height          =   375
         Left            =   -69530
         TabIndex        =   47
         Top             =   5100
         Width           =   2415
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tamanho do Extrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -70320
         TabIndex        =   29
         Top             =   3060
         Width           =   3225
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   2280
            TabIndex        =   32
            Text            =   "90"
            Top             =   330
            Width           =   495
         End
         Begin VB.CommandButton cmdProgramarExtrato 
            Caption         =   "Programar"
            Height          =   375
            Left            =   200
            TabIndex        =   31
            Top             =   800
            Width           =   1335
         End
         Begin VB.CommandButton cmdHabilitarExtrato 
            Caption         =   "Habilitar"
            Height          =   375
            Left            =   1680
            TabIndex        =   30
            Top             =   800
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Número de Linhas:"
            Height          =   255
            Left            =   480
            TabIndex        =   33
            Top             =   370
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdAguardarImpressaoTexto 
         Caption         =   "Aguardar Impressão Texto"
         Height          =   375
         Left            =   -69530
         TabIndex        =   24
         Top             =   4620
         Width           =   2415
      End
      Begin VB.CommandButton cmdCortarPapel 
         Caption         =   "Corte Total do Papel"
         Height          =   375
         Left            =   -72240
         TabIndex        =   23
         Top             =   5100
         Width           =   2415
      End
      Begin VB.CommandButton cmdImprimirCaracterGrafico 
         Caption         =   "Imprimir Caracter gráfico"
         Height          =   375
         Left            =   -72240
         TabIndex        =   22
         Top             =   4620
         Width           =   2415
      End
      Begin VB.Frame Frame6 
         Caption         =   "Status da Impressora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74880
         TabIndex        =   18
         Top             =   4500
         Width           =   2295
         Begin VB.CommandButton cmdStatusImpressora 
            Caption         =   "Status da Impressora"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox Text4 
            Height          =   540
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   300
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdTesteTextoFormatado 
         Caption         =   "T&este texto formatado"
         Height          =   375
         Left            =   -69540
         TabIndex        =   17
         Top             =   2460
         Width           =   2430
      End
      Begin VB.CommandButton cmdImprimeTextoComFormatacao 
         Caption         =   "Imprime texto com &formatação"
         Height          =   375
         Left            =   -72220
         TabIndex        =   16
         Top             =   2430
         Width           =   2415
      End
      Begin VB.CommandButton cmdImprimeTextoSemFormatacao 
         Caption         =   "Imprime te&xto sem formatação"
         Height          =   375
         Left            =   -74880
         TabIndex        =   15
         Top             =   2460
         Width           =   2415
      End
      Begin VB.Frame Frame3 
         Caption         =   "Modos de Formatação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -71070
         TabIndex        =   7
         Top             =   1260
         Width           =   3975
         Begin VB.CheckBox Check4 
            Caption         =   "Expandido"
            Height          =   255
            Left            =   2040
            TabIndex        =   14
            Top             =   615
            Width           =   1335
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Itálico"
            Height          =   195
            Left            =   2040
            TabIndex        =   13
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Sublinhado"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   645
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Negrito"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Modos de Impressão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -74880
         TabIndex        =   6
         Top             =   1260
         Width           =   3735
         Begin VB.OptionButton Option5 
            Caption         =   "Condensado"
            Height          =   195
            Left            =   2130
            TabIndex        =   10
            Top             =   510
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Elite"
            Height          =   255
            Left            =   1320
            TabIndex        =   9
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Normal"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdAcentos 
         Caption         =   "Caracteres A&centuados"
         Height          =   375
         Left            =   -69675
         TabIndex        =   5
         Top             =   765
         Width           =   2565
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   4
         Text            =   "Digite o texto aqui."
         Top             =   720
         Width           =   5055
      End
      Begin VB.Frame Frame4 
         Caption         =   "Programação do presenter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   34
         Top             =   3060
         Width           =   4455
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   1680
            TabIndex        =   37
            Text            =   "5"
            Top             =   270
            Width           =   375
         End
         Begin VB.CommandButton cmdProgramarPresenter 
            Caption         =   "&Programar"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   800
            Width           =   2055
         End
         Begin VB.CommandButton cmdHabilitarPresenter 
            Caption         =   "&Habilitar"
            Height          =   375
            Left            =   2280
            TabIndex        =   35
            Top             =   800
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "segundo(s)"
            Height          =   195
            Left            =   2130
            TabIndex        =   39
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tempo de retração:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Autenticação de Documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   40
         Top             =   3060
         Width           =   4455
         Begin VB.TextBox Text7 
            Height          =   315
            Left            =   2310
            MaxLength       =   48
            TabIndex        =   46
            Text            =   "Teste de Autenticação"
            Top             =   330
            Width           =   1935
         End
         Begin VB.CommandButton cmdVerificaDocInserido 
            Caption         =   "Verificar Documento Inserido"
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   800
            Width           =   2295
         End
         Begin VB.CommandButton cmdAutenticacao 
            Caption         =   "Autenticar Documento"
            Height          =   375
            Left            =   2520
            TabIndex        =   42
            Top             =   800
            Width           =   1815
         End
         Begin VB.TextBox Text6 
            Height          =   315
            Left            =   840
            MaxLength       =   2
            TabIndex        =   41
            Text            =   "5"
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "seg."
            Height          =   195
            Left            =   1250
            TabIndex        =   48
            Top             =   350
            Width           =   300
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Time-out:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   370
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Texto:"
            Height          =   195
            Left            =   1830
            TabIndex        =   44
            Top             =   370
            Width           =   450
         End
      End
      Begin VB.Label Label15 
         Caption         =   "Algorithm"
         Height          =   255
         Left            =   5760
         TabIndex        =   101
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Bitola do papel"
         Height          =   375
         Left            =   5520
         TabIndex        =   99
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image ImageBmp 
         BorderStyle     =   1  'Fixed Single
         Height          =   3015
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   4935
      End
      Begin VB.Label LabelBmpFile 
         Caption         =   "Nome do Arquivo"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   600
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   2145
         Left            =   -70320
         Picture         =   "frmPrincipal.frx":0525
         Top             =   2460
         Visible         =   0   'False
         Width           =   2805
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Idioma/Language"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton Option2 
         Caption         =   "English"
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   270
         Width           =   900
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Português"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FlagStatus As Boolean


Private Sub AjustaBtn_Click()
    If AjustaBtn.Value Then
        BmpWidth.Enabled = False
        BmpHeight.Enabled = False
    Else
        BmpWidth.Enabled = True
        BmpHeight.Enabled = True
    End If
    
    
End Sub

Private Sub cbBarras_Click()
    txtCodigo.Text = ""
    Select Case cbBarras.ListIndex
      Case 0: 'EAN8
            txtCodigo.MaxLength = 7
            txtCodigo.Text = "1234567"
    
        Case 1: 'EAN13
            txtCodigo.MaxLength = 12
            txtCodigo.Text = "1234567890123"
        
        Case 2: 'CODE39
            If optFinas.Value Then 'Barras Finas
                txtCodigo.MaxLength = 15
                txtCodigo.Text = "ABC-123"
            ElseIf optMedias.Value Then  ' Barras Medias
                    txtCodigo.MaxLength = 9
                    txtCodigo.Text = "ABC-12345"
                Else ' Barras Grossas
                    txtCodigo.MaxLength = 6
                    txtCodigo.Text = "AB-123"
        End If

        Case 3:  'CODE 93
            If optFinas.Value Then
                txtCodigo.MaxLength = 15
                txtCodigo.Text = "ABC-123"
            ElseIf optMedias.Value Then
                    txtCodigo.MaxLength = 10
                    txtCodigo.Text = "ABC-123"
                Else
                    txtCodigo.MaxLength = 6
                    txtCodigo.Text = "ABC-12"
            End If

        Case 4: ' CODE 128
            If optFinas.Value Then
                txtCodigo.MaxLength = 48
                txtCodigo.Text = "Bematech"
            ElseIf optMedias.Value Then
                    txtCodigo.MaxLength = 28
                    txtCodigo.Text = "Bematech"
                Else
                    txtCodigo.MaxLength = 16
                    txtCodigo.Text = "Bematech"
            End If
            
        Case 5: ' PDF 417
            txtCodigo.Text = "Bematech. Sempre presente nas melhores soluções.!!!"

        Case 6: ' UPCA
            txtCodigo.MaxLength = 11
            txtCodigo.Text = "12345678901"
        
        Case 7: 'UPCE
            txtCodigo.MaxLength = 6
            txtCodigo.Text = "123456"
        
        Case 8: 'ITF
            If optFinas.Value Then
                txtCodigo.MaxLength = 30
                txtCodigo.Text = "0123456789012345"
            ElseIf optMedias.Value Then
                    txtCodigo.MaxLength = 20
                    txtCodigo.Text = "0123456789012345"
                Else
                    txtCodigo.MaxLength = 14
                    txtCodigo.Text = "01234567890125"
            End If

        Case 9: ' MSI
            If optFinas.Value Then
                txtCodigo.MaxLength = 16
                txtCodigo.Text = "123"
            ElseIf optMedias.Value Then
                    txtCodigo.MaxLength = 10
                    txtCodigo.Text = "123"
                Else
                    txtCodigo.MaxLength = 7
                    txtCodigo.Text = "123"
            End If
         
        Case 10: ' ISBN
            txtCodigo.MaxLength = 19
            txtCodigo.Text = "1-56592-292-X 90000"
         
        Case 11: ' PLESSEY
            If optFinas.Value Then
                txtCodigo.MaxLength = 13
                txtCodigo.Text = "123ABC"
            ElseIf optMedias.Value Then
                    txtCodigo.MaxLength = 7
                    txtCodigo.Text = "123ABC"
                Else
                    txtCodigo.MaxLength = 4
                    txtCodigo.Text = "123B"
            End If

        Case 12: ' CODABAR
            If optFinas.Value Then
                txtCodigo.MaxLength = 20
                txtCodigo.Text = "123-456/001"
            ElseIf optMedias.Value Then
                    txtCodigo.MaxLength = 11
                    txtCodigo.Text = "123-456/001"
                Else
                    txtCodigo.MaxLength = 8
                    txtCodigo.Text = "123-4567"
            End If
    End Select
End Sub

Private Sub cmdAcentos_Click()

   Text1.Text = "âäàáãÃÂÄÁÀ êëèéÊËÉÈ ïíìîÎÍÌÏ öóòôõÖÓÒÔÕ üúùûÜÙÚÛ Çç ÿ Ññ"
   
End Sub

Private Sub cmdAguardarImpressaoTexto_Click()

   If Option1.Value Then
      sDados = InputBox("Impressão de Texto com acionamento da guilhotina", "Quantas vezes você deseja imprimir?")
   Else
      sDados = InputBox("Impression of Text with drive of the guillotine", "How many times you desire to print?")
   End If
     
   If sDados <> "" Then
      iretorno = HabilitaEsperaImpressao(1)
      sBuffer = "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
      sBuffer = sBuffer + "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
      sBuffer = sBuffer + "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
      sBuffer = sBuffer + "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
      sBuffer = sBuffer + "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
      sBuffer = sBuffer + Chr(27) + Chr(119)
   
      For iContador = 1 To CInt(sDados)
          iretorno = BematechTX(sBuffer)
          If iretorno = 0 Then
             If Option1.Value Then
                MsgBox "Problemas na impressão do texto." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
             Else
                MsgBox "Problems in the impression of the text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
                Exit Sub
             End If
          End If
   
          iretorno = EsperaImpressao()
   
          ' Mostra a mensagem para retirar o extrato.
          If Option1.Value Then
             If MsgBox("Retire seu extrato.", vbInformation + vbOKCancel) = vbCancel Then
                iretorno = HabilitaEsperaImpressao(0)
                Exit Sub
             End If
          Else
             If MsgBox("Remove the Coupon.", vbInformation + vbOKCancel) = vbCancel Then
                iretorno = HabilitaEsperaImpressao(0)
                Exit Sub
             End If
          End If
      Next
      iretorno = HabilitaEsperaImpressao(0)
   End If

End Sub

Private Sub cmdAutenticacao_Click()
   'autenticação de documentos
   'document authentication
   iretorno = AutenticaDoc(Text7, CInt(Text6) * 1000)
   If iretorno = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na autenticação do documento." + Chr(10) + "Possíveis causas: Documento não inserido, impressora desligada, off-line ou sem papel", vbOKOnly + vbCritical, "Autenticação de Documentos"
      Else
         MsgBox "Problems in the authentication of the document." + Chr(10) + "Possible causes: Not inserted document, off printer, off-line or without paper", vbOKOnly + vbCritical, "Document Authentication"
      End If
   
   ElseIf iretorno = -1 Then
      If Option1.Value Then
         MsgBox "Tempo maior que o permitido.", vbOKOnly + vbCritical, "Autenticação de Documentos"
      Else
         MsgBox "Bigger time that the allowed.", vbOKOnly + vbCritical, "Document Authentication"
      End If
   End If
End Sub

Private Sub cmdCortarPapel_Click()

   ' Acionamento da guilhotina (cortar o papel)
   ' Drive of the guillotine (to cut the paper)
   iretorno = AcionaGuilhotina(1)
   If iretorno <> 1 Then
      If Option1.Value Then
         MsgBox "Problemas no corte do papel." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
      Else
         MsgBox "Problems in the cut of the paper." + Chr(10) + "Possible causes: Off printer, off-line or without paper", vbInformation + vbOKOnly
      End If
   End If

End Sub

Private Sub cmdDesligarSensorPoucoPapel_Click()

'   ' Sequencia de comando para desligar o sensor de pouco papel
'   sBuffer = Chr(27) + Chr(98) + Chr(1)
'   iretorno = ComandoTX(sBuffer, 3)
'   If iretorno = 0 Then
'      If Option1.Value Then
'         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
'      Else
'         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes: Off printer, off-line or without paper", vbInformation + vbOKOnly
'         exit sub
'      End If
'    End If

End Sub

Private Sub cmdCortarParcial_Click()

   ' Acionamento da guilhotina (corte Parcial o papel)
   ' Drive of the guillotine (Partial cut the paper)
   iretorno = AcionaGuilhotina(0)
   If iretorno <> 1 Then
      If Option1.Value Then
         MsgBox "Problemas no corte do papel." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
      Else
         MsgBox "Problems in the cut of the paper." + Chr(10) + "Possible causes: Off printer, off-line or without paper", vbInformation + vbOKOnly
      End If
   End If

End Sub

Private Sub cmdHabilitarExtrato_Click()

   If cmdHabilitarExtrato.Caption = "Habilitar" Or (cmdHabilitarExtrato.Caption = "Enable") Then
      
      ' Habilita o extrato longo
      iretorno = HabilitaExtratoLongo(1)
      If iretorno = 0 Then
         If Option1.Value Then
            MsgBox "Problemas na habilitação do extrato longo." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Else
           MsgBox "Problems in the qualification of the long extract." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         End If
      Else
         If Option1.Value Then
            cmdHabilitarExtrato.Caption = "Desabilitar"
         Else
            cmdHabilitarExtrato.Caption = "Disable"
         End If
      End If
   Else
         
      ' Desabilita o extrato longo
      iretorno = HabilitaExtratoLongo(0)
      If iretorno = 0 Then
         If Option1.Value Then
            MsgBox "Problemas na desabilitação do extrato longo." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Else
            MsgBox "Problems in the disable of the long extract." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         End If
      Else
         If Option1.Value Then
            cmdHabilitarExtrato.Caption = "Habilitar"
         Else
            cmdHabilitarExtrato.Caption = "Enable"
         End If
      End If
   End If
End Sub

Private Sub cmdHabilitarPresenter_Click()

   If cmdHabilitarPresenter.Caption = "&Habilitar" Or cmdHabilitarPresenter.Caption = "Enable" Then
         
      ' Habilita o presenter retrátil
      iretorno = HabilitaPresenterRetratil(1)
      If iretorno = 0 Then
         If Idioma.Enabled Then
            MsgBox "Problemas na programação do presenter." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Else
            MsgBox "Problems in the programming of presenter." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         End If
      Else
         If Option1.Value Then
            cmdHabilitarPresenter.Caption = "D&esabilitar"
         Else
            cmdHabilitarPresenter.Caption = "Disable"
         End If
      End If
   Else

      ' Desabilita o presenter retrátil
      iretorno = HabilitaPresenterRetratil(0)
      If iretorno = 0 Then
         If Option1.Value Then
            MsgBox "Problemas na programação do presenter." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Else
            MsgBox "Problems in the programming of presenter." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         End If
      Else
         If Option1.Value Then
            cmdHabilitarPresenter.Caption = "&Habilitar"
         Else
            cmdHabilitarPresenter.Caption = "Enable"
         End If
      End If
   End If

End Sub

Private Sub cmdImprimeTextoComFormatacao_Click()

   iNegrito = 0
   iItalico = 0
   iSublinhado = 0
   iExpandido = 0
   If Option1.Value Then
      sDados = InputBox("Impressão de Texto", "Quantas vezes você deseja imprimir ?")
   Else
      sDados = InputBox("Impression of Text", "How many times you desire to print ?")
   End If

   If sDados <> "" Then
      
      ' Verifica modo NORMAL, ELITE ou CONDENSADO.
      
      If Option3.Value = True Then
         iModo = 2
      ElseIf Option4.Value = True Then
         iModo = 3
      Else
         iModo = 1
      End If

      ' Negrito, Itálico, Sublinhado e Expandido
    
      If Check1.Value = 1 Then
         iNegrito = 1
      End If
      If Check2.Value = 1 Then
         iItalico = 1
      End If
      If Check3.Value = 1 Then
         iSublinhado = 1
      End If
      If Check4.Value = 1 Then
         iExpandido = 1
      End If

      For iContador = 1 To CInt(sDados)
          sBuffer = Text1.Text + Chr(13) + Chr(10)
          iretorno = FormataTX(sBuffer, iModo, iItalico, iSublinhado, iExpandido, iNegrito)
        If iretorno = 0 Then
            If Option1.Value Then
                MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
            Else
                MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
            End If
        End If
      Next
   End If

End Sub

Private Sub cmdImprimeTextoSemFormatacao_Click()

   If Option1.Value Then
      sDados = InputBox("Impressão de Texto", "Quantas vezes você deseja imprimir ?")
   Else
      sDados = InputBox("Text Printing", "How many times you desire to print ?")
   End If
    
   If IsNumeric(sDados) Then
    If sDados <> "" Then
       For iContador = 1 To CInt(sDados)
           sBuffer = Text1.Text + Chr(13) + Chr(10)
           iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
           If iretorno = 0 Then
              If Option1.Value Then
                 MsgBox "Problemas na impressão do texto." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
              Else
                 MsgBox "Problems in the impression of the text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
              End If
           End If
       Next
     End If
    Else
        If Option1.Value Then
            MsgBox "Entrada deve ser numérica!!"
         Else
            MsgBox "Only numbers allowed!!"
        End If
    End If


End Sub

Private Sub cmdImprimir_Click()

   Printer.FontBold = CommonDialog1.FontBold
   Printer.FontItalic = CommonDialog1.FontItalic
   Printer.FontName = CommonDialog1.FontName
   Printer.FontSize = CommonDialog1.FontSize
   
   Printer.Print (Text5.Text)
   Printer.EndDoc

End Sub

Private Sub cmdImprimirCaracterGrafico_Click()

   ' DESENHO
   
   '              1 2 3 4 5 6 7 8 9
   ' bit 7 = 128  *               *
   ' bit 6 = 064  * *             *
   ' bit 5 = 032  * * *           *
   ' bit 4 = 016  * * * *         *
   ' bit 3 = 008  * * * * *       *
   ' bit 2 = 004  * * * * * *     *
   ' bit 1 = 002  * * * * * * *   *
   ' bit 0 = 001  * * * * * * * * *
   
   
   ' Comando que habilita o modo grafico com 9 pinos (9 colunas)
   sBuffer = Chr(27) + Chr(94) + Chr(18) + Chr(0)
   iretorno = ComandoTX(sBuffer, 4)
   If iretorno = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do caracter gráfico." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of graphical caracter." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If
   
   ' Sequencia de bytes para a montagem do desenho acima
   sBuffer = Chr(255) + Chr(0) + Chr(0) + Chr(0) + Chr(127) + Chr(0) _
          + Chr(0) + Chr(0) + Chr(63) + Chr(0) + Chr(0) + Chr(0) _
          + Chr(31) + Chr(0) + Chr(0) + Chr(0) + Chr(15) + Chr(0) + Chr(0) _
          + Chr(0) + Chr(7) + Chr(0) + Chr(0) + Chr(0) + Chr(3) + Chr(0) _
          + Chr(0) + Chr(0) + Chr(1) + Chr(0) + Chr(0) + Chr(0) + Chr(255) _
          + Chr(0) + Chr(0) + Chr(0)
      
   ' Descarrega o buffer na impressora.
   sBuffer = sBuffer + Chr(13) + Chr(10)
      
   iretorno = CaracterGrafico(sBuffer, Len(sBuffer))
   
  
'   ' Comando que habilita o modo grafico com 9 pinos (9 colunas)
'   sBuffer = Chr(27) + Chr(94) + Chr(18) + Chr(0)
'   iRetorno = ComandoTX(sBuffer, Len(sBuffer))
'   If iRetorno = 0 Then
'      If Option1.Value Then
'         MsgBox "Problemas na impressão do caracter gráfico." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
'      Else
'         MsgBox "Problems in the impression of graphical caracter." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
'         End
'      End If
'   End If
'
'   ' Sequencia de bytes para a montagem do desenho acima
'   sBuffer = Chr(255) + Chr(0) + Chr(0) + Chr(0) + Chr(127) + Chr(0) _
'          + Chr(0) + Chr(0) + Chr(63) + Chr(0) + Chr(0) + Chr(0) _
'          + Chr(31) + Chr(0) + Chr(0) + Chr(0) + Chr(15) + Chr(0) + Chr(0) _
'          + Chr(0) + Chr(7) + Chr(0) + Chr(0) + Chr(0) + Chr(3) + Chr(0) _
'          + Chr(0) + Chr(0) + Chr(1) + Chr(0) + Chr(0) + Chr(0) + Chr(255) _
'          + Chr(0) + Chr(0) + Chr(0)
'   iRetorno = ComandoTX(sBuffer, Len(sBuffer))
'   If iRetorno = 0 Then
'      If Option1.Value Then
'         MsgBox "Problemas na impressão do caracter gráfico." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
'      Else
'         MsgBox "Problems in the impression of graphical caracter." + Chr(10) + "Possible causes: Off printer, off-line or without paper", vbInformation + vbOKOnly
'         End
'      End If
'   End If

   ' Descarrega o buffer na impressora.
'   sBuffer = Chr(13) + Chr(10)
'   iRetorno = ComandoTX(sBuffer, Len(sBuffer))
'   If iRetorno = 0 Then
'      If Option1.Value Then
'         MsgBox "Problemas na impressão do caracter gráfico." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
'      Else
'         MsgBox "Problems in the impression of graphical caracter." + Chr(10) + "Possible causes: Off printer, off-line or without paper", vbInformation + vbOKOnly
'         End
'      End If
'   End If

End Sub

Private Sub cmdImprimirCodBarras_Click()
Dim Largura, PosixaoCaracter, Fonte

    '==== LARGURA DA LINHA ====
    If optFinas.Value Then
        Largura = 0
    ElseIf optMedias.Value Then
            Largura = 1
        Else
            Largura = 2
    End If
  
    If cbBarras.ListIndex = 1 Then
        Largura = 1
    End If
 

    '==== POSIÇÃO DO CARACTER ====}
    If optAcima.Value Then
        PosicaoCaracter = 1
    ElseIf optAbaixo.Value Then
            PosicaoCaracter = 2
        ElseIf optAcimaAbaixo.Value Then
                PosicaoCaracter = 3
            Else
                PosicaoCaracter = 0
    End If

    '==== FONTE ==
    If optNormal.Value Then
        Fonte = 0
    Else
        Fonte = 1
    End If
  
    Retorno = ConfiguraCodigoBarras(162, Largura, PosicaoCaracter, Fonte, 0)
  
    Select Case cbBarras.ListIndex
        Case 0: Retorno = ImprimeCodigoBarrasEAN8(txtCodigo.Text)
        Case 1: Retorno = ImprimeCodigoBarrasEAN13(txtCodigo.Text)
        Case 2: Retorno = ImprimeCodigoBarrasCODE39(txtCodigo.Text)
        Case 3: Retorno = ImprimeCodigoBarrasCODE93(txtCodigo.Text)
        Case 4: Retorno = ImprimeCodigoBarrasCODE128(txtCodigo.Text)
        Case 5: Retorno = ImprimeCodigoBarrasPDF417(4, 3, 2, 1, txtCodigo.Text)
        Case 6: Retorno = ImprimeCodigoBarrasUPCA(txtCodigo.Text)
        Case 7: Retorno = ImprimeCodigoBarrasUPCE(txtCodigo.Text)
        Case 8: Retorno = ImprimeCodigoBarrasITF(txtCodigo.Text)
        Case 9: Retorno = ImprimeCodigoBarrasMSI(txtCodigo.Text)
        Case 10: Retorno = ImprimeCodigoBarrasISBN(txtCodigo.Text)
        Case 11: Retorno = ImprimeCodigoBarrasPLESSEY(txtCodigo.Text)
        Case 12: Retorno = ImprimeCodigoBarrasCODABAR(txtCodigo.Text)
    End Select
  
End Sub

Private Sub cmdImprimirFigura_Click()

   Printer.PaintPicture Image1.Picture, 1500, 1500, 1500, 1500
   Printer.EndDoc
   
End Sub

Private Sub cmdLigarSensorPoucoPapel_Click()
    Dim sBuffer As String
    
    If cmdLigarSensorPoucoPapel.Caption = "Ligar Sensor de Pouco Papel" Or _
       cmdLigarSensorPoucoPapel.Caption = "Enable Low Paper Sensor" Then
        ' Seqüência de comando para ligar o sensor de pouco papel
        sBuffer = Chr(27) + Chr(98) + Chr(0)
        iretorno = ComandoTX(sBuffer, 3)
        If iretorno = 0 Then
           If Option1.Value Then
              MsgBox "Problemas ao ligar o sensor de pouco papel." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
              Exit Sub
           Else
              MsgBox "Problems when binding the sensor of low paper." + Chr(10) + "Possible causes: Off printer, off-line or without paper", vbInformation + vbOKOnly
              Exit Sub
           End If
        End If
    
        If Option1.Value Then
            cmdLigarSensorPoucoPapel.Caption = "Desligar Sensor de Pouco Papel"
        Else
            cmdLigarSensorPoucoPapel.Caption = "Disable Low Paper Sensor"
        End If
    
    Else
        ' Sequencia de comando para desligar o sensor de pouco papel
        sBuffer = Chr(27) + Chr(98) + Chr(1)
        iretorno = ComandoTX(sBuffer, 3)
        If iretorno = 0 Then
           If Option1.Value Then
              MsgBox "Problemas ao desligar o sensor de pouco papel." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
              Exit Sub
           Else
              MsgBox "Problems when the disconnect the sensor of low paper." + Chr(10) + "Possible causes: Off printer, off-line or without paper", vbInformation + vbOKOnly
              Exit Sub
           End If
         End If
        
        'altera o caption do botão
        If Option1.Value Then
            cmdLigarSensorPoucoPapel.Caption = "Ligar Sensor de Pouco Papel"
        Else
            cmdLigarSensorPoucoPapel.Caption = "Enable Low Paper Sensor"
        End If
    
    End If


End Sub

Private Sub cmdModificarFonte_Click()

   CommonDialog1.Flags = cdlCFNoFaceSel Or cdlCFNoSizeSel Or cdlCFNoStyleSel
   CommonDialog1.Flags = 1
   CommonDialog1.ShowFont
   Text5.FontBold = CommonDialog1.FontBold
   Text5.FontItalic = CommonDialog1.FontItalic
   Text5.FontName = CommonDialog1.FontName
   Text5.FontSize = CommonDialog1.FontSize

End Sub

Private Sub cmdProgramarExtrato_Click()

   ' programa o tamanho do extrato longo
   iretorno = ConfiguraTamanhoExtrato(CInt(Text3.Text))
   If iretorno = 0 Then
      If Option1.Value Then
          MsgBox "Problemas na programação do tamanho do extrato." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
        Else
          MsgBox "Problems in the programming of the size of the extract." + Chr(10) + "Cases Possibles: Printer off, off-line or paper out", vbInformation + vbOKOnly
      End If
   End If
End Sub

Private Sub cmdProgramarPresenter_Click()

  ' Programa o tempo de espera para retração do
  ' papel caso o mesmo não seja retirado do bocal
  ' do presenter.
  
  iretorno = ProgramaPresenterRetratil(CInt(Text2.Text))
  If iretorno = 0 Then
     If Option1.Value Then
        MsgBox "Problemas na programação do presenter." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
     Else
        MsgBox "Problems in the programming of presenter." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
     End If
  End If

End Sub

Private Sub cmdSair_Click()
iPorta = FechaPorta()
   End
End Sub

Private Sub cmdStatusImpressora_Click()

    If cmdStatusImpressora.Caption = "Status da Impressora" Or _
       cmdStatusImpressora.Caption = "Printer Status" Then
        FlagStatus = True
    Else
        FlagStatus = False
    End If
    
   If Option1.Value Then
      cmdStatusImpressora.Caption = "Cancelar"
   Else
      cmdStatusImpressora.Caption = "Cancel"
   End If

   iRetStatus = -2 'inicializa a variavel com um valor qq.
   While FlagStatus
       iretorno = Le_Status()
       DoEvents
       If iRetStatus <> iretorno Then
          If Option1.Value Then
             Select Case iretorno
                    Case 0:
                        If (Combo1.ListIndex = 0) Then  'LPT
                            Text4.Text = "Desligada ou cabo desconectado"
                        Else 'COM
                            Text4.Text = "Off-line"
                        End If
                        
                    Case 32:
                        If (Combo1.ListIndex = 0) Then  'LPT
                            Text4.Text = "Pouco papel e off-line"
                        Else 'COM
                            Text4.Text = "Fim de papel"
                        End If
                    
                    Case 40: Text4.Text = "Fim de papel"
                    Case 4: Text4.Text = "Pouco papel e off-line" 'na COM
                    Case 5, 48: Text4.Text = "Pouco papel e on-line" 'na LPT
                    Case 79: Text4.Text = "Off-line" 'na LPT
                    Case 9, 128: Text4.Text = "Head Up" 'na LPT
                    Case 24, 144: Text4.Text = "Impressora ok" 'on-line na LPT
                    Case 66: Text4.Text = "Temperatura do cabeçote"
                    Case 65: Text4.Text = "Erro no corte do papel"
                    Case 67: Text4.Text = "Papel enroscado"
                    Case 68: Text4.Text = "Impressora Desligada"
                    
                    
                   
                   'Se não for nenhum dos status acima
             Case Else
                  Text4.Text = "Status desconhecido"
             End Select
         Else
           Select Case iretorno
                  Case 0:
                        If (Combo1.ListIndex = 0) Then  'LPT
                            Text4.Text = "Off or detached handle"
                        Else 'COM
                            Text4.Text = "Off-line"
                        End If
                  
                  Case 32:
                        If (Combo1.ListIndex = 0) Then  'LPT
                            Text4.Text = "Low paper and off-line"
                        Else 'COM
                            Text4.Text = "End of paper"
                        End If
                  Case 40: Text4.Text = "End of paper"
                  Case 4: Text4.Text = "Low paper and off-line" 'na COM
                  Case 5, 48: Text4.Text = "Low paper and on-line"
                  Case 79: Text4.Text = "Off-line"
                  Case 9, 128: Text4.Text = "Head Up"
                  Case 24, 144: Text4.Text = "Printer ok" '24 on-line na COM e 144 na LPT
                  
                  ' Se não for nenhum dos status acima
           Case Else
                Text4.Text = "Unknown status"
           End Select
         End If
         iRetStatus = iretorno
         'Text4.Refresh
       End If
       
       DoEvents
   Wend
  
   If Option1.Value Then
      cmdStatusImpressora.Caption = "Status da Impressora"
   Else
      cmdStatusImpressora.Caption = "Printer Status"
   End If
   
   Text4.Text = ""
End Sub

Private Sub cmdTesteTextoFormatado_Click()

   ' Acentos a serem impressos
   sTexto = sTexto + "âäàáãÃÂÄÁÀ êëèéÊËÉÈ ïíìîÎÍÌÏ öóòôõÖÓÒÔÕ üúùûÜÙÚÛ" + Chr(13) + Chr(10) + Chr(10)

   ' Italico
   sFonte = "Itálico" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 1, 0, 0, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Sublinhado
   sFonte = "Sublinhado" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 0, 1, 0, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Expandido
   sFonte = "Expandido" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 0, 0, 1, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Negrito
   sFonte = "Negrito" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 0, 0, 0, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Condensado
   sFonte = "Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 0, 0, 0, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Sublinhado
   sFonte = "Itálico + Sublinhado" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 1, 1, 0, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Expandido
   sFonte = "Itálico + Expandido" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 1, 0, 1, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Negrito
   sFonte = "Itálico + Negrito" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 1, 0, 0, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Condensado
   sFonte = "Itálico + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 1, 0, 0, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Sublinhado + Expandido
   sFonte = "Itálico + Sublinhado + Expandido" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 1, 1, 1, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Sublinhado + Negrito
   sFonte = "Itálico + Sublinhado + Negrito" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 1, 1, 0, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Sublinhado + Condensado
   sFonte = "Itálico + Sublinhado + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 1, 1, 0, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Sublinhado + Expandido + Negrito
   sFonte = "Itálico + Sublinhado + Expandido + Negrito" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 1, 1, 1, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Sublinhado + Expandido + Condensado
   sFonte = "Itálico + Sublinhado + Expandido + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 1, 1, 1, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Itálico + Sublinhado + Expandido + Negrito + Condensado
   sFonte = "Itálico + Sublinhado + Epandido + Negrito + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 1, 1, 1, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Sublinhado + Expandido
   sFonte = "Sublinhado + Expandido" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 0, 1, 1, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Sublinhado + Negrito
   sFonte = "Sublinhado + Negrito" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 0, 1, 0, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Sublinhado + Condensado
   sFonte = "Sublinhado + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 0, 1, 0, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Sublinhado + Expandido + Negrito
   sFonte = "Sublinhado + Expandido + Negrito" + Chr(10)
   If FormataTX(sFonte + sTexto, 3, 0, 1, 1, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Sublinhado + Expandido + Condensado
   sFonte = "Sublinhado + Expandido + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 0, 1, 1, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Sublinhado + Expandido + Negrito + Consensado
   sFonte = "Sublinhado + Expandido + Negrito + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 0, 1, 1, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

  ' Expandido + Negrito
  sFonte = "Expandido + Negrito" + Chr(10)
  If FormataTX(sFonte + sTexto, 3, 0, 0, 1, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Expandido + Condensado
   sFonte = "Expandido + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 0, 0, 1, 0) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Expandido + Negrito + Condensado
   sFonte = "Expandido + Negrito + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 0, 0, 1, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

   ' Negrito + Condensado
   sFonte = "Negrito + Condensado" + Chr(10)
   If FormataTX(sFonte + sTexto, 1, 0, 0, 0, 1) = 0 Then
      If Option1.Value Then
         MsgBox "Problemas na impressão do texto formatado." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbInformation + vbOKOnly
         Exit Sub
      Else
         MsgBox "Problems in the impression of the formatted text." + Chr(10) + "Possible causes:  Off printer, off-line or without paper", vbInformation + vbOKOnly
         Exit Sub
      End If
   End If

End Sub

Private Sub cmdVerificaDocInserido_Click()
    iretorno = DocumentInserted()
    
    If iretorno = 1 Then 'documento inserido
        If Option1.Value = True Then 'idioma português
            MsgBox "Documento inserido.", vbOKOnly + vbInformation, "Verificação de Documento Inserido"
        Else 'idioma inglês
            MsgBox "Inserted document.", vbOKOnly + vbInformation, "Inserted Document verification"
        End If
    
    Else 'documento não inserido
        If Option1.Value = True Then 'idioma português
            MsgBox "Documento não inserido.", vbOKOnly + vbInformation, "Verificação de Documento Inserido"
        Else 'idioma inglês
            MsgBox "Not inserted document.", vbOKOnly + vbInformation, "Inserted Document verification"
        End If
    End If
End Sub

Private Sub cmdVerificarPapelPresenter_Click()
   ' Verifica a presença de papel no presenter
   
    iretorno = VerificaPapelPresenter()
    Select Case iretorno
        Case -1: 'erro de execução da função
            If Option1.Value Then
                MsgBox "Erro de execução da função.", vbCritical + vbOKOnly
            Else
                MsgBox "Error of execution of the function.", vbCritical + vbOKOnly
            End If
        
        Case 0: 'problemas na verificação do papel no presenter
            If Option1.Value Then
                MsgBox "Problemas na verificação do papel no presenter." + Chr(10) + "Possíveis causas: Impressora desligada, off-line ou sem papel", vbCritical + vbOKOnly
            Else
                MsgBox "problems in the verification of the paper in presenter." + Chr(10) + "Possible causes: Off printer, off-line or without paper", vbCritical + vbOKOnly
            End If
        
        Case 1: 'papel posicionado no presenter
            If Option1.Value Then
                MsgBox "Papel posicionado no presenter.", vbInformation + vbOKOnly
            Else
                MsgBox "Paper located in presenter", vbInformation + vbOKOnly
            End If
        
        Case 2: 'papel não posicionado no presenter
            If Option1.Value Then
                MsgBox "Papel não posicionado no presenter.", vbInformation + vbOKOnly
            Else
                MsgBox "Paper not located in presenter.", vbInformation + vbOKOnly
            End If
        
        Case 3: 'erro desconhecido
            If Option1.Value Then
                MsgBox "Erro desconhecido.", vbInformation + vbOKOnly
            Else
                MsgBox "Unknown error.", vbInformation + vbOKOnly
            End If
    End Select
     
End Sub

Private Sub Combo1_Click()

   ' Fecha a porta que está aberta
  
   iPorta = FechaPorta()
   If iPorta <= 0 Then
      If Option1.Value Then
         MsgBox "Problemas ao Fechar a Porta de Comunicação.", vbInformation + vbOKOnly
      Else
         MsgBox "Problems when closing the communication port.", vbInformation + vbOKOnly
      End If
   End If

   If Combo1.Text = "LPT1" Then
      sComunica = "LPT1"
   End If
   If Combo1.Text = "COM1" Then
      sComunica = "COM1"
   End If
   If Combo1.Text = "COM2" Then
      sComunica = "COM2"
   End If
   If Combo1.Text = "COM3" Then
      sComunica = "COM3"
   End If
   If Combo1.Text = "COM4" Then
      sComunica = "COM4"
   End If
   If Combo1.Text = "USB" Then
      sComunica = "USB"
   End If
   

   ' Abre a porta de comunicacao
  
   iPorta = IniciaPorta(sComunica)
   If iPorta <= 0 Then
      If Option1.Value Then
         MsgBox "Problemas ao Abrir a Porta de Comunicação.", vbInformation + vbOKOnly
      Else
         MsgBox "Problems when opening the communication port.", vbInformation + vbOKOnly
      End If
   End If
   
End Sub

Private Sub Combo2_Click()

If Combo2.Text = "MP-20 TH" Or Combo2.Text = "MP-2000 TH" Or Combo2.Text = "MP-2000 CI" Or Combo2.Text = "MP-2100 TH" Then
      iModeloImpressora = 0
   End If
   If Combo2.Text = "MP-20 MI" Then
      iModeloImpressora = 1
   End If
   If Combo2.Text = "MP-4000 TH" Then
      iModeloImpressora = 5
   End If
   If Combo2.Text = "MP-4200 TH" Then
      iModeloImpressora = 7
   End If
   If Combo2.Text = "MP-2500 TH" Then
      iModeloImpressora = 8
   End If


'   If Combo2.Text = "MP-20 CI" Or Combo2.Text = "MP-20 MI" Or Combo2.Text = "MP-2000 CI" Then
    If iModeloImpressora = 0 Or iModeloImpressora = 1 Or iModeloImpressora = 3 Then
      Frame10.Visible = True  'autenticacao
      Frame10.ZOrder 0
      HabilitaProgramacaoPresenter (False)
      HabilitaFuncoesExtrato (False)
      cmdCortarPapel.Enabled = False
      cmdCortarParcial.Enabled = False
      cmdVerificarPapelPresenter.Enabled = False
      cmdAguardarImpressaoTexto.Enabled = False
      frmCodigoBarras.Enabled = False
      frmCodigo.Enabled = False
      cmdImprimirCodBarras.Enabled = False
      frmLarguraBarras.Enabled = False
      frmPosicaoCaracter.Enabled = False
      frmFonte.Enabled = False
      Command3.Enabled = False
      Command2.Enabled = False
      cmdImprimirFigura.Enabled = False
      cmdImprimir.Enabled = False
      
     
   'ElseIf Combo2.Text = "BLOCO TÉRMICO" Then
   ElseIf iModeloImpressora = 5 Or iModeloImpressora = 6 Or iModeloImpressora = 7 Or iModeloImpressora = 8 Then
      Command3.Enabled = True
      Command2.Enabled = True
      cmdImprimirFigura.Enabled = True
      cmdImprimir.Enabled = True
      Frame10.Visible = False 'autenticacao
      Frame10.ZOrder 1
      HabilitaProgramacaoPresenter (True)
      HabilitaFuncoesExtrato (True)
      cmdCortarPapel.Enabled = True
      cmdCortarParcial.Enabled = True
      cmdVerificarPapelPresenter.Enabled = True
      cmdAguardarImpressaoTexto.Enabled = True
      frmCodigoBarras.Enabled = True
      frmCodigo.Enabled = True
      cmdImprimirCodBarras.Enabled = True
      frmLarguraBarras.Enabled = True
      frmPosicaoCaracter.Enabled = True
      frmFonte.Enabled = True
   Else
      Command3.Enabled = True
      Command2.Enabled = True
      cmdImprimirFigura.Enabled = True
      cmdImprimir.Enabled = True
   
      If Combo2.ListIndex = 2 Then
      cmdCortarPapel.Enabled = False
      cmdCortarParcial.Enabled = False
      cmdImprimirCodBarras.Enabled = False
      End If
      
   
      Frame10.Visible = False 'autenticacao
      Frame10.ZOrder 1
      HabilitaProgramacaoPresenter (False)
      HabilitaFuncoesExtrato (False)
      cmdCortarPapel.Enabled = True
      cmdCortarParcial.Enabled = True
      cmdVerificarPapelPresenter.Enabled = False
      cmdAguardarImpressaoTexto.Enabled = False
      If Combo2.ListIndex = 3 Then
        frmCodigoBarras.Enabled = False
        frmCodigo.Enabled = False
        cmdImprimirCodBarras.Enabled = False
        frmLarguraBarras.Enabled = False
        frmPosicaoCaracter.Enabled = False
        frmFonte.Enabled = False
      Else
        frmCodigoBarras.Enabled = True
        frmCodigo.Enabled = True
        cmdImprimirCodBarras.Enabled = True
        frmLarguraBarras.Enabled = True
        frmPosicaoCaracter.Enabled = True
        frmFonte.Enabled = True
      End If
      
      If Combo2.ListIndex = 2 Then
      cmdCortarPapel.Enabled = False
      cmdCortarParcial.Enabled = False
      cmdImprimirCodBarras.Enabled = False
      End If
      
   End If

  ' Configura o modelo da impressora
  
  
  iretorno = ConfiguraModeloImpressora(iModeloImpressora)
  If iretorno = -2 Then
    If Option1.Value Then
       MsgBox "Parâmetro inválido na função ConfiguraModeloImpressora.", vbInformation + vbOKOnly
    Else
       MsgBox "Invalid parameter in the function ConfiguraModeloImpressora.", vbInformation + vbOKOnly
    End If
  End If

  
End Sub

Private Sub Combo3_Click()
   Set Printer = Printers(Combo3.ListIndex)
End Sub

Private Sub ComboAlgorithm_Click()

Dim algorithm As Integer
If ComboAlgorithm.ListIndex = 0 Then
    algorithm = 0
ElseIf ComboAlgorithm.ListIndex = 1 Then
    algorithm = 1
End If

Dim ret As Integer
ret = SelectDithering(algorithm)


End Sub

Private Sub ComboBitola_Click()

Dim bitola As Integer
If ComboBitola.ListIndex = 0 Then
    bitola = 48
ElseIf ComboBitola.ListIndex = 1 Then
    bitola = 58
ElseIf ComboBitola.ListIndex = 2 Then
    bitola = 76
ElseIf ComboBitola.ListIndex = 3 Then
    bitola = 80
ElseIf ComboBitola.ListIndex = 4 Then
    bitola = 112
End If

Dim ret As Integer
ret = AjustaLarguraPapel(bitola)


End Sub

Private Sub Command1_Click()
CommonDialog2.ShowOpen


check = Split(CommonDialog2.FileName, ".", 2)

Dim ret As Integer
ret = StrComp(check(1), "bmp", vbTextCompare)

If ret <> 0 Then
    MsgBox ("Arquivo não é válido")
    Exit Sub
End If

FileName.Text = CommonDialog2.FileName



ImageBmp.Picture = LoadPicture(CommonDialog2.FileName)
        


   
    



End Sub

Private Sub Command2_Click()

If (CommonDialog2.FileName = "") Then
    Exit Sub
End If



Dim ret As Integer

Dim Altura As Integer
Dim Largura As Integer
Dim graus As Integer

If IsNumeric(BmpHeight.Text) = False _
    Or IsNumeric(BmpWidth.Text) = False _
    Or IsNumeric(Degrees) = False Then

    MsgBox ("Dimensões devem conter apenas números")
    Exit Sub

End If


Altura = CInt(BmpHeight)
Largura = CInt(BmpWidth)

graus = CInt(Degrees)


If (AjustaBtn.Value) Then
    Largura = -1

End If




ret = ImprimeBmpEspecial(CommonDialog2.FileName, Largura, Altura, graus)
Buffer = String(4, Chr(10))


ret = ComandoTX(Buffer, 4)
ret = AcionaGuilhotina(1)

End Sub

Private Sub Command3_Click()

If (CommonDialog2.FileName = "") Then
    Exit Sub
End If


Dim ret As Integer
Dim mode As Integer

If (RetratoBtn.Value) Then
    mode = 0
Else
    mode = 1
End If
    


ret = ImprimeBitmap(CommonDialog2.FileName, mode)


Buffer = String(4, Chr(10))


ret = ComandoTX(Buffer, 4)
ret = AcionaGuilhotina(1)

End Sub

Private Sub Form_Load()

   TraduzCaption (0)
   If Dir(App.Path + "\FLGBRAZL.ICO") <> "" Then
       frmPrincipal.Icon = LoadPicture(App.Path + "\FLGBRAZL.ICO")
   End If


   Dim Impressoras As Printer
   
   cbBarras.ListIndex = 0
   
   For Each Impressoras In Printers
       Combo3.AddItem Impressoras.DeviceName
   Next
   
   'Seta a porta de comunicação para COM1
   Combo1.ListIndex = 0
   
   'Seta o modelo da impressora para MP20 CI
   Combo2.ListIndex = 0
   
   ComboBitola.ListIndex = 0
   ComboAlgorithm.ListIndex = 0
   
  
   
   
   SSTab1.Tab = 0
   
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Option1_Click()

    TraduzCaption (0)
    If Dir(App.Path + "\FLGBRAZL.ICO") <> "" Then
        frmPrincipal.Icon = LoadPicture(App.Path + "\FLGBRAZL.ICO")
    End If
   
End Sub

Private Sub Option2_Click()

   TraduzCaption (1)
    If Dir(App.Path + "\FLGUSA02.ICO") <> "" Then
        frmPrincipal.Icon = LoadPicture(App.Path + "\FLGUSA02.ICO")
    End If

End Sub

Private Sub HabilitaFuncoesExtrato(Flag As Boolean)
    Frame5.Enabled = Flag  'tamanho extrato
    Label4.Enabled = Flag
    Text3.Enabled = Flag
    cmdProgramarExtrato.Enabled = Flag
    cmdHabilitarExtrato.Enabled = Flag
End Sub
Private Sub HabilitaProgramacaoPresenter(Flag As Boolean)
    Frame4.Enabled = Flag  'programação do presenter
    Label1.Enabled = Flag
    Label3.Enabled = Flag
    Text2.Enabled = Flag
    cmdProgramarPresenter.Enabled = Flag
    cmdHabilitarPresenter.Enabled = Flag
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 2 Then
      Frame8.Enabled = False
      Frame9.Enabled = False
   Else
      Frame8.Enabled = True
      Frame9.Enabled = True
   End If
End Sub
