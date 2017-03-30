VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMovimentacoes_Emissao_Nota_Totalizador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Nota Totalizador"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   8115
   Begin VB.TextBox txtDados_Adicionais 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4080
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      ToolTipText     =   "Dados Adicionais (Máximo de Caracteres 200)"
      Top             =   6510
      Width           =   3920
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1(0)"
      HotImageList    =   "ImageList1(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Confirmar"
            Object.ToolTipText     =   "Gravar registro - CTRL+G"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Configurar Impressora"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin VB.TextBox txtCupom 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   90
      MaxLength       =   850
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      ToolTipText     =   "Cupons Fiscais das Vendas (Limite 850 caracteres)"
      Top             =   6510
      Width           =   3920
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   4470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frame1 
      Caption         =   "Cálculo do Imposto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   90
      TabIndex        =   85
      Top             =   7080
      Width           =   7905
      Begin VB.TextBox txtValor_Total_Nota 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6300
         TabIndex        =   29
         ToolTipText     =   "Valor Total da Nota "
         Top             =   1110
         Width           =   1470
      End
      Begin VB.TextBox txtValor_IPI 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4755
         TabIndex        =   28
         ToolTipText     =   "Valor Total do IPI"
         Top             =   1110
         Width           =   1500
      End
      Begin VB.TextBox txtOutras_Despesas 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         TabIndex        =   50
         ToolTipText     =   "Outras Despesas Acessórias"
         Top             =   1110
         Width           =   1500
      End
      Begin VB.TextBox txtValor_Seguro 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1665
         TabIndex        =   26
         ToolTipText     =   "Valor do Seguro"
         Top             =   1110
         Width           =   1500
      End
      Begin VB.TextBox txtValor_Frete 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Valor do Frete"
         Top             =   1110
         Width           =   1500
      End
      Begin VB.TextBox txtValor_Total_Produtos 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6300
         TabIndex        =   24
         ToolTipText     =   "Valor Total dos Produtos"
         Top             =   540
         Width           =   1470
      End
      Begin VB.TextBox txtValor_Icms_Substituicao 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4750
         TabIndex        =   23
         ToolTipText     =   "Valor do ICMS Substituição"
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox txtIcms_Substituicao 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         TabIndex        =   22
         ToolTipText     =   "Base de Cálculo ICMS Substituição"
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox txtValor_Icms 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1660
         TabIndex        =   21
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox txtIcms 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Base de Cálculo do ICMS"
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label41 
         Caption         =   "VL Total Nota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6300
         TabIndex        =   93
         ToolTipText     =   "Valor Total dos Produtos"
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Label40 
         Caption         =   "Valor Total IPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4755
         TabIndex        =   92
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label Label39 
         Caption         =   "Outras Despesas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3210
         TabIndex        =   27
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Label38 
         Caption         =   "Valor Seguro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1650
         TabIndex        =   91
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label37 
         Caption         =   "Valor Frete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   90
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Label36 
         Caption         =   "VL Tot. Prod."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6300
         TabIndex        =   89
         ToolTipText     =   "Valor Total dos Produtos"
         Top             =   330
         Width           =   1365
      End
      Begin VB.Label Label35 
         Caption         =   "VL ICMS Sub."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4750
         TabIndex        =   88
         Top             =   330
         Width           =   1305
      End
      Begin VB.Label Label34 
         Caption         =   "ICMS Substit. &"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3210
         TabIndex        =   87
         Top             =   330
         Width           =   1365
      End
      Begin VB.Label Label33 
         Caption         =   "Valor ICMS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1660
         TabIndex        =   0
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label32 
         Caption         =   "ICMS %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   330
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab sstDestinatario_Transportador 
      Height          =   2985
      Left            =   90
      TabIndex        =   60
      Top             =   960
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5265
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Destinatário/Remet. (Cabeçalho)"
      TabPicture(0)   =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label20"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label21"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label24"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label25"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label26"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbbCfop"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtcNatureza_Transporte"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtpHora"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dtpSaida"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dtpEmissao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dtcCliente_Destinatario"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dtcCidade_Destinatario"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNumero_NF"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtSerie_NF"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCliente_Destinatario"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCidade_Destinatario"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtNatureza_Transporte"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Transportador/Volumes (Rodapé)"
      TabPicture(1)   =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPeso_Liquido_Transportador"
      Tab(1).Control(1)=   "txtPeso_Bruto_Transportador"
      Tab(1).Control(2)=   "txtNumero_Transportador"
      Tab(1).Control(3)=   "txtMarca_Transportador"
      Tab(1).Control(4)=   "txtEspecie_Transportador"
      Tab(1).Control(5)=   "txtValor_Frete_Transportador"
      Tab(1).Control(6)=   "txtInscricao_Estadual"
      Tab(1).Control(7)=   "txtUf_Transportador"
      Tab(1).Control(8)=   "txtCidade_Transportador"
      Tab(1).Control(9)=   "txtEndereco"
      Tab(1).Control(10)=   "txtCpf_cnpj"
      Tab(1).Control(11)=   "txtUf_Veiculo"
      Tab(1).Control(12)=   "txtPlaca_Veiculo"
      Tab(1).Control(13)=   "txtNome_Transportador"
      Tab(1).Control(14)=   "cbbFrete_Conta"
      Tab(1).Control(15)=   "dtcCidade_Transportador"
      Tab(1).Control(16)=   "Label31"
      Tab(1).Control(17)=   "Label30"
      Tab(1).Control(18)=   "Label29"
      Tab(1).Control(19)=   "Label28"
      Tab(1).Control(20)=   "Label27"
      Tab(1).Control(21)=   "Label11"
      Tab(1).Control(22)=   "Label10"
      Tab(1).Control(23)=   "Label8"
      Tab(1).Control(24)=   "Label7"
      Tab(1).Control(25)=   "Label6"
      Tab(1).Control(26)=   "Label5"
      Tab(1).Control(27)=   "Label4"
      Tab(1).Control(28)=   "Label3"
      Tab(1).Control(29)=   "Label2"
      Tab(1).Control(30)=   "Label1"
      Tab(1).ControlCount=   31
      Begin VB.TextBox txtPeso_Liquido_Transportador 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68430
         TabIndex        =   46
         Top             =   2520
         Width           =   1185
      End
      Begin VB.TextBox txtPeso_Bruto_Transportador 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69720
         TabIndex        =   45
         Top             =   2520
         Width           =   1240
      End
      Begin VB.TextBox txtNumero_Transportador 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71010
         TabIndex        =   44
         Top             =   2520
         Width           =   1240
      End
      Begin VB.TextBox txtMarca_Transportador 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72300
         TabIndex        =   43
         Top             =   2520
         Width           =   1240
      End
      Begin VB.TextBox txtEspecie_Transportador 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73590
         TabIndex        =   42
         Top             =   2520
         Width           =   1240
      End
      Begin VB.TextBox txtValor_Frete_Transportador 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   41
         Top             =   2520
         Width           =   1240
      End
      Begin VB.TextBox txtInscricao_Estadual 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69270
         TabIndex        =   36
         Top             =   1380
         Width           =   2025
      End
      Begin VB.TextBox txtUf_Transportador 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -67650
         TabIndex        =   40
         Top             =   1950
         Width           =   405
      End
      Begin VB.TextBox txtCidade_Transportador 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71370
         MaxLength       =   5
         TabIndex        =   38
         Top             =   1950
         Width           =   825
      End
      Begin VB.TextBox txtEndereco 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   37
         Top             =   1950
         Width           =   3465
      End
      Begin VB.TextBox txtCpf_cnpj 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71310
         TabIndex        =   35
         Top             =   1380
         Width           =   2000
      End
      Begin VB.TextBox txtUf_Veiculo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71760
         TabIndex        =   34
         Top             =   1380
         Width           =   405
      End
      Begin VB.TextBox txtPlaca_Veiculo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73050
         TabIndex        =   33
         Top             =   1380
         Width           =   1240
      End
      Begin VB.TextBox txtNome_Transportador 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   31
         Top             =   780
         Width           =   7635
      End
      Begin VB.TextBox txtNatureza_Transporte 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Código da Natureza de Transporte"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtCidade_Destinatario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   5
         TabIndex        =   9
         ToolTipText     =   "Código da Cidade"
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox txtCliente_Destinatario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Código do Cliente"
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox txtSerie_NF 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox txtNumero_NF 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   780
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dtcCidade_Destinatario 
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Top             =   1950
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcCliente_Destinatario 
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   1380
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpEmissao 
         Height          =   315
         Left            =   3960
         TabIndex        =   4
         Top             =   780
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   53215233
         CurrentDate     =   38229
      End
      Begin MSComCtl2.DTPicker dtpSaida 
         Height          =   315
         Left            =   5310
         TabIndex        =   5
         Top             =   780
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   53215233
         CurrentDate     =   38229
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   315
         Left            =   6660
         TabIndex        =   6
         Top             =   780
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm"
         Format          =   53215234
         CurrentDate     =   38229
      End
      Begin MSDataListLib.DataCombo dtcNatureza_Transporte 
         Height          =   315
         Left            =   1260
         TabIndex        =   12
         ToolTipText     =   "Descrição da Natureza Transporte"
         Top             =   2520
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AutoCompletar.CbCompleta cbbFrete_Conta 
         Height          =   315
         Left            =   -74880
         TabIndex        =   32
         Top             =   1380
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
      End
      Begin MSDataListLib.DataCombo dtcCidade_Transportador 
         Height          =   315
         Left            =   -70500
         TabIndex        =   39
         Top             =   1950
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AutoCompletar.CbCompleta cbbCfop 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   780
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
      End
      Begin VB.Label Label31 
         Caption         =   "Peso Líquido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68430
         TabIndex        =   84
         Top             =   2310
         Width           =   1125
      End
      Begin VB.Label Label30 
         Caption         =   "Peso Bruto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69720
         TabIndex        =   83
         Top             =   2310
         Width           =   1365
      End
      Begin VB.Label Label29 
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71010
         TabIndex        =   82
         Top             =   2310
         Width           =   1305
      End
      Begin VB.Label Label28 
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72300
         TabIndex        =   81
         Top             =   2310
         Width           =   1365
      End
      Begin VB.Label Label27 
         Caption         =   "Espécie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73590
         TabIndex        =   80
         Top             =   2310
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   2310
         Width           =   1365
      End
      Begin VB.Label Label10 
         Caption         =   "Inscrição Estad."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69270
         TabIndex        =   78
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67650
         TabIndex        =   77
         Top             =   1740
         Width           =   315
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -71370
         TabIndex        =   76
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label6 
         Caption         =   "Endereço"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   75
         Top             =   1740
         Width           =   1905
      End
      Begin VB.Label Label5 
         Caption         =   "CNPJ / CPF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71310
         TabIndex        =   74
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71760
         TabIndex        =   73
         Top             =   1140
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "Frete p/ Conta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   72
         Top             =   1140
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Placa Veículo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73050
         TabIndex        =   71
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Nome / Razão Social"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   70
         Top             =   540
         Width           =   1905
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2160
         TabIndex        =   69
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nat. Transp."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   68
         Top             =   2310
         Width           =   1020
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   67
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label23 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1140
         Width           =   645
      End
      Begin VB.Label Label21 
         Caption         =   "Série"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   65
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Número NF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         TabIndex        =   63
         Top             =   540
         Width           =   1065
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Data Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5310
         TabIndex        =   62
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Hora Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6660
         TabIndex        =   61
         Top             =   540
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdExcluir 
      Height          =   345
      Left            =   7680
      Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":17BA
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Excluir"
      Top             =   4230
      Width           =   345
   End
   Begin VB.CommandButton cmdIncluir 
      Height          =   345
      Left            =   7320
      Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":1904
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Incluir"
      Top             =   4230
      Width           =   345
   End
   Begin VB.TextBox txtValor_Total 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6390
      MaxLength       =   15
      TabIndex        =   17
      Top             =   4230
      Width           =   885
   End
   Begin VB.TextBox txtValor_Unitario 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5460
      MaxLength       =   10
      TabIndex        =   16
      Top             =   4230
      Width           =   885
   End
   Begin VB.TextBox txtQuantidade 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4530
      MaxLength       =   10
      TabIndex        =   15
      Top             =   4230
      Width           =   885
   End
   Begin AutoCompletar.CbCompleta cbbUnidade 
      Height          =   315
      Left            =   3810
      TabIndex        =   14
      Top             =   4230
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
   End
   Begin VB.TextBox txtProduto 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   13
      ToolTipText     =   "Código do Produto"
      Top             =   4230
      Width           =   885
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   315
      Left            =   90
      TabIndex        =   51
      ToolTipText     =   "Empresa"
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgItens 
      Height          =   1635
      Left            =   90
      TabIndex        =   18
      Top             =   4620
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2884
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSDataListLib.DataCombo dtcProduto 
      Height          =   315
      Left            =   1020
      TabIndex        =   47
      Top             =   4230
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   8940
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":2946
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":2C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":2F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":3314
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":36AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":39C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Totalizador.frx":59BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvFiltrar 
      Height          =   510
      Left            =   8490
      TabIndex        =   59
      Top             =   6120
      Width           =   7815
      lastProp        =   500
      _cx             =   13785
      _cy             =   900
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
   Begin VB.Label Label43 
      Caption         =   "Dados Adicionais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   96
      Top             =   6300
      Width           =   1365
   End
   Begin VB.Label lblImpressora_Padrao 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2010
      TabIndex        =   95
      Top             =   315
      Width           =   6015
   End
   Begin VB.Label Label42 
      Caption         =   "Cupons Fiscais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   94
      Top             =   6300
      Width           =   1365
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "VL Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6390
      TabIndex        =   57
      Top             =   3990
      Width           =   690
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "VL Unit."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5460
      TabIndex        =   56
      Top             =   3990
      Width           =   660
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Quantid."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4530
      TabIndex        =   55
      Top             =   3990
      Width           =   705
   End
   Begin VB.Label Label12 
      Caption         =   "Unid."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   54
      Top             =   3990
      Width           =   525
   End
   Begin VB.Label Label9 
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   53
      Top             =   3990
      Width           =   645
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   90
      TabIndex        =   52
      Top             =   360
      Width           =   1290
   End
End
Attribute VB_Name = "frmMovimentacoes_Emissao_Nota_Totalizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador de Vendas                                         '
' Objetivo...............: Movimentação de Emissão de Nota Totalizador                    '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rodrigo Santos                                                 '
' Data de Criação........: 12/06/2006                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public strSql As String
Public strSet As String
Dim strImprime As Integer
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim rstAplicacao As New ADODB.Recordset
Dim conexao_relatorio As New DLLConexao_Sistema.Conexao
Dim log As New DLLSystemManager.log
Dim strNomes As String
Dim strTamanho As String
Dim intLinha As Integer
Dim I As Integer
Dim dblValor_Produto As Double
Dim rstNumero_Nota As New ADODB.Recordset
Option Explicit

Private Sub cmdExcluir_Click()
    If dtcProduto.BoundText = Empty Then
       MsgBox "Não há Produto selecionado. Verifique!", vbInformation, "Only Tech"
       txtProduto.SetFocus
       Exit Sub
    End If

    If hfgItens.Rows <= 2 Then
       hfgItens.Rows = 2: hfgItens.ClearStructure
       'MONTANDO CABEÇALHO GRID
       strNomes = "Código,Produto, UN,Quant.,Valor Unit.,Valor Total"
       strTamanho = "700,3400,400,800,850,950"
        
       Movimentacoes.Monta_HFlex_Grid hfgItens, strTamanho, strNomes, 6, "Otica", Me
    Else
       hfgItens.RemoveItem (hfgItens.Row)
       'ATUALIZANDO VALOR TOTAL DOS PRODUTOS
       If txtValor_Total_Produtos.Text <> Empty Then
           txtValor_Total_Produtos.Text = CDbl(txtValor_Total_Produtos.Text) - dblValor_Produto
        End If
       
       For intLinha = 1 To hfgItens.Rows - 1
           hfgItens.TextMatrix(intLinha, 0) = intLinha
       Next intLinha
    End If

    txtProduto.Text = Empty
    cbbUnidade.Text = Empty
    txtQuantidade.Text = Empty
    txtValor_Unitario.Text = Empty
    txtValor_Total.Text = Empty
    cmdIncluir.ToolTipText = "Incluir"
    intLinha = Empty
    txtProduto.SetFocus
End Sub

Private Sub cmdIncluir_Click()

    'VERIFICANDO SE QUANTIDADE DE ITENS MAIOR QUE 10
    If hfgItens.Rows - 1 >= 10 Then
       MsgBox "Limite de 10 itens por Nota. Verifique!", vbInformation, "OnlyTech"
       txtProduto.Text = Empty
       cbbUnidade.Text = Empty
       txtQuantidade.Text = Empty
       txtValor_Unitario = Empty
       txtValor_Total = Empty
       txtProduto.SetFocus
       Exit Sub
    End If
    
    'ALTERANDO COLUNA DO GRID
    If cmdIncluir.ToolTipText = "Alterar" And dtcProduto.BoundText <> Empty Then
       
       'VERIFICANDO A EXISTÊNCIA DO TELEFONE NO GRID
       For I = 1 To hfgItens.Rows - 1
           hfgItens.Row = I
           hfgItens.Col = 1
           If hfgItens.Text = dtcProduto.BoundText And I <> intLinha Then
              MsgBox "Produto já cadastrado. Verifique.", vbInformation, "OnlyTech"
              txtProduto.Text = Empty
              cbbUnidade.Text = Empty
              txtQuantidade.Text = Empty
              txtValor_Unitario = Empty
              txtValor_Total = Empty
              txtProduto.SetFocus
              Exit Sub
           End If
        Next I
        
        hfgItens.TextMatrix(intLinha, 1) = dtcProduto.BoundText
        hfgItens.TextMatrix(intLinha, 2) = dtcProduto.Text
        hfgItens.TextMatrix(intLinha, 3) = cbbUnidade.Text
        hfgItens.TextMatrix(intLinha, 4) = txtQuantidade.Text
        hfgItens.TextMatrix(intLinha, 5) = txtValor_Unitario.Text
        hfgItens.TextMatrix(intLinha, 6) = txtValor_Total.Text
        
        If txtValor_Total_Produtos.Text <> Empty Then
           Format(txtValor_Total_Produtos.Text, "#,###0.00") = CDbl(txtValor_Total_Produtos.Text) - dblValor_Produto
        End If
        Format(txtValor_Total_Produtos.Text, "#,###0.00") = CDbl(txtValor_Total_Produtos.Text) + CDbl(hfgItens.TextMatrix(intLinha, 6))

    'INCLUINDO COLUNAS NO GRID
    ElseIf cmdIncluir.ToolTipText = "Incluir" And dtcProduto.BoundText <> Empty Then
       'VERIFICANDO A EXISTÊNCIA DO TELEFONE NO GRID
       For I = 1 To hfgItens.Rows - 1
           hfgItens.Row = I
           hfgItens.Col = 1
           If hfgItens.Text = dtcProduto.BoundText Then
              MsgBox "Produto já cadastrado. Verifique.", vbInformation, "OnlyTech"
              txtProduto.Text = Empty
              cbbUnidade.Text = Empty
              txtQuantidade.Text = Empty
              txtValor_Unitario = Empty
              txtValor_Total = Empty
              txtProduto.SetFocus
              Exit Sub
           End If
        Next I
       'INCREMENTANDO GRID ITENS E POSICIONANDO NA LINHA CORRETA
       If hfgItens.Rows = 2 Then
          If hfgItens.TextMatrix(1, 0) = Empty Then
             hfgItens.Row = hfgItens.Rows - 1
          Else
             hfgItens.Rows = hfgItens.Rows + 1
             hfgItens.Row = hfgItens.Rows - 1
          End If
       Else
          hfgItens.Rows = hfgItens.Rows + 1
          hfgItens.Row = hfgItens.Rows - 1
       End If
    
       'ABASTECENDO GRID COM VALORES DOS CAMPOS
       hfgItens.Row = hfgItens.Row: hfgItens.Col = 0
       hfgItens.CellBackColor = &H80FFFF: hfgItens.ColWidth(0) = 500
    
       hfgItens.TextMatrix(hfgItens.Row, 0) = hfgItens.Row
       hfgItens.TextMatrix(hfgItens.Row, 1) = dtcProduto.BoundText
       hfgItens.TextMatrix(hfgItens.Row, 2) = dtcProduto.Text
       hfgItens.TextMatrix(hfgItens.Row, 3) = cbbUnidade.Text
       hfgItens.TextMatrix(hfgItens.Row, 4) = txtQuantidade.Text
       hfgItens.TextMatrix(hfgItens.Row, 5) = txtValor_Unitario.Text
       hfgItens.TextMatrix(hfgItens.Row, 6) = txtValor_Total.Text
       txtValor_Total_Produtos.Text = CDbl(txtValor_Total_Produtos.Text) + CDbl(hfgItens.TextMatrix(hfgItens.Row, 6))
              
    Else
       MsgBox "Não existe dado para inclusão. Verifique!", vbInformation, "Only Tech"
    End If
    
    hfgItens.ColAlignment(2) = 0
    hfgItens.ColAlignment(3) = 4
       
    txtProduto.Text = Empty
    cbbUnidade.Text = Empty
    txtQuantidade.Text = Empty
    txtValor_Unitario.Text = Empty
    txtValor_Total.Text = Empty
    cmdIncluir.ToolTipText = "Incluir"
    intLinha = Empty
    txtValor_Total_Produtos.Text = Format(txtValor_Total_Produtos.Text, "#,###0.00")
    txtProduto.SetFocus
End Sub

Private Sub dtcCidade_Destinatario_LostFocus()
    txtCidade_Destinatario.Text = dtcCidade_Destinatario.BoundText
    If IsNumeric(txtCidade_Destinatario.Text) = False Or dtcCidade_Destinatario.Text = Empty Then txtCidade_Destinatario.Text = Empty: Exit Sub
End Sub

Private Sub dtcCidade_Transportador_GotFocus()
    If Me.txtCidade_Transportador.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCidade_Transportador.Text)
    End If
End Sub

Private Sub dtcCidade_Transportador_LostFocus()
       
    Dim rstCidade As New ADODB.Recordset
    txtCidade_Transportador.Text = dtcCidade_Transportador.BoundText
    
    If IsNumeric(txtCidade_Transportador.Text) = False Or dtcCidade_Transportador.Text = Empty Then txtCidade_Transportador.Text = Empty: Exit Sub
    
    dtcCidade_Transportador.BoundText = txtCidade_Transportador.Text
    
    strSql = "Select TBCidade_Otica.DFUf_TBCidade_Otica FROM TBCidade_Otica " & _
             "WHERE TBCidade_Otica.IXCodigo_Correios_TBCidade_otica = '" & txtCidade_Transportador.Text & "'"
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstCidade, "Otica", Me)
    
    If rstCidade.RecordCount <> 0 Then
       txtUf_Transportador.Text = rstCidade.Fields("DFUf_TBCidade_Otica")
    Else
       txtUf_Transportador.Text = Empty
    End If
    
    rstCidade.Close
End Sub

Private Sub dtcCliente_Destinatario_LostFocus()
    Dim rstInfo_Cliente As New ADODB.Recordset
    
    txtCliente_Destinatario.Text = dtcCliente_Destinatario.BoundText
    If IsNumeric(txtCliente_Destinatario.Text) = False Or dtcCliente_Destinatario.Text = Empty Then txtCliente_Destinatario.Text = Empty: Exit Sub
    
    'BUSCANDO INFORMAÇÕES DA CIDADE
    strSql = "SELECT IXCodigo_TBCliente,DFEndereco_TBCliente," & _
             "DFNumero_TBCliente,DFBairro_TBCliente,DFCep_TBCliente," & _
             "DFInscricao_estadual_TBCliente,DFCpf_TBCliente," & _
             "TBCidade_otica.IXCodigo_Correios_TBCidade_otica," & _
             "TBCidade_otica.DFNome_TBCidade_otica," & _
             "TBCidade_otica.DFUf_TBCidade_otica " & _
             "FROM TBCliente " & _
             "INNER JOIN TBCidade_otica " & _
             "ON TBCliente.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
             "WHERE IXCodigo_TBCliente = '" & dtcCliente_Destinatario.BoundText & "'" & _
             "AND IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
             
     Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstInfo_Cliente, "Otica", Me)
     
     If rstInfo_Cliente.EOF = False Then
        dtcCidade_Destinatario.BoundText = rstInfo_Cliente!IXCodigo_Correios_TBCidade_otica
        txtCidade_Destinatario.Text = dtcCidade_Destinatario.BoundText
     End If
     
     Set rstInfo_Cliente = Nothing
     
End Sub

Private Sub dtcEmpresa_LostFocus()
    dtcEmpresa.Enabled = False
    Call MontaDataCombo
    txtCliente_Destinatario.Text = Empty
    txtNatureza_Transporte.Text = Empty
    txtCidade_Destinatario.Text = Empty
    
End Sub

Private Sub dtcNatureza_Transporte_LostFocus()
    sstDestinatario_Transportador.Tab = 1
    txtNome_Transportador.SetFocus
    txtNatureza_Transporte.Text = dtcNatureza_Transporte.BoundText
    'If IsNumeric(txtNatureza_Transporte.Text) = False Or dtcNatureza_Transporte.Text = Empty Then txtNatureza_Transporte.Text = Empty: Exit Sub
End Sub

Private Sub dtcProduto_LostFocus()
    If dtcProduto.BoundText <> Empty Then txtProduto.Text = dtcProduto.BoundText Else Exit Sub
    If IsNumeric(txtProduto.Text) = False Or dtcProduto.Text = Empty Then txtProduto.Text = Empty: Exit Sub
End Sub

Private Sub dtpEmissao_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos no dataPicker pelo ENTER
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpHora_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos no dataPicker pelo ENTER
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpSaida_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos no dataPicker pelo ENTER
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    
    On Error GoTo erro
    
    'INFORMAÇÕES CONSTANTES PARA O LOG
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Emissão de Nota Cupom"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'INFORMAÇÕES VARIAVEIS PARA O LOG
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
      
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    strSql = "SELECT * FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente_Destinatario, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_Correios_TBCidade_otica,DFNome_TBCidade_otica FROM TBCidade_otica"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_Correios_TBCidade_otica", "DFNome_TBCidade_otica", dtcCidade_Destinatario, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_Correios_TBCidade_otica,DFNome_TBCidade_otica FROM TBCidade_otica"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_Correios_TBCidade_otica", "DFNome_TBCidade_otica", dtcCidade_Transportador, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBNatureza_transporte,DFDescricao_TBNatureza_transporte FROM TBNatureza_transporte"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBNatureza_transporte", "DFDescricao_TBNatureza_transporte", dtcNatureza_Transporte, strSql, "BDRetaguarda", "Otica", Me

    strSql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me
        
    'CAPTURANDO PRÓXIMA NOTA E SÉRIE
    strSql = "SELECT DFProxima_nota_cupom_TBParametros_gerais," & _
             "DFProxima_serie_nota_cupom_TBParametros_gerais FROM TBParametros_gerais " & _
             "WHERE PFKCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
             
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstNumero_Nota, "Otica", Me)
        
    If rstNumero_Nota!DFProxima_nota_cupom_TBParametros_gerais <> Empty Then
       txtNumero_NF.Text = rstNumero_Nota!DFProxima_nota_cupom_TBParametros_gerais
    End If
    If rstNumero_Nota!DFProxima_serie_nota_cupom_TBParametros_gerais <> Empty Then
       txtSerie_NF.Text = rstNumero_Nota!DFProxima_serie_nota_cupom_TBParametros_gerais
    End If
    
    Set rstNumero_Nota = Nothing
    
    'MONTANDO CABEÇALHO GRID
    strNomes = "Código,Produto, UN,Quant.,Valor Unit.,Valor Total"
    strTamanho = "700,3400,400,800,850,950"
    
    Movimentacoes.Monta_HFlex_Grid hfgItens, strTamanho, strNomes, 6, "Otica", Me
    
    dtpEmissao.Value = Date
    dtpSaida.Value = Date
    dtpHora.Value = time
    txtValor_Total_Produtos.Text = "0,00"
    lblImpressora_Padrao.Caption = Funcoes_Gerais.Impressora_Padrao.DeviceName & ", na porta: " & Funcoes_Gerais.Impressora_Padrao.Port
    sstDestinatario_Transportador.Tab = 0
    Call MontaCombo
    
    log.Descricao = "Inicializando Emissão de Nota Cupom"
    'GRAVANDO LOG
    log.Gravar_log "Otica", Me
    
    Exit Sub
    
erro:
    Call erro.erro(Me, "Otica", "Load")
    Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 71: Call Gravar   'CTRL+G
                       Case 83: Unload Me  'CTRL+S
                End Select
    End Select
    If KeyCode = "113" Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub hfgItens_Click()
    
    cmdIncluir.ToolTipText = "Incluir"
    intLinha = Empty
    txtProduto.Text = Empty
    cbbUnidade.Text = Empty
    txtQuantidade.Text = Empty
    txtValor_Unitario.Text = Empty
    txtValor_Total.Text = Empty
    
End Sub

Private Sub hfgItens_DblClick()

    dblValor_Produto = Empty
    hfgItens.Sort = 1
    cmdIncluir.ToolTipText = "Alterar"
    intLinha = hfgItens.Row
    txtProduto.Text = hfgItens.TextMatrix(intLinha, 1)
    dtcProduto.Text = hfgItens.TextMatrix(intLinha, 2)
    cbbUnidade.Text = hfgItens.TextMatrix(intLinha, 3)
    txtQuantidade.Text = hfgItens.TextMatrix(intLinha, 4)
    txtValor_Unitario.Text = hfgItens.TextMatrix(intLinha, 5)
    txtValor_Total.Text = hfgItens.TextMatrix(intLinha, 6)
    dblValor_Produto = hfgItens.TextMatrix(intLinha, 6)
    
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Gravar
           Case 2: Call Cancelar
           Case 4: CommonDialog1.ShowPrinter
    End Select
    
End Sub

Private Function Gravar()

    Dim strLinha_Impressao As String
    Dim strCaminho_impressora As String
    Dim strVelha_Font As String
    Dim strVelho_Size_Font As String
    Dim rstInfo_Cliente As New ADODB.Recordset
    Dim rstCaminho_Impressora As New ADODB.Recordset
    Dim rstNumero_Nota_Final As New ADODB.Recordset
    Dim I As Integer
    Dim intLetras As Integer
    Dim strResto As String
    Dim strInicial As String
    Dim intFrete_Conta As Integer
    Dim strNatureza_Op As String
    Dim strCfop As String
        
    If cbbFrete_Conta.Text = "1-Emitente" Then
       intFrete_Conta = 1
    Else
       intFrete_Conta = 2
    End If
    
    If cbbCfop.Text = "Dentro Estado" Then
       strNatureza_Op = "Lct.Ef.Op.Tb.Reg.Eq.ECF"
       strCfop = "5.929"
    Else
       strNatureza_Op = "Lct.Ef.Op.Tb.Reg.Eq.ECF"
       strCfop = "6.929"
    End If
    
    'VERIFICANDO SE NOTA CUPOM JÁ FOI CADASTRADA
    strSql = "SELECT DFProxima_nota_cupom_TBParametros_gerais " & _
             "FROM TBParametros_gerais " & _
             "WHERE PFKCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
             
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstNumero_Nota, "Otica", Me)
        
    If rstNumero_Nota!DFProxima_nota_cupom_TBParametros_gerais <> Empty Then
       If txtNumero_NF.Text <> rstNumero_Nota!DFProxima_nota_cupom_TBParametros_gerais Then
          MsgBox "Nota difere do valor em Parâmetros Gerais. Verifique!", vbInformation, "OnlyTech"
          txtNumero_NF.SetFocus
          Exit Function
       End If
    End If
    
    Set rstNumero_Nota = Nothing
    
    'BUSCANDO INFORMAÇÕES DO CLIENTE
    strSql = "SELECT IXCodigo_TBCliente,DFEndereco_TBCliente," & _
             "DFNumero_TBCliente,DFBairro_TBCliente,DFCep_TBCliente," & _
             "DFInscricao_estadual_TBCliente,DFCpf_TBCliente," & _
             "TBCidade_otica.IXCodigo_Correios_TBCidade_otica," & _
             "TBCidade_otica.DFNome_TBCidade_otica," & _
             "TBCidade_otica.DFUf_TBCidade_otica " & _
             "FROM TBCliente " & _
             "INNER JOIN TBCidade_otica " & _
             "ON TBCliente.FKId_TBCidade_otica = TBCidade_otica.PKId_TBCidade_otica " & _
             "WHERE IXCodigo_TBCliente = '" & dtcCliente_Destinatario.BoundText & "' " & _
             "AND IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
             
     Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstInfo_Cliente, "Otica", Me)
    
    'BUSCANDO CAMINHO DA IMPRESSORA
    strSql = "SELECT DFCaminho_impressora_via_porta_TBParametros_gerais FROM TBParametros_Gerais WHERE PFKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCaminho_Impressora, "Otica", Me
       
    If IsNull(rstCaminho_Impressora.Fields("DFCaminho_impressora_via_porta_TBParametros_gerais")) = True Then
       strCaminho_impressora = ""
    Else
       strCaminho_impressora = rstCaminho_Impressora.Fields("DFCaminho_impressora_via_porta_TBParametros_gerais")
    End If
    
    Set rstCaminho_Impressora = Nothing
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''  IMPRESSÃO DIRETO NA PORTA  ''''''''''''''''''''''''''''''
     
    'PRESERVANDO FONT E TAMANHO ANTERIOR
    strVelha_Font = Printer.FontName
    strVelho_Size_Font = Printer.FontSize
    
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    
    ''''''''''''''''''''''''''''''''''''''    CABEÇALHO    '''''''''''''''''''''''''''''''''''''
    strLinha_Impressao = "X"
    'Mandando o comando direto na porta da impressora
    Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (strCaminho_impressora)
    'Print #1, 'salta uma linha
    Print #1, Tab(51); strLinha_Impressao; Tab(69); txtNumero_NF.Text & " " & txtSerie_NF.Text  'Imprime caracter Comprimido para matriciais Epson
    Funcoes_Gerais.Fecha_porta_impressora_via_LPT1

                         
    'Mandando o comando direto na porta da impressora
    Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (strCaminho_impressora)
    Print #1, 'salta uma linha
    Print #1, 'salta uma linha
    Print #1, 'salta uma linha
    Print #1, 'salta uma linha
    Print #1, strNatureza_Op; Tab(25); strCfop
    Print #1, 'salta uma linha
    Print #1, 'salta uma linha
    Print #1, dtcCliente_Destinatario.Text; Tab(50); rstInfo_Cliente!DFCpf_TBCliente; Tab(70); dtpEmissao.Value
    Print #1, 'salta uma linha
    Print #1, rstInfo_Cliente!DFEndereco_TBCliente & " - " & rstInfo_Cliente!DFNumero_TBCliente; _
              Tab(40); Left(rstInfo_Cliente!DFBairro_TBCliente, 18); Tab(58); rstInfo_Cliente!DFCep_TBCliente; _
              Tab(70); dtpSaida.Value
    Print #1, 'salta uma linha
    Print #1, rstInfo_Cliente!DFNome_TBCidade_otica; Tab(45); rstInfo_Cliente!DFUf_TBCidade_otica; _
              Tab(50); rstInfo_Cliente!DFInscricao_estadual_TBCliente; Tab(70); dtpHora.Value
    
    Set rstInfo_Cliente = Nothing
       
    '''''''''''''''''''''''''''''''''''''''''    ITENS    '''''''''''''''''''''''''''''''''''''
    Print #1, 'salta uma linha
    Print #1, 'salta uma linha
    Print #1, 'salta uma linha
    Close #1
    'Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (strCaminho_impressora)
    For I = 1 To hfgItens.Rows - 1
        hfgItens.Row = I
        Printer.FontSize = 8
        
        '             CÓDIGO PRODUTO                      DESCRIÇÃO PRODUTO                   UNIDADE                             QUANTIDADE                           VALOR UNITÁRIO                      VALOR TOTAL
        Printer.Print hfgItens.TextMatrix(I, 1); Tab(17); hfgItens.TextMatrix(I, 2); Tab(87); hfgItens.TextMatrix(I, 3); Tab(92); hfgItens.TextMatrix(I, 4); Tab(113); hfgItens.TextMatrix(I, 5); Tab(134); hfgItens.TextMatrix(I, 6)
            
    Next I
    Printer.FontSize = 10
    intLetras = Len(txtCupom.Text)
    strInicial = txtCupom.Text
    
    If Len(txtCupom.Text) > 64 Then
       
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print "VENDAS REFERENTES AOS CUPONS FISCAIS N°: "
       
       'MODIFICAÇÃO SOBRE O MÁXIMO DE CARACTERES POR LINHA (ANTIGO VALOR 115)
       'PEDIDO PELO LEONARDO AZEVEDO(PONTO AZUL)
       Do While intLetras > 110
          intLetras = intLetras - 110
          strResto = Left(strInicial, 110)
          Printer.Print strResto
          strInicial = Right(strInicial, intLetras)
       Loop
       Printer.Print strInicial
    Else
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print "VENDAS REFERENTES AOS CUPONS FISCAIS N°: "; txtCupom.Text
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''   RODAPÉ   ''''''''''''''''''''''''''''''''''''
    Printer.CurrentY = 5570
    Printer.Print txtIcms.Text; Tab(27); txtValor_Icms.Text; Tab(53); txtIcms_Substituicao; Tab(87); txtValor_Icms_Substituicao; Tab(113); txtValor_Total_Produtos.Text
    Printer.Print " "
    Printer.Print txtValor_Frete.Text; Tab(27); txtValor_Seguro.Text; Tab(53); txtOutras_Despesas; Tab(87); txtValor_IPI; Tab(113); txtValor_Total_Nota.Text
    Printer.Print " "
    Printer.Print " "
    Printer.Print " "
    Printer.Print txtNome_Transportador.Text; Tab(70); intFrete_Conta; Tab(78); txtPlaca_Veiculo.Text; Tab(93); txtUf_Veiculo.Text; Tab(99); txtCpf_cnpj.Text
    Printer.Print " "
    Printer.Print txtEndereco.Text; Tab(58); dtcCidade_Transportador.Text; Tab(93); txtUf_Transportador.Text; Tab(99); txtInscricao_Estadual.Text
    Printer.Print " "
    Printer.Print txtValor_Frete_Transportador.Text; Tab(25); txtEspecie_Transportador.Text; Tab(58); txtMarca_Transportador.Text; Tab(80); txtNumero_Transportador.Text; Tab(97); txtPeso_Bruto_Transportador.Text; Tab(114); txtPeso_Liquido_Transportador.Text
    Printer.Print " "
    Printer.Print " "
       
    'DIVIDINDO O CAMPO INFORMAÇÕES COMPLEMENTARES EM 34 CARACTERES POR LINHA
    intLetras = Len(txtDados_Adicionais.Text)
    strInicial = txtDados_Adicionais.Text
    
    Do While intLetras > 34
       intLetras = intLetras - 34
       strResto = Left(strInicial, 34)
       Printer.Print strResto
       strInicial = Right(strInicial, intLetras)
    Loop
    Printer.Print strInicial
    
    Printer.CurrentY = 10650
    Printer.Print Tab(105); txtNumero_NF.Text & " " & txtSerie_NF.Text
    
    'RETORNANDO FONT E TAMANHO ANTERIORES
    Printer.FontName = strVelha_Font
    Printer.FontSize = strVelho_Size_Font
    
    'FINALIZANDO DOCUMENTO
    Printer.EndDoc
    
    '''''''''''''''''''''''''''''''''''''''' FIM IMPRESSÃO '''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'ATUALIZANDO PRÓXIMA NOTA CUPOM NA TBPARAMETROS_GERAIS
    If txtNumero_NF.Text <> Empty Then
       strSet = "SET DFProxima_nota_cupom_TBParametros_gerais = '" & CInt(txtNumero_NF.Text) + 1 & "'"
    End If
    
    Call funcoes_banco.Alterar("TBParametros_gerais", strSet, "PFKCodigo_TBEmpresa", dtcEmpresa.BoundText, "Otica", Me)
    
    
    'CAPTURANDO PRÓXIMA NOTA E SÉRIE
    strSql = "SELECT DFProxima_nota_cupom_TBParametros_gerais," & _
             "DFProxima_serie_nota_cupom_TBParametros_gerais FROM TBParametros_gerais " & _
             "WHERE PFKCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
             
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstNumero_Nota_Final, "Otica", Me)
    
    Call Cancelar
        
    If rstNumero_Nota_Final!DFProxima_nota_cupom_TBParametros_gerais <> Empty Then
       txtNumero_NF.Text = rstNumero_Nota_Final!DFProxima_nota_cupom_TBParametros_gerais
    End If
    If rstNumero_Nota_Final!DFProxima_serie_nota_cupom_TBParametros_gerais <> Empty Then
       txtSerie_NF.Text = rstNumero_Nota_Final!DFProxima_serie_nota_cupom_TBParametros_gerais
    End If
    
    Set rstNumero_Nota_Final = Nothing
    
    cbbCfop.SetFocus
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''' IMPRESSÃO PARA EPSON LX300 ''''''''''''''''''''''''''''
      
    
'    'PRESERVANDO FONT E TAMANHO ANTERIOR
'    strVelha_Font = Printer.FontName
'    strVelho_Size_Font = Printer.FontSize
'
'    Printer.FontName = "Arial"
'    Printer.FontSize = 8
'
'    ''''''''''''''''''''''''''''''''''''''    CABEÇALHO    '''''''''''''''''''''''''''''''''''''
'    strLinha_Impressao = "X"
'    'Mandando o comando direto na porta da impressora
'    Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (strCaminho_impressora)
'    Print #1, 'salta uma linha
'    Print #1, Tab(54); strLinha_Impressao; Tab(71); txtNumero_NF.Text & "-" & txtSerie_NF.Text & " " & txtLetra_NF.Text 'Imprime caracter Comprimido para matriciais Epson
'    Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
'
'    strLinha_Impressao = dtcNatureza_Transporte.BoundText & " - " & dtcNatureza_Transporte.Text
'
'    'Mandando o comando direto na porta da impressora
'    Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (strCaminho_impressora)
'    Print #1, 'salta uma linha
'    Print #1, 'salta uma linha
'    Print #1, 'salta uma linha
'    Print #1, 'salta uma linha
'    Print #1, Spc(1); strLinha_Impressao; Tab(25); txtCfop.Text
'    Print #1, 'salta uma linha
'    Print #1, 'salta uma linha
'    Print #1, Spc(1); dtcCliente_Destinatario.Text; Tab(52); rstInfo_Cliente!DFCpf_TBCliente; Tab(72); dtpEmissao.Value
'
'    Print #1, Spc(1); rstInfo_Cliente!DFEndereco_TBCliente & " - " & rstInfo_Cliente!DFNumero_TBCliente; _
'              Tab(42); rstInfo_Cliente!DFBairro_TBCliente; Tab(60); rstInfo_Cliente!DFCep_TBCliente; _
'              Tab(72); dtpSaida.Value
'
'    Print #1, Spc(1); rstInfo_Cliente!DFNome_TBCidade_otica; Tab(47); rstInfo_Cliente!DFUf_TBCidade_otica; _
'              Tab(52); rstInfo_Cliente!DFInscricao_estadual_TBCliente; Tab(72); dtpHora.Value
'
'    Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
'
'    Set rstInfo_Cliente = Nothing
'
'    '''''''''''''''''''''''''''''''''''''''''    ITENS    '''''''''''''''''''''''''''''''''''''
'
'    Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (strCaminho_impressora)
'    Print #1, 'salta uma linha
'    Print #1, 'salta uma linha
'    Print #1, 'salta uma linha
'    Close #1
'    Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (strCaminho_impressora)
'    For I = 1 To hfgItens.Rows - 1
'        hfgItens.Row = I
'        Printer.FontSize = 8
'
'        '                     CÓDIGO PRODUTO                      DESCRIÇÃO PRODUTO                   UNIDADE                             QUANTIDADE                           VALOR UNITÁRIO                      VALOR TOTAL
'        Printer.Print Spc(2); hfgItens.TextMatrix(I, 1); Tab(17); hfgItens.TextMatrix(I, 2); Tab(92); hfgItens.TextMatrix(I, 3); Tab(97); hfgItens.TextMatrix(I, 4); Tab(113); hfgItens.TextMatrix(I, 5); Tab(134); hfgItens.TextMatrix(I, 6)
'
'    Next I
'    Printer.FontSize = 10
'    intLetras = Len(txtCupom.Text)
'    strInicial = txtCupom.Text
'
'    If Len(txtCupom.Text) > 85 Then
'
'       Printer.Print " "
'       Printer.Print " "
'       Printer.Print " "
'       Printer.Print " "
'       Printer.Print Tab(17); "VENDAS REFERENTES AOS CUPONS FISCAIS N°: "
'
'       Do While intLetras > 85
'          intLetras = intLetras - 85
'          strResto = Left(strInicial, 85)
'          Printer.Print Tab(17); strResto
'          strInicial = Right(strInicial, intLetras)
'       Loop
'       Printer.Print Tab(17); strInicial
'    Else
'       Printer.Print " "
'       Printer.Print " "
'       Printer.Print " "
'       Printer.Print " "
'       Printer.Print Tab(17); "VENDAS REFERENTES AOS CUPONS FISCAIS N°: "; txtCupom.Text
'    End If
'
'    '''''''''''''''''''''''''''''''''''''''''''   RODAPÉ   ''''''''''''''''''''''''''''''''''''
'    Printer.CurrentY = 5600
'    Printer.Print Spc(1); txtIcms.Text; Tab(27); txtValor_Icms.Text; Tab(53); txtIcms_Substituicao; Tab(87); txtValor_Icms_Substituicao; Tab(113); txtValor_Total_Produtos.Text
'    Printer.Print " "
'    Printer.Print Spc(1); txtValor_Frete.Text; Tab(27); txtValor_Seguro.Text; Tab(53); txtOutras_Despesas; Tab(87); txtValor_IPI; Tab(113); txtValor_Total_Nota.Text
'    Printer.Print " "
'    Printer.Print " "
'    Printer.Print " "
'    Printer.Print Spc(1); txtNome_Transportador.Text; Tab(77); intFrete_Conta; Tab(85); txtPlaca_Veiculo.Text; Tab(100); txtUf_Veiculo.Text; Tab(106); txtCpf_cnpj.Text
'    Printer.Print " "
'    Printer.Print Spc(1); txtEndereco.Text; Tab(62); dtcCidade_Transportador.Text; Tab(100); txtUf_Transportador.Text; Tab(106); txtInscricao_Estadual.Text
'    Printer.Print " "
'    Printer.Print Spc(1); txtValor_Frete_Transportador.Text; Tab(25); txtEspecie_Transportador.Text; Tab(62); txtMarca_Transportador.Text; Tab(90); txtNumero_Transportador.Text; Tab(105); txtPeso_Bruto_Transportador.Text; Tab(121); txtPeso_Liquido_Transportador.Text
'    Printer.Print " "
'    Printer.Print " "
'
'    'DIVIDINDO O CAMPO INFORMAÇÕES COMPLEMENTARES EM 34 CARACTERES POR LINHA
'    intLetras = Len(txtDados_Adicionais.Text)
'    strInicial = txtDados_Adicionais.Text
'
'    Do While intLetras > 34
'       intLetras = intLetras - 34
'       strResto = Left(strInicial, 34)
'       Printer.Print Spc(1); strResto
'       strInicial = Right(strInicial, intLetras)
'    Loop
'    Printer.Print Spc(1); strInicial
'
'    Printer.CurrentY = 10650
'    Printer.Print Tab(115); txtNumero_NF.Text & "-" & txtSerie_NF.Text
'
'    'RETORNANDO FONT E TAMANHO ANTERIORES
'    Printer.FontName = strVelha_Font
'    Printer.FontSize = strVelho_Size_Font
'
'    'FINALIZANDO DOCUMENTO
'    Close #1
'    Printer.EndDoc
'
'    '''''''''''''''''''''''''''''''''''''''' FIM IMPRESSÃO '''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    'ATUALIZANDO PRÓXIMA NOTA CUPOM NA TBPARAMETROS_GERAIS
'    If txtNumero_NF.Text <> Empty Then
'       strSet = "SET DFProxima_nota_cupom_TBParametros_gerais = '" & CInt(txtNumero_NF.Text) + 1 & "'"
'    End If
'
'    Call funcoes_banco.Alterar("TBParametros_gerais", strSet, "PFKCodigo_TBEmpresa", dtcEmpresa.BoundText, "Otica", Me)
'
'
'    'CAPTURANDO PRÓXIMA NOTA E SÉRIE
'    strSql = "SELECT DFProxima_nota_cupom_TBParametros_gerais," & _
'             "DFProxima_serie_nota_cupom_TBParametros_gerais FROM TBParametros_gerais " & _
'             "WHERE PFKCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
'
'    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstNumero_Nota, "Otica", Me)
'
'    Call Cancelar
'
'    If rstNumero_Nota!DFProxima_nota_cupom_TBParametros_gerais <> Empty Then
'       txtNumero_NF.Text = rstNumero_Nota!DFProxima_nota_cupom_TBParametros_gerais
'    End If
'    If rstNumero_Nota!DFProxima_serie_nota_cupom_TBParametros_gerais <> Empty Then
'       txtSerie_NF.Text = rstNumero_Nota!DFProxima_serie_nota_cupom_TBParametros_gerais
'    End If
'
'    Set rstNumero_Nota = Nothing
'
'    txtLetra_NF.SetFocus


End Function
Private Function Cancelar()

    cmdIncluir.ToolTipText = "Incluir"
    txtNumero_NF.Text = Empty
    txtSerie_NF.Text = Empty
    txtCidade_Destinatario.Text = Empty
    txtCidade_Transportador.Text = Empty
    txtCliente_Destinatario.Text = Empty
    txtCpf_cnpj.Text = Empty
    txtCupom.Text = Empty
    txtDados_Adicionais.Text = Empty
    txtEndereco.Text = Empty
    txtIcms.Text = Empty
    txtIcms_Substituicao.Text = Empty
    txtInscricao_Estadual.Text = Empty
    cbbCfop.Text = Empty
    txtMarca_Transportador.Text = Empty
    txtNatureza_Transporte.Text = Empty
    txtNome_Transportador.Text = Empty
    txtNumero_Transportador.Text = Empty
    txtOutras_Despesas.Text = Empty
    txtPeso_Bruto_Transportador.Text = Empty
    txtPeso_Liquido_Transportador.Text = Empty
    txtPlaca_Veiculo.Text = Empty
    txtUf_Transportador.Text = Empty
    txtUf_Veiculo.Text = Empty
    txtValor_Frete.Text = Empty
    txtValor_Icms.Text = Empty
    txtValor_Icms_Substituicao.Text = Empty
    txtValor_IPI.Text = Empty
    txtValor_Seguro.Text = Empty
    txtValor_Total.Text = Empty
    txtValor_Total_Produtos.Text = "0,00"
    txtValor_Unitario.Text = Empty
    cbbFrete_Conta.Text = Empty
    cbbUnidade.Text = Empty
    intLinha = Empty
    txtProduto.Text = Empty
    cbbUnidade.Text = Empty
    txtQuantidade.Text = Empty
    txtValor_Unitario.Text = Empty
    txtValor_Total.Text = Empty
    txtValor_Total_Nota.Text = Empty
    txtValor_Frete_Transportador.Text = Empty
    txtEspecie_Transportador.Text = Empty
    
    sstDestinatario_Transportador.Tab = 0
    
    'REMONTANDO GRID
    hfgItens.Rows = 2: hfgItens.ClearStructure
    
    strNomes = "Código,Produto, UN,Quant.,Valor Unit.,Valor Total"
    strTamanho = "700,3400,400,800,850,950"
        
    Movimentacoes.Monta_HFlex_Grid hfgItens, strTamanho, strNomes, 6, "Otica", Me
    
End Function

Private Sub txtCidade_Destinatario_Change()
    dtcCidade_Destinatario.BoundText = txtCidade_Destinatario.Text
    If IsNumeric(txtCidade_Destinatario.Text) = False Then
       txtCidade_Destinatario.Text = Empty
       Exit Sub
    End If
End Sub

Private Sub txtCidade_Destinatario_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCidade_Destinatario_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCidade_Transportador_Change()
    dtcCidade_Transportador.BoundText = txtCidade_Transportador.Text
    If IsNumeric(txtCidade_Transportador.Text) = False Then
       txtCidade_Transportador.Text = Empty
       Exit Sub
    End If
End Sub

Private Sub txtCidade_Transportador_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCliente_Destinatario_Change()
    dtcCliente_Destinatario.BoundText = txtCliente_Destinatario.Text
    If IsNumeric(txtCliente_Destinatario.Text) = False Then
       txtCliente_Destinatario.Text = Empty
       Exit Sub
    End If
End Sub

Private Sub txtCliente_Destinatario_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCliente_Destinatario_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCupom_LostFocus()
    txtCupom.Text = UCase(txtCupom)
End Sub

Private Sub txtDados_Adicionais_LostFocus()
    txtDados_Adicionais.Text = UCase(txtDados_Adicionais)
End Sub

Private Sub txtEndereco_LostFocus()
    txtEndereco.Text = UCase(txtEndereco.Text)
End Sub

Private Sub txtEspecie_Transportador_LostFocus()
    txtEspecie_Transportador.Text = UCase(txtEspecie_Transportador.Text)
End Sub

Private Sub txtIcms_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtIcms_LostFocus()
    txtIcms.Text = Format(txtIcms.Text, "#,###0.00")
End Sub

Private Sub txtIcms_Substituicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtIcms_Substituicao_LostFocus()
    txtIcms_Substituicao.Text = Format(txtIcms_Substituicao.Text, "#,###0.00")
End Sub

Private Sub txtInscricao_estadual_LostFocus()
    txtInscricao_Estadual.Text = UCase(txtInscricao_Estadual.Text)
End Sub

Private Sub txtMarca_Transportador_LostFocus()
    txtMarca_Transportador.Text = UCase(txtMarca_Transportador.Text)
End Sub

Private Sub txtNatureza_Transporte_Change()
    dtcNatureza_Transporte.BoundText = txtNatureza_Transporte.Text
    If IsNumeric(txtNatureza_Transporte.Text) = False Then
       txtNatureza_Transporte.Text = Empty
       Exit Sub
    End If
End Sub

Private Sub txtNatureza_Transporte_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNatureza_Transporte_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNome_Transportador_LostFocus()
    txtNome_Transportador.Text = UCase(txtNome_Transportador.Text)
End Sub

Private Sub txtNumero_NF_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_NF_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumero_Transportador_LostFocus()
    txtNumero_Transportador.Text = UCase(txtNumero_Transportador.Text)
End Sub

Private Sub txtOutras_Despesas_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtOutras_Despesas_LostFocus()
    txtOutras_Despesas.Text = Format(txtOutras_Despesas.Text, "#,###0.00")
End Sub

Private Sub txtPeso_Bruto_Transportador_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPeso_Bruto_Transportador_LostFocus()
    txtPeso_Bruto_Transportador.Text = Format(txtPeso_Bruto_Transportador.Text, "#,###0.00")
End Sub

Private Sub txtPeso_Liquido_Transportador_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPeso_Liquido_Transportador_LostFocus()
    txtPeso_Liquido_Transportador.Text = Format(txtPeso_Liquido_Transportador.Text, "#,###0.00")
End Sub

Private Sub txtPlaca_Veiculo_LostFocus()
    txtPlaca_Veiculo.Text = UCase(txtPlaca_Veiculo.Text)
End Sub

Private Sub txtProduto_Change()
    dtcProduto.BoundText = txtProduto.Text
    If IsNumeric(txtProduto.Text) = False Then
       txtProduto.Text = Empty
       Exit Sub
    End If
End Sub

Private Sub txtProduto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
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
End Sub

Private Sub txtSerie_NF_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtSerie_NF_LostFocus()
    txtSerie_NF.Text = UCase(txtSerie_NF.Text)
End Sub

Private Sub txtUf_Veiculo_LostFocus()
    txtUf_Veiculo.Text = UCase(txtUf_Veiculo)
End Sub

Private Sub txtValor_Frete_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Frete_LostFocus()
    txtValor_Frete.Text = Format(txtValor_Frete.Text, "#,###0.00")
End Sub

Private Sub txtValor_Frete_Transportador_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Frete_Transportador_LostFocus()
    txtValor_Frete_Transportador.Text = Format(txtValor_Frete_Transportador.Text, "#,###0.00")
End Sub

Private Sub txtValor_Icms_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Icms_LostFocus()
    txtValor_Icms.Text = Format(txtValor_Icms.Text, "#,###0.00")
End Sub

Private Sub txtValor_Icms_Substituicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Icms_Substituicao_LostFocus()
    txtValor_Icms_Substituicao.Text = Format(txtValor_Icms_Substituicao.Text, "#,###0.00")
End Sub

Private Sub txtValor_IPI_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_IPI_LostFocus()
    txtValor_IPI.Text = Format(txtValor_IPI.Text, "#,###0.00")
End Sub

Private Sub txtValor_Seguro_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Seguro_LostFocus()
    txtValor_Seguro.Text = Format(txtValor_Seguro.Text, "#,###0.00")
End Sub

Private Sub txtValor_Total_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Total_LostFocus()
    txtValor_Total.Text = Format(txtValor_Total.Text, "#,###0.00")
End Sub

Private Sub txtValor_Total_Nota_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Total_Nota_LostFocus()
    txtValor_Total_Nota.Text = Format(txtValor_Total_Nota.Text, "#,###0.00")
End Sub

Private Sub txtValor_Total_Produtos_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Total_Produtos_LostFocus()
    txtValor_Total_Produtos.Text = Format(txtValor_Total_Produtos.Text, "#,###0.00")
End Sub

Private Sub txtValor_Unitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Unitario_LostFocus()
    If txtQuantidade.Text <> Empty And txtValor_Unitario.Text <> Empty Then
       txtValor_Total.Text = CDbl(txtQuantidade.Text) * CDbl(txtValor_Unitario.Text)
    End If
    txtValor_Unitario.Text = Format(txtValor_Unitario.Text, "#,###0.000")
    txtValor_Total.Text = Format(txtValor_Total.Text, "#,###0.00")
    End Sub

Private Function MontaCombo()
    
    cbbFrete_Conta.Clear
    cbbFrete_Conta.AddItem ("1-Emitente")
    cbbFrete_Conta.AddItem ("2-Destinatário")
    
    cbbCfop.Clear
    cbbCfop.AddItem ("Dentro Estado")
    cbbCfop.AddItem ("Fora Estado")
    
    cbbUnidade.Clear
    cbbUnidade.AddItem ("LT")
    cbbUnidade.AddItem ("KG")
    cbbUnidade.AddItem ("UN")
    
End Function

Private Function MontaDataCombo()

    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente_Destinatario, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_Correios_TBCidade_otica,DFNome_TBCidade_otica FROM TBCidade_otica"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_Correios_TBCidade_otica", "DFNome_TBCidade_otica", dtcCidade_Destinatario, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_Correios_TBCidade_otica,DFNome_TBCidade_otica FROM TBCidade_otica"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_Correios_TBCidade_otica", "DFNome_TBCidade_otica", dtcCidade_Transportador, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBNatureza_transporte,DFDescricao_TBNatureza_transporte FROM TBNatureza_transporte"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBNatureza_transporte", "DFDescricao_TBNatureza_transporte", dtcNatureza_Transporte, strSql, "BDRetaguarda", "Otica", Me

    strSql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me
     
End Function

