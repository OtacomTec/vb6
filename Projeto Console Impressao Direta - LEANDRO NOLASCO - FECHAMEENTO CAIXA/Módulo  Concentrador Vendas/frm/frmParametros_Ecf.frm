VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmParametros_Ecf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parâmetros ECF"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
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
   ScaleHeight     =   5700
   ScaleWidth      =   8280
   Begin TabDlg.SSTab sstParametros_Ecf 
      Height          =   5355
      Left            =   0
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   330
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   9446
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
      TabCaption(0)   =   "&Geral"
      TabPicture(0)   =   "frmParametros_Ecf.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtSenha"
      Tab(0).Control(1)=   "txtSerie_proximo_cupom"
      Tab(0).Control(2)=   "txtProximo_cupom"
      Tab(0).Control(3)=   "txtEndereco_ip"
      Tab(0).Control(4)=   "txtProxima_Serie_Orcamento_Balcao"
      Tab(0).Control(5)=   "txtProximo_Orcamento_Balcao"
      Tab(0).Control(6)=   "txtProduto_Associado_Desconto"
      Tab(0).Control(7)=   "txtPercentual_Taxa"
      Tab(0).Control(8)=   "txtProduto_Associado_Taxa"
      Tab(0).Control(9)=   "txtCodigo"
      Tab(0).Control(10)=   "cbbTipo"
      Tab(0).Control(11)=   "cbbNumeros_Decimais"
      Tab(0).Control(12)=   "cbbDesconto"
      Tab(0).Control(13)=   "cbbCodigo_Inicial"
      Tab(0).Control(14)=   "cbbPreco_Peso"
      Tab(0).Control(15)=   "dtcProduto_Associado_Taxa"
      Tab(0).Control(16)=   "dtcProduto_Associado_Desconto"
      Tab(0).Control(17)=   "cbbControla_Vendedor"
      Tab(0).Control(18)=   "dtcEmpresa"
      Tab(0).Control(19)=   "cbbIntegracao"
      Tab(0).Control(20)=   "cbbAtualizacao"
      Tab(0).Control(21)=   "cbbComissao"
      Tab(0).Control(22)=   "cbbPerfil"
      Tab(0).Control(23)=   "Label27"
      Tab(0).Control(24)=   "Label23"
      Tab(0).Control(25)=   "Label22"
      Tab(0).Control(26)=   "Label21"
      Tab(0).Control(27)=   "Label20"
      Tab(0).Control(28)=   "Label19"
      Tab(0).Control(29)=   "Label18"
      Tab(0).Control(30)=   "Label17"
      Tab(0).Control(31)=   "Label16"
      Tab(0).Control(32)=   "Label15"
      Tab(0).Control(33)=   "Label14"
      Tab(0).Control(34)=   "Label11"
      Tab(0).Control(35)=   "Label10"
      Tab(0).Control(36)=   "Label9"
      Tab(0).Control(37)=   "Label8"
      Tab(0).Control(38)=   "Label7"
      Tab(0).Control(39)=   "Label4"
      Tab(0).Control(40)=   "Label5"
      Tab(0).Control(41)=   "Label3"
      Tab(0).Control(42)=   "Label1"
      Tab(0).Control(43)=   "Label2"
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "&Finalizadoras"
      TabPicture(1)   =   "frmParametros_Ecf.frx":179E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label24"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label25"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label26"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label28"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label29"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "dtcCancelamento"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "dtcX"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "dtcFinalizadora_Dia"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "dtcFinalizadora_Operador"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "dtcFinalizadora_Cartao"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "dtcFinalizadora_Sangria"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "dtcFinalizadora_Abertura"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtFinalizadora_Cartao"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtFinalizadora_Sangria"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtFinalizadora_Abertura"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtFinalizadora_Dia"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtFinalizadora_Operador"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtX"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtCancelamento"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "&Listagem"
      TabPicture(2)   =   "frmParametros_Ecf.frx":17BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdConsulta"
      Tab(2).Control(1)=   "cmdRefresh"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtConsulta"
      Tab(2).Control(3)=   "HfgParametros_Ecf"
      Tab(2).Control(4)=   "cbbCampos"
      Tab(2).Control(5)=   "cbbConsulta"
      Tab(2).Control(6)=   "Label6"
      Tab(2).ControlCount=   7
      Begin VB.TextBox txtCancelamento 
         Height          =   360
         Left            =   120
         TabIndex        =   41
         Top             =   4770
         Width           =   2055
      End
      Begin VB.TextBox txtX 
         Height          =   360
         Left            =   120
         TabIndex        =   38
         Top             =   4110
         Width           =   2055
      End
      Begin VB.TextBox txtFinalizadora_Operador 
         Height          =   360
         Left            =   120
         TabIndex        =   33
         Top             =   2790
         Width           =   2055
      End
      Begin VB.TextBox txtFinalizadora_Dia 
         Height          =   360
         Left            =   120
         TabIndex        =   35
         Top             =   3450
         Width           =   2055
      End
      Begin VB.TextBox txtFinalizadora_Abertura 
         Height          =   360
         Left            =   120
         TabIndex        =   27
         Top             =   780
         Width           =   2055
      End
      Begin VB.TextBox txtFinalizadora_Sangria 
         Height          =   360
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtFinalizadora_Cartao 
         Height          =   360
         Left            =   120
         TabIndex        =   31
         Top             =   2100
         Width           =   2055
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   -67650
         Picture         =   "frmParametros_Ecf.frx":17D6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   -67260
         Picture         =   "frmParametros_Ecf.frx":34D0
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72450
         TabIndex        =   1
         Top             =   780
         Width           =   4725
      End
      Begin VB.TextBox txtSenha 
         Height          =   360
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   26
         Top             =   4800
         Width           =   1155
      End
      Begin VB.TextBox txtSerie_proximo_cupom 
         Height          =   375
         Left            =   -74880
         MaxLength       =   3
         TabIndex        =   19
         Top             =   4110
         Width           =   2055
      End
      Begin VB.TextBox txtProximo_cupom 
         Height          =   360
         Left            =   -68100
         MaxLength       =   6
         TabIndex        =   18
         Top             =   3420
         Width           =   1215
      End
      Begin VB.TextBox txtEndereco_ip 
         Height          =   360
         Left            =   -69540
         MaxLength       =   15
         TabIndex        =   25
         ToolTipText     =   "Endereço IP do Servidor"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtProxima_Serie_Orcamento_Balcao 
         Height          =   375
         Left            =   -70680
         TabIndex        =   21
         Top             =   4110
         Width           =   2055
      End
      Begin VB.TextBox txtProximo_Orcamento_Balcao 
         Height          =   375
         Left            =   -72780
         TabIndex        =   20
         Top             =   4110
         Width           =   2055
      End
      Begin VB.TextBox txtProduto_Associado_Desconto 
         Height          =   360
         Left            =   -74880
         TabIndex        =   16
         Top             =   3420
         Width           =   2055
      End
      Begin VB.TextBox txtPercentual_Taxa 
         Height          =   360
         Left            =   -68130
         TabIndex        =   15
         Top             =   2760
         Width           =   1245
      End
      Begin VB.TextBox txtProduto_Associado_Taxa 
         Height          =   360
         Left            =   -74880
         TabIndex        =   13
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtCodigo 
         Height          =   360
         Left            =   -74880
         TabIndex        =   5
         Top             =   1440
         Width           =   2055
      End
      Begin AutoCompletar.CbCompleta cbbTipo 
         Height          =   360
         Left            =   -72780
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
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
         Left            =   -70800
         TabIndex        =   7
         ToolTipText     =   "Número de Casas Decimais"
         Top             =   1440
         Width           =   1905
         _ExtentX        =   3360
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
         Left            =   -72780
         TabIndex        =   10
         ToolTipText     =   "por Valor/ por Porcentagem"
         Top             =   2100
         Width           =   1935
         _ExtentX        =   3413
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
         Left            =   -68850
         TabIndex        =   8
         ToolTipText     =   "Primeiro Número do Código de Barras"
         Top             =   1440
         Width           =   1965
         _ExtentX        =   3466
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
         Left            =   -74880
         TabIndex        =   9
         Top             =   2100
         Width           =   2055
         _ExtentX        =   3625
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
         Left            =   -72780
         TabIndex        =   14
         Top             =   2760
         Width           =   4605
         _ExtentX        =   8123
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
         Left            =   -72780
         TabIndex        =   17
         Top             =   3420
         Width           =   4635
         _ExtentX        =   8176
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
         Left            =   -70800
         TabIndex        =   11
         ToolTipText     =   "Indicar o Vendedor ao final da Venda"
         Top             =   2100
         Width           =   1905
         _ExtentX        =   3360
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
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   360
         Left            =   -74880
         TabIndex        =   4
         Top             =   780
         Width           =   7995
         _ExtentX        =   14102
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
      Begin AutoCompletar.CbCompleta cbbIntegracao 
         Height          =   360
         Left            =   -74880
         TabIndex        =   23
         ToolTipText     =   "Gravação online dos Pontos de Venda no Banco Retaguarda"
         Top             =   4800
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
      Begin AutoCompletar.CbCompleta cbbAtualizacao 
         Height          =   360
         Left            =   -72420
         TabIndex        =   24
         ToolTipText     =   "Atualização online dos Preços nos Pontos de Vendas"
         Top             =   4800
         Width           =   2835
         _ExtentX        =   5001
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
      Begin AutoCompletar.CbCompleta cbbComissao 
         Height          =   360
         Left            =   -68850
         TabIndex        =   12
         ToolTipText     =   "Comissiona Vendedor"
         Top             =   2100
         Width           =   1965
         _ExtentX        =   3466
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgParametros_Ecf 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   3
         Top             =   1230
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   7011
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   -74880
         TabIndex        =   0
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
      Begin AutoCompletar.CbCompleta cbbConsulta 
         Height          =   360
         Left            =   -72450
         TabIndex        =   67
         Top             =   780
         Width           =   2235
         _ExtentX        =   3942
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
         Left            =   2220
         TabIndex        =   28
         Top             =   780
         Width           =   5895
         _ExtentX        =   10398
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
         Left            =   2220
         TabIndex        =   30
         Top             =   1440
         Width           =   5895
         _ExtentX        =   10398
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
      Begin MSDataListLib.DataCombo dtcFinalizadora_Cartao 
         Height          =   360
         Left            =   2220
         TabIndex        =   32
         Top             =   2100
         Width           =   5895
         _ExtentX        =   10398
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
      Begin MSDataListLib.DataCombo dtcFinalizadora_Operador 
         Height          =   360
         Left            =   2220
         TabIndex        =   34
         Top             =   2790
         Width           =   5895
         _ExtentX        =   10398
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
      Begin MSDataListLib.DataCombo dtcFinalizadora_Dia 
         Height          =   360
         Left            =   2220
         TabIndex        =   36
         Top             =   3450
         Width           =   5895
         _ExtentX        =   10398
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
      Begin AutoCompletar.CbCompleta cbbPerfil 
         Height          =   360
         Left            =   -68580
         TabIndex        =   22
         Top             =   4110
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSDataListLib.DataCombo dtcX 
         Height          =   360
         Left            =   2220
         TabIndex        =   39
         Top             =   4110
         Width           =   5895
         _ExtentX        =   10398
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
      Begin MSDataListLib.DataCombo dtcCancelamento 
         Height          =   360
         Left            =   2220
         TabIndex        =   42
         Top             =   4770
         Width           =   5895
         _ExtentX        =   10398
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
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Cancelamento"
         Height          =   240
         Left            =   120
         TabIndex        =   40
         Top             =   4530
         Width           =   1215
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Leitura X"
         Height          =   240
         Left            =   120
         TabIndex        =   37
         Top             =   3870
         Width           =   765
      End
      Begin VB.Label Label27 
         Caption         =   "Perfil Varejo"
         Height          =   240
         Left            =   -68580
         TabIndex        =   73
         Top             =   3870
         Width           =   1095
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Fechamento Operador"
         Height          =   240
         Left            =   120
         TabIndex        =   72
         Top             =   2550
         Width           =   1920
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Fechamento Dia"
         Height          =   240
         Left            =   120
         TabIndex        =   71
         Top             =   3210
         Width           =   1380
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Abertura"
         Height          =   240
         Left            =   120
         TabIndex        =   70
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Sangria"
         Height          =   240
         Left            =   120
         TabIndex        =   69
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Cartão Afinidade"
         Height          =   240
         Left            =   120
         TabIndex        =   68
         Top             =   1860
         Width           =   1425
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   66
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label23 
         Caption         =   "Senha"
         Height          =   240
         Left            =   -68040
         TabIndex        =   63
         Top             =   4560
         Width           =   765
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Comissão Vendedor"
         Height          =   240
         Left            =   -68850
         TabIndex        =   62
         Top             =   1860
         Width           =   1710
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Próx. Série Cupom"
         Height          =   240
         Left            =   -74880
         TabIndex        =   61
         Top             =   3870
         Width           =   1605
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Próx. Cupom"
         Height          =   240
         Left            =   -68100
         TabIndex        =   60
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "IP Concentrador"
         Height          =   240
         Left            =   -69540
         TabIndex        =   59
         Top             =   4560
         Width           =   1545
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Atualiza Pr. Online Retag. - PDV"
         Height          =   240
         Left            =   -72420
         TabIndex        =   58
         Top             =   4560
         Width           =   2730
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Integra Online PDV - Retag."
         Height          =   240
         Left            =   -74880
         TabIndex        =   57
         Top             =   4560
         Width           =   2370
      End
      Begin VB.Label Label16 
         Caption         =   "Próx. Série Orç. Balcão"
         Height          =   240
         Left            =   -70680
         TabIndex        =   64
         Top             =   3870
         Width           =   2025
      End
      Begin VB.Label Label15 
         Caption         =   "Próx. Orç. Balcão"
         Height          =   240
         Left            =   -72780
         TabIndex        =   56
         Top             =   3870
         Width           =   1755
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   240
         Left            =   -74880
         TabIndex        =   55
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Controla Vendedor"
         Height          =   240
         Left            =   -70800
         TabIndex        =   54
         Top             =   1860
         Width           =   1605
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Produto Associado Desc."
         Height          =   240
         Left            =   -74880
         TabIndex        =   53
         Top             =   3180
         Width           =   2100
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "% Taxa"
         Height          =   240
         Left            =   -68130
         TabIndex        =   52
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Produto Associado Taxa"
         Height          =   240
         Left            =   -74880
         TabIndex        =   51
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cód. Parâmetro ECF"
         Height          =   240
         Left            =   -74880
         TabIndex        =   50
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Inicial Balança"
         Height          =   240
         Left            =   -68850
         TabIndex        =   49
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Preço Peso Balança"
         Height          =   240
         Left            =   -74850
         TabIndex        =   48
         Top             =   1860
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Desconto"
         Height          =   240
         Left            =   -72780
         TabIndex        =   47
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Decimais"
         Height          =   240
         Left            =   -70800
         TabIndex        =   46
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Quantidade"
         Height          =   240
         Left            =   -72780
         TabIndex        =   44
         Top             =   1200
         Width           =   1515
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
      TabIndex        =   45
      Top             =   0
      Width           =   8280
      _ExtentX        =   14605
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
      Left            =   8640
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
            Picture         =   "frmParametros_Ecf.frx":4512
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":482C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":4B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":4EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":527A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":5594
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Ecf.frx":58AE
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
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Cadastro de Parâmetros Ecf                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rafael Gomes                                                   '
' Data de Criação........: 14/01/2005                                                     '
' Desenvolvedor..........: Leandro Nolasco Ferreira                                       '
' Data última manutenção.: 14/08/2006                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strTamanho As String
Dim strNomes As String
Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Dim booAlterar As Boolean
Public strSQL As String
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
Dim log As New DLLSystemManager.log

Private Sub Imprimir()
    On Error GoTo Erro
    'Tratamento de erro
    If strSQL = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Sub
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Relatorio_Parametros_Ecf.Show
    
    Unload frmAguarde
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Sub
End Sub

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty
    cbbConsulta.Text = Empty
    
    If cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       cbbConsulta.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Integração Online" Or cbbCampos.Text = "Atualização Preço" Or cbbCampos.Text = "Comissão Vendedor" Or cbbCampos.Text = "Controla Vendedor" Then
        txtConsulta.Visible = False
        cbbConsulta.Visible = True
        cbbConsulta.SetFocus
    Else
       cbbConsulta.Visible = False
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cbbPerfil_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub cbbTipo_GotFocus()
sstParametros_Ecf.Tab = 0
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub dtcEmpresa_LostFocus()
    dtcEmpresa.Enabled = False
End Sub

Private Sub dtcFinalizadora_Abertura_GotFocus()
    If txtFinalizadora_Abertura.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora_Abertura)
    End If
    sstParametros_Ecf.Tab = 1
End Sub

Private Sub dtcFinalizadora_Abertura_LostFocus()
    txtFinalizadora_Abertura.Text = dtcFinalizadora_Abertura.BoundText
    If IsNumeric(txtFinalizadora_Abertura.Text) = False Or dtcFinalizadora_Abertura.Text = Empty Then txtFinalizadora_Abertura.Text = Empty: Exit Sub
End Sub

Private Sub dtcFinalizadora_Dia_GotFocus()
    If txtFinalizadora_Dia.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora_Dia)
    End If
End Sub

Private Sub dtcFinalizadora_Dia_LostFocus()
    txtFinalizadora_Dia.Text = dtcFinalizadora_Dia.BoundText
    If IsNumeric(txtFinalizadora_Dia.Text) = False Or dtcFinalizadora_Dia.Text = Empty Then txtFinalizadora_Dia.Text = Empty: Exit Sub
End Sub

Private Sub dtcFinalizadora_Operador_GotFocus()
    If txtFinalizadora_Operador.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora_Operador)
    End If
End Sub

Private Sub dtcFinalizadora_Operador_LostFocus()
    txtFinalizadora_Operador.Text = dtcFinalizadora_Operador.BoundText
    If IsNumeric(txtFinalizadora_Operador.Text) = False Or dtcFinalizadora_Operador.Text = Empty Then txtFinalizadora_Operador.Text = Empty: Exit Sub
End Sub

Private Sub dtcFinalizadora_Cartao_GotFocus()
    If txtFinalizadora_Cartao.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora_Cartao)
    End If
End Sub

Private Sub dtcFinalizadora_Cartao_LostFocus()
    txtFinalizadora_Cartao.Text = dtcFinalizadora_Cartao.BoundText
    If IsNumeric(txtFinalizadora_Cartao.Text) = False Or dtcFinalizadora_Cartao.Text = Empty Then txtFinalizadora_Cartao.Text = Empty: Exit Sub
End Sub

Private Sub dtcFinalizadora_Sangria_GotFocus()
    If txtFinalizadora_Sangria.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora_Sangria)
    End If
End Sub

Private Sub dtcFinalizadora_Sangria_LostFocus()
    txtFinalizadora_Sangria.Text = dtcFinalizadora_Sangria.BoundText
    If IsNumeric(txtFinalizadora_Sangria.Text) = False Or dtcFinalizadora_Sangria.Text = Empty Then txtFinalizadora_Sangria.Text = Empty: Exit Sub
End Sub

Private Sub dtcProduto_Associado_Desconto_GotFocus()
    If txtProduto_Associado_Desconto.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcProduto_Associado_Desconto)
    End If
End Sub

Private Sub dtcProduto_Associado_Desconto_LostFocus()
    txtProduto_Associado_Desconto.Text = dtcProduto_Associado_Desconto.BoundText
    If IsNumeric(txtProduto_Associado_Desconto.Text) = False Or dtcProduto_Associado_Desconto.Text = Empty Then txtProduto_Associado_Desconto.Text = Empty: Exit Sub
End Sub

Private Sub dtcProduto_Associado_Taxa_GotFocus()
    If txtProduto_Associado_Taxa.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcProduto_Associado_Taxa)
    End If
End Sub

Private Sub dtcProduto_Associado_Taxa_LostFocus()
    txtProduto_Associado_Taxa.Text = dtcProduto_Associado_Taxa.BoundText
    If IsNumeric(txtProduto_Associado_Taxa.Text) = False Or dtcProduto_Associado_Taxa.Text = Empty Then txtProduto_Associado_Taxa.Text = Empty: Exit Sub
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
    If KeyCode = "113" And booAlterar = False Then
        Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
    End If
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
    log.Programa = "Cadastro de Parâmetros ECF"
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
    
    log.Descricao = "Inicializando o cadastro de Parâmetros ECF"
    'Gravando o log
    log.Gravar_log "Otica", Me
        
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    Call Reposicao
    
    sstParametros_Ecf.TabEnabled(0) = False
    sstParametros_Ecf.TabEnabled(1) = False
    sstParametros_Ecf.Tab = 2
      
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
    
    If frmIntegracao.Visible = True Then
        Unload frmIntegracao
    End If
    
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
      'Integração
       If booIntegra_Portal = True Then
          tlbBotoes.Buttons.Item(9).Enabled = True
       End If
        
      frmAguarde.Show
      DoEvents
       
      txtCodigo.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 1))
      cbbTipo.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 2))
      cbbNumeros_Decimais.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 3))
      cbbCodigo_Inicial.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 4))
      cbbPreco_Peso.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 5))
      cbbDesconto.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 6))
      cbbControla_Vendedor.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 7))
      cbbComissao.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 8))
      txtProduto_Associado_Taxa.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 9))
      txtPercentual_Taxa.Text = Format(HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 10)), "#,###0.00")
      txtProduto_Associado_Desconto.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 11))
      txtFinalizadora_Abertura.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 12))
      txtFinalizadora_Sangria.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 13))
      txtX.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 14))
      txtFinalizadora_Cartao.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 15))
      
      txtFinalizadora_Operador.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 16))
      txtFinalizadora_Dia.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 17))
      
      txtProximo_cupom.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 18))
      txtSerie_proximo_cupom.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 19))
      txtProximo_Orcamento_Balcao = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 20))
      txtProxima_Serie_Orcamento_Balcao = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 21))
      cbbPerfil.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 22))
      
      cbbIntegracao.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 23))
      cbbAtualizacao.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 24))
      txtEndereco_ip.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 25))
      txtSenha.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 26))
      txtCancelamento.Text = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 27))
      dtcEmpresa.BoundText = HfgParametros_Ecf.TextArray((HfgParametros_Ecf.Row * HfgParametros_Ecf.Cols + HfgParametros_Ecf.Col + 28))
     
      txtCodigo.Enabled = False
      booAlterar = True
      txtConsulta.Text = Empty
      sstParametros_Ecf.TabEnabled(0) = True
      sstParametros_Ecf.TabEnabled(1) = True
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
       If frmIntegracao.Visible = True Then
           Unload frmIntegracao
       End If
       txtFinalizadora_Abertura.SetFocus
    ElseIf sstParametros_Ecf.Tab = 2 Then
        If strCombo <> Empty And strCombo <> "Todos" Then
           If txtConsulta.Visible = True Then
              cbbCampos.Text = strCombo
              txtConsulta.SetFocus
           Else
              cbbCampos.Text = strCombo
              cbbConsulta.SetFocus
           End If
        ElseIf strCombo = "Todos" Then
           HfgParametros_Ecf.Row = 1
           HfgParametros_Ecf.Col = 0
           HfgParametros_Ecf.SetFocus
        End If
    End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
    End Select
End Sub

Private Sub Gravar()
    On Error GoTo Erro
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim strTipo As String
    Dim strDesconto As String
    Dim strPreco_Peso As String
    Dim strControla_Vendedor As String
    Dim strIntegracao As String
    Dim strAtualizacao As String
    Dim strComissao As String
    Dim strPerfil As String
    
    If txtCodigo.Text = Empty Then
       MsgBox "O campo Código Parâmetro ECF não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtCodigo.SetFocus
       Exit Sub
    ElseIf dtcEmpresa.Text = Empty Then
       MsgBox "O campo Empresa do Parâmetro ECF não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       Exit Sub
    End If
    
    If cbbTipo.Text = "Inteiro" Then: strTipo = "I": Else: strTipo = "F"
    If cbbDesconto.Text = "$ - Dinheiro" Then: strDesconto = "$": Else: strDesconto = "%"
    If cbbPreco_Peso.Text = "0 - Preço" Then: strPreco_Peso = 0: Else: strPreco_Peso = 1
    If cbbControla_Vendedor.Text = "Sim" Then: strControla_Vendedor = 1: Else: strControla_Vendedor = 0
    If cbbIntegracao.Text = "Sim" Then: strIntegracao = 1: Else: strIntegracao = 0
    If cbbAtualizacao.Text = "Sim" Then: strAtualizacao = 1: Else: strAtualizacao = 0
    If cbbComissao.Text = "Sim" Then: strComissao = 1: Else: strComissao = 0
    If cbbPerfil.Text = "Auto-Atendimento" Then
       strPerfil = 1
    ElseIf cbbPerfil.Text = "Posto Gasolina" Then
       strPerfil = 2
    Else
       strPerfil = 0
    End If

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
               "DFProximo_serie_orcamento_balcao_TBParametros_ecf," & _
               "DFIntegracao_online_pdv_retaguarda_TBParametros_ecf," & _
               "DFAtualizacao_preco_online_retaguarda_pdv_TBParametros_ecf," & _
               "DFEndereco_ip_concentrador_TBParametros_ecf," & _
               "DFProximo_cupom_TBParametros_ecf," & _
               "DFProximo_serie_cupom_TBParametros_ecf," & _
               "DFComissao_vendedor_TBParametros_ecf," & _
               "DFSenha_seguranca_TBparametros_ecf," & _
               "DFFinalizadora_cartao_afinidade_TBParametros_ecf," & _
               "DFFinalizadora_fechamento_operador_TBParametros_ecf,"
               
    strCampo = strCampo & "DFFinalizadora_fechamento_dia_TBParametros_ecf,DFPerfil_varejo_TBParametros_ecf,DFFinalizadora_X_TBParametros_ecf," & _
                          "DFFinalizadora_cancelamento_TBParametros_ecf,DFData_alteracao_TBParametros_ecf,DFIntegrado_filiais_TBParametros_ecf"
               
    If booIntegra_Portal = True Then
       strCampo = strCampo & ",DFIntegrado_portal_TBParametros_ecf"
    End If
                          
    strValores = "'" & txtCodigo.Text & "'," & _
                 "'" & strTipo & "'," & _
                 "'" & cbbNumeros_Decimais.Text & "'," & _
                 "'" & strDesconto & "'," & _
                 "'" & cbbCodigo_Inicial.Text & "'," & _
                 "'" & strPreco_Peso & "'," & _
                 "'" & txtProduto_Associado_Taxa.Text & "'," & _
                 " " & Funcoes_Gerais.Grava_Moeda(txtPercentual_Taxa.Text) & "," & _
                 "'" & txtProduto_Associado_Desconto.Text & "'," & _
                 "'" & strControla_Vendedor & "'," & _
                 "'" & txtFinalizadora_Abertura.Text & "'," & _
                 "'" & txtFinalizadora_Sangria.Text & "'," & _
                 " " & dtcEmpresa.BoundText & ", " & _
                 "'" & txtProximo_Orcamento_Balcao.Text & "'," & _
                 "'" & txtProxima_Serie_Orcamento_Balcao.Text & "'," & _
                 "'" & strIntegracao & "'," & _
                 "'" & strAtualizacao & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtEndereco_ip.Text) & "'," & _
                 "'" & txtProximo_cupom.Text & "'," & _
                 "'" & txtSerie_proximo_cupom.Text & "'," & _
                 "'" & strComissao & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtSenha.Text) & "'," & _
                 "'" & txtFinalizadora_Cartao.Text & "'," & _
                 "'" & txtFinalizadora_Operador.Text & "',"
                 
    strValores = strValores & "'" & txtFinalizadora_Dia.Text & "'," & strPerfil & ",'" & txtX.Text & "'," & _
                              "'" & txtCancelamento.Text & "', '" & Format(Date, "YYYYMMDD") & "',0"
                              
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
    End If
                        
    If booAlterar = True Then
       
       log.Evento = "Alterar"
       strSet = "SET DFTipo_quantidade_TBParametros_ecf = '" & strTipo & "'," & _
                "DFNumero_decimais_TBParametros_ecf = '" & cbbNumeros_Decimais.Text & "'," & _
                "DFTipo_desconto_TBParametros_ecf = '" & strDesconto & "'," & _
                "DFCodigo_inicial_peso_variavel_TBParametros_ecf = '" & cbbCodigo_Inicial.Text & "'," & _
                "DFPreco_peso_balanca_TBParametros_ecf = '" & strPreco_Peso & "'," & _
                "DFProduto_associado_taxa_TBParametros_ecf = " & txtProduto_Associado_Taxa.Text & "," & _
                "DFPercentual_taxa_TBParametros_ecf = " & Funcoes_Gerais.Grava_Moeda(txtPercentual_Taxa.Text) & "," & _
                "DFProduto_Desconto_associado_TBParametros_ecf = " & txtProduto_Associado_Desconto.Text & "," & _
                "DFControla_vendedor_TBParametros_ecf = '" & strControla_Vendedor & "'," & _
                "DFFinalizadora_abertura_TBParametros_ecf = '" & txtFinalizadora_Abertura.Text & "'," & _
                "DFFinalizadora_sangria_TBParametros_ecf = '" & txtFinalizadora_Sangria.Text & "'," & _
                "FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & "," & _
                "DFProximo_orcamento_balcao_TBParametros_ecf = " & txtProximo_Orcamento_Balcao.Text & "," & _
                "DFProximo_serie_orcamento_balcao_TBParametros_ecf = '" & txtProxima_Serie_Orcamento_Balcao.Text & "'," & _
                "DFIntegracao_online_pdv_retaguarda_TBParametros_ecf = '" & strIntegracao & "', " & _
                "DFAtualizacao_preco_online_retaguarda_pdv_TBParametros_ecf = '" & strAtualizacao & "'," & _
                "DFEndereco_ip_concentrador_TBParametros_ecf = '" & Funcoes_Gerais.Grava_String(txtEndereco_ip.Text) & "'," & _
                "DFProximo_cupom_TBParametros_ecf = '" & txtProximo_cupom.Text & "'," & _
                "DFProximo_serie_cupom_TBParametros_ecf = '" & txtSerie_proximo_cupom.Text & "'," & _
                "DFComissao_vendedor_TBParametros_ecf = '" & strComissao & "'," & _
                "DFSenha_seguranca_TBparametros_ecf = '" & Funcoes_Gerais.Grava_String(txtSenha.Text) & "'," & _
                "DFFinalizadora_cartao_afinidade_TBParametros_ecf = '" & txtFinalizadora_Cartao.Text & "'," & _
                "DFFinalizadora_fechamento_operador_TBParametros_ecf = '" & txtFinalizadora_Operador.Text & "'," & _
                "DFFinalizadora_fechamento_dia_TBParametros_ecf = '" & txtFinalizadora_Dia.Text & "',"
                
                
       strSet = strSet & "DFPerfil_varejo_TBParametros_ecf = " & strPerfil & "," & _
                         "DFFinalizadora_X_TBParametros_ecf = '" & txtX.Text & "', " & _
                         "DFFinalizadora_cancelamento_TBParametros_ecf = '" & txtCancelamento.Text & "', " & _
                         "DFData_alteracao_TBParametros_ecf = '" & Format(Date, "YYYYMMDD") & "'," & _
                         "DFIntegrado_filiais_TBParametros_ecf = 0"
                         
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBParametros_ecf = 0"
       End If
                 
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
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       HfgParametros_Ecf.Visible = False
    End If
    
    sstParametros_Ecf.TabEnabled(0) = False
    sstParametros_Ecf.TabEnabled(1) = False
    sstParametros_Ecf.Tab = 2
    
    Exit Sub
    
Erro:

    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Sub
    
End Sub

Private Sub Excluir()
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
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       HfgParametros_Ecf.Visible = False
    End If
        
    sstParametros_Ecf.TabEnabled(0) = False
    sstParametros_Ecf.TabEnabled(1) = False
    sstParametros_Ecf.Tab = 2
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Excluir")
    Exit Sub
End Sub

Private Sub Cancelar()
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
    sstParametros_Ecf.TabEnabled(1) = False
    sstParametros_Ecf.Tab = 2
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Sub
End Sub

Private Sub Novo()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
    
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
    sstParametros_Ecf.TabEnabled(1) = True
    sstParametros_Ecf.Tab = 0
    
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    booAlterar = False
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Sub
End Sub

Private Sub Reposicao()

    On Error GoTo Erro
    
    strNomes = "Código,Tipo Quantidade,Nº Decimais,Código Inicial," & _
               "Preço Peso,Tipo Desconto," & _
               "Controla Vendedor,Comissão Vendedor," & _
               "Prod.Associado Taxa,% Taxa," & _
               "Prod.Associado Desc.,Finalizadora Abertura,Finalizadora Sangria," & _
               "Finalizadora X, Finaliz. Cartão Afinidade," & _
               "Finalizadora Operador,Finalizadora Dia," & _
               "Prox. Cupom,Prox. Série Cupom,Prox. Orç. Balcão," & _
               "Prox. Série Orç. Balcão,Perfil Varejo,Int. Online PDV-Ret.," & _
               "At. Pr. Online Ret.-PDV,IP Concentrador,Senha," & _
               "Cancelamento,Empresa,Nome"
    
    strTamanho = "1000,1500,1500,1500," & _
                 "1500,1600," & _
                 "1800,1800," & _
                 "1900,1200," & _
                 "1900,2000,2000," & _
                 "2000,2000," & _
                 "2000,2000," & _
                 "1200,2000,1600," & _
                 "2200,1700,1900," & _
                 "2100,1600,0," & _
                 "2000,900,3000"
     
    Movimentacoes.Monta_HFlex_Grid HfgParametros_Ecf, strTamanho, strNomes, 29, "Otica", Me
    
    Call Monta_Combo
    Call Monta_DataCombo
              
    Exit Sub
    
Erro:
    Call Erro.Erro(Me, "OTICA", "Reposicao")
    Exit Sub
    Resume
End Sub

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
    cbbCampos.AddItem ("Código Inicial")
    cbbCampos.AddItem ("Preço Peso")
    cbbCampos.AddItem ("Tipo Desconto")
    cbbCampos.AddItem ("Controla Vendedor")
    cbbCampos.AddItem ("Comissão Vendedor")
    cbbCampos.AddItem ("Produto Associado Taxa")
    cbbCampos.AddItem ("% Taxa")
    cbbCampos.AddItem ("Produto Associado Desc.")
    cbbCampos.AddItem ("Finalizadora Abertura")
    cbbCampos.AddItem ("Finalizadora Sangria")
    cbbCampos.AddItem ("Finalizadora X")
    cbbCampos.AddItem ("Finaliz. Cartão Afinidade")
    cbbCampos.AddItem ("Finaliz. Operador")
    cbbCampos.AddItem ("Finaliz. Dia")
    cbbCampos.AddItem ("Prox. Cupom")
    cbbCampos.AddItem ("Prox. Série Cupom")
    cbbCampos.AddItem ("Prox. Orç. Balcão")
    cbbCampos.AddItem ("Prox. Série Orç. Balcão")
    cbbCampos.AddItem ("Perfil Varejo")
    cbbCampos.AddItem ("Integração Online")
    cbbCampos.AddItem ("Atualização Preço")
    cbbCampos.AddItem ("Endereço de IP")
    cbbCampos.AddItem ("Cancelamento")
    cbbCampos.AddItem ("Código da Empresa")
    cbbCampos.AddItem ("Nome Empresa")
    
    cbbTipo.Clear
    cbbTipo.AddItem ("Inteiro")
    cbbTipo.AddItem ("Fração")
    
    cbbPerfil.Clear
    cbbPerfil.AddItem ("Auto-Atendimento")
    cbbPerfil.AddItem ("Posto Gasolina")
    
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
    
    cbbAtualizacao.Clear
    cbbAtualizacao.AddItem ("Sim")
    cbbAtualizacao.AddItem ("Não")
    
    cbbIntegracao.Clear
    cbbIntegracao.AddItem ("Sim")
    cbbIntegracao.AddItem ("Não")
    
    cbbConsulta.Clear
    cbbConsulta.AddItem ("Sim")
    cbbConsulta.AddItem ("Não")
    
    cbbComissao.Clear
    cbbComissao.AddItem ("Sim")
    cbbComissao.AddItem ("Não")
    
End Function

Private Sub Consulta()
    Dim strPercentual As String
    Dim strControla_Vendedor As String
    Dim strIntegracao As String
    Dim strAtualizacao As String
    Dim strComissao As String
    Dim strPerfil As Integer
    
    If cbbCampos.Text = Empty Or cbbCampos.Text <> "Todos" And txtConsulta.Visible = True Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Sub
       End If
    ElseIf cbbCampos.Text = "Integração Online" Or cbbCampos.Text = "Atualização Preço" Or cbbCampos.Text = "Comissão Vendedor" Or cbbCampos.Text = "Controla Vendedor" Then
       If cbbConsulta.Text = Empty Then
          MsgBox "Selecione uma opção para consulta.", vbCritical, "Only Tech"
          cbbConsulta.SetFocus
          Exit Sub
       End If
    End If
    
'    If cbbcampos.Text = "% Taxa" Then
'       strPercentual = Format(txtConsulta.Text, "#,###.00")
'    End If
    
    If cbbCampos.Text = "Controla Vendedor" Then
       If cbbConsulta.Text = "Não" Then
          strControla_Vendedor = 0
       Else
          strControla_Vendedor = 1
       End If
    ElseIf cbbCampos.Text = "Integração Online" Then
       If cbbConsulta.Text = "Sim" Then
          strIntegracao = "1"
       Else
          strIntegracao = "0"
       End If
    ElseIf cbbCampos.Text = "Perfil Varejo" Then
       If cbbConsulta.Text = "Auto-Atendimento" Then
          strPerfil = "1"
       ElseIf cbbConsulta.Text = "Posto Gasolina" Or cbbConsulta.Text = "POSTO GASOLINA" Then
          strPerfil = "2"
       End If
    ElseIf cbbCampos.Text = "Atualização Preço" Then
       If cbbConsulta.Text = "Sim" Then
          strAtualizacao = "1"
       Else
          strAtualizacao = "0"
       End If
    ElseIf cbbCampos.Text = "Comissão Vendedor" Then
       If cbbConsulta.Text = "Sim" Then
          strComissao = "1"
       Else
          strComissao = "0"
       End If
    End If
    
    strSQL = "SELECT TBParametros_ecf.PKCodigo_TBParametros_ecf," & _
             "TBParametros_ecf.DFTipo_quantidade_TBParametros_ecf," & _
             "TBParametros_ecf.DFNumero_decimais_TBParametros_ecf," & _
             "TBParametros_ecf.DFCodigo_inicial_peso_variavel_TBParametros_ecf," & _
             "TBParametros_ecf.DFPreco_peso_balanca_TBParametros_ecf," & _
             "TBParametros_ecf.DFTipo_desconto_TBParametros_ecf," & _
             "TBParametros_ecf.DFControla_vendedor_TBParametros_ecf," & _
             "TBParametros_ecf.DFComissao_vendedor_TBParametros_ecf," & _
             "TBParametros_ecf.DFProduto_associado_taxa_TBParametros_ecf," & _
             "TBParametros_ecf.DFPercentual_taxa_TBParametros_ecf," & _
             "TBParametros_ecf.DFProduto_Desconto_associado_TBParametros_ecf," & _
             "TBParametros_ecf.DFFinalizadora_abertura_TBParametros_ecf," & _
             "TBParametros_ecf.DFFinalizadora_sangria_TBParametros_ecf," & _
             "TBParametros_ecf.DFFinalizadora_X_TBParametros_ecf," & _
             "TBParametros_ecf.DFFinalizadora_cartao_afinidade_TBParametros_ecf," & _
             "TBParametros_ecf.DFFinalizadora_fechamento_operador_TBParametros_ecf," & _
             "TBParametros_ecf.DFFinalizadora_fechamento_dia_TBParametros_ecf,"
             
             
    strSQL = strSQL + "TBParametros_ecf.DFProximo_cupom_TBParametros_ecf," & _
             "TBParametros_ecf.DFProximo_serie_cupom_TBParametros_ecf, " & _
             "TBParametros_ecf.DFProximo_orcamento_balcao_TBParametros_ecf," & _
             "TBParametros_ecf.DFProximo_serie_orcamento_balcao_TBParametros_ecf," & _
             "TBParametros_ecf.DFPerfil_varejo_TBParametros_ecf," & _
             "TBParametros_ecf.DFIntegracao_online_pdv_retaguarda_TBParametros_ecf," & _
             "TBParametros_ecf.DFAtualizacao_preco_online_retaguarda_pdv_TBParametros_ecf," & _
             "TBParametros_ecf.DFEndereco_ip_concentrador_TBParametros_ecf," & _
             "TBParametros_ecf.DFSenha_seguranca_TBparametros_ecf," & _
             "TBparametros_ecf.DFFinalizadora_cancelamento_TBParametros_ecf, " & _
             "TBParametros_ecf.FKCodigo_TBEmpresa," & _
             "TBEmpresa.DFRazao_Social_TBEmpresa " & _
             "FROM TBParametros_ecf " & _
             "INNER JOIN TBEmpresa ON TBParametros_ecf.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa "
                     
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código" Then
          strSQL = strSQL & " WHERE convert(nvarchar,PKCodigo_TBParametros_ecf) = " & txtConsulta.Text & ""
       ElseIf cbbCampos.Text = "Tipo Quantidade" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFTipo_quantidade_TBParametros_ecf) LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Nº Decimais" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFNumero_decimais_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Código Inicial" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFCodigo_inicial_peso_variavel_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Preço Peso" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFPreco_peso_balanca_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Tipo Desconto" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFTipo_desconto_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Controla Vendedor" Then
          strSQL = strSQL & " WHERE convert(bit,DFControla_vendedor_TBParametros_ecf) = '" & strControla_Vendedor & "'"
       ElseIf cbbCampos.Text = "Comissão Vendedor" Then
          strSQL = strSQL & " WHERE convert(bit,DFComissao_vendedor_TBParametros_ecf) = '" & strComissao & "'"
       ElseIf cbbCampos.Text = "Produto Associado Taxa" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFProduto_associado_taxa_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "% Taxa" Then
          txtConsulta.Text = Format(txtConsulta.Text, "#,###0.00")
          strSQL = strSQL & " WHERE DFPercentual_taxa_TBParametros_ecf = " & Funcoes_Gerais.Grava_Moeda(txtConsulta) & " "
       ElseIf cbbCampos.Text = "Produto Associado Desc." Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFProduto_Desconto_associado_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Finalizadora Abertura" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFFinalizadora_abertura_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Finalizadora Sangria" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFFinalizadora_sangria_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Finalizadora X" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFFinalizadora_X_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Cancelamento" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFFinalizadora_cancelamento_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Prox. Cupom" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFProximo_cupom_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Prox. Série Cupom" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFProximo_serie_cupom_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Prox. Orç. Balcão" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFProximo_orcamento_balcao_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Prox. Série Orç. Balcão" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFProximo_serie_orcamento_balcao_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Integração Online" Then
          strSQL = strSQL & " WHERE convert(bit,DFIntegracao_online_pdv_retaguarda_TBParametros_ecf) = '" & strIntegracao & "'"
       ElseIf cbbCampos.Text = "Atualização Preço" Then
          strSQL = strSQL & " WHERE convert(bit,DFAtualizacao_preco_online_retaguarda_pdv_TBParametros_ecf) = '" & strAtualizacao & "'"
       ElseIf cbbCampos.Text = "Endereço de IP" Then
          strSQL = strSQL & " WHERE convert(nvarchar,DFEndereco_ip_concentrador_TBParametros_ecf) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Código da Empresa" Then
          strSQL = strSQL & " WHERE convert(nvarchar,FKCodigo_TBEmpresa) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Nome Empresa" Then
          strSQL = strSQL & " WHERE DFRazao_Social_TBEmpresa LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Finaliz. Cartão Afinidade" Then
          strSQL = strSQL & " WHERE DFFinalizadora_cartao_afinidade_TBParametros_ecf = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Finaliz. Operador" Then
          strSQL = strSQL & " WHERE DFFinalizadora_fechamento_operador_TBParametros_ecf = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Finaliz. Dia" Then
          strSQL = strSQL & " WHERE DFFinalizadora_fechamento_dia_TBParametros_ecf = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Perfil Varejo" Then
          strSQL = strSQL & " WHERE DFPerfil_varejo_TBParametros_ecf = '" & strPerfil & "' "
       End If
    End If

    frmAguarde.Show
    DoEvents
    
    Movimentacoes.Movimenta_HFlex_Grid strSQL, HfgParametros_Ecf, strTamanho, strNomes, "BDRetaguarda", "Otica", Me, "S"
    
    HfgParametros_Ecf.Col = 1
    HfgParametros_Ecf.Row = 1
    If HfgParametros_Ecf.Rows > 1 And HfgParametros_Ecf.Text <> Empty Then
       For I = 1 To HfgParametros_Ecf.Rows - 1
           HfgParametros_Ecf.Row = I
           HfgParametros_Ecf.Col = 2
           If HfgParametros_Ecf.Text = "I" Then
              HfgParametros_Ecf.Text = "Inteiro"
           Else
              HfgParametros_Ecf.Text = "Fração"
           End If
           HfgParametros_Ecf.Col = 6
           If HfgParametros_Ecf.Text = "$" Then
              HfgParametros_Ecf.Text = "$ - Dinheiro"
           Else
              HfgParametros_Ecf.Text = "% - Percentagem"
           End If
           HfgParametros_Ecf.Col = 5
           If HfgParametros_Ecf.Text = "Não" Then
              HfgParametros_Ecf.Text = "0 - Preço"
           Else
              HfgParametros_Ecf.Text = "1 - Peso"
           End If
           HfgParametros_Ecf.Col = 21
           If HfgParametros_Ecf.Text = "1" Then
              HfgParametros_Ecf.Text = "Auto-Atendimento"
           ElseIf HfgParametros_Ecf.Text = "2" Then
              HfgParametros_Ecf.Text = "Posto Gasolina"
           ElseIf HfgParametros_Ecf.Text = "0" Then
              HfgParametros_Ecf.Text = ""
           End If
       Next I
    Else
       HfgParametros_Ecf.Rows = 2
       Movimentacoes.Monta_HFlex_Grid HfgParametros_Ecf, strTamanho, strNomes, 29, "Otica", Me
    End If
    
    Unload frmAguarde
    
    HfgParametros_Ecf.Col = 0
    HfgParametros_Ecf.Row = 1
    
    HfgParametros_Ecf.Refresh
    HfgParametros_Ecf.SetFocus
    
End Sub

Private Function Monta_DataCombo()
    
    strSQL = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT * FROM TBProduto"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto_Associado_Desconto, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT * FROM TBProduto"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto_Associado_Desconto, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT * FROM TBProduto"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto_Associado_Taxa, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_Abertura, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_Sangria, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_Cartao, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_Operador, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora_Dia, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcX, strSQL, "BDRetaguarda", "Otica", Me

    strSQL = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcCancelamento, strSQL, "BDRetaguarda", "Otica", Me

End Function

Private Sub txtEndereco_ip_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtEndereco_ip_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtFinalizadora_Abertura_Change()
    dtcFinalizadora_Abertura.BoundText = txtFinalizadora_Abertura.Text
    If IsNumeric(txtFinalizadora_Abertura.Text) = False Then txtFinalizadora_Abertura.Text = Empty: Exit Sub
End Sub

Private Sub txtFinalizadora_Abertura_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
        sstParametros_Ecf.Tab = 1
End Sub

Private Sub txtFinalizadora_Cartao_Change()
    dtcFinalizadora_Cartao.BoundText = txtFinalizadora_Cartao.Text
    If IsNumeric(txtFinalizadora_Cartao.Text) = False Then txtFinalizadora_Cartao.Text = Empty: Exit Sub
End Sub

Private Sub txtFinalizadora_Cartao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFinalizadora_Dia_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFinalizadora_Operador_Change()
    dtcFinalizadora_Operador.BoundText = txtFinalizadora_Operador.Text
    If IsNumeric(txtFinalizadora_Operador.Text) = False Then txtFinalizadora_Operador.Text = Empty: Exit Sub
End Sub

Private Sub txtFinalizadora_Dia_Change()
    dtcFinalizadora_Dia.BoundText = txtFinalizadora_Dia.Text
    If IsNumeric(txtFinalizadora_Dia.Text) = False Then txtFinalizadora_Dia.Text = Empty: Exit Sub
End Sub

Private Sub txtFinalizadora_Operador_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFinalizadora_Sangria_Change()
    dtcFinalizadora_Sangria.BoundText = txtFinalizadora_Sangria.Text
    If IsNumeric(txtFinalizadora_Sangria.Text) = False Then txtFinalizadora_Sangria.Text = Empty: Exit Sub
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
    If IsNumeric(txtProduto_Associado_Desconto.Text) = False Then txtProduto_Associado_Desconto.Text = Empty: Exit Sub
End Sub

Private Sub txtProduto_Associado_Desconto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtProduto_Associado_Taxa_Change()
    dtcProduto_Associado_Taxa.BoundText = txtProduto_Associado_Taxa.Text
    If IsNumeric(txtProduto_Associado_Taxa.Text) = False Then txtProduto_Associado_Taxa.Text = Empty: Exit Sub
End Sub

Private Sub txtProduto_Associado_Taxa_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtProxima_Serie_Orcamento_Balcao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtProximo_cupom_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtProximo_cupom_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtProximo_Orcamento_Balcao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtProximo_Orcamento_Balcao_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtSenha_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtSerie_proximo_cupom_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Limpa_Combos()
    cbbTipo.Text = Empty
    cbbNumeros_Decimais.Text = Empty
    cbbCodigo_Inicial.Text = Empty
    cbbPreco_Peso.Text = Empty
    cbbDesconto.Text = Empty
    cbbControla_Vendedor.Text = Empty
    cbbComissao.Text = Empty
    cbbIntegracao.Text = Empty
    cbbAtualizacao.Text = Empty
    cbbPerfil.Text = Empty
End Function

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo_TBParametros_ecf", txtCodigo.Text, "DFIntegrado_filiais_TBParametros_ecf", "TBParametros_ecf", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBParametros_ecf", Me.Top, Me.Left, Me.width, Me.Height, "Parâmetros ECF")
    
End Function

Private Sub txtX_Change()
    dtcX.BoundText = txtX.Text
    If IsNumeric(txtX.Text) = False Then txtX.Text = Empty: Exit Sub
End Sub

Private Sub txtX_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub dtcX_GotFocus()
    If txtX.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcX)
    End If
End Sub

Private Sub dtcX_LostFocus()
    txtX.Text = dtcX.BoundText
    If IsNumeric(txtX.Text) = False Or dtcX.Text = Empty Then txtX.Text = Empty: Exit Sub
End Sub

Private Sub txtCancelamento_Change()
    dtcCancelamento.BoundText = txtCancelamento.Text
    If IsNumeric(txtCancelamento.Text) = False Then txtCancelamento.Text = Empty: Exit Sub
End Sub

Private Sub txtCancelamento_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub dtcCancelamento_GotFocus()
    If txtCancelamento.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCancelamento)
    End If
End Sub

Private Sub dtcCancelamento_LostFocus()
    txtCancelamento.Text = dtcCancelamento.BoundText
    If IsNumeric(txtCancelamento.Text) = False Or dtcCancelamento.Text = Empty Then txtCancelamento.Text = Empty: Exit Sub
End Sub
