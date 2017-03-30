VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNota_Saida_Cupom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Fiscal de Saída Cupom"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNota_Saida_Cupom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11100
   Begin TabDlg.SSTab sstNota 
      Height          =   6585
      Left            =   0
      TabIndex        =   23
      Top             =   330
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   11615
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabPicture(0)   =   "frmNota_Saida_Cupom.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCancelada"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label33"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label31"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblCliente_Fornecedor"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label18"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label34"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtcCupom"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DTPicker2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "DTPicker1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbTipo_operacao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dtcPlano_pagamento"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dtcCliente"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dtcEmpresa"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtSerie"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtNumero_Nota"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame5"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame4"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Frame3"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtPlano_pagamento"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtCliente"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Frame1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdPesquisar"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdInformacoes_cupons"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "chkPrevisao"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdConsulta_cupom"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtCupom"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdIncluir_Item"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "Listagem"
      TabPicture(1)   =   "frmNota_Saida_Cupom.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtConsulta"
      Tab(1).Control(1)=   "cmdRefresh"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdConsulta"
      Tab(1).Control(3)=   "dtpFinal"
      Tab(1).Control(4)=   "dtpInicial"
      Tab(1).Control(5)=   "adgNota"
      Tab(1).Control(6)=   "cbbCampos"
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(8)=   "lblA"
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdIncluir_Item 
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
         Left            =   6780
         Picture         =   "frmNota_Saida_Cupom.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Adicionar item"
         Top             =   2475
         Width           =   375
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
         Height          =   315
         Left            =   120
         MaxLength       =   40
         TabIndex        =   15
         ToolTipText     =   "Código do Produto"
         Top             =   2475
         Width           =   1000
      End
      Begin VB.CommandButton cmdConsulta_cupom 
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
         Left            =   6390
         Picture         =   "frmNota_Saida_Cupom.frx":1904
         Style           =   1  'Graphical
         TabIndex        =   81
         TabStop         =   0   'False
         ToolTipText     =   "Consulta detalhada do produto "
         Top             =   2475
         Width           =   375
      End
      Begin VB.CheckBox chkPrevisao 
         Caption         =   "&Previsão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10230
         TabIndex        =   79
         Top             =   6060
         Width           =   225
      End
      Begin VB.CommandButton cmdInformacoes_cupons 
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
         Left            =   6780
         Picture         =   "frmNota_Saida_Cupom.frx":1C8E
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         ToolTipText     =   "Projeção de Pagamentos"
         Top             =   1935
         Width           =   375
      End
      Begin VB.CommandButton cmdPesquisar 
         Height          =   315
         Left            =   6390
         Picture         =   "frmNota_Saida_Cupom.frx":2018
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Consultar"
         Top             =   1935
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Itens da Nota (F10)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1995
         Left            =   120
         TabIndex        =   75
         Top             =   2820
         Width           =   10815
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgItem_NOTA 
            Height          =   1545
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   10545
            _ExtentX        =   18600
            _ExtentY        =   2725
            _Version        =   393216
            FixedCols       =   0
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -71520
         TabIndex        =   1
         Top             =   780
         Width           =   6645
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   -64440
         Picture         =   "frmNota_Saida_Cupom.frx":3D12
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   3705
         MaxLength       =   40
         TabIndex        =   8
         ToolTipText     =   "Código do Cliente"
         Top             =   1350
         Width           =   975
      End
      Begin VB.TextBox txtPlano_pagamento 
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
         MaxLength       =   40
         TabIndex        =   10
         ToolTipText     =   "Código da condição de pagamento"
         Top             =   1935
         Width           =   1000
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resumo Financeiro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1635
         Left            =   2940
         TabIndex        =   45
         Top             =   4830
         Width           =   5385
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Total Praticado.:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2970
            TabIndex        =   60
            Top             =   990
            Width           =   1200
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Total Tabela....:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2970
            TabIndex        =   59
            Top             =   660
            Width           =   1185
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Total Produtos.......:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   330
            Width           =   1530
         End
         Begin VB.Label lblTotal_tabela 
            Alignment       =   1  'Right Justify
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   4200
            TabIndex        =   57
            Top             =   660
            Width           =   1050
         End
         Begin VB.Label lblTotal_praticado 
            Alignment       =   1  'Right Justify
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   4200
            TabIndex        =   56
            Top             =   990
            Width           =   1050
         End
         Begin VB.Label lblTotal_Produtos 
            Alignment       =   1  'Right Justify
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1740
            TabIndex        =   55
            ToolTipText     =   "Total bruto dos itens"
            Top             =   330
            Width           =   1050
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Desc.Esp.+Indeni...:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   990
            Width           =   1530
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "IPI + Desp.Aces.....:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   53
            ToolTipText     =   "Total de IPI  + Total de despesas  acessórios"
            Top             =   660
            Width           =   1545
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Total Nota............:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   1530
         End
         Begin VB.Label lblDescontos_especiais 
            Alignment       =   1  'Right Justify
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1740
            TabIndex        =   51
            ToolTipText     =   "Total de descontos especias"
            Top             =   990
            Width           =   1050
         End
         Begin VB.Label lblIPI 
            Alignment       =   1  'Right Justify
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1740
            TabIndex        =   50
            Top             =   660
            Width           =   1050
         End
         Begin VB.Label lblTotal_Nota 
            Alignment       =   1  'Right Justify
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1740
            TabIndex        =   49
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Line Line1 
            BorderStyle     =   6  'Inside Solid
            DrawMode        =   2  'Blackness
            X1              =   120
            X2              =   2800
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label lblDivergencia 
            Alignment       =   1  'Right Justify
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   4200
            TabIndex        =   48
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Divergência.....:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2970
            TabIndex        =   47
            Top             =   1320
            Width           =   1200
         End
         Begin VB.Line Line2 
            X1              =   2970
            X2              =   5200
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label lblTotal_Itens 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   3360
            TabIndex        =   46
            Top             =   330
            Width           =   45
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Resumo Logistíco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   8370
         TabIndex        =   40
         Top             =   4830
         Width           =   2565
         Begin VB.Label lblPeso_Bruto 
            Alignment       =   1  'Right Justify
            Caption         =   "0.000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1140
            TabIndex        =   44
            Top             =   660
            Width           =   1305
         End
         Begin VB.Label lblPeso_Liquido 
            Alignment       =   1  'Right Justify
            Caption         =   "0.000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1140
            TabIndex        =   43
            Top             =   330
            Width           =   1305
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Peso Liquido:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   330
            Width           =   945
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Peso Bruto..:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   660
            Width           =   960
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Outros (F12)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1635
         Left            =   120
         TabIndex        =   35
         Top             =   4830
         Width           =   2775
         Begin VB.TextBox txtIndenizacao 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1530
            MaxLength       =   40
            TabIndex        =   18
            ToolTipText     =   "Indenização no pedido"
            Top             =   210
            Width           =   1095
         End
         Begin VB.TextBox txtIPI 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1530
            MaxLength       =   40
            TabIndex        =   21
            ToolTipText     =   "Valor do IPI"
            Top             =   1220
            Width           =   1095
         End
         Begin VB.TextBox txtDespesas_acessorios 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1530
            MaxLength       =   40
            TabIndex        =   20
            ToolTipText     =   "Despesas acessórios do pedido"
            Top             =   890
            Width           =   1095
         End
         Begin VB.TextBox txtDesconto_especial 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1530
            MaxLength       =   40
            TabIndex        =   19
            ToolTipText     =   "Desconto especial no pedido"
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Desconto Especial:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   660
            Width           =   1350
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Indenização.......:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Desp. Acessórios.:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   990
            Width           =   1350
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "IPI..................:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   1350
         End
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   -64830
         Picture         =   "frmNota_Saida_Cupom.frx":4D54
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Informações Adicionais (F11)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2115
         Left            =   7260
         TabIndex        =   24
         Top             =   690
         Width           =   3675
         Begin MSComCtl2.DTPicker dtpData_Saida_nota 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "Data de vigência final"
            Top             =   1680
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   58589185
            CurrentDate     =   37957
         End
         Begin MSComCtl2.DTPicker dtpData_Emissao_nota 
            Height          =   315
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Data de vigência inicial"
            Top             =   1110
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   58589185
            CurrentDate     =   37957
         End
         Begin MSComCtl2.DTPicker dtpHora_Emissao 
            Height          =   315
            Left            =   1410
            TabIndex        =   27
            Top             =   1680
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   58589186
            CurrentDate     =   37957
         End
         Begin VB.Label lblDigitador 
            Caption         =   "lblDigitador"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1440
            TabIndex        =   34
            ToolTipText     =   "Digitador"
            Top             =   330
            Width           =   2115
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Digitador.........:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   330
            Width           =   1245
         End
         Begin VB.Label lblFaturista 
            Caption         =   "lblFaturista"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1440
            TabIndex        =   32
            ToolTipText     =   "Digitador"
            Top             =   645
            Width           =   2115
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Faturista.........:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   650
            Width           =   1245
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Data Emissão"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   900
            Width           =   960
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Data Saída"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1470
            Width           =   780
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Hora Emissão"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1410
            TabIndex        =   28
            Top             =   1470
            Width           =   960
         End
      End
      Begin VB.TextBox txtNumero_Nota 
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
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Código do Cliente"
         Top             =   1350
         Width           =   1425
      End
      Begin VB.TextBox txtSerie 
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
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   6
         ToolTipText     =   "Código do Cliente"
         Top             =   1350
         Width           =   465
      End
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Empresa Padrão"
         Top             =   750
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
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
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   360
         Left            =   -68010
         TabIndex        =   62
         Top             =   780
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   58589185
         CurrentDate     =   37923
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   360
         Left            =   -71520
         TabIndex        =   63
         Top             =   780
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   58589185
         CurrentDate     =   37923
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid adgNota 
         Height          =   5235
         Left            =   -74880
         TabIndex        =   3
         Top             =   1200
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   9234
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
         Width           =   3315
         _ExtentX        =   5847
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
         Height          =   315
         Left            =   4740
         TabIndex        =   9
         ToolTipText     =   "Cliente"
         Top             =   1350
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
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
      Begin MSDataListLib.DataCombo dtcPlano_pagamento 
         Height          =   315
         Left            =   1170
         TabIndex        =   11
         ToolTipText     =   "Condição de pagamento"
         Top             =   1935
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
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
      Begin AutoCompletar.CbCompleta cmbTipo_operacao 
         Height          =   315
         Left            =   2100
         TabIndex        =   7
         Top             =   1350
         Width           =   1560
         _ExtentX        =   2752
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
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5100
         TabIndex        =   13
         ToolTipText     =   "Data de vigência final"
         Top             =   1935
         Width           =   1245
         _ExtentX        =   2196
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
         Format          =   58589185
         CurrentDate     =   37957
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3690
         TabIndex        =   12
         ToolTipText     =   "Data de vigência inicial"
         Top             =   1935
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   58589185
         CurrentDate     =   37957
      End
      Begin MSDataListLib.DataCombo dtcCupom 
         Height          =   315
         Left            =   1170
         TabIndex        =   16
         ToolTipText     =   "Produto"
         Top             =   2475
         Width           =   5175
         _ExtentX        =   9128
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cupom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   82
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "&Previsão...........:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8820
         TabIndex        =   80
         Top             =   6060
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4950
         TabIndex        =   77
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Período"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3690
         TabIndex        =   76
         Top             =   1725
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   71
         Top             =   540
         Width           =   435
      End
      Begin VB.Label lblCliente_Fornecedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3705
         TabIndex        =   70
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plano Pagamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   1725
         Width           =   1245
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "até"
         Height          =   240
         Left            =   -68970
         TabIndex        =   68
         Top             =   900
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Operação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2100
         TabIndex        =   67
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° Nota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   1140
         Width           =   570
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1590
         TabIndex        =   65
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lblCancelada 
         AutoSize        =   -1  'True
         Caption         =   "CANCELADA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   6000
         TabIndex        =   64
         Top             =   780
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12570
      Top             =   1200
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
            Picture         =   "frmNota_Saida_Cupom.frx":6A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNota_Saida_Cupom.frx":6D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNota_Saida_Cupom.frx":7082
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNota_Saida_Cupom.frx":741C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNota_Saida_Cupom.frx":77B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNota_Saida_Cupom.frx":7AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNota_Saida_Cupom.frx":7DEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
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
            Description     =   "Consulta Detalhada"
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consulta Detalhada - CTRL + T"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgApoio_Vencimento_Pedido 
      Height          =   3615
      Left            =   11340
      TabIndex        =   74
      Top             =   330
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
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
      _Band(0).Cols   =   1
   End
End
Attribute VB_Name = "frmNota_Saida_Cupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Only Tech                                                                               '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Sistema................: Only Tech                                                      '
'' Módulo.................: Concentrador de Vendas                                         '
'' Objetivo...............: Nota Fiscal de Saída Cupom                                     '
'' Data de Criação........: 10/06/2006                                                     '
'' Equipe Responsável.....: Only Tech Solutions                                            '
'' Desenvolvedor..........: Jones Sá Peixoto                                               '
'' Data Criação...........: 10/06/2006                                                     '
'' Desenvolvedor..........: Jones Sá Peixoto                                               '
'' Data última manutenção.:                                                                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim log As New DLLSystemManager.log
'Dim strSql As String
'Dim strTamanhos As String
'Dim strNomes As String
'Dim intContador As Integer
''RECORDSETS
'Dim rstAplicacao As New ADODB.Recordset
'Dim rstPlano_Pagamento As New ADODB.Recordset
'Dim rstCliente As New ADODB.Recordset
''CONEXOES
'Dim cnGravacao As New DLLConexao_Sistema.Conexao
''VARIAVEIS DE GRAVACAO
'Dim intIDVendedor As String
'Dim intIDPlano As String
'Dim intCodigo_Transportadora As Integer
'Dim lngIDCfop As Integer
'Dim strUnidade As String
'Dim lngIDProduto As String
'Dim strCST1 As String
'Dim strCST2 As String
''VALORES PARAMETRO
'Dim strValor_Min_IR As String
'Dim strValor_Min_Contribuicao As String
'Option Explicit
'
'Public Sub cbbEmissao_Click()
'
'    'FORMATANDO A TELA DE ACORDO COM O TIPO DE EMISSAO
'    If cbbEmissao.Text = "Lote" Then
'
'       lblCliente.Caption = "Ramo de Atividade"
'       txtCliente.Text = Empty
'       txtPlano_pagamento.Text = Empty
'
'       strSql = "SELECT PKCodigo_TBRamo_atividade,DFDescricao_TBRamo_atividade FROM TBRamo_atividade"
'       Movimentacoes.Movimenta_DataCombo "PKCodigo_TBRamo_atividade", "DFDescricao_TBRamo_atividade", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
'
'       txtPlano_pagamento.Enabled = False
'       dtcPlano_pagamento.Enabled = False
'
'       lblObservacao.Top = 1530
'       txtObservacao.Top = 1740
'
'       Me.Height = 2510
'       frItens_Nota.Visible = False
'
'    Else
'
'       lblCliente.Caption = "Cliente"
'       txtCliente.Text = Empty
'
'       strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente " & _
'                "FROM TBCliente " & _
'                "INNER JOIN TBContrato_cliente ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente  " & _
'                "WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
'
'       txtPlano_pagamento.Enabled = True
'       dtcPlano_pagamento.Enabled = True
'
'       lblObservacao.Top = 6390
'       txtObservacao.Top = 6600
'
'       Me.Height = 7395
'       frItens_Nota.Visible = True
'
'    End If
'
'End Sub
'
'Private Sub cmdIncluir_Item_Click()
'
'    Dim strIndice As String
'    Dim strTotal As String
'    Dim strQuantidade As String
'
'    If txtProduto.Text = Empty Then
'       MsgBox "Produto inválido. Verifique.", vbInformation, "Only Tech"
'       txtProduto.SetFocus
'       Exit Sub
'    ElseIf txtQuantidade_produto.Text = Empty Or txtQuantidade_produto.Text = "0,00" Then
'       MsgBox "Quantidade inválida. Verifique.", vbInformation, "Only Tech"
'       txtQuantidade_produto.SetFocus
'       Exit Sub
'    ElseIf txtUnidade.Text = Empty Then
'       MsgBox "Unidade de Produto inválida. Verifique o cadastro de produto.", vbInformation, "Only Tech"
'       txtProduto.SetFocus
'       Exit Sub
'    ElseIf txtPreco_unitario.Text = Empty Then
'       MsgBox "Preço unitário inválido. Verifique o cadastro de produto.", vbInformation, "Only Tech"
'       txtPreco_unitario.SetFocus
'       Exit Sub
'    End If
'
'    'VERIFICACAO QUANTO AO NUMERO DE ITENS INCLUIDOS NA ORDEM DE SERVICO, O MAXIMO SETADO VIA RELATORIO É 4
'    If hfgProduto.Rows >= 5 Then
'       MsgBox "Esta Ordem já excedeu o número máximo de Concentrador de Vendass permitidos. Verifique.", vbInformation, "Only Tech"
'       Exit Sub
'    End If
'
'    'Verifica se o produto já foi incluso para o cupom de acerto, se sim, ele é somado ao já incluso
'    intContador = 1
'    Do While intContador <= hfgProduto.Rows - 1
'       hfgProduto.Row = intContador
'       hfgProduto.Col = 1
'       If hfgProduto.Text = txtProduto.Text Then
'          MsgBox "Produto já incluído. Verifique.", vbInformation, "Only Tech"
'          txtQuantidade_produto.Text = Empty
'          txtUnidade.Text = Empty
'          txtPreco_unitario.Text = Empty
'          txtTotal_item.Text = Empty
'          txtProduto.SetFocus
'          txtProduto.SetFocus
'          Exit Sub
'       End If
'       intContador = intContador + 1
'    Loop
'
'    hfgProduto.Row = 1
'    hfgProduto.Col = 0
'    If hfgProduto.Text <> Empty Then
'       strIndice = hfgProduto.Rows
'       hfgProduto.Rows = hfgProduto.Rows + 1
'    Else
'       strIndice = 1
'    End If
'
'    hfgProduto.Row = strIndice
'
'    hfgProduto.Col = 0
'    hfgProduto.ColWidth(0) = 380
'    hfgProduto.Font.Name = "Tahoma"
'    hfgProduto.CellFontSize = 7
'    hfgProduto.CellFontBold = False
'    hfgProduto.CellBackColor = &H80FFFF
'    hfgProduto.Text = strIndice
'
'    hfgProduto.Col = 1
'    hfgProduto.Text = txtProduto.Text
'    hfgProduto.Col = 2
'    hfgProduto.Text = dtcProduto.Text
'    hfgProduto.Col = 3
'    hfgProduto.Text = txtQuantidade_produto.Text
'    hfgProduto.Col = 4
'    hfgProduto.Text = txtUnidade.Text
'    hfgProduto.Col = 5
'    hfgProduto.Text = txtPreco_unitario.Text
'    hfgProduto.Col = 6
'    hfgProduto.Text = txtTotal_item.Text
'
'    Call Calcula_Resumos
'
'    txtProduto.Text = Empty
'    txtQuantidade_produto.Text = Empty
'    txtUnidade.Text = Empty
'    txtPreco_unitario.Text = Empty
'    txtTotal_item.Text = Empty
'
'    txtProduto.SetFocus
'End Sub
'
'Private Sub cmdRemover_Item_Click()
'
'    If hfgProduto.Col <> 0 Or hfgProduto.Text = Empty Then
'       MsgBox "Não há produto selecionado para exclusão. Verifique", vbInformation, "Only Tech"
'       txtProduto.SetFocus
'       Exit Sub
'    End If
'
'    If hfgProduto.Rows <= 2 Then
'       hfgProduto.Clear
'       Movimentacoes.Monta_HFlex_Grid hfgProduto, strTamanhos, strNomes, 6, "OTICA", Me
'    Else
'       hfgProduto.RemoveItem (hfgProduto.Row)
'       intContador = 1
'       hfgProduto.Col = 0
'       Do While intContador <= hfgProduto.Rows - 1
'          hfgProduto.Row = intContador
'          hfgProduto.Text = intContador
'          intContador = intContador + 1
'       Loop
'    End If
'    Call Calcula_Resumos
'End Sub
'
'Private Sub dtcEmpresa_Change()
'    txtProduto.Text = Empty: txtCliente.Text = Empty: txtPlano_pagamento.Text = Empty
'End Sub
'
'Private Sub dtcEmpresa_LostFocus()
'
'    If Not IsNumeric(dtcEmpresa.BoundText) Then dtcEmpresa.Text = Empty
'    If IsNumeric(dtcEmpresa.Text) Then dtcEmpresa.Text = Empty
'
'    If dtcEmpresa.Text <> Empty Then
'       strSql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me
'
'       If cbbEmissao.Text = "Individual" Then
'          strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente " & _
'                   "FROM TBCliente " & _
'                   "INNER JOIN TBContrato_cliente ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente  " & _
'                   "WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'          Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
'       End If
'
'       strSql = "SELECT IXCodigo_TBPlano_Pagamento,DFDescricao_TBPlano_Pagamento FROM TBPlano_Pagamento WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBPlano_Pagamento", "DFDescricao_TBPlano_Pagamento", dtcPlano_pagamento, strSql, "BDRetaguarda", "Otica", Me
'    End If
'
'    dtcEmpresa.Enabled = False
'
'End Sub
'
'Private Sub dtcProduto_GotFocus()
'    If txtProduto.Text = Empty Then
'       Call Movimentacoes.Verifica_DataCombo(dtcProduto.Text)
'    End If
'End Sub
'
'Private Sub dtcProduto_LostFocus()
'
'    txtProduto.Text = dtcProduto.BoundText
'    If txtProduto.Text <> Empty Then
'       strSql = "SELECT DFUnidade_varejo_TBProduto " & _
'                "FROM TBProduto " & _
'                "WHERE TBProduto.IXCodigo_TBProduto = " & txtProduto.Text & " " & _
'                "AND TBProduto.IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " "
'
'       Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'       If rstAplicacao.RecordCount <> 0 Then
'          If IsNull(rstAplicacao.Fields("DFUnidade_varejo_TBProduto")) = False Then
'             txtUnidade.Text = rstAplicacao.Fields("DFUnidade_varejo_TBProduto")
'          Else
'             txtUnidade.Text = Empty
'             txtTotal_item.Text = Empty
'          End If
'       Else
'          txtUnidade.Text = Empty
'          txtTotal_item.Text = Empty
'       End If
'       Set rstAplicacao = Nothing
'    End If
'
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    'Teclas de Atalho da TOOLBAR
'    Select Case Shift
'           Case 2
'                Select Case KeyCode
'                       Case 71: Call Gravar      'CTRL+G
'                       Case 67: Call Cancelar    'CTRL+C
'                       Case 83: Unload Me        'CTRL+S
'                End Select
'    End Select
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    'Habilita a troca de campos pelo ENTER
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub Form_Load()
'
'    On Error GoTo erro
'
'    'Informações Constantes para o log
'    log.Usuario = MDIPrincipal.OCXUsuario.Nome
'    log.Programa = "Nota Fiscal de Saída Cupom"
'    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
'
'    'Informações Variaveis para o log
'    log.Evento = "Load"
'    log.Tipo = 1
'    log.Data = Date
'    log.Hora = Format(Now, "hh:mm:ss")
'
'    If MDIPrincipal.booDesign_time = False Then
'       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
'    End If
'
'    log.Descricao = "Inicializando Nota Fiscal de Saída Cupom"
'    'Gravando o log
'    log.Gravar_log "Otica", Me
'
'    'Montando os datacombo de tela
'    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa "
'    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
'
'    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
'
'    strSql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me
'
'    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente " & _
'             "FROM TBCliente " & _
'             "INNER JOIN TBContrato_cliente ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente  " & _
'             "WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
'
'    strSql = "SELECT IXCodigo_TBPlano_Pagamento,DFDescricao_TBPlano_Pagamento FROM TBPlano_Pagamento WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBPlano_Pagamento", "DFDescricao_TBPlano_Pagamento", dtcPlano_pagamento, strSql, "BDRetaguarda", "Otica", Me
'
'    'MONTANDO GRID DE PRODUTOS
'    strTamanhos = "800,3600,950,350,1200,1100"
'    strNomes = "Código,Produto,Quantidade,UN,Pr. Praticado,Total"
'
'    Monta_HFlex_Grid hfgProduto, strTamanhos, strNomes, 6, "Otica", Me
'
'    cbbEmissao.Clear
'    cbbEmissao.AddItem ("Lote")
'    cbbEmissao.AddItem ("Individual")
'    cbbEmissao.Text = "Lote"
'
'    Call cbbEmissao_Click
'
'    'ABASTECENDO IMPOSTOS
'    strSql = "SELECT DFPercentual_iss_TBParametros_fiscais," & _
'             "DFPercentual_irrf_TBParametros_fiscais," & _
'             "DFPercentual_contribuicao_social_TBParametros_fiscais," & _
'             "DFPercentual_cofins_TBParametros_fiscais," & _
'             "DFPercentual_pis_TBParametros_fiscais," & _
'             "DFValor_minimo_calculo_irrf_TBParametros_fiscais," & _
'             "DFValor_minimo_calculo_contribuicao_TBParametros_fiscais " & _
'             "FROM TBParametros_fiscais " & _
'             "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'    If rstAplicacao.RecordCount <> 0 Then
'       If IsNull(rstAplicacao.Fields("DFPercentual_irrf_TBParametros_fiscais")) = False Then
'          txtImposto_Renda.Text = Format(rstAplicacao.Fields("DFPercentual_irrf_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtImposto_Renda.Text = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFPercentual_iss_TBParametros_fiscais")) = False Then
'          txtIss.Text = Format(rstAplicacao.Fields("DFPercentual_iss_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtIss.Text = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFPercentual_contribuicao_social_TBParametros_fiscais")) = False Then
'          txtContribuicao_Social.Text = Format(rstAplicacao.Fields("DFPercentual_contribuicao_social_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtContribuicao_Social.Text = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFPercentual_cofins_TBParametros_fiscais")) = False Then
'          txtCofins.Text = Format(rstAplicacao.Fields("DFPercentual_cofins_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtCofins.Text = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFPercentual_pis_TBParametros_fiscais")) = False Then
'          txtPis.Text = Format(rstAplicacao.Fields("DFPercentual_pis_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtPis.Text = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFValor_minimo_calculo_irrf_TBParametros_fiscais")) = False Then
'          strValor_Min_IR = Format(rstAplicacao.Fields("DFValor_minimo_calculo_irrf_TBParametros_fiscais"), "#,###0.00")
'       Else
'          strValor_Min_IR = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFValor_minimo_calculo_contribuicao_TBParametros_fiscais")) = False Then
'          strValor_Min_Contribuicao = Format(rstAplicacao.Fields("DFValor_minimo_calculo_contribuicao_TBParametros_fiscais"), "#,###0.00")
'       Else
'          strValor_Min_Contribuicao = "0,00"
'       End If
'    Else
'       txtImposto_Renda.Text = "0,00"
'       txtIss.Text = "0,00"
'       txtCofins.Text = "0,00"
'       txtPis.Text = "0,00"
'       txtContribuicao_Social.Text = "0,00"
'       strValor_Min_IR = "0,00"
'       strValor_Min_Contribuicao = "0,00"
'    End If
'
'    Set rstAplicacao = Nothing
'
'    Exit Sub
'erro:
'    Call erro.erro(Me, "Otica", "Load")
'    Exit Sub
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    On Error GoTo erro
'
'    log.Hora = Format(Now, "hh:mm:ss")
'
'    'Gravando Log
'    log.Gravar_log "Otica", Me
'
'    Exit Sub
'erro:
'    Call erro.erro(Me, "Otica", "Unload")
'    Exit Sub
'End Sub
'
'Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
'    Select Case Button.Index
'           Case 1: Call Gravar
'           Case 2: Call Cancelar
'           Case 4: Unload Me
'    End Select
'End Sub
'
'Private Sub txtCliente_Change()
'    dtcCliente.BoundText = txtCliente.Text
'    If IsNumeric(txtCliente.Text) = False Then txtCliente.Text = Empty: Exit Sub
'End Sub
'
'Private Sub txtCliente_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
'End Sub
'
'Private Sub dtcCliente_GotFocus()
'    If txtCliente.Text = Empty Then
'       Call Movimentacoes.Verifica_DataCombo(dtcCliente.Text)
'    End If
'End Sub
'
'Private Sub dtcCliente_LostFocus()
'    txtCliente.Text = dtcCliente.BoundText
'    If IsNumeric(txtCliente.Text) = False Or dtcCliente.Text = Empty Then
'       txtCliente.Text = Empty: Exit Sub
'    Else
'      If cbbEmissao.Text = "Individual" Then
'
'          strSql = "SELECT IXCodigo_TBPlano_pagamento " & _
'                   "FROM TBCliente " & _
'                   "INNER JOIN TBPlano_Pagamento ON TBCliente.FKId_TBPlano_pagamento = TBPlano_Pagamento.PKId_TBPlano_pagamento " & _
'                   "WHERE IXCodigo_TBCliente = " & txtCliente.Text & " " & _
'                   "AND TBCliente.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'          Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'          If rstAplicacao.RecordCount <> 0 And IsNull(rstAplicacao.Fields("IXCodigo_TBPlano_pagamento")) = False Then
'             txtPlano_pagamento.Text = rstAplicacao.Fields("IXCodigo_TBPlano_pagamento")
'          End If
'          Set rstAplicacao = Nothing
'       End If
'    End If
'End Sub
'
'Private Sub txtCliente_LostFocus()
'    If dtcCliente.Text = Empty Then
'       txtCliente.Text = Empty
'    Else
'       Call dtcCliente_LostFocus
'    End If
'End Sub
'
'Private Sub txtPlano_Pagamento_Change()
'    dtcPlano_pagamento.BoundText = txtPlano_pagamento.Text
'    If IsNumeric(txtPlano_pagamento.Text) = False Then txtPlano_pagamento.Text = Empty: Exit Sub
'End Sub
'
'Private Sub txtPlano_Pagamento_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
'End Sub
'
'Private Sub dtcPlano_Pagamento_GotFocus()
'    If txtPlano_pagamento.Text = Empty Then
'       Call Movimentacoes.Verifica_DataCombo(dtcPlano_pagamento.Text)
'    End If
'End Sub
'
'Private Sub dtcPlano_Pagamento_LostFocus()
'    txtPlano_pagamento.Text = dtcPlano_pagamento.BoundText
'    If IsNumeric(txtPlano_pagamento.Text) = False Or dtcPlano_pagamento.Text = Empty Then txtPlano_pagamento.Text = Empty: Exit Sub
'End Sub
'
'Private Sub txtDesconto_Especial_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
'End Sub
'
'Private Sub txtDesconto_Especial_KeyPress(KeyAscii As Integer)
'    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
'       Exit Sub
'    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
'       KeyAscii = 0
'    End If
'End Sub
'
'Private Sub txtDesconto_Especial_LostFocus()
'    txtDesconto_especial.Text = Format(txtDesconto_especial.Text, "#,###0.00")
'    Call Calcula_Resumos
'End Sub
'
'Private Sub txtObservacao_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
'End Sub
'
'Private Sub txtObservacao_LostFocus()
'    txtObservacao.Text = UCase(txtObservacao.Text)
'End Sub
'
'Private Sub txtPreco_unitario_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
'End Sub
'
'Private Sub txtPreco_unitario_KeyPress(KeyAscii As Integer)
'    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
'       Exit Sub
'    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
'       KeyAscii = 0
'    End If
'End Sub
'
'Private Sub txtPreco_Unitario_LostFocus()
'    txtPreco_unitario.Text = Format(txtPreco_unitario.Text, "#,###0.00")
'    If txtPreco_unitario.Text = Empty Or txtQuantidade_produto.Text = Empty Then
'       txtTotal_item.Text = Empty
'    Else
'       txtTotal_item.Text = Format(CDbl(txtPreco_unitario.Text) * CDbl(txtQuantidade_produto.Text), "#,###0.00")
'    End If
'End Sub
'
'Private Sub txtProduto_Change()
'    dtcProduto.BoundText = txtProduto.Text
'    If IsNumeric(txtProduto.Text) = False Then
'       txtProduto.Text = Empty
'       Exit Sub
'    End If
'End Sub
'
'Private Sub txtProduto_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
'End Sub
'
'Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
'        KeyAscii = 0
'    End If
'End Sub
'
'Private Sub txtProduto_LostFocus()
'    If dtcProduto.Text = Empty Then
'       txtProduto.Text = Empty
'       txtUnidade.Text = Empty
'       txtPreco_unitario.Text = Empty
'       txtTotal_item.Text = Empty
'    End If
'End Sub
'
'Private Sub txtQuantidade_produto_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
'End Sub
'
'Private Sub txtQuantidade_produto_KeyPress(KeyAscii As Integer)
'    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
'       Exit Sub
'    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
'       KeyAscii = 0
'    End If
'End Sub
'
'Private Sub txtQuantidade_produto_LostFocus()
'    txtQuantidade_produto.Text = Format(txtQuantidade_produto.Text, "#,###0.00")
'    If txtPreco_unitario.Text <> Empty And txtQuantidade_produto.Text <> Empty Then
'       txtTotal_item.Text = Format(CDbl(txtQuantidade_produto.Text) * CDbl(txtPreco_unitario.Text), "#,###0.00")
'    Else
'       txtTotal_item.Text = Empty
'    End If
'End Sub
'
'Private Function Gravar()
'
'    If txtCliente.Text = Empty And cbbEmissao.Text = "Individual" Then
'       MsgBox "O Campo código do Cliente não pode ser nulo. Verifique.", vbInformation, "Only Tech"
'       txtCliente.SetFocus
'       Exit Function
'    ElseIf txtPlano_pagamento.Text = Empty And cbbEmissao.Text = "Individual" Then
'       MsgBox "O Campo código do Plano de Pagamento não pode ser nulo. Verifique.", vbInformation, "Only Tech"
'       txtCliente.SetFocus
'       Exit Function
'    End If
'
'    'VENDEDOR
'    strSql = "SELECT PKID_TBVendedor " & _
'             "FROM TBVendedor " & _
'             "WHERE IXCodigo_TBVendedor = 9999 " & _
'             "AND IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'    If rstAplicacao.RecordCount <> 0 Then
'       intIDVendedor = rstAplicacao.Fields("PKID_TBVendedor")
'    Else
'       MsgBox "É necessário que conste no sistema um vendedor de código 9999. A operação está impossibilitada de continuar.", vbInformation, "Only Tech"
'       Set rstAplicacao = Nothing
'       Exit Function
'    End If
'
'    Set rstAplicacao = Nothing
'
'    'TRANSPORTADORA
'    strSql = "SELECT PKCodigo_TBTransportadora " & _
'             "FROM TBTransportadora " & _
'             "WHERE PKCodigo_TBTransportadora = 9999"
'
'    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'    If rstAplicacao.RecordCount <> 0 Then
'       intCodigo_Transportadora = rstAplicacao.Fields("PKCodigo_TBTransportadora")
'    Else
'       MsgBox "É necessário que conste no sistema uma transportadora de código 9999. A operação está impossibilitada de continuar.", vbInformation, "Only Tech"
'       Set rstAplicacao = Nothing
'       Exit Function
'    End If
'
'    Set rstAplicacao = Nothing
'
'
'    Call Grava_Corpo_Nota
'
'End Function
'
'Private Function Cancelar()
'
'    lblTotal_Produtos.Caption = Empty
'    lblTotal_Pedido.Caption = Empty
'    lblDescontos_especiais.Caption = Empty
'    lblImpostos.Caption = Empty
'
'    Call Limpa_TXT(Me)
'
'    hfgProduto.Rows = 2
'    Monta_HFlex_Grid hfgProduto, strTamanhos, strNomes, 6, "Otica", Me
'
'    'ABASTECENDO IMPOSTOS
'    strSql = "SELECT DFPercentual_iss_TBParametros_fiscais," & _
'             "DFPercentual_irrf_TBParametros_fiscais," & _
'             "DFPercentual_contribuicao_social_TBParametros_fiscais," & _
'             "DFPercentual_cofins_TBParametros_fiscais," & _
'             "DFPercentual_pis_TBParametros_fiscais " & _
'             "FROM TBParametros_fiscais " & _
'             "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'    If rstAplicacao.RecordCount <> 0 Then
'       If IsNull(rstAplicacao.Fields("DFPercentual_irrf_TBParametros_fiscais")) = False Then
'          txtImposto_Renda.Text = Format(rstAplicacao.Fields("DFPercentual_irrf_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtImposto_Renda.Text = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFPercentual_iss_TBParametros_fiscais")) = False Then
'          txtIss.Text = Format(rstAplicacao.Fields("DFPercentual_iss_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtIss.Text = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFPercentual_contribuicao_social_TBParametros_fiscais")) = False Then
'          txtContribuicao_Social.Text = Format(rstAplicacao.Fields("DFPercentual_contribuicao_social_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtContribuicao_Social.Text = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFPercentual_cofins_TBParametros_fiscais")) = False Then
'          txtCofins.Text = Format(rstAplicacao.Fields("DFPercentual_cofins_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtCofins.Text = "0,00"
'       End If
'       If IsNull(rstAplicacao.Fields("DFPercentual_pis_TBParametros_fiscais")) = False Then
'          txtPis.Text = Format(rstAplicacao.Fields("DFPercentual_pis_TBParametros_fiscais"), "#,###0.00")
'       Else
'          txtPis.Text = "0,00"
'       End If
'    Else
'       txtImposto_Renda.Text = "0,00"
'       txtIss.Text = "0,00"
'       txtCofins.Text = "0,00"
'       txtPis.Text = "0,00"
'       txtContribuicao_Social.Text = "0,00"
'    End If
'
'    Set rstAplicacao = Nothing
'
'End Function
'
'Private Function Grava_Corpo_Nota()
'    Dim dblIndenizacao As Double
'    Dim dblDesconto_Especial As Double
'    Dim dblTotal_Pedido As Double
'    Dim dblValor_Contrato As Double
'    Dim intEmitente As Integer
'    Dim intCodigo_Tabela_Vigente As Integer
'    Dim intDia_Vencimento As Integer
'    Dim dblImpostos As Double
'    Dim strObservacao As String
'    Dim datVencimento As Date
'
'    frmAguarde.Show
'
'    'BUSCANDO INFORMACOES PERTINENTES PARA GRAVAÇÃO
'
'    'TABELA VIGENTE
'    strSql = "SELECT DFNumero_tabela_vigente_TBParametros_venda " & _
'             "FROM TBParametros_venda " & _
'             "WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'    If rstAplicacao.RecordCount <> 0 Then
'       intCodigo_Tabela_Vigente = rstAplicacao.Fields("DFNumero_tabela_vigente_TBParametros_venda")
'    End If
'
'    Set rstAplicacao = Nothing
'
'    'CFOP PARAMETRO FISCAL
'    strSql = "SELECT PKID_TBCfop " & _
'             "FROM TBParametros_fiscais " & _
'             "INNER JOIN TBCFOP " & _
'             "ON TBParametros_fiscais.DFProximo_cfop_venda_dentro_estado_TBParametros_fiscais = TBCFOP.DFCodigo_TBCfop " & _
'             "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'    If rstAplicacao.RecordCount <> 0 Then
'       lngIDCfop = rstAplicacao.Fields("PKID_TBCfop")
'    End If
'
'    Set rstAplicacao = Nothing
'
'    'INFORMACOES DO CLIENTE E DO CONTRATO
'    strSql = "SELECT PKId_TBCliente,IXCodigo_TBCliente,FKId_TBPlano_pagamento," & _
'             "DFDia_vencimento_TBCliente," & _
'             "DFValor_TBContrato_cliente,DFDescricao_TBPlano_pagamento " & _
'             "FROM TBCliente " & _
'             "INNER JOIN TBContrato_cliente ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente  " & _
'             "INNER JOIN TBPlano_Pagamento ON TBCliente.FKId_TBPlano_pagamento = TBPlano_pagamento.PKId_TBPlano_pagamento " & _
'             "INNER JOIN TBCidade_Otica ON TBCliente.FKID_TBCidade_Otica = TBCidade_Otica.PKID_TBCidade_Otica " & _
'             "WHERE TBCliente.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'    If cbbEmissao.Text = "Individual" Then
'
'       strSql = strSql + " AND IXCodigo_TBCliente = " & txtCliente.Text & ""
'
'       Select_geral strSql, "BDRetaguarda", rstCliente, "Otica", Me
'
'       'PLANO DE PAGAMENTO
'       strSql = "SELECT PKId_TBPlano_pagamento,DFDigita_vencimento_TBPlano_pagamento " & _
'                "FROM TBPlano_pagamento " & _
'                "WHERE IXCodigo_TBPlano_pagamento = " & txtPlano_pagamento.Text & " " & _
'                "AND IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'       Select_geral strSql, "BDRetaguarda", rstPlano_Pagamento, "Otica", Me
'
'       If rstPlano_Pagamento.RecordCount <> 0 Then
'          intIDPlano = rstPlano_Pagamento.Fields("PKId_TBPlano_pagamento")
'       End If
'
'    Else
'
'       If txtCliente.Text <> Empty Then
'          strSql = strSql + " AND TBCliente.FKCodigo_TBRamo_atividade = " & txtCliente.Text & " "
'       End If
'
'       Select_geral strSql, "BDRetaguarda", rstCliente, "Otica", Me
'
'       'ID DO PRODUTO PADRAO
'       strSql = "SELECT FKId_contrato_TBProduto,DFUnidade_venda_TBProduto,DFCst1_TBProduto, " & _
'                "DFCst2_TBProduto " & _
'                "FROM TBParametros_servicos " & _
'                "INNER JOIN TBProduto ON TBParametros_servicos.FKId_contrato_TBProduto = TBProduto.PKId_TBProduto " & _
'                "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'
'       Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'       If rstAplicacao.RecordCount <> 0 Then
'          lngIDProduto = rstAplicacao.Fields("FKId_contrato_TBProduto")
'          strUnidade = rstAplicacao.Fields("DFUnidade_venda_TBProduto")
'          strCST1 = rstAplicacao.Fields("DFCst1_TBProduto")
'          strCST2 = rstAplicacao.Fields("DFCst2_TBProduto")
'       Else
'          Unload frmAguarde
'          MsgBox "Produto Padrão não definido no cadastro de Parâmetros de Concentrador de Vendass. A operação está impossibilitada de continuar.", vbInformation, "Only Tech"
'          Set rstAplicacao = Nothing
'          Set rstPlano_Pagamento = Nothing
'          Set rstCliente = Nothing
'          Exit Function
'       End If
'
'       Set rstAplicacao = Nothing
'
'    End If
'
'    On Error GoTo erro
'
'    'ABRINDO CONEXAO
'    cnGravacao.Initial_Catalog = "BDRetaguarda"
'    cnGravacao.Abrir_conexao "Otica"
'    cnGravacao.CNconexao.BeginTrans
'
'    Do While rstCliente.EOF = False
'
'       '''''''''''CAPTURANDO O DIA DE VENCIMENTO'''''''''''''''
'       intDia_Vencimento = Format(rstCliente.Fields("DFDia_vencimento_TBCliente"), "00")
'
'       If intDia_Vencimento = 0 Then
'          intDia_Vencimento = 15
'       End If
'
'       If intDia_Vencimento <= Format(Now, "DD") Then
'          datVencimento = intDia_Vencimento & "/" & Format(DateAdd("M", 1, Now), "MM/YYYY")
'       Else
'          datVencimento = intDia_Vencimento & "/" & Format(Now, "MM/YYYY")
'       End If
'       '''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'       If cbbEmissao.Text = "Individual" Then
'
'          dblDesconto_Especial = lblDescontos_especiais.Caption
'
'          If txtObservacao.Text = Empty Then
'             strObservacao = "VENC. " & datVencimento
'          Else
'             strObservacao = txtObservacao.Text
'          End If
'
'          dblValor_Contrato = CDbl(lblTotal_Produtos.Caption)
'
'          dblImpostos = Format(lblImpostos.Caption, "#,###0.00")
'
'       Else
'
'          'PLANO DE PAGAMENTO
'          intIDPlano = rstCliente.Fields("FKId_TBPlano_pagamento")
'
'          'MONTANDO OBSERVACAO
'          strObservacao = "VENC. " & datVencimento
'
'          dblDesconto_Especial = 0
'
'          If IsNull(rstCliente.Fields("DFValor_TBContrato_cliente")) = False Then
'             dblValor_Contrato = rstCliente.Fields("DFValor_TBContrato_cliente")
'          Else
'             dblValor_Contrato = 0
'          End If
'
'          dblImpostos = 0
'
'          'Montando o total de Impostos
'          If dblValor_Contrato > CDbl(strValor_Min_IR) Then
'             dblImpostos = Format(CDbl(txtImposto_Renda.Text), "#,###0.00")
'          End If
'
'          If dblValor_Contrato > CDbl(strValor_Min_Contribuicao) Then
'             dblImpostos = Format(dblImpostos + CDbl(txtCofins.Text) + CDbl(txtContribuicao_Social.Text) + CDbl(txtPis.Text), "#,###0.00")
'          End If
'
'          dblImpostos = Format(dblValor_Contrato * CDbl(dblImpostos) / 100, "#,###0.00")
'
'       End If
'
'       'CALCULANDO TOTAL DO PEDIDO
'       dblTotal_Pedido = Format(dblValor_Contrato - dblImpostos - dblDesconto_Especial, "#,###0.00")
'
'       'TIPO EMITENTE
'       intEmitente = 0
'
'       'GRAVANDO CORPO DO PEDIDO
'       strSql = "INSERT INTO TBPedido(FKCodigo_TBEmpresa, " & _
'                "FKCodigo_TBTabela_preco, " & _
'                "FKId_TBVendedor," & _
'                "FKId_TBPlano_pagamento," & _
'                "FKCodigo_TBTransportadora," & _
'                "DFTipo_operacao_TBPedido," & _
'                "DFEmitente_TBPedido," & _
'                "DFTotal_itens_TBPedido," & _
'                "DFTotal_pedido_TBPedido," & _
'                "DFTotal_pedido_tabelaTBPedido," & _
'                "DFDesconto_especial_TBPedido," & _
'                "DFDesconto_indenizacao_TBPedido," & _
'                "DFData_Digitacao_TBPedido," & _
'                "DFUsuario_TBPedido," & _
'                "DFFaturado_TBPedido," & _
'                "DFPrevisao_TBPedido," & _
'                "DFValor_ipi_TBPedido," & _
'                "DFBloqueado_TBPedido," & _
'                "DFDespesas_acessorias_TBPedido,"
'
'        strSql = strSql + "DFTotal_descontos_itens_TBPedido," & _
'                "DFTotal_peso_liquido_TBPedido," & _
'                "DFTotal_peso_bruto_TBPedido," & _
'                "DFTipo_emitente_TBPedido, " & _
'                "DFObservacao_TBPedido,DFBase_calculo_subst_tributaria_TBPedido," & _
'                "DFValor_subst_tributaria_TBPedido,DFValor_Frete_TBPedido,DFTipo_Frete_TBPedido) " & _
'                "VALUES (" & _
'                " " & dtcEmpresa.BoundText & "," & _
'                " " & intCodigo_Tabela_Vigente & "," & _
'                " " & intIDVendedor & "," & _
'                " " & intIDPlano & "," & _
'                " " & intCodigo_Transportadora & "," & _
'                " " & 1 & "," & _
'                " " & rstCliente.Fields("IXCodigo_TBCliente") & " ," & _
'                " " & Funcoes_Gerais.Grava_Moeda(dblValor_Contrato) & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(dblTotal_Pedido) & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(dblValor_Contrato) & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(dblDesconto_Especial) & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(dblImpostos) & "," & _
'                " '" & Format(Now, "YYYYMMDD") & "'," & _
'                " '" & MDIPrincipal.OCXUsuario.Nome & "'," & _
'                " " & 0 & "," & _
'                " " & 0 & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                " " & 0 & ","
'
'        strSql = strSql + " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                " " & intEmitente & "," & _
'                " '" & Funcoes_Gerais.Grava_String(strObservacao) & "'," & _
'                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                " " & Funcoes_Gerais.Grava_Moeda(0) & ",0)"
'
'        'Gravando o corpo do Pedido
'        cnGravacao.CNconexao.Execute strSql
'
'        Call Grava_Itens
'
'        'Gravando o CFO na tabela CFO-PEDIDO
'        strSql = "INSERT INTO TBCfop_pedido(FKId_TBCfop,FKId_TBPedido) " & _
'                 "SELECT " & lngIDCfop & ",MAX(PKID_TBPedido) FROM TBPedido "
'
'        cnGravacao.CNconexao.Execute strSql
'
'        rstCliente.MoveNext
'
'    Loop
'
'    cnGravacao.CNconexao.CommitTrans
'    cnGravacao.Fechar_conexao
'
'    Unload frmAguarde
'
'    MsgBox "" & rstCliente.RecordCount & " Ordem(ns) gravada(s) corretamente."
'
'    Set rstCliente = Nothing
'    Set rstAplicacao = Nothing
'    Set rstPlano_Pagamento = Nothing
'
'    Call Cancelar
'
'    Exit Function
'erro:
'    'TRATAMENTO DE ERRO
'    Unload frmAguarde
'    cnGravacao.CNconexao.RollbackTrans
'    cnGravacao.Fechar_conexao
'    Call erro.erro(Me, "Otica", "Load")
'End Function
'
'Private Function Grava_Itens()
'
'    Dim intRotina As Integer
'    Dim dblValor_Item As Double
'    Dim dblTotal_Item As Double
'    Dim dblQuantidade_Item As Double
'
'    If cbbEmissao.Text = "Individual" Then
'       intRotina = hfgProduto.Rows - 1
'    Else
'       intRotina = 1
'    End If
'
'    intContador = 1
'
'    Do While intContador <= intRotina
'
'       If cbbEmissao.Text = "Individual" Then
'
'          hfgProduto.Row = intContador
'          hfgProduto.Col = 3
'          dblQuantidade_Item = CDbl(hfgProduto.Text)
'          hfgProduto.Col = 4
'          strUnidade = hfgProduto.Text
'          hfgProduto.Col = 5
'          dblValor_Item = CDbl(hfgProduto.Text)
'          hfgProduto.Col = 6
'          dblTotal_Item = CDbl(hfgProduto.Text)
'
'          hfgProduto.Col = 1
'          strSql = "SELECT PKID_TBProduto " & _
'                   "FROM TBProduto " & _
'                   "WHERE IXCodigo_TBProduto = " & hfgProduto.Text & " " & _
'                   "AND IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
'
'          Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'          If rstAplicacao.RecordCount <> 0 Then
'             lngIDProduto = rstAplicacao.Fields("PKID_TBProduto")
'          End If
'
'          Set rstAplicacao = Nothing
'
'       Else
'          dblValor_Item = rstCliente.Fields("DFValor_TBContrato_cliente")
'          dblQuantidade_Item = 1
'          dblTotal_Item = dblValor_Item
'       End If
'
'       strSql = "INSERT INTO TBItens_pedido(" & _
'                "FKId_TBPedido," & _
'                "FKId_TBProduto," & _
'                "FKId_TBCfop," & _
'                "DFCst1_TBItens_pedido," & _
'                "DFCst2_TBItens_pedido," & _
'                "DFQuantidade_TBItens_pedido," & _
'                "DFTipo_preco_TBItens_pedido," & _
'                "DFPreco_tabela_TBItens_pedido," & _
'                "DFPercentual_desconto_TBItens_pedido," & _
'                "DFPreco_praticado_TBItens_pedido," & _
'                "DFValor_total_tabela_TBItens_pedido," & _
'                "DFValor_total_praticado_TBItens_pedido," & _
'                "DFPercentual_icms_TBItens_pedido," & _
'                "DFValor_total_icms_TBItens_pedido," & _
'                "DFUnidade_TBItens_pedido," & _
'                "DFPeso_liquido_TBItens_pedido," & _
'                "DFPeso_bruto_TBItens_pedido," & _
'                "DFQuantidade_baixa_estoque_TBItens_pedido," & _
'                "DFDivisor_baixa_estoque_TBItens_pedido," & _
'                "FKId_TBVendedor,"
'
'        strSql = strSql + "DFValor_total_item_TBItens_pedido,DFBase_calculo_subst_tributaria_TBItens_pedido," & _
'                          "DFValor_subst_tributaria_TBItens_pedido,DFValor_cotacao_dia_TBItens_pedido) " & _
'                          "SELECT " & _
'                          "MAX(PKID_TBPedido)," & _
'                          "" & lngIDProduto & "," & _
'                          "" & lngIDCfop & "," & _
'                          "'" & strCST1 & "'," & _
'                          "'" & strCST2 & "'," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(dblQuantidade_Item) & "," & _
'                          "" & 1 & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(dblValor_Item) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(dblValor_Item) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Item) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Item) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                          "'" & strUnidade & "'," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(dblQuantidade_Item) & "," & _
'                          "" & 1 & "," & _
'                          "" & intIDVendedor & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Item) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(0) & ","
'
'        strSql = strSql + "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
'                          "" & Funcoes_Gerais.Grava_Moeda(0) & " " & _
'                          "FROM TBPedido "
'
'        'Gravando o item do Pedido
'        cnGravacao.CNconexao.Execute strSql
'
'        intContador = intContador + 1
'    Loop
'
'End Function
'
'Private Function Calcula_Resumos()
'
'    Dim dblImpostos As Double
'    Dim dblTotal_Produtos As Double
'
'    hfgProduto.Col = 6
'    intContador = 1
'    Do While intContador <= hfgProduto.Rows - 1
'       hfgProduto.Row = intContador
'       If hfgProduto.Text <> Empty Then
'          dblTotal_Produtos = dblTotal_Produtos + hfgProduto.Text
'       End If
'       intContador = intContador + 1
'    Loop
'
'    If dblTotal_Produtos > CDbl(strValor_Min_IR) Then
'       dblImpostos = CDbl(txtImposto_Renda.Text)
'    End If
'
'    If dblTotal_Produtos > CDbl(strValor_Min_Contribuicao) Then
'       dblImpostos = Format(dblImpostos + CDbl(txtCofins.Text) + CDbl(txtContribuicao_Social.Text) + CDbl(txtPis.Text), "#,###0.00")
'    End If
'
'    lblDescontos_especiais.Caption = txtDesconto_especial.Text
'    lblTotal_Produtos.Caption = Format(dblTotal_Produtos, "#,###0.00")
'
'    If lblTotal_Produtos.Caption = Empty Then lblTotal_Produtos.Caption = "0,00"
'    If lblDescontos_especiais.Caption = Empty Then lblDescontos_especiais.Caption = "0,00"
'    If lblImpostos.Caption = Empty Then lblImpostos.Caption = "0,00"
'    If lblTotal_Pedido.Caption = Empty Then lblTotal_Pedido.Caption = "0,00"
'
'    lblImpostos.Caption = Format(CDbl(dblImpostos) * CDbl(lblTotal_Produtos.Caption) / 100, "#,###0.00")
'
'    lblTotal_Pedido.Caption = Format(CDbl(lblTotal_Produtos.Caption) - CDbl(lblImpostos.Caption) - CDbl(lblDescontos_especiais.Caption), "#,###0.00")
'
'End Function
'
'
'


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub
