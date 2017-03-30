VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFechamento_caixa_posto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fechamento de Caixa"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFechamento_caixa_posto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   10335
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid objFlex_Aux 
      Height          =   645
      Left            =   6120
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1138
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
   Begin TabDlg.SSTab sstFechamento 
      Height          =   8040
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   14182
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "&Geral"
      TabPicture(0)   =   "frmFechamento_caixa_posto.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpFechamento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtcOperador"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtOperador"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraConferencia"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtObservacao"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraSecao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdOK"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraFinalizadora"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdConsulta_Encerrante"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdImprimir_Fechamento"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "&Conferência Produto"
      TabPicture(1)   =   "frmFechamento_caixa_posto.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "label15(2)"
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label17"
      Tab(1).Control(7)=   "cbbTipo_Preco"
      Tab(1).Control(8)=   "dtcProduto"
      Tab(1).Control(9)=   "hfgProduto"
      Tab(1).Control(10)=   "txtProduto"
      Tab(1).Control(11)=   "cmdLimpar"
      Tab(1).Control(12)=   "cmdRemover_Item"
      Tab(1).Control(13)=   "cmdIncluir_Item"
      Tab(1).Control(14)=   "txtUnidade"
      Tab(1).Control(15)=   "txtPreco_unitario"
      Tab(1).Control(16)=   "txtQuantidade_produto"
      Tab(1).Control(17)=   "txtTotal_item"
      Tab(1).Control(18)=   "txtEstoque_Atual"
      Tab(1).Control(19)=   "Frame1"
      Tab(1).Control(20)=   "fraVendedor"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "&Listagem"
      TabPicture(2)   =   "frmFechamento_caixa_posto.frx":17BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdConsulta"
      Tab(2).Control(1)=   "cmdRefresh"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtConsulta"
      Tab(2).Control(3)=   "hfgFechamento"
      Tab(2).Control(4)=   "cbbCampos"
      Tab(2).Control(5)=   "dtpInicial"
      Tab(2).Control(6)=   "dtpFinal"
      Tab(2).Control(7)=   "lblA"
      Tab(2).Control(8)=   "Label29"
      Tab(2).ControlCount=   9
      Begin VB.Frame fraVendedor 
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   -74850
         TabIndex        =   59
         Top             =   5280
         Width           =   10035
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgVendedor 
            Height          =   2175
            Left            =   120
            TabIndex        =   60
            Top             =   300
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            ScrollBars      =   2
            SelectionMode   =   1
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
      End
      Begin VB.CommandButton cmdImprimir_Fechamento 
         Caption         =   "Imprimir Fechamento"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8790
         TabIndex        =   34
         ToolTipText     =   "Incluir"
         Top             =   7290
         Width           =   1425
      End
      Begin VB.Frame Frame1 
         Caption         =   "Totalizador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74850
         TabIndex        =   54
         Top             =   4530
         Width           =   10035
         Begin VB.Label lblTotal_Vendas 
            Caption         =   "lblTotal_Vendas"
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
            Height          =   225
            Left            =   7770
            TabIndex        =   58
            ToolTipText     =   "Total de Valor pego pelo Cliente"
            Top             =   330
            Width           =   1995
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total:"
            Height          =   240
            Left            =   6600
            TabIndex        =   57
            ToolTipText     =   "Total de IPI  + Total de despesas  acessórios"
            Top             =   330
            Width           =   1020
         End
         Begin VB.Label lblTotal_Quantidade 
            Caption         =   "lblTotal_Quantidade"
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
            Height          =   225
            Left            =   1875
            TabIndex        =   56
            ToolTipText     =   "Total de Títulos pegos pelo Cliente"
            Top             =   330
            Width           =   1995
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Total Quantidade.:"
            Height          =   240
            Left            =   120
            TabIndex        =   55
            Top             =   330
            Width           =   1605
         End
      End
      Begin VB.CommandButton cmdConsulta_Encerrante 
         Height          =   360
         Left            =   9390
         Picture         =   "frmFechamento_caixa_posto.frx":17D6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Digitação de informações de encerrante"
         Top             =   780
         Width           =   375
      End
      Begin VB.Frame fraFinalizadora 
         Caption         =   "Finalizadoras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   21
         Top             =   2595
         Width           =   10095
         Begin VB.TextBox txtValor_Finalizadora 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   6150
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            ToolTipText     =   "Troco Recebido"
            Top             =   510
            Width           =   1485
         End
         Begin VB.CommandButton cmdIncluir_Finalizadora 
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
            Height          =   375
            Left            =   7710
            TabIndex        =   27
            ToolTipText     =   "Incluir"
            Top             =   510
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemover_Finalizadora 
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
            Height          =   375
            Left            =   8850
            TabIndex        =   28
            ToolTipText     =   "Remover"
            Top             =   510
            Width           =   1095
         End
         Begin VB.TextBox txtFinalizadora 
            Height          =   360
            Left            =   120
            MaxLength       =   20
            TabIndex        =   23
            ToolTipText     =   "Código da Finalizadora"
            Top             =   510
            Width           =   1365
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgFinalizadora 
            Height          =   1455
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   9825
            _ExtentX        =   17330
            _ExtentY        =   2566
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
         Begin MSDataListLib.DataCombo dtcFinalizadora 
            Height          =   360
            Left            =   1530
            TabIndex        =   24
            Top             =   510
            Width           =   4575
            _ExtentX        =   8070
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
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   240
            Left            =   6150
            TabIndex        =   25
            Top             =   270
            Width           =   450
         End
         Begin VB.Label label15 
            AutoSize        =   -1  'True
            Caption         =   "Finalizadora"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.TextBox txtEstoque_Atual 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   360
         Left            =   -67830
         TabIndex        =   49
         ToolTipText     =   "Estoque atual do item"
         Top             =   1440
         Width           =   1750
      End
      Begin VB.TextBox txtTotal_item 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   360
         Left            =   -69630
         TabIndex        =   47
         ToolTipText     =   "Total do item"
         Top             =   1440
         Width           =   1750
      End
      Begin VB.TextBox txtQuantidade_produto 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   -73380
         TabIndex        =   41
         ToolTipText     =   "Quantidade do Item"
         Top             =   1440
         Width           =   1395
      End
      Begin VB.TextBox txtPreco_unitario 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   -71430
         TabIndex        =   45
         ToolTipText     =   "Preço unitário do item"
         Top             =   1440
         Width           =   1750
      End
      Begin VB.TextBox txtUnidade 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   360
         Left            =   -71940
         TabIndex        =   43
         ToolTipText     =   "Unidade do Item"
         Top             =   1440
         Width           =   465
      End
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
         Height          =   345
         Left            =   -66000
         Picture         =   "frmFechamento_caixa_posto.frx":1B60
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Adicionar item"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdRemover_Item 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -65220
         Picture         =   "frmFechamento_caixa_posto.frx":1CAA
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Remove Item"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdLimpar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -65610
         Picture         =   "frmFechamento_caixa_posto.frx":21EC
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Cancelar"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtProduto 
         Height          =   360
         Left            =   -74850
         MaxLength       =   20
         TabIndex        =   36
         ToolTipText     =   "Código do Produto"
         Top             =   780
         Width           =   1395
      End
      Begin VB.CommandButton cmdOK 
         Height          =   360
         Left            =   9810
         Picture         =   "frmFechamento_caixa_posto.frx":2336
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   -65610
         Picture         =   "frmFechamento_caixa_posto.frx":4030
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   -65220
         Picture         =   "frmFechamento_caixa_posto.frx":5D2A
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72570
         TabIndex        =   63
         Top             =   780
         Width           =   6885
      End
      Begin VB.Frame fraSecao 
         Caption         =   "Seções de Venda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   120
         TabIndex        =   30
         Top             =   5175
         Width           =   10095
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSecao 
            Height          =   1455
            Left            =   120
            TabIndex        =   31
            Top             =   270
            Width           =   9825
            _ExtentX        =   17330
            _ExtentY        =   2566
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            ScrollBars      =   2
            SelectionMode   =   1
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
      End
      Begin VB.TextBox txtObservacao 
         Height          =   600
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         ToolTipText     =   "Observação"
         Top             =   7305
         Width           =   8595
      End
      Begin VB.Frame fraConferencia 
         Caption         =   "Conferência"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   120
         TabIndex        =   9
         Top             =   1215
         Width           =   10095
         Begin VB.Frame Frame6 
            Caption         =   "Resultado do Caixa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   7800
            TabIndex        =   19
            Top             =   240
            Width           =   2175
            Begin VB.TextBox txtResultado_Caixa 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   120
               MaxLength       =   100
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               ToolTipText     =   "Resultado do Caixa"
               Top             =   510
               Width           =   1920
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Total de Vendas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   5760
            TabIndex        =   17
            Top             =   240
            Width           =   2025
            Begin VB.TextBox txtTotal_Vendas 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   360
               Left            =   120
               MaxLength       =   100
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               ToolTipText     =   "Total de Vendas Grupos"
               Top             =   510
               Width           =   1770
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Finalizadora"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   5625
            Begin VB.TextBox txtSubTotal 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   360
               Left            =   3720
               MaxLength       =   100
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               ToolTipText     =   "Sub-Total (Total Finalizadora - Troco Recebido)"
               Top             =   510
               Width           =   1770
            End
            Begin VB.TextBox txtTroco_Recebido 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   360
               Left            =   1920
               MaxLength       =   100
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               ToolTipText     =   "Troco Recebido"
               Top             =   510
               Width           =   1770
            End
            Begin VB.TextBox txtTotal_Finalizadora 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   360
               Left            =   120
               MaxLength       =   100
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               ToolTipText     =   "Total das Finalizadoras"
               Top             =   510
               Width           =   1770
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Sub-Total"
               Height          =   240
               Left            =   3720
               TabIndex        =   15
               Top             =   270
               Width           =   840
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Débitos"
               Height          =   240
               Left            =   1920
               TabIndex        =   13
               Top             =   270
               Width           =   630
            End
            Begin VB.Label lblObservacao 
               AutoSize        =   -1  'True
               Caption         =   "Créditos"
               Height          =   240
               Left            =   120
               TabIndex        =   11
               Top             =   270
               Width           =   705
            End
         End
      End
      Begin VB.TextBox txtOperador 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Código do Operador"
         Top             =   780
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dtcOperador 
         Height          =   360
         Left            =   1530
         TabIndex        =   4
         Top             =   780
         Width           =   6375
         _ExtentX        =   11245
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
      Begin MSComCtl2.DTPicker dtpFechamento 
         Height          =   360
         Left            =   7950
         TabIndex        =   6
         Top             =   780
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
         Format          =   20643841
         CurrentDate     =   37881
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgFechamento 
         Height          =   6675
         Left            =   -74880
         TabIndex        =   66
         Top             =   1230
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   11774
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   -74880
         TabIndex        =   62
         Top             =   780
         Width           =   2265
         _ExtentX        =   3995
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
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   360
         Left            =   -72570
         TabIndex        =   67
         Top             =   780
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   20643841
         CurrentDate     =   37881
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   360
         Left            =   -68160
         TabIndex        =   68
         Top             =   780
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   20643841
         CurrentDate     =   37881
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgProduto 
         Height          =   2655
         Left            =   -74850
         TabIndex        =   53
         Top             =   1860
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4683
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
         Height          =   360
         Left            =   -73410
         TabIndex        =   37
         Top             =   780
         Width           =   8565
         _ExtentX        =   15108
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
      Begin AutoCompletar.CbCompleta cbbTipo_Preco 
         Height          =   360
         Left            =   -74850
         TabIndex        =   39
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Preço"
         Height          =   240
         Left            =   -74850
         TabIndex        =   38
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque Atual"
         Height          =   240
         Left            =   -67830
         TabIndex        =   48
         Top             =   1185
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Item"
         Height          =   240
         Left            =   -69630
         TabIndex        =   46
         Top             =   1185
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         Height          =   240
         Left            =   -73380
         TabIndex        =   40
         Top             =   1185
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pr. Unitário"
         Height          =   240
         Left            =   -71430
         TabIndex        =   44
         Top             =   1185
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unid."
         Height          =   240
         Left            =   -71940
         TabIndex        =   42
         Top             =   1185
         Width           =   435
      End
      Begin VB.Label label15 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         Height          =   240
         Index           =   2
         Left            =   -74850
         TabIndex        =   35
         Top             =   540
         Width           =   660
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         Height          =   240
         Left            =   -69600
         TabIndex        =   69
         Top             =   930
         Width           =   105
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   61
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   240
         Left            =   120
         TabIndex        =   32
         Top             =   7065
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fechamento"
         Height          =   240
         Left            =   7950
         TabIndex        =   5
         Top             =   540
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Operador"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   810
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10050
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
            Picture         =   "frmFechamento_caixa_posto.frx":6D6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_caixa_posto.frx":7086
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_caixa_posto.frx":73A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_caixa_posto.frx":773A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_caixa_posto.frx":7AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_caixa_posto.frx":7DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_caixa_posto.frx":8108
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
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
      EndProperty
   End
End
Attribute VB_Name = "frmFechamento_caixa_posto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador de Vendas                                         '
' Objetivo...............: Cadastro de Fechamento de Caixa                                '
' Data de Criação........: 11/07/2003                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......: Relatório de fechamento modelo c (mazzarope por fechamento)    '
' Desenvolvedor..........: Leandro Nolasco Ferreira                                       '
' Data última manutenção.: 30/08/2006                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL_Listagem As String

Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim strIDFechamento As String

Dim rstAplicacao As New ADODB.Recordset
Dim rstTipo_Preco As New ADODB.Recordset

Dim rstProdutos As ADODB.Recordset

'array para controle de alterações de produtos dos bicos das bombas
Dim arrEncerrantes_Bicos() As String

'variáveis para montagem de grids
Dim strNomes As String
Dim strTamanho As String

'Declaração das variaveis da acessibilidade
Dim acesso As New DLLSystemManager.Acessibilidade
Dim strID_Acessibilidade As String
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean

'Ferramentas do sistema
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim log As New DLLSystemManager.log

'Portal
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean

'Seção do encerrante
Dim strCodSecaoCombustivel As String
Dim strDescSecaoCombustivel As String
Dim dblValorSecaoCombustivel As Double

Dim booDigita_Encerrante As Boolean

Option Explicit

Public Function setLista_Encerrante_Bico(objFlex As MSHFlexGrid) As Boolean

    Dim intI As Integer
    Dim intJ As Integer
    Dim dblValor As Double
    Dim conConexao As Conexao
    Dim strCodSecao As String

    setLista_Encerrante_Bico = False

    ReDim arrEncerrantes_Bicos(objFlex.Cols, objFlex.Rows - 2)

    If objFlex.TextMatrix(1, 1) = Empty Then
        setLista_Encerrante_Bico = True
        Exit Function
    End If

    For intI = 1 To objFlex.Rows - 1
        
        For intJ = 0 To objFlex.Cols - 1
            With objFlex
                arrEncerrantes_Bicos(intJ, intI - 1) = .TextMatrix(intI, intJ)
            End With
        Next intJ
        
        dblValor = dblValor + CDbl(arrEncerrantes_Bicos(12, intI - 1))
        strCodSecao = arrEncerrantes_Bicos(13, intI - 1)
        strDescSecaoCombustivel = arrEncerrantes_Bicos(14, intI - 1)
    
    Next intI
    
    For intI = 0 To UBound(arrEncerrantes_Bicos, 2)
        If arrEncerrantes_Bicos(6, 0) <> 0 Then
            cmdOk.Enabled = True
            Exit For
        End If
    Next intI
    
    strCodSecaoCombustivel = strCodSecao
    dblValorSecaoCombustivel = dblValor
    
    For intI = 1 To hfgSecao.Rows - 1
        If hfgSecao.TextMatrix(intI, 1) = strCodSecaoCombustivel Then
            hfgSecao.TextMatrix(intI, 3) = Format(dblValorSecaoCombustivel, "##,##0.00")
            Exit For
        End If
    Next intI

    setLista_Encerrante_Bico = True

End Function

Private Sub Imprimir()

    Dim strFormulas As String
    Dim strValores As String
    Dim frmAux_Imp As Form
    
    On Error GoTo Erro
    'Tratamento de Erro
    If hfgFechamento.TextMatrix(1, 1) = Empty Or strSQL_Listagem = Empty Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.", vbInformation, "Only Tech"
       txtConsulta.SetFocus
       Exit Sub
    End If

    frmAguarde.Show
    DoEvents

    Set frmAux_Imp = New frmConsole_Geral
    
    strFormulas = "Cliente;Tipo_relatorio"
    strValores = Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me) & ";Listagem Geral"
    
    frmAux_Imp.setParametros strSQL_Listagem, "rptListagem_Fechamento_caixa_posto.rpt", strFormulas, strValores
    
    frmAux_Imp.Show
    
    Unload frmAguarde
    
    Set frmAux_Imp = Nothing

    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Sub
End Sub

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty

    If cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       lblA.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Data Fechamento" Then
       txtConsulta.Visible = False
       dtpInicial.Visible = True
       dtpFinal.Visible = True
       lblA.Visible = True
       dtpInicial.SetFocus
    Else
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       lblA.Visible = False
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If
    
End Sub

Private Sub cbbTipo_Preco_GotFocus()
    If cbbTipo_Preco.Text = Empty Then
       SendKeys "{F4}"
    Else
       On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
    End If
End Sub

Private Sub cbbTipo_Preco_LostFocus()

    Dim strSQL As String

    If txtProduto.Text <> Empty Then
        strSQL = "SELECT DFPreco_avista_TBItens_tabela_preco," & _
                 "DFPreco_promocao_TBItens_tabela_preco," & _
                 "DFPreco_revenda_TBItens_tabela_preco," & _
                 "DFPreco_especial_TBItens_tabela_preco," & _
                 "DFPreco_varejo_TBItens_tabela_preco " & _
                 "FROM TBItens_tabela_preco " & _
                 "INNER JOIN TBProduto ON TBItens_tabela_preco.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                 "WHERE FKCodigo_TBTabela_preco IN (SELECT DFNumero_tabela_vigente_TBParametros_venda " & _
                 "FROM TBParametros_venda WHERE IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ") " & _
                 "AND TBProduto.IXCodigo_TBProduto = " & txtProduto.Text & " " & _
                 "AND TBProduto.IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " "
        
        Select_geral strSQL, "BDRetaguarda", rstAplicacao, "Otica", Me
       
        If rstAplicacao.RecordCount = 0 Then
            Set rstAplicacao = Nothing
            MsgBox "Produto não cadastrado na Tabela de Preço. Verifique.", vbInformation, "Only Tech"
            'cbbTipo_Preco.Text = Empty
            txtUnidade.Text = Empty
            txtEstoque_Atual.Text = Empty
            txtProduto.Text = Empty
            txtProduto.SetFocus
            Exit Sub
        End If
       
        strSQL = "SELECT DFNome_Preco_avista_TBTipo_preco," & _
                 "DFNome_Preco_promocao_TBTipo_preco," & _
                 "DFNome_Preco_revenda_TBTipo_preco," & _
                 "DFNome_Preco_especial_TBTipo_preco," & _
                 "DFNome_Preco_varejo_TBTipo_preco " & _
                 "FROM TBTipo_preco"
                 
        Select_geral strSQL, "BDRetaguarda", rstTipo_Preco, "Otica", Me
       
        If cbbTipo_Preco.Text = rstTipo_Preco.Fields("DFNome_Preco_avista_TBTipo_preco") Then
            If IsNull(rstAplicacao.Fields("DFPreco_avista_TBItens_tabela_preco")) = False Then
                txtPreco_unitario.Text = rstAplicacao.Fields("DFPreco_avista_TBItens_tabela_preco")
            End If
        ElseIf cbbTipo_Preco.Text = rstTipo_Preco.Fields("DFNome_Preco_promocao_TBTipo_preco") Then
            If IsNull(rstAplicacao.Fields("DFPreco_promocao_TBItens_tabela_preco")) = False Then
                txtPreco_unitario.Text = rstAplicacao.Fields("DFPreco_promocao_TBItens_tabela_preco")
            End If
        ElseIf cbbTipo_Preco.Text = rstTipo_Preco.Fields("DFNome_Preco_revenda_TBTipo_preco") Then
            If IsNull(rstAplicacao.Fields("DFPreco_revenda_TBItens_tabela_preco")) = False Then
                txtPreco_unitario.Text = rstAplicacao.Fields("DFPreco_revenda_TBItens_tabela_preco")
            End If
        ElseIf cbbTipo_Preco.Text = rstTipo_Preco.Fields("DFNome_Preco_especial_TBTipo_preco") Then
            If IsNull(rstAplicacao.Fields("DFPreco_especial_TBItens_tabela_preco")) = False Then
                txtPreco_unitario.Text = rstAplicacao.Fields("DFPreco_especial_TBItens_tabela_preco")
            End If
        ElseIf cbbTipo_Preco.Text = rstTipo_Preco.Fields("DFNome_Preco_varejo_TBTipo_preco") Then
            If IsNull(rstAplicacao.Fields("DFPreco_varejo_TBItens_tabela_preco")) = False Then
                txtPreco_unitario.Text = rstAplicacao.Fields("DFPreco_varejo_TBItens_tabela_preco")
            End If
        End If
        
        txtPreco_unitario.Text = Format(txtPreco_unitario.Text, "##,##0.00")
        
        Set rstAplicacao = Nothing
        Set rstTipo_Preco = Nothing
    Else
        txtQuantidade_produto.Text = Empty
        txtPreco_unitario.Text = Empty
        txtTotal_item.Text = Empty
        txtEstoque_Atual.Text = Empty
        txtUnidade.Text = Empty
    End If
    
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdConsulta_Encerrante_Click()
    
    Dim frmAux As Form
    
    If txtOperador.Text <> Empty Then
        frmAguarde.Show
       
        Set frmAux = New frmFechamento_caixa_posto_consulta_encerrante
       
        If UBound(arrEncerrantes_Bicos, 2) > 0 Then
            frmAux.setParametros arrEncerrantes_Bicos, False
        Else
            frmAux.setParametros arrEncerrantes_Bicos, True
        End If
       
        'frmAux.setParametros ""
        Unload frmAguarde
        frmAguarde.Show
        frmAux.Show 1, MDIPrincipal
        Unload frmAguarde
    
        Set frmAux = Nothing
    
    End If
    
End Sub

Private Sub cmdImprimir_Fechamento_Click()

'''''    On Error GoTo Erro
'''''
'''''    Dim strSQL As String
'''''
'''''    Dim strSQLSub1 As String
'''''    Dim strSQLSub2 As String
'''''    Dim strSQLSub3 As String
'''''    Dim strSQLSub4 As String
'''''
'''''    Dim strAliasTabelaSub1 As String
'''''    Dim strAliasTabelaSub2 As String
'''''    Dim strAliasTabelaSub3 As String
'''''    Dim strAliasTabelaSub4 As String
'''''
'''''    Dim strArqSubRpt1 As String
'''''    Dim strArqSubRpt2 As String
'''''    Dim strArqSubRpt3 As String
'''''    Dim strArqSubRpt4 As String
'''''
'''''    Dim strFormulas As String
'''''    Dim strValores As String
'''''    Dim frmAux_Imp As Form
'''''
'''''    Dim conConexao As Conexao
'''''    Dim rstAux As ADODB.Recordset
'''''
'''''    Dim strSufixo_Tabela As String
'''''
'''''
'''''    strArqSubRpt1 = "rptFechamento_Caixa_Posto_Finalizadoras.rpt"
'''''    strArqSubRpt2 = "rptFechamento_Caixa_Posto_Secao.rpt"
'''''    strArqSubRpt3 = "rptFechamento_Caixa_Posto_Produto.rpt"
'''''    strArqSubRpt4 = "rptFechamento_Caixa_Posto_Encerrante.rpt"
'''''
'''''
'''''    strSufixo_Tabela = Format(Now, "ddMMyyHHnnss")
'''''    'strSufixo_Tabela = Empty
'''''
'''''    'fechamento
'''''    strSQL = strSQL & "IF OBJECT_ID('TBTEMP_FechCxPst" & strSufixo_Tabela & "') IS NOT NULL BEGIN DROP TABLE TBTEMP_FechCxPst" & strSufixo_Tabela & " END " & _
'''''        "SELECT PKId_TBFechamento_caixa_posto, " & _
'''''               "DFData_TBFechamento_caixa_posto, " & _
'''''               "FKCodigo_TBOperadores_ecf, " & _
'''''               "DFNome_TBOperadores_ecf, " & _
'''''               "DFTotal_finalizadoras_TBFechamento_caixa_posto, " & _
'''''               "DFTotal_troco_TBFechamento_caixa_posto, " & _
'''''               "DFTotal_vendas_grupo_TBFechamento_caixa_posto, " & _
'''''               "DFResultado_TBFechamento_caixa_posto, " & _
'''''               "DFObservacao_TBFechamento_caixa_posto " & _
'''''          "/*INTO TBTEMP_Fech_Caixa_Posto*/ " & _
'''''          "INTO TBTEMP_FechCxPst" & strSufixo_Tabela & " " & _
'''''          "FROM TBFechamento_caixa_posto " & _
'''''    "INNER JOIN TBOperadores_ecf " & _
'''''            "ON TBFechamento_caixa_posto.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
'''''         "WHERE PKId_TBFechamento_caixa_posto = " & strIDFechamento & " "
'''''
'''''    'finalizadora
'''''    strSQL = strSQL & "IF OBJECT_ID('TBTEMP_FechCxPstFin" & strSufixo_Tabela & "') IS NOT NULL BEGIN DROP TABLE TBTEMP_FechCxPstFin" & strSufixo_Tabela & " END " & _
'''''        "SELECT IXCodigo_TBFinalizadora, " & _
'''''               "DFDescricao_TBFinalizadora, " & _
'''''               "DFValor_total_TBFechamento_caixa_posto_finalizadora, " & _
'''''               "CONVERT(NVARCHAR, DFDebito_credito_TBFinalizadora) AS DFDebito_credito_TBFinalizadora, " & _
'''''               "CONVERT(NVARCHAR, DFDebito_Credito_TBfinalizadora) + '|' + CONVERT(NVARCHAR, DFDescricao_TBfinalizadora) AS Fin_DebCred, " & _
'''''               "FKID_TBfechamento_caixa_posto " & _
'''''          "INTO TBTEMP_FechCxPstFin" & strSufixo_Tabela & " " & _
'''''          "FROM TBFechamento_caixa_posto_finalizadora " & _
'''''    "INNER JOIN TBFinalizadora " & _
'''''            "ON TBFechamento_caixa_posto_finalizadora.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
'''''         "WHERE FKId_TBfechamento_caixa_posto = " & strIDFechamento & " " & _
'''''      "ORDER BY 5 "
'''''
'''''    'seção
'''''    strSQL = strSQL & "IF OBJECT_ID('TBTEMP_FechCxPstSec" & strSufixo_Tabela & "') IS NOT NULL BEGIN DROP TABLE TBTEMP_FechCxPstSec" & strSufixo_Tabela & " END " & _
'''''        "SELECT FKCodigo_TBSecao, " & _
'''''               "DFDescricao_TBsecao, " & _
'''''               "DFValor_total_TBFechamento_caixa_posto_venda_grupo, " & _
'''''               "FKID_TBfechamento_caixa_posto " & _
'''''          "INTO TBTEMP_FechCxPstSec" & strSufixo_Tabela & " " & _
'''''          "FROM TBFechamento_caixa_posto_venda_grupo " & _
'''''    "INNER JOIN TBSecao " & _
'''''            "ON TBFechamento_caixa_posto_venda_grupo.FKCodigo_TBSecao = TBSecao.PKCodigo_TBSecao " & _
'''''         "WHERE FKId_TBfechamento_caixa_posto = " & strIDFechamento & " " & _
'''''      "ORDER BY DFDescricao_TBsecao "
'''''
'''''    'produto
'''''    strSQL = strSQL & "IF OBJECT_ID('TBTEMP_FechCxPstPrd" & strSufixo_Tabela & "') IS NOT NULL BEGIN DROP TABLE TBTEMP_FechCxPstPrd" & strSufixo_Tabela & " END " & _
'''''        "SELECT IXCodigo_TBProduto, " & _
'''''               "DFDescricao_TBProduto, " & _
'''''               "DFQuantidade_TBFechamento_caixa_produto_posto, " & _
'''''               "ISNULL(ISNULL(DFUnidade_varejo_TBProduto, DFUnidade_venda_TBProduto), '') AS DFunidade, " & _
'''''               "DFValor_unitario_TBFechamento_caixa_produto_posto, " & _
'''''               "DFValor_Total_TBFechamento_caixa_produto_posto, " & _
'''''               "PKCodigo_TBSecao, " & _
'''''               "FKID_TBfechamento_caixa_posto " & _
'''''          "INTO TBTEMP_FechCxPstPrd" & strSufixo_Tabela & " " & _
'''''          "FROM TBFechamento_caixa_produto_posto " & _
'''''    "INNER JOIN TBFechamento_caixa_posto " & _
'''''            "ON TBFechamento_caixa_produto_posto.FKId_TBfechamento_caixa_posto = TBFechamento_caixa_posto.PKId_TBfechamento_caixa_posto " & _
'''''    "INNER JOIN TBProduto " & _
'''''            "ON dbo.TBFechamento_caixa_produto_posto.FKId_TBproduto = TBProduto.PKId_TBproduto " & _
'''''    "LEFT  JOIN TBsecao " & _
'''''            "ON TBproduto.FKCodigo_TBSecao = TBsecao.PKCodigo_TBSecao " & _
'''''         "WHERE PKId_TBfechamento_caixa_posto = " & strIDFechamento & " " & _
'''''      "ORDER BY DFDescricao_TBProduto "
'''''
'''''    'encerrante
'''''    strSQL = strSQL & "IF OBJECT_ID('TBTEMP_FechCxPstEnc" & strSufixo_Tabela & "') IS NOT NULL BEGIN DROP TABLE TBTEMP_FechCxPstEnc" & strSufixo_Tabela & " END " & _
'''''        "SELECT PKId_TBbomba_bico, " & _
'''''               "IXCodigo_Bomba, " & _
'''''               "IXCodigo_TBBomba_bico, " & _
'''''               "TBproduto.IXCodigo_TBproduto, " & _
'''''               "TBproduto.DFdescricao_TBproduto, " & _
'''''               "DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
'''''               "DFEncerrante_final_TBEncerrante_caixa_posto, " & _
'''''               "DFAfericao_TBEncerrante_caixa_posto, " & _
'''''               "DFQuantidade_DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
'''''               "DFValor_unitario_DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
'''''               "DFCusto_unitario_DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
'''''               "DFValor_total_TBEncerrante_caixa_posto, " & _
'''''               "PKCodigo_TBsecao, " & _
'''''               "DFDescricao_TBsecao, " & _
'''''               "FKID_TBfechamento_caixa_posto "
'''''    strSQL = strSQL & _
'''''          "INTO TBTEMP_FechCxPstEnc" & strSufixo_Tabela & " " & _
'''''          "FROM TBbomba_bico " & _
'''''    "INNER JOIN TBbomba " & _
'''''            "ON TBbomba.PKId_TBbomba = TBbomba_bico.FKId_TBbomba " & _
'''''    "INNER JOIN TBEncerrante_caixa_posto " & _
'''''            "ON PKId_TBbomba_bico = FKId_TBbomba_bico " & _
'''''           "AND FKId_TBfechamento_caixa_posto = " & strIDFechamento & " " & _
'''''    "INNER JOIN TBProduto " & _
'''''            "ON TBbomba_bico.FKId_TBproduto = TBProduto.PKId_TBProduto " & _
'''''           "AND TBProduto.IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
'''''    "INNER JOIN TBitens_tabela_preco " & _
'''''            "ON PKId_TBProduto = TBitens_tabela_preco.FKId_TBproduto " & _
'''''           "AND FKCodigo_TBTabela_preco IN ( SELECT DFNumero_tabela_vigente_TBParametros_venda " & _
'''''                                              "From TBparametros_venda " & _
'''''                                             "WHERE IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " ) " & _
'''''    "LEFT  JOIN TBsecao " & _
'''''            "ON TBproduto.FKCodigo_TBsecao = TBsecao.PKCodigo_TBsecao "
'''''
'''''
'''''    Set conConexao = New Conexao
'''''    Conexao.Initial_Catalog = "BDRetaguarda"
'''''    Call conConexao.Abrir_conexao("Otica")
'''''
'''''    Call conConexao.CNconexao.Execute(strSQL)
'''''
'''''
'''''    'Verifica se pode ser impresso relatório
'''''    Set rstAux = New ADODB.Recordset
'''''    Call Movimentacoes.Select_geral("SELECT * FROM TBTEMP_FechCxPstSec" & strSufixo_Tabela, "BDRetaguarda", rstAux, "Otica", Me)
'''''    If rstAux.RecordCount <= 0 Then
'''''       MsgBox "Não existem informações suficientes para a geração deste relatório.", vbInformation, "Only Tech"
'''''       txtConsulta.SetFocus
'''''       Exit Sub
'''''    End If
'''''
'''''    If Not rstAux Is Nothing Then
'''''        If rstAux.State <> adStateClosed Then
'''''            rstAux.Close
'''''        End If
'''''        Set rstAux = Nothing
'''''    End If
'''''
'''''    strSQLSub1 = "SELECT * FROM TBTEMP_FechCxPstFin" & strSufixo_Tabela
'''''    strSQLSub2 = "SELECT * FROM TBTEMP_FechCxPstSec" & strSufixo_Tabela
'''''    strSQLSub3 = "SELECT * FROM TBTEMP_FechCxPstPrd" & strSufixo_Tabela
'''''    strSQLSub4 = "SELECT * FROM TBTEMP_FechCxPstEnc" & strSufixo_Tabela
'''''
'''''    strAliasTabelaSub1 = "TBTEMP_Fech_Caixa_Posto_Finalizadora"
'''''    strAliasTabelaSub2 = "TBTEMP_Fech_Caixa_Posto_Finalizadora"
'''''    strAliasTabelaSub3 = "TBTEMP_Fech_Caixa_Posto_Finalizadora"
'''''    strAliasTabelaSub4 = "TBTEMP_Fech_Caixa_Posto_Finalizadora"
'''''
'''''    frmAguarde.Show
'''''    DoEvents
'''''
'''''    Set frmAux_Imp = New frmConsole_Geral
'''''
'''''    strFormulas = "Cliente;Tipo_Relatorio"
'''''    strValores = Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me) & ";Fechamento de Caixa Analítico"
'''''
'''''    If booDigita_Encerrante Then
'''''        frmAux_Imp.setParametros "SELECT * FROM TBTEMP_FechCxPst" & strSufixo_Tabela, "rptFechamento_caixa_posto.rpt", strFormulas, strValores, True, strSQLSub1 & ";" & strSQLSub2 & ";" & strSQLSub3 & ";" & strSQLSub4, strArqSubRpt1 & ";" & strArqSubRpt2 & ";" & strArqSubRpt3 & ";" & strArqSubRpt4, strAliasTabelaSub1 & ";" & strAliasTabelaSub2 & ";" & strAliasTabelaSub3 & ";" & strAliasTabelaSub4
'''''    Else
'''''        frmAux_Imp.setParametros "SELECT * FROM TBTEMP_FechCxPst" & strSufixo_Tabela, "rptFechamento_caixa_posto.rpt", strFormulas, strValores, False, , strSQLSub1 & ";" & strSQLSub2 & ";" & strSQLSub3 & ";" & strSQLSub4, strArqSubRpt1 & ";" & strArqSubRpt2 & ";" & strArqSubRpt3 & ";" & strArqSubRpt4, strAliasTabelaSub1 & ";" & strAliasTabelaSub2 & ";" & strAliasTabelaSub3 & ";" & strAliasTabelaSub4
'''''    End If
'''''
''''''    If booDigita_Encerrante Then
''''''        frmAux_Imp.setParametros "SELECT * FROM TBTEMP_FechCxPst" & strSufixo_Tabela, "rptFechamento_caixa_posto.rpt", strFormulas, strValores, True
''''''    Else
''''''        frmAux_Imp.setParametros "SELECT * FROM TBTEMP_FechCxPst" & strSufixo_Tabela, "rptFechamento_caixa_posto.rpt", strFormulas, strValores
''''''    End If
'''''
'''''    frmAux_Imp.Show
'''''
'''''    'strSQL = _
'''''        "DROP TABLE TBTEMP_FechCxPst" & strSufixo_Tabela & " " & _
'''''        "DROP TABLE TBTEMP_FechCxPstFin" & strSufixo_Tabela & " " & _
'''''        "DROP TABLE TBTEMP_FechCxPstSec" & strSufixo_Tabela & " " & _
'''''        "DROP TABLE TABLE TBTEMP_FechCxPstPrd" & strSufixo_Tabela & " " & _
'''''        "DROP TABLE TABLE TBTEMP_FechCxPstEnc" & strSufixo_Tabela
'''''    'conConexao.CNconexao.Execute strSQL
'''''
'''''    Call conConexao.Fechar_conexao
'''''    Set conConexao = Nothing
'''''
'''''    Unload frmAguarde
'''''
'''''    Set frmAux_Imp = Nothing
'''''
'''''    Exit Sub
    
'''''Erro:
'''''    Call Erro.Erro(Me, "Otica")
'''''    Call conConexao.Fechar_conexao
'''''    Set conConexao = Nothing
'''''    Exit Sub
'''''    Resume


    Dim Array_Impressao() As String
    Dim Minha_Colecao As New Collection
    Dim int_Linhas_Cabecalho As Integer
    Dim int_Cont As Integer
    Dim str_Spaco As String
    Dim lngIdx As Long
    
    Dim bol_Cabecalho As Boolean
    Dim str_Filtro As String
    Dim intSpcs As Integer
    
    Dim Dados_Relatorio As clsDados_Relatorios
    Dim lng_Cont_Grupo As Long
    Dim byt_Linhas_Cab_Grupo As Byte
    Dim byt_Linhas_Grupo_Atual As Byte
    
    Dim lngTamanhoArray As Long
    Dim lngQtde_Fin_Deb As Long
    Dim lngQtde_Fin_Crd As Long
    Dim lngQtde_Tot_Liq As Long
    
    On Error GoTo TrataErro
    
    'Variável que conta os gupos do relatório
    lng_Cont_Grupo = 0
    'Variável que conta as linhas de cabeçalho por grupo
    byt_Linhas_Grupo_Atual = 0
    'Número de linhas correspondentes ao cabeçalho do Grupo
    byt_Linhas_Cab_Grupo = 1
    
    int_Cont = 0
    
    str_Spaco = Space(3) 'Máximo tamanho do campo
    
    'Totaliza linhas de finalizadoras
    For lngIdx = 1 To hfgFinalizadora.Rows - 1
        'Créditos menos abertura do caixa
        If hfgFinalizadora.TextMatrix(lngIdx, 4) = "C" Then
            If InStr(UCase(hfgFinalizadora.TextMatrix(lngIdx, 2)), "ABERTURA") Then
                lngQtde_Fin_Crd = lngQtde_Fin_Crd + 1
            End If
        'Débitos
        Else
            lngQtde_Fin_Deb = lngQtde_Fin_Deb + 1
        End If
        
        'Compõe finalizadora
        If hfgFinalizadora.TextMatrix(lngIdx, 6) = "Sim" Then
            lngQtde_Tot_Liq = lngQtde_Tot_Liq + 1
        End If
    Next lngIdx
    
    int_Linhas_Cabecalho = 7
    lngTamanhoArray = 2 + (hfgVendedor.Rows - 1) + 6 + lngQtde_Fin_Deb + 6 + lngQtde_Fin_Crd + 6 + lngQtde_Tot_Liq + 3
    
    Set Impressao = New clsImpressao
    
    ReDim Array_Impressao(lngTamanhoArray + int_Linhas_Cabecalho, 1)
        
    ' // Cabeçalho da página ----------------------------------------
    
    
    
'    Array_Impressao(int_Cont, 0) = "Only Tech Solutions" & Space(56) & "Data de Criacao: 30/08/2006"
'    Array_Impressao(int_Cont, 1) = "N"
'    int_Cont = int_Cont + 1
'
'    Array_Impressao(int_Cont, 0) = String(102, "-")
'    Array_Impressao(int_Cont, 1) = "N"
'    int_Cont = int_Cont + 1
'
'    Array_Impressao(int_Cont, 0) = Trim(Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me)) & Space(102 - (Len(Trim(Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me))) + Len("Data de Geracao: " & Format(Date, "dd/MM/yyyy")))) & "Data de Geracao: " & Format(Date, "dd/MM/yyyy")
'    Array_Impressao(int_Cont, 1) = "N"
'    int_Cont = int_Cont + 1
'
'    Array_Impressao(int_Cont, 0) = "Relatório de Fechamento de Caixa - Modelo C" & Space(102 - Len("Relatório de Fechamento de Caixa - Modelo C") + (Len("Hora de Geracao: " & Format(Now, "HH:nn:ss  ")))) & "Hora de Geracao: " & Format(Now, "HH:nn:ss  ")
'    Array_Impressao(int_Cont, 1) = "N"
'    int_Cont = int_Cont + 1
'
'    Array_Impressao(int_Cont, 0) = "Operador: " & Trim(txtOperador.Text) & "-" & Trim(dtcOperador.Text) & Space(102 - (Len("Operador: " & Trim(txtOperador.Text) & "-" & Trim(dtcOperador.Text)) + Len("Páginas        :           "))) & "Páginas        :           "
'    Array_Impressao(int_Cont, 1) = "N"
'    int_Cont = int_Cont + 1
'
'    Array_Impressao(int_Cont, 0) = String(102, "-")
'    Array_Impressao(int_Cont, 1) = "N"
'    int_Cont = int_Cont + 1
'
'    Array_Impressao(int_Cont, 0) = str_Spaco
'    Array_Impressao(int_Cont, 1) = "N"
'    int_Cont = int_Cont + 1
    ' Fim do Cabeçalho da página --------------------------------- //
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
'        If Not Vazio(str_xPesquisa(lngIdx, 5)) Then
'
'            Array_Impressao(int_Cont, 0) = Array_Impressao(int_Cont, 0) & str_xPesquisa(lngIdx, 0) & Space(7 - Len(str_xPesquisa(lngIdx, 0))) & str_Spaco
'
'            intSpcs = 20 - Len(str_xPesquisa(lngIdx, 1))
'            If intSpcs < 1 Then
'                Array_Impressao(int_Cont, 0) = Array_Impressao(int_Cont, 0) & Left(str_xPesquisa(lngIdx, 1), 20) & str_Spaco
'            Else
'                Array_Impressao(int_Cont, 0) = Array_Impressao(int_Cont, 0) & str_xPesquisa(lngIdx, 1) & Space(intSpcs) & str_Spaco
'            End If
'
'            Array_Impressao(int_Cont, 0) = Array_Impressao(int_Cont, 0) & _
'                str_xPesquisa(lngIdx, 2) & Space(18 - Len(str_xPesquisa(lngIdx, 2))) & str_Spaco & _
'                str_xPesquisa(lngIdx, 3) & Space(18 - Len(str_xPesquisa(lngIdx, 3))) & str_Spaco
'
'            intSpcs = 30 - Len(str_xPesquisa(lngIdx, 4))
'            If intSpcs < 1 Then
'                Array_Impressao(int_Cont, 0) = Array_Impressao(int_Cont, 0) & Left(str_xPesquisa(lngIdx, 4), 30) & str_Spaco
'            Else
'                Array_Impressao(int_Cont, 0) = Array_Impressao(int_Cont, 0) & str_xPesquisa(lngIdx, 4) & Space(intSpcs) & str_Spaco
'            End If
'
'            Array_Impressao(int_Cont, 0) = Array_Impressao(int_Cont, 0) & _
'                str_xPesquisa(lngIdx, 5) & Space(18 - Len(str_xPesquisa(lngIdx, 5))) & str_Spaco & _
'                str_xPesquisa(lngIdx, 6) & Space(18 - Len(str_xPesquisa(lngIdx, 6))) & str_Spaco & _
'                str_xPesquisa(lngIdx, 7) & Space(18 - Len(str_xPesquisa(lngIdx, 7)))
'
'            Array_Impressao(int_Cont, 1) = "N"
'        End If
'
'        'contador de linhas
'        int_Cont = int_Cont + 1
            
    'Next
    
    'informa as orientacaoes
    Impressao.Orientacao = 1                'Retrato
    Impressao.Papel = "A4"                  'Tamanho do Papel
    Impressao.Destino = 0                   'Tela
    Impressao.Fonte = "COURIER NEW"
    Impressao.Tamanho_Fonte_Cabecalho = 10
    Impressao.Tamanho_Fonte_Corpo = 10
    Impressao.Linhas_Cabecalho = int_Linhas_Cabecalho
    
    Impressao.Data_Criacao = "30/08/2006"
    Impressao.Titulo_Empresa = "Only Tech Solutions"
    Impressao.Titulo_Cliente = Trim(Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me))
    Impressao.Titulo = "Relatório de Fechamento de Caixa - Modelo C"
    Impressao.Sub_Titulo = "Operador: " & Trim(txtOperador.Text) & "-" & Trim(dtcOperador.Text) & " / Data do Fechamento: " & Format(dtpFechamento.Value, "dd/MM/yyyy")
    ' Fim do Cabeçalho da página --------------------------------- //
    
    
    Impressao.Rodape_Pagina = "Only Tech Retaguarda / Concentrador de Vendas: Cadastro >> Fechamento de Caixa"
    
    
    Array_Impressao(int_Cont, 0) = Impressao.Linha_Relatorio
    Array_Impressao(int_Cont, 1) = "N"
    int_Cont = int_Cont + 1
                     
    Impressao.Zera_Colecao
    
    For lngIdx = 0 To UBound(Array_Impressao)
    
        Set Dados_Relatorio = New clsDados_Relatorios
        With Dados_Relatorio
    
            If lngIdx >= int_Linhas_Cabecalho Then
                
                If Left(Array_Impressao(lngIdx, 0), 7) <> Empty Then
                    .IGrupo = lng_Cont_Grupo
                    byt_Linhas_Grupo_Atual = byt_Linhas_Grupo_Atual + 1
                    If byt_Linhas_Grupo_Atual = byt_Linhas_Cab_Grupo Then
                        lng_Cont_Grupo = lng_Cont_Grupo + 1
                        byt_Linhas_Grupo_Atual = 0
                    End If
                Else
                    .IGrupo = -1
                End If
                
            End If
        
            .Linha = Array_Impressao(lngIdx, 0)
            .Negrito = Array_Impressao(lngIdx, 1)
        End With
        
        Minha_Colecao.Add Dados_Relatorio, CStr(lngIdx)
        Set Dados_Relatorio = Nothing
            
    Next
        
    'configurar e executar a impressão
    Impressao.Conteudo = Minha_Colecao
    
    Impressao.Configura_Tela
    
'    str_Filtro = "Operador: " & txtOperador.Text & " - " & dtcOperador.Text & Space(20) & "Data do Fechamento: " & Format(dtpFechamento.Value, "dd/MM/yyyy")
'
'    Impressao.Sub_Titulo = str_Filtro
    frmVisualiza_Impressao.Show 1, MDIPrincipal
    
    Set Impressao = Nothing
    
    Exit Sub

TrataErro:
    If Err.Number <> 0 Then
         MsgBox Err.Number & " - " & Err.Description, vbCritical, wNomeSistema
    End If
    If Not Dados_Relatorio Is Nothing Then
        Set Dados_Relatorio = Nothing
    End If


End Sub

Private Sub cmdIncluir_Finalizadora_Click()

    'Sim = Débito
    'Não e Null = Crédito

    Dim I As Integer
    Dim strSQL As String
    Dim strDebCred As String
    Dim strTotal_Liquido As String
    
    hfgFinalizadora.Sort = 5

    If txtFinalizadora.Text = Empty Then
       MsgBox "Finalizadora inválida. Verifique.", vbInformation, "Only Tech"
       txtFinalizadora.SetFocus
       Exit Sub
    ElseIf txtValor_Finalizadora.Text = Empty Then
       MsgBox "Valor inválido. Verifique.", vbInformation, "Only Tech"
       txtValor_Finalizadora.SetFocus
       Exit Sub
    End If
    
    'verifica cadastro duplicado
    For I = 1 To hfgFinalizadora.Rows - 1
        If hfgFinalizadora.TextMatrix(I, 1) = txtFinalizadora.Text Then
            MsgBox "Finalizadora já cadastrada.", vbExclamation + vbOKOnly, "Only Tech"
            txtFinalizadora.SetFocus
            Exit Sub
        End If
    Next I
    
    Set rstAplicacao = Nothing
    Set rstAplicacao = New ADODB.Recordset
    
    'Verifica se é Débito ou Crédito
    strSQL = "SELECT DFDebito_credito_TBFinalizadora, CONVERT(NVARCHAR, ISNULL(DFDebito_credito_TBFinalizadora, 0)) AS DebCred, CONVERT(BIT, ISNULL(DFCompoe_total_liquido_TBFinalizadora, 0)) AS TL " & _
             "FROM TBFinalizadora " & _
             "WHERE IXCodigo_TBFinalizadora = " & txtFinalizadora.Text & ""
    
    Select_geral strSQL, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       ' no ado, false é igual ao bit 1 do sql server pra essa tabela = então será débito no sistema
       strDebCred = rstAplicacao!DebCred
       strTotal_Liquido = IIf(rstAplicacao!TL = True, "Sim", "Não")
       
       If rstAplicacao!DFDebito_Credito_TBfinalizadora = False Then
          txtValor_Finalizadora.Text = Format(CDbl(txtValor_Finalizadora) * -1, "##,##0.00")
       Else
          txtValor_Finalizadora.Text = Format(CDbl(txtValor_Finalizadora), "##,##0.00")
       End If
    End If
    
    rstAplicacao.Close
    Set rstAplicacao = Nothing
    
    If hfgFinalizadora.Rows >= 2 And hfgFinalizadora.TextMatrix(1, 1) <> Empty Then
        hfgFinalizadora.AddItem ""
    End If
    
    hfgFinalizadora.Row = hfgFinalizadora.Rows - 1
    hfgFinalizadora.Text = hfgFinalizadora.Row
    
    'formato
    hfgFinalizadora.Col = 0
    hfgFinalizadora.ColWidth(0) = 500
    hfgFinalizadora.Font.Name = "Tahoma"
    hfgFinalizadora.CellFontSize = 7
    hfgFinalizadora.CellFontBold = False
    hfgFinalizadora.CellBackColor = &H80FFFF
    

    hfgFinalizadora.Col = 1
    hfgFinalizadora.Text = txtFinalizadora.Text
    
    hfgFinalizadora.Col = 2
    hfgFinalizadora.Text = dtcFinalizadora.Text
    
    hfgFinalizadora.Col = 3
    hfgFinalizadora.Text = txtValor_Finalizadora.Text

    hfgFinalizadora.Col = 4
    hfgFinalizadora.Text = IIf(strDebCred = "1", "Sim", "Não")
    
    hfgFinalizadora.Col = 5
    hfgFinalizadora.Text = strDebCred & "|" & dtcFinalizadora.Text

    hfgFinalizadora.Col = 6
    hfgFinalizadora.Text = strTotal_Liquido


    hfgFinalizadora.Col = 0
    txtFinalizadora.Text = Empty
    txtValor_Finalizadora.Text = Empty

    Call Recalcula_Totais_Finalizadora
    Call Recalcula_Totais_Gerais
    Call QuickSortStringsAscending(hfgFinalizadora, 5, 1, hfgFinalizadora.Rows - 1)

    txtFinalizadora.SetFocus

End Sub

Private Sub cmdIncluir_Item_Click()
    
    Dim I As Integer
    
    If txtProduto.Text = Empty Then
        MsgBox "Selecione o produto.", vbExclamation + vbOKOnly, "Only Tech"
        txtProduto.SetFocus
        Exit Sub
    ElseIf txtQuantidade_produto.Text = Empty Or txtQuantidade_produto.Text = "0,00" Then
        MsgBox "Informe a quantidade.", vbExclamation + vbOKOnly, "Only Tech"
        txtQuantidade_produto.SetFocus
        Exit Sub
    ElseIf txtPreco_unitario.Text = Empty Or txtPreco_unitario.Text = "0,00" Then
        MsgBox "Preço unitário do produto inválido. Verifique o cadastro do produto.", vbExclamation + vbOKOnly, "Only Tech"
        txtQuantidade_produto.SetFocus
        Exit Sub
    ElseIf txtUnidade.Text = Empty Then
        MsgBox "Unidade inválida. Verifique o cadastro de produto.", vbExclamation + vbOKOnly, "Only Tech"
        txtProduto.SetFocus
        Exit Sub
    End If
    
    With hfgProduto
    
        'verifica cadastro duplicado
        For I = 1 To .Rows - 1
            If .TextMatrix(I, 1) = txtProduto.Text Then
                MsgBox "Produto já incluso.", vbExclamation + vbOKOnly, "Only Tech"
                txtProduto.SetFocus
                Exit Sub
            End If
        Next I
        
        If .Rows >= 2 And .TextMatrix(.Rows - 1, 1) <> Empty Then
            .AddItem Empty, .Rows
        End If
        
        .Row = .Rows - 1
        .Text = .Row
    
        'formato
        .Col = 0
        .ColWidth(0) = 500
        .Font.Name = "Tahoma"
        .CellFontSize = 7
        .CellFontBold = False
        .CellBackColor = &H80FFFF
        
        .TextMatrix(.Row, 1) = txtProduto.Text
        .TextMatrix(.Row, 2) = dtcProduto.Text
        .TextMatrix(.Row, 3) = Format(txtQuantidade_produto.Text, "##,##0.00")
        .TextMatrix(.Row, 4) = txtUnidade.Text
        .TextMatrix(.Row, 5) = Format(txtPreco_unitario.Text, "##,##0.00")
        .TextMatrix(.Row, 6) = Format(txtTotal_item.Text, "##,##0.00")
        .TextMatrix(.Row, 7) = rstProdutos!FKCodigo_TBSecao
    
        .Col = 0
    
        'AJUSTANDO A SECAO E FINALIZADORA COM RESPECTIVOS TOTAIS
        For I = 1 To hfgSecao.Rows - 1
            'INSERE VALOR EM SEÇÃO JÁ EXISTENTE
            If hfgSecao.TextMatrix(I, 1) = .TextMatrix(.Row, 7) Then
                hfgSecao.TextMatrix(I, 3) = Format(CDbl(hfgSecao.TextMatrix(I, 3)) + CDbl(.TextMatrix(.Row, 6)), "##,##0.00")
                Exit For
            End If
            'INSERE VALOR COM NOVA SEÇÃO
            If I = hfgSecao.Rows - 1 Then
                If hfgSecao.TextMatrix(I, 1) <> Empty Then
                    I = I + 1
                    hfgSecao.AddItem "", I
                End If
                hfgSecao.TextMatrix(I, 1) = .TextMatrix(.Row, 7)
                hfgSecao.TextMatrix(I, 2) = rstProdutos!DFDescricao_TBsecao
                hfgSecao.TextMatrix(I, 3) = .TextMatrix(.Row, 6)
            End If
        Next I
    
    End With
    
    Call Recalcula_Totais_Produtos
    Call Reindexa_Grid(hfgProduto)
    Call Reindexa_Grid(hfgSecao)
    Call Recalcula_Totais_Finalizadora
    Call Recalcula_Totais_Secao
    Call Recalcula_Totais_Gerais

    Call QuickSortStringsAscending(hfgProduto, 2, 1, hfgProduto.Rows - 1)
    Call QuickSortStringsAscending(hfgSecao, 2, 1, hfgSecao.Rows - 1)

    txtProduto.Text = Empty
    txtQuantidade_produto.Text = Empty
    txtUnidade.Text = Empty
    txtPreco_unitario.Text = Empty
    txtTotal_item.Text = Empty
    txtEstoque_Atual.Text = Empty
    'cbbTipo_Preco.Text = Empty
    
    txtProduto.SetFocus
    
End Sub

Private Sub cmdLimpar_Click()
    txtProduto.Text = Empty
    txtQuantidade_produto.Text = Empty
    txtTotal_item.Text = Empty
    txtEstoque_Atual.Text = Empty
    txtPreco_unitario.Text = Empty
    txtUnidade.Text = Empty
    'cbbTipo_Preco.Text = Empty
    txtProduto.SetFocus
End Sub

Private Sub cmdOk_Click()

    Dim strSQL As String
    Dim intContador As Integer

    frmAguarde.Show
    
    If booAlterar = False And Trim(txtOperador.Text) <> Empty Then
        
        If Trim(hfgFinalizadora.TextMatrix(1, 1)) <> Empty Or Trim(hfgSecao.TextMatrix(1, 1)) <> Empty Or Trim(hfgVendedor.TextMatrix(1, 1)) <> Empty Then
            If MsgBox("Deseja que todas as informações de venda e totalizadores sejam atualizadas? Quaisquer modificações serão perdidas.", vbYesNo, "Only Tech") = vbNo Then
               txtOperador.SetFocus
               Exit Sub
            End If
        End If
       
       'Abastecendo as Finalizadoras
       strSQL = "SELECT IXCodigo_TBFinalizadora," & _
                       "DFDescricao_TBFinalizadora," & _
                       "SUM(DFValor_TBOperacao_caixa) as Total," & _
                       "DFDebito_credito_TBFinalizadora, " & _
                       "CONVERT(NVARCHAR, DFDebito_Credito_TBfinalizadora) + '|' + CONVERT(NVARCHAR, DFDescricao_TBfinalizadora) AS Fin_DebCred, " & _
                       "CONVERT(BIT, ISNULL(DFCompoe_total_liquido_TBFinalizadora, 0)) AS TL " & _
                  "FROM TBOperacao_caixa " & _
            "INNER JOIN TBFinalizadora " & _
                    "ON TBOperacao_caixa.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                 "WHERE FKCodigo_TBOperadores_ecf = " & txtOperador.Text & " " & _
                   "AND DFData_TBOperacao_caixa = '" & Format(dtpFechamento.Value, "YYYYMMDD") & "' " & _
                   "AND TBOperacao_caixa.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
                   "AND TBOperacao_caixa.DFStatus_aberto_fechado_TBOperacao_caixa = 0 " & _
              "GROUP BY IXCodigo_TBFinalizadora, DFDescricao_TBFinalizadora, DFDebito_credito_TBFinalizadora "
       'Pegando finalizadora de fechamento de operacao pra abater das demais
       'ISSO SERÁ FEITO PRA SABER COM QUANTO O OPERADOR ANTERIOR ENTREGOU O CAIXA PRA ELA
       'SEM O DINHEIRO DE TROCO ENTREGUE PELO GERENTE PARA A OPERAÇÃO
       strSQL = strSQL & " UNION " & _
            "SELECT IXCodigo_TBFinalizadora, " & _
                   "DFDescricao_TBFinalizadora, " & _
                   "SUM(DFValor_TBOperacao_caixa) as Total, " & _
                   "DFDebito_credito_TBFinalizadora, " & _
                   "CONVERT(NVARCHAR, DFDebito_Credito_TBfinalizadora) + '|' + CONVERT(NVARCHAR, DFDescricao_TBfinalizadora) AS Fin_DebCred " & _
              "FROM TBOperacao_caixa " & _
        "INNER JOIN TBFinalizadora " & _
                "ON TBOperacao_caixa.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
             "WHERE FKCodigo_TBOperadores_ecf = '" & txtOperador.Text & "' " & _
               "AND DFData_TBOperacao_caixa = '" & Format(dtpFechamento.Value, "YYYYMMDD") & "' " & _
               "AND TBOperacao_caixa.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
               "AND IXCodigo_TBFinalizadora IN (SELECT DFFinalizadora_fechamento_operador_TBParametros_ecf FROM TBParametros_ecf WHERE FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ") " & _
          "GROUP BY IXCodigo_TBFinalizadora, " & _
                   "DFDescricao_TBFinalizadora, " & _
                   "DFDebito_credito_TBFinalizadora " & _
          "ORDER BY DFDebito_credito_TBFinalizadora ASC, DFDescricao_TBFinalizadora ASC "
       
        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgFinalizadora, "1000,6000,2000,0,0,0", "Código,Finalizadora,Valor,Debito_credito,Fin_DebCred,Total_Liquido", "BDRetaguarda", "Otica", Me, , 2
        hfgFinalizadora.Col = 0
        hfgFinalizadora.Row = 1
        If hfgFinalizadora.Text = Empty Then
           hfgFinalizadora.Rows = 2
           Movimentacoes.Monta_HFlex_Grid hfgFinalizadora, "1000,6000,2000,0,0,0", "Código,Finalizadora,Valor,Debito_credito,Fin_DebCred,Total_Liquido", 6, "OTICA", Me
        End If
        
        'Abastecendo as Vendas por Grupo - desconsiderando combustíveis (apenas produtos diversos, ou seja, fora da seção de combustíveis)
        strSQL = "SELECT FKCodigo_TBSecao,DFDescricao_TBsecao," & _
                        "SUM(DFValor_total_praticado_TBItens_cupom) as Total " & _
                   "FROM TBItens_cupom " & _
             "INNER JOIN TBCupom " & _
                     "ON TBItens_cupom.FKId_TBCupom = TBCupom.PKId_TBCupom " & _
             "INNER JOIN TBProduto " & _
                     "ON TBItens_cupom.DFCodigo_TBProduto = TBProduto.IXCodigo_TBProduto " & _
             "INNER JOIN TBSecao " & _
                     "ON TBProduto.FKCodigo_TBSecao = TBSecao.PKCodigo_TBSecao " & _
                  "WHERE FKCodigo_TBOperadores_ecf = " & txtOperador.Text & " " & _
                    "AND DFData_Saida_TBCupom = '" & Format(dtpFechamento.Value, "YYYYMMDD") & "' " & _
                    "AND TBCupom.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
                    "AND TBProduto.IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
                    "AND PKId_TBProduto NOT IN (SELECT FKId_TBProduto FROM TBBomba_Bico) " & _
               "GROUP BY FKCodigo_TBSecao, DFDescricao_TBsecao ORDER BY DFDescricao_TBSecao ASC "
                 
        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgSecao, "1000,6000,2000", "Código,Seção,Valor", "BDRetaguarda", "Otica", Me, , 2
        
        hfgSecao.Col = 0
        hfgSecao.Row = 1
        If hfgSecao.Text = Empty Then
           hfgSecao.Rows = 2
           Movimentacoes.Monta_HFlex_Grid hfgSecao, "1000,6000,2000", "Código,Seção,Valor", 3, "OTICA", Me
        End If
        
        'Abastecendo total por vendedores
        strSQL = "SELECT TBcupom.FKId_TBVendedor," & _
                 "DFNome_TBVendedor," & _
                 "SUM(DFTotal_cupom_TBCupom) As DFtotal_cupom " & _
                 "From TBcupom " & _
                 "INNER JOIN TBvendedor " & _
                 "ON TBCupom.FKId_TBVendedor = TBVendedor.PKId_TBVendedor " & _
                 "WHERE DFdata_saida_TBcupom = '" & Format(dtpFechamento.Value, "YYYYMMDD") & "' " & _
                 "AND TBcupom.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
                 "AND FKCodigo_TBOperadores_ecf = " & txtOperador.Text & " " & _
                 "GROUP BY TBcupom.FKId_TBVendedor, DFNome_TBVendedor ORDER BY DFNome_TBVendedor"
        
        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgVendedor, "1000,6000,2000", "Código,Vendedor,Valor", "BDRetaguarda", "Otica", Me, , 2
        
        hfgVendedor.Col = 0
        hfgVendedor.Row = 1
        If hfgVendedor.Text = Empty Then
           hfgVendedor.Rows = 2
           Movimentacoes.Monta_HFlex_Grid hfgVendedor, "1000,6000,2000", "Código,Vendedor,Valor", 3, "OTICA", Me
        End If
        
        'GUIA DE PRODUTOS DE CONFERÊNCIA
        strSQL = "SELECT IXCodigo_TBProduto," & _
                 "DFDescricao_TBProduto,DFQuantidade_TBItens_cupom,DFUnidade_TBItens_cupom," & _
                 "DFPreco_praticado_TBItens_cupom,DFValor_total_item_TBItens_cupom, PKCodigo_TBSecao " & _
                 "FROM TBItens_cupom " & _
                 "INNER JOIN TBCupom " & _
                 "ON TBItens_cupom.FKId_TBCupom = TBCupom.PKId_TBCupom " & _
                 "INNER JOIN TBProduto " & _
                 "ON TBItens_cupom.DFCodigo_TBProduto = TBProduto.IXCodigo_TBProduto " & _
                 "LEFT  JOIN TBsecao " & _
                 "ON TBproduto.FKCodigo_TBSecao = TBsecao.PKCodigo_TBSecao " & _
                 "WHERE TBCupom.FKCodigo_TBOperadores_ecf = " & frmFechamento_caixa_posto.txtOperador.Text & " " & _
                 "AND DFData_Saida_TBCupom = '" & Format(frmFechamento_caixa_posto.dtpFechamento.Value, "YYYYMMDD") & "' " & _
                 "AND TBCupom.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
                 "AND TBProduto.IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
                 "AND TBCupom.DFCancelado_TBCupom = 0 " & _
                 "AND PKId_TBProduto NOT IN (SELECT FKId_TBProduto FROM TBBomba_Bico) " & _
                 "ORDER BY DFNumero_TBCupom"

        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgProduto, "700,4000,1000,600,1300,1500,0", "Código,Produto,Quantidade,UN,Pr. Praticado,Total,Seção", "BDRetaguarda", "Otica", Me, "N", 2
        
        hfgProduto.Col = 0
        hfgProduto.Row = 1
        If hfgProduto.Text = Empty Then
           hfgProduto.Rows = 2
           Movimentacoes.Monta_HFlex_Grid hfgProduto, "700,4000,1000,600,1300,1500,0", "Código,Produto,Quantidade,UN,Pr. Praticado,Total,Seção", 7, "OTICA", Me
        End If
        
        intContador = 1
        txtTotal_Finalizadora.Text = "0,00"
        txtTroco_Recebido.Text = "0,00"
        txtTotal_Vendas.Text = "0,00"
        hfgSecao.Col = 3
        hfgFinalizadora.Col = 3
        
        Do While intContador <> 0
        
           If intContador <= hfgFinalizadora.Rows - 1 Then
              hfgFinalizadora.Row = intContador
              If hfgFinalizadora.Text = Empty Then hfgFinalizadora.Text = "0,00"
              hfgFinalizadora.Col = 4
              If hfgFinalizadora.Text = "Não" Then
                 hfgFinalizadora.Col = 3
                 hfgFinalizadora.Text = Format(CDbl(hfgFinalizadora.Text) * (-1), "##,##0.00")
              Else
                 hfgFinalizadora.Col = 3
              End If
              txtTotal_Finalizadora.Text = CDbl(txtTotal_Finalizadora.Text) + CDbl(hfgFinalizadora.Text)
           End If
           
           If intContador <= hfgSecao.Rows - 1 And intContador <= hfgVendedor.Rows - 1 Then
              hfgSecao.Row = intContador
              hfgVendedor.Row = intContador
              If hfgSecao.Text = Empty Then hfgSecao.Text = "0,00"
              If hfgVendedor.Text = Empty Then hfgVendedor.Text = "0,00"
              txtTotal_Vendas.Text = CDbl(txtTotal_Vendas.Text) + CDbl(hfgSecao.Text) + CDbl(hfgVendedor.Text)
           End If
           
           intContador = intContador + 1
           
           If intContador > hfgSecao.Rows - 1 And intContador > hfgFinalizadora.Rows - 1 And intContador > hfgVendedor.Rows - 1 Then
              intContador = 0
              txtTotal_Finalizadora.Text = Format(txtTotal_Finalizadora.Text, "##,##0.00")
              txtTotal_Vendas.Text = Format(txtTotal_Vendas.Text, "##,##0.00")
              txtSubTotal.Text = txtTotal_Finalizadora.Text
              txtResultado_Caixa.Text = Format(CDbl(txtSubTotal.Text) - CDbl(txtTotal_Vendas.Text), "##,##0.00")
           End If
           
        Loop
        
        hfgSecao.Row = 1: hfgSecao.Col = 0
        hfgVendedor.Row = 1: hfgVendedor.Col = 0
        hfgFinalizadora.Row = 1: hfgFinalizadora.Col = 0
        
        'Limpando o texto do valor do grid se ele estiver preenchido
        If hfgSecao.Text = Empty Then hfgSecao.Row = 1: hfgSecao.Col = 3: hfgSecao.Text = Empty
        If hfgVendedor.Text = Empty Then hfgVendedor.Row = 1: hfgVendedor.Col = 3: hfgVendedor.Text = Empty
        If hfgFinalizadora.Text = Empty Then hfgFinalizadora.Row = 1: hfgFinalizadora.Col = 3: hfgFinalizadora.Text = Empty
        
        hfgSecao.Row = 1: hfgSecao.Col = 0
        hfgVendedor.Row = 1: hfgVendedor.Col = 0
        hfgFinalizadora.Row = 1: hfgFinalizadora.Col = 0
                
        
        If hfgSecao.Rows = 2 And hfgSecao.TextMatrix(1, 1) = Empty Then
            hfgSecao.TextMatrix(1, 1) = strCodSecaoCombustivel
            hfgSecao.TextMatrix(1, 2) = strDescSecaoCombustivel
            If strCodSecaoCombustivel <> Empty Then
                hfgSecao.TextMatrix(1, 3) = dblValorSecaoCombustivel
            End If
        Else
            hfgSecao.AddItem ""
            hfgSecao.TextMatrix(hfgSecao.Rows - 1, 0) = hfgSecao.Rows - 1
            hfgSecao.TextMatrix(hfgSecao.Rows - 1, 1) = strCodSecaoCombustivel
            hfgSecao.TextMatrix(hfgSecao.Rows - 1, 2) = strDescSecaoCombustivel
            hfgSecao.TextMatrix(hfgSecao.Rows - 1, 3) = Format(dblValorSecaoCombustivel, "##,##0.00")
            hfgSecao.ColWidth(0) = 500
            hfgSecao.Font.Name = "Tahoma"
            hfgSecao.CellFontSize = 7
            hfgSecao.CellFontBold = False
            hfgSecao.CellBackColor = &H80FFFF
        End If
        
        Call Recalcula_Totais_Produtos
        Call Recalcula_Totais_Finalizadora
        Call Recalcula_Totais_Secao
        If strCodSecaoCombustivel <> Empty Then
            Call Reindexa_Grid(hfgSecao)
        End If
        Call Recalcula_Totais_Gerais
        
        If txtTotal_Finalizadora.Text <> Empty Or txtTotal_Finalizadora.Text <> "0,00" Then
            fraConferencia.Enabled = True
            fraFinalizadora.Enabled = True
        End If
        
        If txtTroco_Recebido.Text <> Empty Or txtTroco_Recebido.Text <> "0,00" Then
            fraSecao.Enabled = True
        End If
        
        sstFechamento.TabEnabled(1) = True
        
    Else
    
        'Limpa listas de alterações
        Call Limpa_Listas
        
        If Trim(hfgFinalizadora.TextMatrix(1, 1)) <> Empty Or Trim(hfgSecao.TextMatrix(1, 1)) <> Empty Or Trim(hfgVendedor.TextMatrix(1, 1)) <> Empty Then
            If MsgBox("Deseja que todas as informações de venda e totalizadores sejam atualizadas? Quaisquer modificações serão perdidas.", vbYesNo, "Only Tech") = vbNo Then
               txtOperador.SetFocus
               Exit Sub
            End If
        End If
        
        Call hfgFechamento_Click
        
    End If
    
    Call QuickSortStringsAscending(hfgFinalizadora, 5, 1, hfgFinalizadora.Rows - 1)
    Call QuickSortStringsAscending(hfgProduto, 2, 1, hfgProduto.Rows - 1)
    Call QuickSortStringsAscending(hfgSecao, 2, 1, hfgSecao.Rows - 1)
    
    Unload frmAguarde

End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    Call Consulta
End Sub

Private Sub cmdRemover_Finalizadora_Click()

    With hfgFinalizadora

        If .TextMatrix(.Row, 0) = Empty Then
            MsgBox "Selecione a finalizadora que deseja remover.", vbOKOnly + vbExclamation, "Only Tech"
            Exit Sub
        End If
    
        If .Rows > 2 Then
            .RemoveItem .Row
        Else
            Movimentacoes.Monta_HFlex_Grid hfgFinalizadora, "1000,6000,2000,0,0,0", "Código,Finalizadora,Valor,Debito_credito,Fin_DebCred,Total_Liquido", 6, "OTICA", Me
        End If
        
        Call Reindexa_Grid(hfgFinalizadora)
        Call Recalcula_Totais_Finalizadora
        Call Recalcula_Totais_Gerais

    End With

End Sub

Private Sub cmdRemover_Item_Click()

    Dim I As Integer
    Dim J As Integer

    With hfgProduto

        If .TextMatrix(.Row, 1) = Empty Then
            MsgBox "Selecione o produto que deseja remover.", vbOKOnly + vbExclamation, "Only Tech"
            Exit Sub
        End If
        
        'AJUSTANDO A SECAO E FINALIZADORA COM RESPECTIVOS TOTAIS
        For I = 1 To hfgSecao.Rows - 1
            'REMOVE VALOR EM SEÇÃO JÁ EXISTENTE
            If hfgSecao.TextMatrix(I, 1) = .TextMatrix(.Row, 7) Then
                hfgSecao.TextMatrix(I, 3) = Format(CDbl(hfgSecao.TextMatrix(I, 3)) - CDbl(.TextMatrix(.Row, 6)), "##,##0.00")
                If CDbl(hfgSecao.TextMatrix(I, 3)) = 0 Then
                    If hfgSecao.Rows > 2 Then
                        hfgSecao.RemoveItem I
                    Else
                        For J = 0 To hfgSecao.Cols - 1
                            hfgSecao.TextMatrix(1, J) = Empty
                        Next J
                    End If
                End If
                Exit For
            End If
        Next I
    
        If .Rows > 2 Then
            .RemoveItem .Row
        Else
            For I = 0 To .Cols - 1
                .TextMatrix(1, I) = Empty
            Next I
        End If
    
    
    End With
    
    Call Reindexa_Grid(hfgProduto)
    Call Reindexa_Grid(hfgSecao)
    Call Recalcula_Totais_Finalizadora
    Call Recalcula_Totais_Secao
    Call Recalcula_Totais_Gerais

End Sub

Private Sub dtcFinalizadora_GotFocus()
    If txtFinalizadora.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFinalizadora.Text)
    End If
End Sub

Private Sub dtcFinalizadora_LostFocus()
    txtFinalizadora.Text = dtcFinalizadora.BoundText
    If IsNumeric(txtFinalizadora.Text) = False Or dtcFinalizadora.Text = Empty Then txtFinalizadora.Text = Empty: Exit Sub
End Sub

Private Sub dtcOperador_GotFocus()
    If txtOperador.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcOperador.Text)
    End If
End Sub

Private Sub dtcOperador_LostFocus()
    txtOperador.Text = dtcOperador.BoundText
End Sub

Private Sub dtcProduto_GotFocus()
    If txtProduto.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcProduto.Text)
    End If
End Sub

Private Sub dtcProduto_LostFocus()

    txtProduto.Text = dtcProduto.BoundText
    
    If dtcProduto.BoundText <> Empty Then
        rstProdutos.Filter = "IXCodigo_TBproduto = " & txtProduto.Text
    
        If rstProdutos.RecordCount > 0 Then
        
            'unidade varejo ou unidade de venda
            If Not IsNull(rstProdutos!DFUnidade_varejo_TBProduto) Then
                txtUnidade.Text = rstProdutos!DFUnidade_varejo_TBProduto
            Else
                txtUnidade.Text = rstProdutos!DFUnidade_venda_TBProduto
            End If
            
            'Estoque atual
            txtEstoque_Atual.Text = Format(rstProdutos!DFEstoque_Atual_TBProduto, "##,##0.00")
            
            If Trim(cbbTipo_Preco.Text) <> Empty Then
                Call cbbTipo_Preco_LostFocus
            End If
            
        End If
    End If

End Sub

Private Sub dtpFechamento_GotFocus()
    If txtOperador.Text = Empty Then txtOperador.SetFocus
End Sub

Private Sub dtpFechamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = vbKeyTab
    ElseIf KeyCode = vbKeySpace And Shift = 0 Then
        KeyCode = vbKeyRight
    ElseIf KeyCode = vbKeySpace And Shift = 1 Then
        KeyCode = vbKeyLeft
    End If
End Sub

Private Sub dtpFechamento_LostFocus()
    
    If booDigita_Encerrante Then
        cmdConsulta_Encerrante.Enabled = (Trim(txtOperador.Text) <> Empty And IsDate(dtpFechamento.Value))
    Else
        cmdOk.Enabled = (Trim(txtOperador.Text) <> Empty And IsDate(dtpFechamento.Value))
    End If

End Sub

Private Sub dtpFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
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
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
  
Private Sub Form_Load()
    
    On Error GoTo Erro
    
    Dim rstAux As ADODB.Recordset
    
    ReDim arrEncerrantes_Bicos(0, 0)
    
    Set rstProdutos = Nothing
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Fechamento Caixa"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Controle de encerrante
    Set rstAux = New ADODB.Recordset
    Call Movimentacoes.Select_geral("SELECT DFControla_encerrante_TBParametros_venda FROM TBparametros_venda WHERE IXCodigo_TBempresa = '" & MDIPrincipal.OCXUsuario.Empresa & "'", "BDRetaguarda", rstAux, "Otica", Me)
    If rstAux.RecordCount = 0 Then
        booDigita_Encerrante = False
    Else
        booDigita_Encerrante = IIf(IsNull(rstAux!DFControla_encerrante_TBParametros_venda), False, rstAux!DFControla_encerrante_TBParametros_venda)
    End If
    
    If Not rstAux Is Nothing Then
        If Not rstAux.State = adStateClosed Then
            rstAux.Close
        End If
        Set rstAux = Nothing
    End If
   
    If MDIPrincipal.booDesign_time = False Then
        Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstFechamento, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
        booPrivilegio_Incluir = True
        booPrivilegio_Alterar = True
        booPrivilegio_Excluir = True
        booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando Cadastro de Fechamento Caixa"
    'Gravando o log
    log.Gravar_log "Otica", Me
            
    Call Reposicao
    
    dtpInicial.Value = Date
    dtpFinal.Value = Date
    dtpFechamento.Value = Date
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    
    'Controles desabilitados enquanto não digitar operador, data e encerrrante
    cmdConsulta_Encerrante.Enabled = False
    cmdOk.Enabled = False
    fraConferencia.Enabled = False
    fraFinalizadora.Enabled = False
    fraSecao.Enabled = False
    sstFechamento.TabEnabled(1) = False
    
    'tratamento padrão na abertura do form... deveria tratar a tab 1 também, mas ela tem suas particularidades, como acima
    sstFechamento.TabEnabled(0) = False
    sstFechamento.Tab = 2

    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Load")
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    If Not rstProdutos Is Nothing Then
        If rstProdutos.State <> adStateClosed Then
            rstProdutos.Close
        End If
        Set rstProdutos = Nothing
    End If
    
    Unload frmFechamento_caixa_posto_conferencia
    'Unload frmFechamento_caixa_posto_consulta_encerrante
    Unload frmFechamento_caixa_posto_informacoes_adicionais
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    strCombo = Empty
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgFechamento_Click()

    Dim strSQL As String
    Dim rstAux As ADODB.Recordset
    
    Static booExecutando As Boolean
    
    If booExecutando Then
        Exit Sub
    End If
    booExecutando = True
    
    
    If hfgFechamento.Col = 0 And hfgFechamento.Text <> Empty Then
        
        Call Limpa_TXT(Me)
        
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

        strIDFechamento = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 1))
        dtpFechamento.Value = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 2))
        txtOperador.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 3))
        txtTotal_Finalizadora.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 5))
        txtTroco_Recebido.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 6))
        txtTotal_Vendas.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 7))
        txtResultado_Caixa.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 8))
        txtObservacao.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 9))
        
        'Abastecendo o Subtotal
        txtSubTotal.Text = Format(CDbl(txtTotal_Finalizadora.Text) - CDbl(txtTroco_Recebido.Text), "##,##0.00")
        
        'Abastecendo as finalizadoras
        strSQL = "SELECT IXCodigo_TBFinalizadora," & _
                 "DFDescricao_TBFinalizadora," & _
                 "DFValor_total_TBFechamento_caixa_posto_finalizadora, " & _
                 "DFDebito_credito_TBFinalizadora, " & _
                 "CONVERT(NVARCHAR, DFDebito_Credito_TBfinalizadora) + '|' + CONVERT(NVARCHAR, DFDescricao_TBfinalizadora) AS Fin_DebCred, " & _
                 "CONVERT(BIT, ISNULL(DFCompoe_total_liquido_TBFinalizadora, 0)) AS TL " & _
                 "FROM TBFechamento_caixa_posto_finalizadora " & _
                 "INNER JOIN TBFinalizadora ON TBFechamento_caixa_posto_finalizadora.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                 "WHERE FKId_TBFechamento_caixa_posto = " & strIDFechamento & " " & _
                 "ORDER BY 5"
        
        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgFinalizadora, "1000,6000,2000,0,0,0", "Código,Finalizadora,Valor,Debito_credito,Fin_DebCred,Total_Liquido", "BDRetaguarda", "Otica", Me, "N", 2

        hfgFinalizadora.Col = 0
        hfgFinalizadora.Row = 1
        If hfgFinalizadora.Text = Empty Then
           hfgFinalizadora.Rows = 2
           Movimentacoes.Monta_HFlex_Grid hfgFinalizadora, "1000,6000,2000,0,0,0", "Código,Finalizadora,Valor,Debito_credito,Fin_DebCred,Total_Liquido", 6, "OTICA", Me
        End If
        
        'Abastecendo as vendas por grupo
        strSQL = "SELECT FKCodigo_TBSecao,DFDescricao_TBsecao," & _
                 "DFValor_total_TBFechamento_caixa_posto_venda_grupo " & _
                 "FROM TBFechamento_caixa_posto_venda_grupo " & _
                 "INNER JOIN TBSecao ON TBFechamento_caixa_posto_venda_grupo.FKCodigo_TBSecao = TBSecao.PKCodigo_TBSecao " & _
                 "WHERE FKId_TBFechamento_caixa_posto = " & strIDFechamento & " ORDER BY DFDescricao_TBsecao "
        
        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgSecao, "1000,6000,2000", "Código,Seção,Valor", "BDRetaguarda", "Otica", Me, "N", 2
        
        hfgSecao.Col = 0
        hfgSecao.Row = 1
        If hfgSecao.Text = Empty Then
           hfgSecao.Rows = 2
           Movimentacoes.Monta_HFlex_Grid hfgSecao, "1000,6000,2000", "Código,Seção,Valor", 3, "OTICA", Me
        End If
        
        'abastecendo vendedores
        strSQL = "SELECT TBcupom.FKId_TBVendedor," & _
                 "DFNome_TBVendedor," & _
                 "SUM(DFTotal_cupom_TBCupom) As DFtotal_cupom " & _
                 "From TBcupom " & _
                 "INNER JOIN TBvendedor " & _
                 "ON TBCupom.FKId_TBVendedor = TBVendedor.PKId_TBVendedor " & _
                 "WHERE DFdata_saida_TBcupom = '" & Format(dtpFechamento.Value, "YYYYMMDD") & "' " & _
                 "AND TBcupom.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
                 "AND FKCodigo_TBOperadores_ecf = " & txtOperador.Text & " " & _
                 "GROUP BY TBcupom.FKId_TBVendedor, DFNome_TBVendedor ORDER BY DFNome_TBVendedor"
        
        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgVendedor, "1000,6000,2000", "Código,Vendedor,Valor", "BDRetaguarda", "Otica", Me, , 2
        
        hfgVendedor.Col = 0
        hfgVendedor.Row = 1
        If hfgVendedor.Text = Empty Then
            hfgVendedor.Rows = 2
            Movimentacoes.Monta_HFlex_Grid hfgVendedor, "1000,6000,2000", "Código,Vendedor,Valor", 3, "OTICA", Me
        End If
        
        'GUIA DE PRODUTOS DE CONFERÊNCIA
        strSQL = _
        "SELECT IXCodigo_TBProduto, " & _
               "DFDescricao_TBProduto, " & _
               "DFQuantidade_TBFechamento_caixa_produto_posto, " & _
               "ISNULL(ISNULL(DFUnidade_varejo_TBProduto, DFUnidade_venda_TBProduto), '') AS DFunidade, " & _
               "DFValor_unitario_TBFechamento_caixa_produto_posto, " & _
               "DFValor_Total_TBFechamento_caixa_produto_posto, " & _
               "PKCodigo_TBSecao " & _
          "FROM TBFechamento_caixa_produto_posto " & _
    "INNER JOIN TBFechamento_caixa_posto " & _
            "ON TBFechamento_caixa_produto_posto.FKId_TBfechamento_caixa_posto = TBFechamento_caixa_posto.PKId_TBfechamento_caixa_posto " & _
    "INNER JOIN TBProduto " & _
            "ON dbo.TBFechamento_caixa_produto_posto.FKId_TBproduto = TBProduto.PKId_TBproduto " & _
    "LEFT  JOIN TBsecao " & _
            "ON TBproduto.FKCodigo_TBSecao = TBsecao.PKCodigo_TBSecao " & _
         "WHERE PKId_TBfechamento_caixa_posto = " & strIDFechamento & " " & _
      "ORDER BY DFDescricao_TBProduto "

        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgProduto, "700,4000,1000,600,1300,1500,0", "Código,Produto,Quantidade,UN,Pr. Praticado,Total,Seção", "BDRetaguarda", "Otica", Me, "N", 2
        
        hfgProduto.Col = 0
        hfgProduto.Row = 1
        If hfgProduto.Text = Empty Then
           hfgProduto.Rows = 2
           Movimentacoes.Monta_HFlex_Grid hfgProduto, "700,4000,1000,600,1300,1500,0", "Código,Produto,Quantidade,UN,Pr. Praticado,Total,Seção", 7, "OTICA", Me
        End If
        
        If booDigita_Encerrante Then
            'Encerrante
            strSQL = _
            "SELECT PKId_TBbomba_bico, " & _
                   "IXCodigo_Bomba, " & _
                   "IXCodigo_TBBomba_bico, " & _
                   "TBproduto.IXCodigo_TBproduto, " & _
                   "TBproduto.DFdescricao_TBproduto, " & _
                   "DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
                   "DFEncerrante_final_TBEncerrante_caixa_posto, " & _
                   "DFAfericao_TBEncerrante_caixa_posto, " & _
                   "DFQuantidade_DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
                   "DFValor_unitario_DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
                   "DFCusto_unitario_DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
                   "DFValor_total_TBEncerrante_caixa_posto, " & _
                   "PKCodigo_TBsecao, " & _
                   "DFDescricao_TBsecao "
        strSQL = strSQL & _
              "FROM TBbomba_bico " & _
        "INNER JOIN TBbomba " & _
                "ON TBbomba.PKId_TBbomba = TBbomba_bico.FKId_TBbomba " & _
        "INNER JOIN TBEncerrante_caixa_posto " & _
                "ON PKId_TBbomba_bico = FKId_TBbomba_bico " & _
               "AND FKId_TBfechamento_caixa_posto = " & strIDFechamento & " " & _
        "INNER JOIN TBProduto " & _
                "ON TBbomba_bico.FKId_TBproduto = TBProduto.PKId_TBProduto " & _
               "AND TBProduto.IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
        "INNER JOIN TBitens_tabela_preco " & _
                "ON PKId_TBProduto = TBitens_tabela_preco.FKId_TBproduto " & _
               "AND FKCodigo_TBTabela_preco IN ( SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBparametros_venda WHERE IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " ) " & _
        "LEFT  JOIN TBsecao " & _
                "ON TBproduto.FKCodigo_TBsecao = TBsecao.PKCodigo_TBsecao "
            
            'Abastecendo os itens
            Movimentacoes.Movimenta_HFlex_Grid strSQL, objFlex_Aux, "0,600,600,0,1600,1250,1250,950,1100,900,950,1100,0,0", "ID_Bomba_Bico,Bomba,Bico,Cod_Produto,Combustível,Inicial,Final,Aferição,Vendas(L),Pr. Varejo,Custo,Venda ($),Cod_Secao,Desc_Secao", "BDRetaguarda", "Otica", Me, "N", 3
            If setLista_Encerrante_Bico(objFlex_Aux) Then
                Call cmdConsulta_Encerrante_Click
            End If
        
        End If
        
        booAlterar = True
        cmdImprimir_Fechamento.Enabled = True
        txtConsulta.Text = Empty
        
        Call Recalcula_Totais_Finalizadora
        Call Recalcula_Totais_Secao
        Call Recalcula_Totais_Produtos
        Call Recalcula_Totais_Gerais
        
        fraConferencia.Enabled = True
        fraFinalizadora.Enabled = True
        fraSecao.Enabled = True
        
        sstFechamento.TabEnabled(0) = True
        sstFechamento.TabEnabled(1) = True
        sstFechamento.Tab = 0
        Me.txtOperador.SetFocus
   End If
   Unload frmAguarde
   
   booExecutando = False
   
End Sub

Private Sub hfgFechamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgFechamento_Click
    End If
End Sub

Private Sub hfgFinalizadora_EnterCell()
    If hfgFinalizadora.Col <> 3 Then
        Exit Sub
    End If
    hfgFinalizadora.CellFontBold = True
    'hfgFinalizadora.CellForeColor = &H808080
End Sub

Private Sub hfgFinalizadora_GotFocus()
    If hfgFinalizadora.Col = 3 Then
        Call hfgFinalizadora_EnterCell
    End If
End Sub

Private Sub hfgFinalizadora_LostFocus()
    Call hfgFinalizadora_LeaveCell
    Call Recalcula_Totais_Finalizadora
    Call Recalcula_Totais_Gerais
End Sub

Private Sub hfgSecao_DblClick()
    If hfgSecao.Text <> Empty And hfgSecao.Col = 0 And hfgSecao.Row <> 0 Then
        hfgSecao.Col = 1
        Unload frmFechamento_caixa_posto_informacoes_adicionais
        frmAguarde.Show
        frmFechamento_caixa_posto_informacoes_adicionais.Show
        Unload frmAguarde
    End If
End Sub

Private Sub sstFechamento_Click(PreviousTab As Integer)
    If sstFechamento.Tab = 0 Then
        txtOperador.SetFocus
    ElseIf sstFechamento.Tab = 1 Then
        txtProduto.SetFocus
    ElseIf sstFechamento.Tab = 2 Then
        If frmIntegracao.Visible = True Then
            Unload frmIntegracao
        End If
        If strCombo <> Empty And strCombo <> "Todos" And strCombo <> "Data Fechamento" Then
           cbbCampos.Text = strCombo
           txtConsulta.SetFocus
        ElseIf strCombo = "Todos" Then
           hfgFechamento.Row = 1
           hfgFechamento.Col = 0
           hfgFechamento.SetFocus
        ElseIf strCombo = "Data Fechamento" Then
           cbbCampos.Text = strCombo
           dtpInicial.SetFocus
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
    
    Dim strCampo As String
    Dim strValores As String
    Dim strCodigo_Secao As String
    Dim strID_Fechamento_Caixa As String
    Dim strSQL As String
    Dim intContador As Integer
    Dim strIDFinalizadora As String
    
    On Error GoTo Erro
    
    If sstFechamento.Tab = 2 Then Exit Sub
    
    'Verifica se os campos necessarios para gravar não estão nulos
    If txtOperador.Text = Empty Then
       MsgBox "O campo código do Operador não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtOperador.SetFocus
       Exit Sub
    ElseIf txtTroco_Recebido.Text = Empty Then
       MsgBox "O campo Troco Recebido não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtTroco_Recebido.SetFocus
       Exit Sub
    ElseIf txtResultado_Caixa.Text = Empty Then
       MsgBox "O campo Resultado não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtResultado_Caixa.SetFocus
       Exit Sub
    End If
    
    Call Objetos.Retira_Espaco_Lateral(Me)
    Call Objetos.Maiusculo_TXT(Me)

    strCampo = "DFData_TBFechamento_caixa_posto,FKCodigo_TBOperadores_ecf," & _
               "DFTotal_finalizadoras_TBFechamento_caixa_posto,DFTotal_troco_TBFechamento_caixa_posto," & _
               "DFTotal_vendas_grupo_TBFechamento_caixa_posto,DFResultado_TBFechamento_caixa_posto," & _
               "DFObservacao_TBFechamento_caixa_posto,DFData_alteracao_TBFechamento_caixa_posto," & _
               "DFIntegrado_filiais_TBFechamento_caixa_posto"

    If booIntegra_Portal = True Then
       strCampo = strCampo & ",DFIntegrado_portal_TBFechamento_caixa_posto"
    End If
    
    strValores = "'" & Format(dtpFechamento.Value, "YYYYMMDD") & "'," & txtOperador.Text & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtTotal_Finalizadora.Text) & "," & Funcoes_Gerais.Grava_Moeda(txtTroco_Recebido.Text) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtTotal_Vendas.Text) & "," & Funcoes_Gerais.Grava_Moeda(txtResultado_Caixa.Text) & "," & _
                 "'" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "','" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
    End If
    
    'Abrindo conexao
    Conexao.Initial_Catalog = "BDRetaguarda"
    Conexao.Abrir_conexao "Otica"
    Conexao.CNconexao.BeginTrans
       
    On Error GoTo Erro_gravacao
    frmAguarde.Show
    If booAlterar = True Then

       log.Evento = "Alterar"

       strSQL = "UPDATE TBFechamento_caixa_posto " & _
                "SET DFData_TBFechamento_caixa_posto = '" & Format(dtpFechamento.Value, "YYYYMMDD") & "'," & _
                "FKCodigo_TBOperadores_ecf = " & txtOperador.Text & "," & _
                "DFTotal_finalizadoras_TBFechamento_caixa_posto = " & Funcoes_Gerais.Grava_Moeda(txtTotal_Finalizadora.Text) & "," & _
                "DFTotal_troco_TBFechamento_caixa_posto = " & Funcoes_Gerais.Grava_Moeda(txtTroco_Recebido.Text) & "," & _
                "DFTotal_vendas_grupo_TBFechamento_caixa_posto = " & Funcoes_Gerais.Grava_Moeda(txtTotal_Vendas.Text) & "," & _
                "DFResultado_TBFechamento_caixa_posto = " & Funcoes_Gerais.Grava_Moeda(txtResultado_Caixa.Text) & "," & _
                "DFObservacao_TBFechamento_caixa_posto = '" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "', " & _
                "DFData_alteracao_TBFechamento_caixa_posto = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBFechamento_caixa_posto = 0 "

       If booIntegra_Portal = True Then
          strSQL = strSQL & ",DFIntegrado_portal_TBFechamento_caixa_posto = 0"
       End If
          
       strSQL = strSQL & "WHERE PKId_TBFechamento_caixa_posto = " & strIDFechamento
                
       Conexao.CNconexao.Execute strSQL
       
       'Deletando registros antes da nova gravacao'
       strSQL = "DELETE FROM TBFechamento_caixa_posto_finalizadora WHERE FKId_TBFechamento_caixa_posto = " & strIDFechamento & ""
       
       Conexao.CNconexao.Execute strSQL
       
       If hfgFinalizadora.TextMatrix(1, 1) <> Empty Then
          intContador = 1
          Do While intContador <= hfgFinalizadora.Rows - 1
             hfgFinalizadora.Row = intContador
             
             hfgFinalizadora.Col = 1
             strIDFinalizadora = Funcoes_Gerais.Localiza_ID("PKId_TBFinalizadora", "IXCodigo_TBFinalizadora", hfgFinalizadora.Text, "TBFinalizadora", "Otica", Me, "BDRetaguarda")
                
             hfgFinalizadora.Col = 3
           
             strSQL = Empty
             strSQL = "INSERT INTO TBFechamento_caixa_posto_finalizadora (FKId_TBFinalizadora," & _
                      "FKId_TBFechamento_caixa_posto,DFValor_total_TBFechamento_caixa_posto_finalizadora," & _
                      "DFData_alteracao_TBFechamento_caixa_posto_finalizadora," & _
                      "DFIntegrado_filiais_TBFechamento_caixa_posto_finalizadora) "
                      
             If booIntegra_Portal = True Then
                strSQL = Replace(strSQL, ")", Empty)
                strSQL = strSQL & ",DFIntegrado_portal_TBFechamento_caixa_posto_finalizadora) "
             End If
                      
             strSQL = strSQL & " VALUES ( " & strIDFinalizadora & "," & strIDFechamento & "," & _
                               "" & Funcoes_Gerais.Grava_Moeda(hfgFinalizadora.Text) & "," & _
                               "'" & Format(Date, "YYYYMMDD") & "',0 "
                    
             If booIntegra_Portal = True Then
                strSQL = strSQL & ",0"
             End If
             strSQL = strSQL & " ) "
                    
             Conexao.CNconexao.Execute strSQL
             
             intContador = intContador + 1
           Loop
       End If
       
       'Deletando registros antes da nova gravacao'
       strSQL = "DELETE FROM TBFechamento_caixa_posto_venda_grupo WHERE FKId_TBFechamento_caixa_posto = " & strIDFechamento & ""
       
       Conexao.CNconexao.Execute strSQL
       
       hfgSecao.Col = 0: hfgSecao.Row = 1
       If hfgSecao.Text <> Empty Then
          intContador = 1
          Do While intContador <= hfgSecao.Rows - 1
             
             hfgSecao.Row = intContador
             
             hfgSecao.Col = 1
             strCodigo_Secao = hfgSecao.Text
             hfgSecao.Col = 3
                
             strSQL = Empty
             strSQL = "INSERT INTO TBFechamento_caixa_posto_venda_grupo (FKCodigo_TBSecao," & _
                      "FKId_TBFechamento_caixa_posto,DFValor_total_TBFechamento_caixa_posto_venda_grupo," & _
                      "DFData_alteracao_TBFechamento_caixa_posto_venda_grupo," & _
                      "DFIntegrado_filiais_TBFechamento_caixa_posto_venda_grupo ) "
                      
             If booIntegra_Portal = True Then
                strSQL = Replace(strSQL, ")", Empty)
                strSQL = strSQL & ",DFIntegrado_portal_TBFechamento_caixa_posto_venda_grupo ) "
             End If
                      
             strSQL = strSQL & " VALUES ( " & strCodigo_Secao & "," & strIDFechamento & "," & _
                               "" & Funcoes_Gerais.Grava_Moeda(hfgSecao.Text) & "," & _
                               "'" & Format(Date, "YYYYMMDD") & "',0 "
                               
             If booIntegra_Portal = True Then
                strSQL = strSQL & ",0"
             End If
                     
             strSQL = strSQL & " ) "
                     
             Conexao.CNconexao.Execute strSQL

             intContador = intContador + 1
           Loop
       End If
       
       log.Descricao = "Alterando o registro fechamento de data: " & dtpFechamento.Value
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "Otica", Me
       
    Else
    
       log.Evento = "Incluir Novo"
       
       strSQL = "INSERT INTO TBFechamento_caixa_posto (" & strCampo & ") VALUES(" & strValores & ")"
       
       Conexao.CNconexao.Execute strSQL

       hfgFinalizadora.Col = 0: hfgFinalizadora.Row = 1
       If hfgFinalizadora.Text <> Empty Then
          intContador = 1
          Do While intContador <= hfgFinalizadora.Rows - 1
             hfgFinalizadora.Row = intContador
             hfgFinalizadora.Col = 1
             strIDFinalizadora = Funcoes_Gerais.Localiza_ID("PKId_TBFinalizadora", "IXCodigo_TBFinalizadora", hfgFinalizadora.Text, "TBFinalizadora", "Otica", Me, "BDRetaguarda")
                
             hfgFinalizadora.Col = 3
            
             strSQL = Empty
             strSQL = "INSERT INTO TBFechamento_caixa_posto_finalizadora (FKId_TBFechamento_caixa_posto," & _
                      "FKId_TBFinalizadora,DFValor_total_TBFechamento_caixa_posto_finalizadora," & _
                      "DFData_alteracao_TBFechamento_caixa_posto_finalizadora," & _
                      "DFIntegrado_filiais_TBFechamento_caixa_posto_finalizadora"
                      
             If booIntegra_Portal = True Then
                strSQL = strSQL & ",DFIntegrado_portal_TBFechamento_caixa_posto_finalizadora)"
             Else
                strSQL = strSQL & ")"
             End If
                      
             strSQL = strSQL & "SELECT MAX(PKId_TBFechamento_caixa_posto)," & strIDFinalizadora & "," & _
                               "" & Funcoes_Gerais.Grava_Moeda(hfgFinalizadora.Text) & "," & _
                               "'" & Format(Date, "YYYYMMDD") & "',0"
                               
             If booIntegra_Portal = True Then
                strSQL = strSQL & ",0"
             End If
                
             strSQL = strSQL & "FROM TBFechamento_caixa_posto "
             
             Conexao.CNconexao.Execute strSQL
             
             intContador = intContador + 1
             
           Loop
       End If
       
       hfgSecao.Col = 0: hfgSecao.Row = 1
       If hfgSecao.Text <> Empty Then
          intContador = 1
          Do While intContador <= hfgSecao.Rows - 1
             hfgSecao.Row = intContador
             
             hfgSecao.Col = 1
             strCodigo_Secao = hfgSecao.Text
             hfgSecao.Col = 3
               
             strSQL = Empty
             strSQL = "INSERT INTO TBFechamento_caixa_posto_venda_grupo (FKId_TBFechamento_caixa_posto," & _
                      "FKCodigo_TBSecao,DFValor_total_TBFechamento_caixa_posto_venda_grupo) " & _
                      "SELECT MAX(PKId_TBFechamento_caixa_posto)," & _
                      "" & strCodigo_Secao & "," & Funcoes_Gerais.Grava_Moeda(hfgSecao.Text) & " FROM TBFechamento_caixa_posto"
                      
             Conexao.CNconexao.Execute strSQL
             intContador = intContador + 1
           Loop
       End If

    End If
    
    'Recupera fechamento gravado nessa transação
    strID_Fechamento_Caixa = Conexao.CNconexao.Execute("SELECT MAX(PKId_TBfechamento_caixa_posto) AS ID FROM TBfechamento_caixa_posto")!ID
    
    'deleta produtos desse fechamento de caixa e reinsere
    Conexao.CNconexao.Execute "DELETE FROM TBFechamento_caixa_produto_posto WHERE FKId_TBFechamento_caixa_posto = " & strID_Fechamento_Caixa
    For intContador = 1 To hfgProduto.Rows - 1
        If hfgProduto.TextMatrix(intContador, 1) <> Empty Then
            strSQL = "INSERT INTO TBFechamento_caixa_produto_posto ( " & _
                        "FKId_TBFechamento_caixa_posto, " & _
                        "FKId_TBProduto, " & _
                        "DFQuantidade_TBFechamento_caixa_produto_posto, " & _
                        "DFValor_unitario_TBFechamento_caixa_produto_posto, " & _
                        "DFValor_Total_TBFechamento_caixa_produto_posto  "
            
            If booIntegra_Portal = True Then
               strSQL = strSQL & ", DFIntegrado_portal_TBFechamento_caixa_produto_posto ) " & _
               "SELECT " & strID_Fechamento_Caixa & ", (SELECT PKId_TBproduto FROM TBproduto WHERE IXCodigo_TBproduto = " & hfgProduto.TextMatrix(intContador, 1) & " AND IXCODIGO_TBEMPRESA = " & MDIPrincipal.OCXUsuario.Empresa & "), " & Grava_Moeda(hfgProduto.TextMatrix(intContador, 3)) & ", " & Grava_Moeda(hfgProduto.TextMatrix(intContador, 5)) & ", " & Grava_Moeda(hfgProduto.TextMatrix(intContador, 6)) & ", 0 "
            Else
               strSQL = strSQL & ") " & _
               "SELECT " & strID_Fechamento_Caixa & ", (SELECT PKId_TBproduto FROM TBproduto WHERE IXCodigo_TBproduto = " & hfgProduto.TextMatrix(intContador, 1) & " AND IXCODIGO_TBEMPRESA = " & MDIPrincipal.OCXUsuario.Empresa & "), " & Grava_Moeda(hfgProduto.TextMatrix(intContador, 3)) & ", " & Grava_Moeda(hfgProduto.TextMatrix(intContador, 5)) & ", " & Grava_Moeda(hfgProduto.TextMatrix(intContador, 6))
            End If
        
            Conexao.CNconexao.Execute strSQL
        End If
    Next intContador
    
    
    'deleta encerrantes desse fechamento de caixa e reinsere
    If booDigita_Encerrante Then
        Conexao.CNconexao.Execute "DELETE FROM TBEncerrante_caixa_posto WHERE FKId_TBFechamento_caixa_posto = " & strID_Fechamento_Caixa
        For intContador = 0 To UBound(arrEncerrantes_Bicos, 2)
            
            strSQL = "INSERT INTO TBEncerrante_caixa_posto ( " & _
                                 "FKId_TBBomba_bico, " & _
                                 "FKId_TBFechamento_caixa_posto, " & _
                                 "DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
                                 "DFEncerrante_final_TBEncerrante_caixa_posto, " & _
                                 "DFAfericao_TBEncerrante_caixa_posto, " & _
                                 "DFQuantidade_DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
                                 "DFValor_unitario_DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
                                 "DFCusto_unitario_DFEncerrante_inicial_TBEncerrante_caixa_posto, " & _
                                 "DFValor_total_TBEncerrante_caixa_posto "
    
            If booIntegra_Portal = True Then
               strSQL = strSQL & ", DFIntegrado_portal_TBEncerrante_caixa_posto ) " & _
               "VALUES (" & arrEncerrantes_Bicos(1, intContador) & ", " & strID_Fechamento_Caixa & ", " & Grava_Moeda(arrEncerrantes_Bicos(6, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(7, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(8, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(9, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(10, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(11, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(12, intContador)) & ", 0 )"
            Else
               strSQL = strSQL & ") " & _
               "VALUES (" & arrEncerrantes_Bicos(1, intContador) & ", " & strID_Fechamento_Caixa & ", " & Grava_Moeda(arrEncerrantes_Bicos(6, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(7, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(8, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(9, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(10, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(11, intContador)) & ", " & Grava_Moeda(arrEncerrantes_Bicos(12, intContador)) & " )"
            End If
        
            Conexao.CNconexao.Execute strSQL
        Next intContador
    End If
    
    log.Descricao = "Gravando o registro de Fechamento de Data: " & dtpFechamento.Value
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'Fechando a conexao
    Conexao.CNconexao.CommitTrans
    Conexao.Fechar_conexao
    
    Call Objetos.Limpa_TXT(Me)
       
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgFechamento.Visible = False
    End If
    Unload frmFechamento_caixa_posto_conferencia
    Unload frmFechamento_caixa_posto_consulta_encerrante
    Unload frmFechamento_caixa_posto_informacoes_adicionais
    Unload frmAguarde
    sstFechamento.TabEnabled(0) = False
    sstFechamento.TabEnabled(1) = False
    sstFechamento.Tab = 2

    Call Limpa_Listas

    Exit Sub
    
Erro_gravacao:
    'cancelando as alteracoes
    Conexao.CNconexao.RollbackTrans
    'fechando conexao
    Conexao.Fechar_conexao
    Unload frmAguarde
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Sub
    Resume
End Sub

Private Sub Excluir()
    
    Dim strSQL As String
    
    On Error GoTo Erro
    
    If booAlterar = False Then Exit Sub
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro de Fechamento de Data: " & Me.dtpFechamento.Value
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
        
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'abrindo conexao
    Conexao.Abrir_conexao "Otica"
    Conexao.CNconexao.BeginTrans

    'Deletando registros filhos
    strSQL = "DELETE FROM TBFechamento_caixa_posto_finalizadora WHERE FKId_TBFechamento_caixa_posto = " & strIDFechamento & ""
    
    Conexao.CNconexao.Execute strSQL
    
    'Deletando registros filhos
    strSQL = "DELETE FROM TBFechamento_caixa_posto_venda_grupo WHERE FKId_TBFechamento_caixa_posto = " & strIDFechamento & ""
    
    Conexao.CNconexao.Execute strSQL
    
    If booDigita_Encerrante Then
        'Deletando registros filhos
        strSQL = "DELETE FROM TBEncerrante_caixa_posto WHERE FKId_TBFechamento_caixa_posto = " & strIDFechamento & ""
        
        Conexao.CNconexao.Execute strSQL
    End If
    
    'Deletando registros filhos
    strSQL = "DELETE FROM TBFechamento_caixa_produto_posto WHERE FKId_TBFechamento_caixa_posto = " & strIDFechamento & ""
    
    Conexao.CNconexao.Execute strSQL
    
    'Deletando registro
    strSQL = "DELETE FROM TBFechamento_caixa_posto WHERE PKId_TBFechamento_caixa_posto = " & strIDFechamento & ""
    
    Conexao.CNconexao.Execute strSQL
    
    'Fechando a conexao
    Conexao.CNconexao.CommitTrans
    Conexao.Fechar_conexao
    
    Call Limpa_Listas
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
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgFechamento.Visible = False
    End If
    
    sstFechamento.TabEnabled(0) = False
    sstFechamento.TabEnabled(1) = False
    sstFechamento.Tab = 2
    
    Exit Sub
Erro:
    Conexao.CNconexao.RollbackTrans
    Conexao.Fechar_conexao
    
    Call Erro.Erro(Me, "OTICA", "Excluir")
    Exit Sub
    Resume
End Sub

Private Sub Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    
    strCodSecaoCombustivel = Empty
    strDescSecaoCombustivel = Empty
    dblValorSecaoCombustivel = 0
    
    Call Limpa_Listas
    Call Recalcula_Totais_Produtos
    
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
    'tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgFechamento.Visible = False
    End If
    
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Unload frmFechamento_caixa_posto_conferencia
    Unload frmFechamento_caixa_posto_consulta_encerrante
    Unload frmFechamento_caixa_posto_informacoes_adicionais
    
    'Controles desabilitados enquanto não digitar operador, data e encerrrante
    cmdConsulta_Encerrante.Enabled = False
    cmdOk.Enabled = False
    fraConferencia.Enabled = False
    fraFinalizadora.Enabled = False
    fraSecao.Enabled = False
    fraVendedor.Enabled = False
    sstFechamento.TabEnabled(1) = False
    
    sstFechamento.TabEnabled(0) = False
    sstFechamento.Tab = 2
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Sub
End Sub

Private Sub Novo()
    On Error GoTo Erro
   
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
                
    sstFechamento.TabEnabled(0) = True
    'sstFechamento.TabEnabled(1) = True
    sstFechamento.Tab = 0
    
    hfgFinalizadora.Rows = 2
    hfgSecao.Rows = 2
    hfgProduto.Rows = 2
    Movimentacoes.Monta_HFlex_Grid hfgFinalizadora, "1000,6000,2000,0,0,0", "Código,Finalizadora,Valor,Debito_credito,Fin_DebCred,Total_Liquido", 6, "OTICA", Me
    Movimentacoes.Monta_HFlex_Grid hfgSecao, "1000,6000,2000", "Código,Seção,Valor", 3, "OTICA", Me
    Movimentacoes.Monta_HFlex_Grid hfgVendedor, "1000,6000,2000", "Código,Vendedor,Valor", 3, "OTICA", Me
    Movimentacoes.Monta_HFlex_Grid hfgProduto, "700,4000,1000,600,1300,1500,0", "Código,Produto,Quantidade,UN,Pr. Praticado,Total,Seção", 7, "OTICA", Me
    Movimentacoes.Monta_HFlex_Grid objFlex_Aux, "0,600,600,0,1600,1250,1250,950,1100,900,950,1100", "ID_Bomba_Bico,Bomba,Bico,Cod_Produto,Combustível,Inicial,Final,Aferição,Vendas(L),Pr. Varejo,Custo,Venda ($)", 12, "OTICA", Me
    dtpFechamento.Value = Date
    
    
    Call Recalcula_Totais_Produtos
    
    
    txtOperador.SetFocus
    booAlterar = False
    cmdImprimir_Fechamento.Enabled = False
    
    Call Objetos.Limpa_TXT(Me)
    
    cmdOk.Enabled = False
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Sub
    Resume
End Sub

Private Sub txtFinalizadora_Change()
    dtcFinalizadora.BoundText = txtFinalizadora.Text
    If IsNumeric(txtFinalizadora.Text) = False Then txtFinalizadora.Text = Empty: Exit Sub
End Sub

Private Sub txtFinalizadora_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtOperador_Change()
    dtcOperador.BoundText = txtOperador.Text
    If IsNumeric(txtOperador.Text) = False Then
       txtOperador.Text = Empty
    End If
    
    If booDigita_Encerrante Then
        cmdConsulta_Encerrante.Enabled = (Trim(txtOperador.Text) <> Empty And IsDate(dtpFechamento.Value))
    Else
        cmdOk.Enabled = (Trim(txtOperador.Text) <> Empty And IsDate(dtpFechamento.Value))
    End If
End Sub

Private Sub txtOperador_LostFocus()
    If dtcOperador.Text = Empty Then txtOperador.Text = Empty
End Sub

Private Sub txtPreco_unitario_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPreco_unitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_Unitario_LostFocus()
    txtPreco_unitario.Text = Format(txtPreco_unitario.Text, "##,##0.00")
    If txtPreco_unitario.Text = Empty Or txtQuantidade_produto.Text = Empty Then
       txtTotal_item.Text = Empty
    Else
       txtTotal_item.Text = Format(CDbl(txtPreco_unitario.Text) * CDbl(txtQuantidade_produto.Text), "##,##0.00")
    End If
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

Private Sub txtProduto_LostFocus()
    If dtcProduto.Text = Empty Then
       txtProduto.Text = Empty
    End If
    Call dtcProduto_LostFocus
End Sub

Private Sub txtQuantidade_produto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtQuantidade_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtQuantidade_produto_LostFocus()
    txtQuantidade_produto.Text = Format(txtQuantidade_produto.Text, "##,##0.00")
    If txtPreco_unitario.Text <> Empty And txtQuantidade_produto.Text <> Empty Then
       txtTotal_item.Text = Format(CDbl(txtQuantidade_produto.Text) * CDbl(txtPreco_unitario.Text), "##,##0.00")
    Else
       txtTotal_item.Text = Empty
    End If
End Sub

Private Sub txtResultado_Caixa_LostFocus()
    txtResultado_Caixa.Text = Format(txtResultado_Caixa, "##,##0.00")
End Sub

Private Sub Reposicao()
    
    On Error GoTo Erro
    
    Dim strSQL As String
    
    Call Recalcula_Totais_Produtos
    
    strNomes = "ID,Data Fechamento,Código,Operador,Total Finaliz.," & _
               "Troco Recebido,Total Vendas,Resultado,Observação"

    strTamanho = "0,1800,1200,3000,1500," & _
                 "1500,1500,1500,2000"
    
    Movimentacoes.Monta_HFlex_Grid hfgFechamento, strTamanho, strNomes, 9, "OTICA", Me
    Movimentacoes.Monta_HFlex_Grid hfgFinalizadora, "1000,6000,2000,0,0,0", "Código,Finalizadora,Valor,Debito_credito,Fin_DebCred,Total_Liquido", 6, "OTICA", Me
    Movimentacoes.Monta_HFlex_Grid hfgSecao, "1000,6000,2000", "Código,Seção,Valor", 3, "OTICA", Me
    Movimentacoes.Monta_HFlex_Grid hfgVendedor, "1000,6000,2000", "Código,Vendedor,Valor", 3, "OTICA", Me
    Movimentacoes.Monta_HFlex_Grid hfgProduto, "700,4000,1000,600,1300,1500,0", "Código,Produto,Quantidade,UN,Pr. Praticado,Total,Seção", 7, "OTICA", Me
    Movimentacoes.Monta_HFlex_Grid objFlex_Aux, "0,600,600,0,1600,1250,1250,950,1100,900,950,1100", "ID_Bomba_Bico,Bomba,Bico,Cod_Produto,Combustível,Inicial,Final,Aferição,Vendas(L),Pr. Varejo,Custo,Venda ($)", 12, "OTICA", Me
    
    'Para correção do problema de alinhamento no caso de finalizadoras iniciando com numéricos
    hfgFinalizadora.ColAlignment(2) = 1
    hfgSecao.ColAlignment(2) = 1
    hfgProduto.ColAlignment(2) = 1
    hfgVendedor.ColAlignment(2) = 1
    
    strSQL = "SELECT TBOperadores_ecf.PKCodigo_TBOperadores_ecf,TBOperadores_ecf.DFNome_TBOperadores_ecf FROM TBOperadores_ecf WHERE FKCodigo_TBEmpresa  = " & MDIPrincipal.OCXUsuario.Empresa & ""
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT TBFinalizadora.IXCodigo_TBFinalizadora,TBFinalizadora.DFDescricao_TBFinalizadora FROM TBFinalizadora "
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSQL, "BDRetaguarda", "Otica", Me
    
    
    'Combo de Produtos...
    strSQL = "SELECT IXCodigo_TBProduto, DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa  = " & MDIPrincipal.OCXUsuario.Empresa & " AND PKId_TBProduto NOT IN (SELECT FKId_TBProduto FROM TBBomba_Bico)"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSQL, "BDRetaguarda", "Otica", Me
    
    'ABASTECENDO PRODUTOS DE CONFERÊNCIA
    Set rstProdutos = New ADODB.Recordset
    
    strSQL = "SELECT PKId_TBproduto, " & _
                    "IXCodigo_TBProduto, " & _
                    "DFDescricao_TBProduto, " & _
                    "DFUnidade_varejo_TBProduto, " & _
                    "DFUnidade_venda_TBProduto, " & _
                    "DFPreco_venda_TBProduto, " & _
                    "DFPreco_avista_TBItens_tabela_preco, " & _
                    "DFPreco_promocao_TBItens_tabela_preco, " & _
                    "DFPreco_revenda_TBItens_tabela_preco, " & _
                    "DFPreco_especial_TBItens_tabela_preco, " & _
                    "DFPreco_varejo_TBItens_tabela_preco, " & _
                    "DFTipo_preco_TBProduto, " & _
                    "DFEstoque_atual_TBProduto, " & _
                    "FKCodigo_TBSecao, " & _
                    "DFDescricao_TBsecao " & _
               "FROM TBProduto " & _
         "LEFT  JOIN TBItens_tabela_preco " & _
                 "ON PKId_TBproduto = FKID_TBProduto " & _
         "LEFT  JOIN TBsecao " & _
                 "ON FKCodigo_TBSecao = PKCodigo_TBSecao " & _
              "WHERE TBProduto.IXCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' " & _
                "AND FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBParametros_Venda WHERE IXCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' ) " & _
                "AND PKId_TBProduto NOT IN (SELECT FKId_TBProduto FROM TBBomba_Bico) "
    Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstProdutos, "Otica", Me)
    
    Call Monta_Combo
    
    'MONTANDO COMBO DE PREÇOS
    strSQL = "SELECT * FROM TBTipo_preco"
    Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstTipo_Preco, "Otica", Me
     
    If rstTipo_Preco.BOF = True And rstTipo_Preco.EOF = True Then
       MsgBox "Tipo de preço não cadastrado!Verifque.", vbCritical, "Only Tech"
       Set rstTipo_Preco = Nothing
       Exit Sub
    End If
     
    rstTipo_Preco.MoveFirst
    cbbTipo_Preco.Clear
    If IsNull(rstTipo_Preco!DFNome_Preco_avista_TBTipo_preco) = False Then
       cbbTipo_Preco.AddItem ("" & rstTipo_Preco!DFNome_Preco_avista_TBTipo_preco & "")
    End If
    If IsNull(rstTipo_Preco!DFNome_Preco_promocao_TBTipo_preco) = False Then
       cbbTipo_Preco.AddItem ("" & rstTipo_Preco!DFNome_Preco_promocao_TBTipo_preco & "")
    End If
    If IsNull(rstTipo_Preco!DFNome_Preco_revenda_TBTipo_preco) = False Then
       cbbTipo_Preco.AddItem ("" & rstTipo_Preco!DFNome_Preco_revenda_TBTipo_preco & "")
    End If
    If IsNull(rstTipo_Preco!DFNome_Preco_especial_TBTipo_preco) = False Then
       cbbTipo_Preco.AddItem ("" & rstTipo_Preco!DFNome_Preco_especial_TBTipo_preco & "")
    End If
    If IsNull(rstTipo_Preco!DFNome_Preco_varejo_TBTipo_preco) = False Then
       cbbTipo_Preco.AddItem ("" & rstTipo_Preco!DFNome_Preco_varejo_TBTipo_preco & "")
    End If
    
    cbbTipo_Preco.Text = rstTipo_Preco!DFNome_Preco_avista_TBTipo_preco & ""
    
    Set rstTipo_Preco = Nothing

    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Reposicao")
    Exit Sub
    Resume
End Sub

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Sub txtObservacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtObservacao_LostFocus()
    txtObservacao.Text = UCase(txtObservacao.Text)
End Sub

Private Sub Consulta()

    Dim strSQL As String

    If cbbCampos.Text <> "Todos" And cbbCampos.Text <> "Data Fechamento" Then
        If Trim(cbbCampos.Text) = Empty Then
            MsgBox "Selecione um campo para realizar a consulta.", vbInformation + vbOKOnly, "Only Tech"
            cbbCampos.SetFocus
            Exit Sub
        ElseIf Trim(txtConsulta.Text) = Empty Then
            MsgBox "Digite uma informação para realizar a consulta.", vbInformation + vbOKOnly, "Only Tech"
            txtConsulta.SetFocus
            Exit Sub
        End If
    End If
    
    strSQL = "SELECT PKId_TBFechamento_caixa_posto," & _
                    "DFData_TBFechamento_caixa_posto," & _
                    "FKCodigo_TBOperadores_ecf," & _
                    "DFNome_TBOperadores_ecf," & _
                    "DFTotal_finalizadoras_TBFechamento_caixa_posto," & _
                    "DFTotal_troco_TBFechamento_caixa_posto," & _
                    "DFTotal_vendas_grupo_TBFechamento_caixa_posto," & _
                    "DFResultado_TBFechamento_caixa_posto, " & _
                    "DFObservacao_TBFechamento_caixa_posto " & _
               "FROM TBFechamento_caixa_posto " & _
         "INNER JOIN TBOperadores_ecf " & _
                 "ON TBFechamento_caixa_posto.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf "
                              
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = Funcoes_Gerais.Grava_String(txtConsulta.Text)
             
    If cbbCampos.Text <> "Todos" Then
        If cbbCampos.Text = "Data Fechamento" Then
            strSQL = strSQL & " WHERE DFData_TBFechamento_caixa_posto BETWEEN '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                              " AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "'"
        ElseIf cbbCampos.Text = "Código Operador" Then
            If Not IsNumeric(strConsulta) Then
                MsgBox "Informe um código numérico.", vbInformation + vbOKOnly, "Only Tech"
                txtConsulta.Text = Empty
                strConsulta = Empty
                txtConsulta.SetFocus
                Exit Sub
            End If
            strSQL = strSQL & " WHERE FKCodigo_TBOperadores_ecf = '" & strConsulta & "' "
        ElseIf cbbCampos.Text = "Operador" Then
            strSQL = strSQL & " WHERE DFNome_TBOperadores_ecf LIKE '%" & strConsulta & "%' "
        ElseIf cbbCampos.Text = "Total Finaliz." Then
            strSQL = strSQL & " WHERE DFTotal_finalizadoras_TBFechamento_caixa_posto = " & Funcoes_Gerais.Grava_Moeda(strConsulta) & " "
        ElseIf cbbCampos.Text = "Troco Recebido" Then
            strSQL = strSQL & " WHERE DFTotal_troco_TBFechamento_caixa_posto = " & Funcoes_Gerais.Grava_Moeda(strConsulta) & " "
        ElseIf cbbCampos.Text = "Total Vendas" Then
            strSQL = strSQL & " WHERE DFTotal_vendas_grupo_TBFechamento_caixa_posto = " & Funcoes_Gerais.Grava_Moeda(strConsulta) & " "
        ElseIf cbbCampos.Text = "Resultado" Then
            strSQL = strSQL & " WHERE DFResultado_TBFechamento_caixa_posto = " & Funcoes_Gerais.Grava_Moeda(strConsulta) & " "
        ElseIf cbbCampos.Text = "Observação" Then
            strSQL = strSQL & " WHERE DFObservacao_TBFechamento_caixa_posto = '" & strConsulta & "' "
        End If
    End If
    
    strSQL = strSQL & " ORDER BY FKCodigo_TBOperadores_ecf"
    strSQL_Listagem = strSQL
    
    frmAguarde.Show
    DoEvents
           
    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgFechamento, strTamanho, strNomes, "BDRetaguarda", "OTICA", Me, , 2
    
    If Trim(hfgFechamento.TextMatrix(1, 1)) = Empty Then
        hfgFechamento.Rows = 2
        Movimentacoes.Monta_HFlex_Grid hfgFechamento, strTamanho, strNomes, 9, "OTICA", Me
    End If
    
    Unload frmAguarde
    
    hfgFechamento.Refresh
    
    hfgFechamento.Row = 1
    hfgFechamento.Col = 0
    hfgFechamento.SetFocus
    
End Sub

Private Sub Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Data Fechamento")
    cbbCampos.AddItem ("Código Operador")
    cbbCampos.AddItem ("Operador")
    cbbCampos.AddItem ("Total Finaliz.")
    cbbCampos.AddItem ("Troco Recebido")
    cbbCampos.AddItem ("Total Vendas")
    cbbCampos.AddItem ("Resultado")
    cbbCampos.AddItem ("Observação")
End Sub

Private Sub txtTroco_Recebido_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtTroco_Recebido_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtResultado_Caixa_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtResultado_Caixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub Integracao()

    Call frmIntegracao.Verifica_Integracao("PKId_TBFechamento_caixa_posto", strIDFechamento, "DFIntegrado_filiais_TBFechamento_caixa_posto", "TBFechamento_caixa_posto", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBFechamento_caixa_posto", Me.Top, Me.Left, Me.width, Me.Height, "Fechamento Caixa")
    
End Sub

Private Sub Recalcula_Totais_Finalizadora()

    Dim intIdx As Integer
    Dim dblValorCredito As Double
    Dim dblValorDebito As Double
    
    If hfgFinalizadora.Rows = 2 And hfgFinalizadora.TextMatrix(1, 2) = Empty Then
        Exit Sub
    End If
    
    For intIdx = 1 To hfgFinalizadora.Rows - 1
        If hfgFinalizadora.TextMatrix(intIdx, 3) > 0 Then
            dblValorCredito = dblValorCredito + CDbl(hfgFinalizadora.TextMatrix(intIdx, 3))
        Else
            dblValorDebito = dblValorDebito + CDbl(hfgFinalizadora.TextMatrix(intIdx, 3))
        End If
    Next

    txtTotal_Finalizadora.Text = Format(dblValorCredito, "##,##0.00")
    txtTroco_Recebido.Text = Format(dblValorDebito, "##,##0.00")

End Sub

Private Sub Recalcula_Totais_Gerais()
    
    If txtTotal_Finalizadora.Text = Empty Or txtTroco_Recebido.Text = Empty Or txtTotal_Vendas.Text = Empty Then
        Exit Sub
    End If
    
    txtSubTotal.Text = Format(CDbl(txtTotal_Finalizadora.Text) + CDbl(txtTroco_Recebido.Text), "##,##0.00")
    txtResultado_Caixa.Text = Format(CDbl(txtSubTotal.Text) - CDbl(txtTotal_Vendas.Text), "##,##0.00")
    If CDbl(txtResultado_Caixa.Text) > 0 Then
        txtResultado_Caixa.ForeColor = RGB(0, 0, 255)
    ElseIf CDbl(txtResultado_Caixa.Text) < 0 Then
        txtResultado_Caixa.ForeColor = RGB(255, 0, 0)
    Else
        txtResultado_Caixa.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub Limpa_Listas()

    ReDim arrEncerrantes_Bicos(0, 0)

End Sub

Private Sub Reindexa_Grid(ByRef HFlexGrid As MSHFlexGrid)

    Dim I As Integer

    With HFlexGrid
    
        For I = 1 To .Rows - 1
            .Row = I
            .TextMatrix(I, 0) = I
            'formato
            .Col = 0
            .ColWidth(0) = 500
            .Font.Name = "Tahoma"
            .CellFontSize = 7
            .CellFontBold = False
            .CellBackColor = &H80FFFF
        Next I
    
    End With

End Sub

Private Sub EscreveNaGrid(ByRef Grid As MSHFlexGrid, Coluna As Integer, Linha As Integer, Key As Integer, SoNumero As Boolean, Optional ByVal AceitaNegativo As Boolean = True)
    'FUNÇÃO GERADA PARA PERMITIR A INSERÇÃO DIRETO NO GRID
    
    On Error Resume Next
    If Key = 8 Then
        Grid.TextMatrix(Linha, Coluna) = Left(Grid.TextMatrix(Linha, Coluna), Len(Grid.TextMatrix(Linha, Coluna)) - 1)
    ElseIf Key = 13 Then
        CallByName Grid, "LeaveCell", VbMethod
        Grid.Row = Grid.Row + 1
        CallByName Grid, "EnterCell", VbMethod
    Else
        If SoNumero = True Then
           If AceitaNegativo Then
               If Not IsNumeric(Chr(Key)) And (InStr("44,45", Key) = 0) Then Exit Sub
           Else
               If Not IsNumeric(Chr(Key)) And (InStr("44", Key) = 0) Then Exit Sub
           End If
        End If
        
        If InStr(",-", Chr(Key)) > 0 Then
            If InStr(Grid.Text, Chr(Key)) > 0 Then
                Exit Sub
            End If
        End If
        
        'SE cinza, limpa tudo pra começar uma nova digitação... enquanto estiver azul
        If Grid.CellFontBold And Grid.CellForeColor <> vbBlue Then
            Grid.Text = Empty
            Grid.CellForeColor = vbBlue
           'Grid.TextMatrix(Linha, Coluna) = Chr(Key)
        End If
       
        If Key = 45 Then
            If Len(Grid.Text) > 0 Then
                Exit Sub
            End If
        End If
       
        Grid.TextMatrix(Linha, Coluna) = Grid.TextMatrix(Linha, Coluna) & Chr(Key)
       
    End If
End Sub

Private Sub hfgFinalizadora_KeyPress(KeyAscii As Integer)
    If hfgFinalizadora.Rows >= 2 Then
        If hfgFinalizadora.TextMatrix(hfgFinalizadora.Row, 1) <> Empty Then
            If hfgFinalizadora.Col = 3 And hfgFinalizadora.Row > 0 Then
                EscreveNaGrid hfgFinalizadora, 3, hfgFinalizadora.Row, KeyAscii, True
            End If
        End If
    End If
End Sub

Private Sub hfgFinalizadora_LeaveCell()
    
    With hfgFinalizadora
        If .Col <> 3 Then
            Exit Sub
        End If
        
        .Text = Format(.Text, "##,##0.00")
                 
        .CellFontBold = False
        .CellForeColor = &H0
        
    End With
    
    Call Recalcula_Totais_Finalizadora
    Call Recalcula_Totais_Gerais
    
End Sub

Private Sub Recalcula_Totais_Secao(Optional ByVal Digitacao As Boolean = False)

    Dim I As Integer
    Dim J As Integer
    Dim dblValor As Double
    
    If Digitacao <> True Then
        
        For J = 1 To hfgSecao.Rows - 1
            If hfgSecao.TextMatrix(J, 1) = strCodSecaoCombustivel And strCodSecaoCombustivel <> Empty Then
                hfgSecao.TextMatrix(J, 3) = Format(dblValorSecaoCombustivel, "##,##0.00")
            Else
                If hfgSecao.TextMatrix(J, 1) <> Empty Then
                    hfgSecao.TextMatrix(J, 3) = "0,00"
                End If
            End If
        Next J
        
        J = 0
              
        
        If hfgProduto.TextMatrix(1, 1) <> Empty Then
            For I = 1 To hfgProduto.Rows - 1
                            
                For J = 1 To hfgSecao.Rows - 1
                    If hfgSecao.TextMatrix(J, 1) = hfgProduto.TextMatrix(I, 7) Then
                        hfgSecao.TextMatrix(J, 3) = Format(CDbl(hfgSecao.TextMatrix(J, 3)) + CDbl(hfgProduto.TextMatrix(I, 6)), "##,##0.00")
                    End If
                Next J
                
            Next I
        End If
        
    End If
    
    For I = 1 To hfgSecao.Rows - 1
        dblValor = dblValor + CDbl(IIf(hfgSecao.TextMatrix(I, 3) = Empty, 0, hfgSecao.TextMatrix(I, 3)))
    Next I


    txtTotal_Vendas.Text = Format(dblValor, "##,##0.00")

End Sub

Private Sub hfgProduto_KeyPress(KeyAscii As Integer)
    If hfgProduto.Rows >= 2 Then
        If hfgProduto.TextMatrix(hfgProduto.Row, 1) <> Empty Then
            If (hfgProduto.Col = 3 Or hfgProduto.Col = 5) And hfgProduto.Row > 0 Then
                EscreveNaGrid hfgProduto, hfgProduto.Col, hfgProduto.Row, KeyAscii, True, False
            End If
        End If
    End If
End Sub

Private Sub hfgProduto_EnterCell()
    If (hfgProduto.Col <> 3 And hfgProduto.Col <> 5) Then
        Exit Sub
    End If
    
    hfgProduto.CellFontBold = True
    'hfgProduto.CellForeColor = &H808080
End Sub

Private Sub hfgProduto_LeaveCell()
    
    With hfgProduto
        If (.Col <> 3 And .Col <> 5) Then
            Exit Sub
        End If
        
        .Text = Format(.Text, "##,##0.00")
        .CellForeColor = &H0
        
        If .Name = "hfgProduto" Then
            .TextMatrix(.Row, 6) = Format(CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 5)), "##,##0.00")
        End If
        
        Call Recalcula_Totais_Produtos
        Call Recalcula_Totais_Secao(True)
    End With
    
    Call Recalcula_Totais_Gerais
    
End Sub

Private Sub hfgProduto_GotFocus()
    If hfgProduto.Col = 3 Or hfgProduto.Col = 5 Then
        Call hfgProduto_EnterCell
    End If
End Sub

Private Sub hfgProduto_LostFocus()
    Call hfgProduto_LeaveCell
    Call Recalcula_Totais_Produtos
    Call Recalcula_Totais_Secao
    Call Recalcula_Totais_Gerais
End Sub

Private Sub Recalcula_Totais_Produtos()

    Dim I As Integer

    With hfgProduto
    
        lblTotal_Quantidade.Caption = "0,00"
        lblTotal_Vendas.Caption = "0,00"
    
        If .TextMatrix(1, 1) = Empty Then
            lblTotal_Quantidade.Caption = "0,00"
            lblTotal_Vendas.Caption = "0,00"
            Exit Sub
        End If
    
        For I = 1 To .Rows - 1
            lblTotal_Quantidade.Caption = Format(CDbl(lblTotal_Quantidade.Caption) + .TextMatrix(I, 3), "##,##0.00")
            lblTotal_Vendas.Caption = Format(CDbl(lblTotal_Vendas.Caption) + .TextMatrix(I, 6), "##,##0.00")
        Next I
    
    End With

End Sub

Private Sub QuickSortStringsAscending(ByRef objFlex As MSHFlexGrid, ByVal intCol As Integer, ByVal inLow As Long, ByVal inHi As Long)

    'transferir todas as colunas na hora que ele faz a troca

    Dim pivot As String
    Dim tmpSwap As String
    Dim idxTmpSwap As Long
    Dim tmpLow As Long
    Dim tmpHi As Long
    
    'variáveis pra controle de colunas do flex - transferir valores
    ReDim arrApoio(1, objFlex.Cols - 1) As String
    Dim I As Long

    tmpLow = inLow
    tmpHi = inHi

    pivot = objFlex.TextMatrix(((inLow + inHi) \ 2), intCol)

    While (tmpLow <= tmpHi)

        While (objFlex.TextMatrix(tmpLow, intCol) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend

        While (pivot < objFlex.TextMatrix(tmpHi, intCol) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend

        If (tmpLow <= tmpHi) Then
            
            For I = 1 To objFlex.Cols - 1
                arrApoio(1, I) = objFlex.TextMatrix(tmpLow, I)
            Next I
            
            For I = 1 To objFlex.Cols - 1
                objFlex.TextMatrix(tmpLow, I) = objFlex.TextMatrix(tmpHi, I)
                objFlex.TextMatrix(tmpHi, I) = arrApoio(1, I)
            Next I
            
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If

    Wend

    If (inLow < tmpHi) Then QuickSortStringsAscending objFlex, intCol, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSortStringsAscending objFlex, intCol, tmpLow, inHi

End Sub
