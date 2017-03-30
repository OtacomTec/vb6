VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmMovimentacoes_Geracao_Ordem_Servico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Ordem Serviço"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9090
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
      Height          =   1515
      Left            =   90
      TabIndex        =   37
      Top             =   4830
      Width           =   5325
      Begin VB.TextBox txtDesconto_Especial 
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
         Left            =   3540
         TabIndex        =   15
         Top             =   1080
         Width           =   1665
      End
      Begin VB.TextBox txtCofins 
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
         Height          =   300
         Left            =   3540
         TabIndex        =   42
         Top             =   540
         Width           =   1665
      End
      Begin VB.TextBox txtIss 
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
         Height          =   300
         Left            =   1860
         TabIndex        =   41
         ToolTipText     =   "ISS"
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txtPis 
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
         Height          =   300
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtContribuicao_Social 
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
         Height          =   300
         Left            =   1860
         TabIndex        =   39
         Top             =   1080
         Width           =   1635
      End
      Begin VB.TextBox txtImposto_Renda 
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
         Height          =   300
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Imposto de Renda"
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desconto Especial"
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
         Left            =   3540
         TabIndex        =   51
         Top             =   870
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "% Cofins"
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
         Left            =   3540
         TabIndex        =   47
         Top             =   330
         Width           =   660
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "% Contribuição Social"
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
         Left            =   1860
         TabIndex        =   46
         Top             =   870
         Width           =   1560
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "% PIS"
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
         TabIndex        =   45
         Top             =   870
         Width           =   450
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "% Imposto Renda"
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
         TabIndex        =   44
         Top             =   330
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "% ISS"
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
         Left            =   1860
         TabIndex        =   43
         Top             =   330
         Width           =   450
      End
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
      Height          =   1515
      Left            =   5430
      TabIndex        =   28
      Top             =   4830
      Width           =   3555
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Total Produtos................:"
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
         Top             =   285
         Width           =   2070
      End
      Begin VB.Label lblTotal_Produtos 
         Alignment       =   1  'Right Justify
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
         Left            =   2280
         TabIndex        =   35
         ToolTipText     =   "Total bruto dos itens"
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Cofins + PIS + Contr. + IR.:"
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
         TabIndex        =   34
         Top             =   810
         Width           =   2070
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Desconto Especial............:"
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
         TabIndex        =   33
         ToolTipText     =   "Total de IPI  + Total de despesas  acessórios"
         Top             =   555
         Width           =   2070
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Total Nota.....................:"
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
         TabIndex        =   32
         Top             =   1185
         Width           =   2070
      End
      Begin VB.Label lblImpostos 
         Alignment       =   1  'Right Justify
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
         Left            =   2280
         TabIndex        =   31
         ToolTipText     =   "Total de descontos especias"
         Top             =   810
         Width           =   1050
      End
      Begin VB.Label lblDescontos_especiais 
         Alignment       =   1  'Right Justify
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
         Left            =   2280
         TabIndex        =   30
         Top             =   555
         Width           =   1050
      End
      Begin VB.Label lblTotal_Pedido 
         Alignment       =   1  'Right Justify
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
         Left            =   2280
         TabIndex        =   29
         Top             =   1185
         Width           =   1050
      End
      Begin VB.Line Line1 
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   2  'Blackness
         X1              =   90
         X2              =   3400
         Y1              =   1110
         Y2              =   1110
      End
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
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Código do Cliente"
      Top             =   1140
      Width           =   945
   End
   Begin VB.TextBox txtPlano_Pagamento 
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
      Left            =   4590
      TabIndex        =   3
      ToolTipText     =   "Código da condição de pagamento"
      Top             =   1140
      Width           =   945
   End
   Begin VB.CommandButton cmdInformacoes_condicao_pagamento 
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
      Left            =   8610
      Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Projeção de Pagamentos"
      Top             =   1140
      Width           =   375
   End
   Begin VB.CommandButton cmdConsulta_Cliente 
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
      Left            =   4170
      Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":1B0C
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Consulta detalhada do produto "
      Top             =   1140
      Width           =   375
   End
   Begin VB.TextBox txtObservacao 
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
      MaxLength       =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      ToolTipText     =   "Observação"
      Top             =   6600
      Width           =   8895
   End
   Begin VB.Frame frItens_Nota 
      Caption         =   "Itens da Nota"
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
      Height          =   3255
      Left            =   90
      TabIndex        =   17
      Top             =   1530
      Width           =   8895
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
         Left            =   120
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Código do Produto"
         Top             =   450
         Width           =   945
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
         Height          =   315
         Left            =   8010
         Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":1E96
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancelar"
         Top             =   450
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
         Height          =   315
         Left            =   8400
         Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":1FE0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Remove Item"
         Top             =   450
         Width           =   375
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
         Height          =   315
         Left            =   7620
         Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":2522
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Adicionar item"
         Top             =   450
         Width           =   375
      End
      Begin VB.TextBox txtUnidade 
         Alignment       =   2  'Center
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
         Left            =   5130
         TabIndex        =   8
         ToolTipText     =   "Unidade do Item"
         Top             =   450
         Width           =   405
      End
      Begin VB.TextBox txtPreco_unitario 
         Alignment       =   1  'Right Justify
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
         Left            =   5580
         TabIndex        =   9
         ToolTipText     =   "Preço unitário do item"
         Top             =   450
         Width           =   945
      End
      Begin VB.TextBox txtQuantidade_produto 
         Alignment       =   1  'Right Justify
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
         Left            =   4200
         TabIndex        =   7
         ToolTipText     =   "Quantidade do Item"
         Top             =   450
         Width           =   885
      End
      Begin VB.TextBox txtTotal_item 
         Alignment       =   1  'Right Justify
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
         Left            =   6570
         TabIndex        =   10
         ToolTipText     =   "Total do item"
         Top             =   450
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgProduto 
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4048
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
      Begin MSDataListLib.DataCombo dtcProduto 
         Height          =   315
         Left            =   1110
         TabIndex        =   6
         Top             =   450
         Width           =   3045
         _ExtentX        =   5371
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
      Begin VB.Label label15 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
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
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unid."
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
         Left            =   5130
         TabIndex        =   21
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pr. Unitário"
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
         Left            =   5580
         TabIndex        =   20
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quant."
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
         Left            =   4200
         TabIndex        =   19
         Top             =   255
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Item"
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
         Left            =   6570
         TabIndex        =   18
         Top             =   255
         Width           =   735
      End
   End
   Begin MSDataListLib.DataCombo dtcCliente 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Cliente"
      Top             =   1140
      Width           =   3045
      _ExtentX        =   5371
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
   Begin MSDataListLib.DataCombo dtcPlano_Pagamento 
      Height          =   315
      Left            =   5580
      TabIndex        =   4
      ToolTipText     =   "Condição de pagamento"
      Top             =   1140
      Width           =   2985
      _ExtentX        =   5265
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
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   315
      Left            =   90
      TabIndex        =   48
      ToolTipText     =   "Empresa"
      Top             =   570
      Width           =   7695
      _ExtentX        =   13573
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
   Begin AutoCompletar.CbCompleta cbbEmissao 
      Height          =   315
      Left            =   7830
      TabIndex        =   0
      Top             =   570
      Width           =   1155
      _ExtentX        =   2037
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
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":266C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":2986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":2CA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":303A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":33D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Geracao_Ordem_Servico.frx":36EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Emissão"
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
      Left            =   7830
      TabIndex        =   50
      Top             =   360
      Width           =   915
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
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
      Left            =   90
      TabIndex        =   49
      Top             =   360
      Width           =   1050
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
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
      Left            =   90
      TabIndex        =   27
      Top             =   930
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Condição Pagamento"
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
      Left            =   4590
      TabIndex        =   26
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label lblObservacao 
      AutoSize        =   -1  'True
      Caption         =   "Observação"
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
      Left            =   90
      TabIndex        =   23
      Top             =   6390
      Width           =   870
   End
End
Attribute VB_Name = "frmMovimentacoes_Geracao_Ordem_Servico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviço                                                        '
' Objetivo...............: Geração de Ordem de Serviço                                    '
' Data de Criação........: 26/04/2006                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Sá Peixoto                                               '
' Data Criação...........: 26/04/2006                                                     '
' Desenvolvedor..........: Jones Sá Peixoto                                               '
' Data última manutenção.:                                                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim log As New DLLSystemManager.log
Dim strsql As String
Dim strTamanhos As String
Dim strNomes As String
Dim intContador As Integer
'RECORDSETS
Dim rstAplicacao As New ADODB.Recordset
Dim rstPlano_Pagamento As New ADODB.Recordset
Dim rstCliente As New ADODB.Recordset
'CONEXOES
Dim cnGravacao As New DLLConexao_Sistema.conexao
'VARIAVEIS DE GRAVACAO
Dim intIDVendedor As String
Dim intIDPlano As String
Dim intCodigo_Transportadora As Integer
Dim lngIDCfop As Integer
Dim strUnidade As String
Dim lngIDProduto As String
Dim strCST1 As String
Dim strCST2 As String
'VALORES PARAMETRO
Dim strValor_Min_IR As String
Dim strValor_Min_Contribuicao As String
Option Explicit

Public Sub cbbEmissao_Click()

    'FORMATANDO A TELA DE ACORDO COM O TIPO DE EMISSAO
    If cbbEmissao.Text = "Lote" Then
    
       lblCliente.Caption = "Ramo de Atividade"
       txtCliente.Text = Empty
       txtPlano_Pagamento.Text = Empty
       
       strsql = "SELECT PKCodigo_TBRamo_atividade,DFDescricao_TBRamo_atividade FROM TBRamo_atividade"
       Movimentacoes.Movimenta_DataCombo "PKCodigo_TBRamo_atividade", "DFDescricao_TBRamo_atividade", dtcCliente, strsql, "BDRetaguarda", "Otica", Me
    
       txtPlano_Pagamento.Enabled = False
       dtcPlano_Pagamento.Enabled = False
       
       lblObservacao.Top = 1530
       txtObservacao.Top = 1740
       
       Me.Height = 2510
       frItens_Nota.Visible = False

    Else
    
       lblCliente.Caption = "Cliente"
       txtCliente.Text = Empty
       
       strsql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente " & _
                "FROM TBCliente " & _
                "INNER JOIN TBContrato_cliente ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente  " & _
                "WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
        
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strsql, "BDRetaguarda", "Otica", Me
       
       txtPlano_Pagamento.Enabled = True
       dtcPlano_Pagamento.Enabled = True
       
       lblObservacao.Top = 6390
       txtObservacao.Top = 6600
       
       Me.Height = 7395
       frItens_Nota.Visible = True

    End If
    
End Sub

Private Sub cmdIncluir_Item_Click()

    Dim strIndice As String
    Dim strTotal As String
    Dim strQuantidade As String
    
    If txtProduto.Text = Empty Then
       MsgBox "Produto inválido. Verifique.", vbInformation, "Only Tech"
       txtProduto.SetFocus
       Exit Sub
    ElseIf txtQuantidade_produto.Text = Empty Or txtQuantidade_produto.Text = "0,00" Then
       MsgBox "Quantidade inválida. Verifique.", vbInformation, "Only Tech"
       txtQuantidade_produto.SetFocus
       Exit Sub
    ElseIf txtUnidade.Text = Empty Then
       MsgBox "Unidade de Produto inválida. Verifique o cadastro de produto.", vbInformation, "Only Tech"
       txtProduto.SetFocus
       Exit Sub
    ElseIf txtPreco_unitario.Text = Empty Then
       MsgBox "Preço unitário inválido. Verifique o cadastro de produto.", vbInformation, "Only Tech"
       txtPreco_unitario.SetFocus
       Exit Sub
    End If
    
    'VERIFICACAO QUANTO AO NUMERO DE ITENS INCLUIDOS NA ORDEM DE SERVICO, O MAXIMO SETADO VIA RELATORIO É 4
    If hfgProduto.Rows >= 5 Then
       MsgBox "Esta Ordem já excedeu o número máximo de serviços permitidos. Verifique.", vbInformation, "Only Tech"
       Exit Sub
    End If
    
    'Verifica se o produto já foi incluso para o cupom de acerto, se sim, ele é somado ao já incluso
    intContador = 1
    Do While intContador <= hfgProduto.Rows - 1
       hfgProduto.Row = intContador
       hfgProduto.Col = 1
       If hfgProduto.Text = txtProduto.Text Then
          MsgBox "Produto já incluído. Verifique.", vbInformation, "Only Tech"
          txtQuantidade_produto.Text = Empty
          txtUnidade.Text = Empty
          txtPreco_unitario.Text = Empty
          txtTotal_item.Text = Empty
          txtProduto.SetFocus
          txtProduto.SetFocus
          Exit Sub
       End If
       intContador = intContador + 1
    Loop
    
    hfgProduto.Row = 1
    hfgProduto.Col = 0
    If hfgProduto.Text <> Empty Then
       strIndice = hfgProduto.Rows
       hfgProduto.Rows = hfgProduto.Rows + 1
    Else
       strIndice = 1
    End If
    
    hfgProduto.Row = strIndice
    
    hfgProduto.Col = 0
    hfgProduto.ColWidth(0) = 380
    hfgProduto.Font.Name = "Tahoma"
    hfgProduto.CellFontSize = 7
    hfgProduto.CellFontBold = False
    hfgProduto.CellBackColor = &H80FFFF
    hfgProduto.Text = strIndice
    
    hfgProduto.Col = 1
    hfgProduto.Text = txtProduto.Text
    hfgProduto.Col = 2
    hfgProduto.Text = dtcProduto.Text
    hfgProduto.Col = 3
    hfgProduto.Text = txtQuantidade_produto.Text
    hfgProduto.Col = 4
    hfgProduto.Text = txtUnidade.Text
    hfgProduto.Col = 5
    hfgProduto.Text = txtPreco_unitario.Text
    hfgProduto.Col = 6
    hfgProduto.Text = txtTotal_item.Text

    Call Calcula_Resumos
    
    txtProduto.Text = Empty
    txtQuantidade_produto.Text = Empty
    txtUnidade.Text = Empty
    txtPreco_unitario.Text = Empty
    txtTotal_item.Text = Empty
    
    txtProduto.SetFocus
End Sub

Private Sub cmdRemover_Item_Click()

    If hfgProduto.Col <> 0 Or hfgProduto.Text = Empty Then
       MsgBox "Não há produto selecionado para exclusão. Verifique", vbInformation, "Only Tech"
       txtProduto.SetFocus
       Exit Sub
    End If

    If hfgProduto.Rows <= 2 Then
       hfgProduto.Clear
       Movimentacoes.Monta_HFlex_Grid hfgProduto, strTamanhos, strNomes, 6, "OTICA", Me
    Else
       hfgProduto.RemoveItem (hfgProduto.Row)
       intContador = 1
       hfgProduto.Col = 0
       Do While intContador <= hfgProduto.Rows - 1
          hfgProduto.Row = intContador
          hfgProduto.Text = intContador
          intContador = intContador + 1
       Loop
    End If
    Call Calcula_Resumos
End Sub

Private Sub dtcEmpresa_Change()
    txtProduto.Text = Empty: txtCliente.Text = Empty: txtPlano_Pagamento.Text = Empty
End Sub

Private Sub dtcEmpresa_LostFocus()

    If Not IsNumeric(dtcEmpresa.BoundText) Then dtcEmpresa.Text = Empty
    If IsNumeric(dtcEmpresa.Text) Then dtcEmpresa.Text = Empty

    If dtcEmpresa.Text <> Empty Then
       strsql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strsql, "BDRetaguarda", "Otica", Me
        
       If cbbEmissao.Text = "Individual" Then
          strsql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente " & _
                   "FROM TBCliente " & _
                   "INNER JOIN TBContrato_cliente ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente  " & _
                   "WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
                   
          Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strsql, "BDRetaguarda", "Otica", Me
       End If
       
       strsql = "SELECT IXCodigo_TBPlano_Pagamento,DFDescricao_TBPlano_Pagamento FROM TBPlano_Pagamento WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBPlano_Pagamento", "DFDescricao_TBPlano_Pagamento", dtcPlano_Pagamento, strsql, "BDRetaguarda", "Otica", Me
    End If
    
    dtcEmpresa.Enabled = False
    
End Sub

Private Sub dtcProduto_GotFocus()
    If txtProduto.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcProduto.Text)
    End If
End Sub

Private Sub dtcProduto_LostFocus()

    txtProduto.Text = dtcProduto.BoundText
    If txtProduto.Text <> Empty Then
       strsql = "SELECT DFUnidade_varejo_TBProduto " & _
                "FROM TBProduto " & _
                "WHERE TBProduto.IXCodigo_TBProduto = " & txtProduto.Text & " " & _
                "AND TBProduto.IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " "
                
       Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
       
       If rstAplicacao.RecordCount <> 0 Then
          If IsNull(rstAplicacao.Fields("DFUnidade_varejo_TBProduto")) = False Then
             txtUnidade.Text = rstAplicacao.Fields("DFUnidade_varejo_TBProduto")
          Else
             txtUnidade.Text = Empty
             txtTotal_item.Text = Empty
          End If
       Else
          txtUnidade.Text = Empty
          txtTotal_item.Text = Empty
       End If
       Set rstAplicacao = Nothing
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 71: Call Gravar      'CTRL+G
                       Case 67: Call Cancelar    'CTRL+C
                       Case 83: Unload Me        'CTRL+S
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
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Geração de Ordem de Serviço"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Geração de Ordem de Serviço"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    'Montando os datacombo de tela
    strsql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strsql, "BDRetaguarda", "Otica", Me

    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    strsql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strsql, "BDRetaguarda", "Otica", Me
    
    strsql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente " & _
             "FROM TBCliente " & _
             "INNER JOIN TBContrato_cliente ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente  " & _
             "WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strsql, "BDRetaguarda", "Otica", Me
    
    strsql = "SELECT IXCodigo_TBPlano_Pagamento,DFDescricao_TBPlano_Pagamento FROM TBPlano_Pagamento WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBPlano_Pagamento", "DFDescricao_TBPlano_Pagamento", dtcPlano_Pagamento, strsql, "BDRetaguarda", "Otica", Me
    
    'MONTANDO GRID DE PRODUTOS
    strTamanhos = "800,3600,950,350,1200,1100"
    strNomes = "Código,Produto,Quantidade,UN,Pr. Praticado,Total"

    Monta_HFlex_Grid hfgProduto, strTamanhos, strNomes, 6, "Otica", Me
    
    cbbEmissao.Clear
    cbbEmissao.AddItem ("Lote")
    cbbEmissao.AddItem ("Individual")
    cbbEmissao.Text = "Lote"
    Call cbbEmissao_Click
    
    'ABASTECENDO IMPOSTOS
    strsql = "SELECT DFPercentual_iss_TBParametros_fiscais," & _
             "DFPercentual_irrf_TBParametros_fiscais," & _
             "DFPercentual_contribuicao_social_TBParametros_fiscais," & _
             "DFPercentual_cofins_TBParametros_fiscais," & _
             "DFPercentual_pis_TBParametros_fiscais," & _
             "DFValor_minimo_calculo_irrf_TBParametros_fiscais," & _
             "DFValor_minimo_calculo_contribuicao_TBParametros_fiscais " & _
             "FROM TBParametros_fiscais " & _
             "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    
    Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       If IsNull(rstAplicacao.Fields("DFPercentual_irrf_TBParametros_fiscais")) = False Then
          txtImposto_Renda.Text = Format(rstAplicacao.Fields("DFPercentual_irrf_TBParametros_fiscais"), "#,###0.00")
       Else
          txtImposto_Renda.Text = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFPercentual_iss_TBParametros_fiscais")) = False Then
          txtIss.Text = Format(rstAplicacao.Fields("DFPercentual_iss_TBParametros_fiscais"), "#,###0.00")
       Else
          txtIss.Text = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFPercentual_contribuicao_social_TBParametros_fiscais")) = False Then
          txtContribuicao_Social.Text = Format(rstAplicacao.Fields("DFPercentual_contribuicao_social_TBParametros_fiscais"), "#,###0.00")
       Else
          txtContribuicao_Social.Text = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFPercentual_cofins_TBParametros_fiscais")) = False Then
          txtCofins.Text = Format(rstAplicacao.Fields("DFPercentual_cofins_TBParametros_fiscais"), "#,###0.00")
       Else
          txtCofins.Text = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFPercentual_pis_TBParametros_fiscais")) = False Then
          txtPis.Text = Format(rstAplicacao.Fields("DFPercentual_pis_TBParametros_fiscais"), "#,###0.00")
       Else
          txtPis.Text = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFValor_minimo_calculo_irrf_TBParametros_fiscais")) = False Then
          strValor_Min_IR = Format(rstAplicacao.Fields("DFValor_minimo_calculo_irrf_TBParametros_fiscais"), "#,###0.00")
       Else
          strValor_Min_IR = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFValor_minimo_calculo_contribuicao_TBParametros_fiscais")) = False Then
          strValor_Min_Contribuicao = Format(rstAplicacao.Fields("DFValor_minimo_calculo_contribuicao_TBParametros_fiscais"), "#,###0.00")
       Else
          strValor_Min_Contribuicao = "0,00"
       End If
    Else
       txtImposto_Renda.Text = "0,00"
       txtIss.Text = "0,00"
       txtCofins.Text = "0,00"
       txtPis.Text = "0,00"
       txtContribuicao_Social.Text = "0,00"
       strValor_Min_IR = "0,00"
       strValor_Min_Contribuicao = "0,00"
    End If
    
    Set rstAplicacao = Nothing
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro

    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
        
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Unload")
    Exit Sub
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Gravar
           Case 2: Call Cancelar
           Case 4: Unload Me
    End Select
End Sub

Private Sub txtCliente_Change()
    dtcCliente.BoundText = txtCliente.Text
    If IsNumeric(txtCliente.Text) = False Then txtCliente.Text = Empty: Exit Sub
End Sub

Private Sub txtCliente_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub dtcCliente_GotFocus()
    If txtCliente.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCliente.Text)
    End If
End Sub

Private Sub dtcCliente_LostFocus()
    txtCliente.Text = dtcCliente.BoundText
    If IsNumeric(txtCliente.Text) = False Or dtcCliente.Text = Empty Then
       txtCliente.Text = Empty: Exit Sub
    Else
      If cbbEmissao.Text = "Individual" Then
       
          strsql = "SELECT IXCodigo_TBPlano_pagamento " & _
                   "FROM TBCliente " & _
                   "INNER JOIN TBPlano_Pagamento ON TBCliente.FKId_TBPlano_pagamento = TBPlano_Pagamento.PKId_TBPlano_pagamento " & _
                   "WHERE IXCodigo_TBCliente = " & txtCliente.Text & " " & _
                   "AND TBCliente.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
          
          Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
          
          If rstAplicacao.RecordCount <> 0 And IsNull(rstAplicacao.Fields("IXCodigo_TBPlano_pagamento")) = False Then
             txtPlano_Pagamento.Text = rstAplicacao.Fields("IXCodigo_TBPlano_pagamento")
          End If
          Set rstAplicacao = Nothing
       End If
    End If
End Sub

Private Sub txtCliente_LostFocus()
    If dtcCliente.Text = Empty Then
       txtCliente.Text = Empty
    Else
       Call dtcCliente_LostFocus
    End If
End Sub

Private Sub txtPlano_Pagamento_Change()
    dtcPlano_Pagamento.BoundText = txtPlano_Pagamento.Text
    If IsNumeric(txtPlano_Pagamento.Text) = False Then txtPlano_Pagamento.Text = Empty: Exit Sub
End Sub

Private Sub txtPlano_Pagamento_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub dtcPlano_Pagamento_GotFocus()
    If txtPlano_Pagamento.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcPlano_Pagamento.Text)
    End If
End Sub

Private Sub dtcPlano_Pagamento_LostFocus()
    txtPlano_Pagamento.Text = dtcPlano_Pagamento.BoundText
    If IsNumeric(txtPlano_Pagamento.Text) = False Or dtcPlano_Pagamento.Text = Empty Then txtPlano_Pagamento.Text = Empty: Exit Sub
End Sub

Private Sub txtContribuicao_Social_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtContribuicao_Social_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtContribuicao_Social_LostFocus()
    txtContribuicao_Social.Text = Format(txtContribuicao_Social.Text, "#,###0.00")
End Sub

Private Sub txtDesconto_Especial_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDesconto_Especial_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtDesconto_Especial_LostFocus()
    txtDesconto_Especial.Text = Format(txtDesconto_Especial.Text, "#,###0.00")
    Call Calcula_Resumos
End Sub

Private Sub txtImposto_Renda_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtImposto_Renda_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtImposto_Renda_LostFocus()
    txtImposto_Renda.Text = Format(txtImposto_Renda.Text, "#,###0.00")
End Sub

Private Sub txtIss_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtIss_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtIss_LostFocus()
    txtIss.Text = Format(txtIss.Text, "#,###0.00")
End Sub

Private Sub txtCofins_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCofins_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtCofins_LostFocus()
    txtCofins.Text = Format(txtCofins.Text, "#,###0.00")
End Sub

Private Sub txtObservacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtObservacao_LostFocus()
    txtObservacao.Text = UCase(txtObservacao.Text)
End Sub

Private Sub txtPis_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPis_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Or KeyAscii = Asc("-") Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPis_LostFocus()
    txtPis.Text = Format(txtPis.Text, "#,###0.00")
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
    txtPreco_unitario.Text = Format(txtPreco_unitario.Text, "#,###0.00")
    If txtPreco_unitario.Text = Empty Or txtQuantidade_produto.Text = Empty Then
       txtTotal_item.Text = Empty
    Else
       txtTotal_item.Text = Format(CDbl(txtPreco_unitario.Text) * CDbl(txtQuantidade_produto.Text), "#,###0.00")
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
       txtUnidade.Text = Empty
       txtPreco_unitario.Text = Empty
       txtTotal_item.Text = Empty
    End If
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
    txtQuantidade_produto.Text = Format(txtQuantidade_produto.Text, "#,###0.00")
    If txtPreco_unitario.Text <> Empty And txtQuantidade_produto.Text <> Empty Then
       txtTotal_item.Text = Format(CDbl(txtQuantidade_produto.Text) * CDbl(txtPreco_unitario.Text), "#,###0.00")
    Else
       txtTotal_item.Text = Empty
    End If
End Sub

Private Function Gravar()
    
    If txtCliente.Text = Empty And cbbEmissao.Text = "Individual" Then
       MsgBox "O Campo código do Cliente não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtCliente.SetFocus
       Exit Function
    ElseIf txtPlano_Pagamento.Text = Empty And cbbEmissao.Text = "Individual" Then
       MsgBox "O Campo código do Plano de Pagamento não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtCliente.SetFocus
       Exit Function
    End If
    
    'VENDEDOR
    strsql = "SELECT PKID_TBVendedor " & _
             "FROM TBVendedor " & _
             "WHERE IXCodigo_TBVendedor = 9999 " & _
             "AND IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    
    Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       intIDVendedor = rstAplicacao.Fields("PKID_TBVendedor")
    Else
       MsgBox "É necessário que conste no sistema um vendedor de código 9999. A operação está impossibilitada de continuar.", vbInformation, "Only Tech"
       Set rstAplicacao = Nothing
       Exit Function
    End If
    
    Set rstAplicacao = Nothing
    
    'TRANSPORTADORA
    strsql = "SELECT PKCodigo_TBTransportadora " & _
             "FROM TBTransportadora " & _
             "WHERE PKCodigo_TBTransportadora = 9999"
    
    Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       intCodigo_Transportadora = rstAplicacao.Fields("PKCodigo_TBTransportadora")
    Else
       MsgBox "É necessário que conste no sistema uma transportadora de código 9999. A operação está impossibilitada de continuar.", vbInformation, "Only Tech"
       Set rstAplicacao = Nothing
       Exit Function
    End If
    
    Set rstAplicacao = Nothing
    
    
    Call Grava_Corpo_Nota
    
End Function

Private Function Cancelar()

    lblTotal_Produtos.Caption = Empty
    lblTotal_Pedido.Caption = Empty
    lblDescontos_especiais.Caption = Empty
    lblImpostos.Caption = Empty
    
    Call Limpa_TXT(Me)
    
    hfgProduto.Rows = 2
    Monta_HFlex_Grid hfgProduto, strTamanhos, strNomes, 6, "Otica", Me

    'ABASTECENDO IMPOSTOS
    strsql = "SELECT DFPercentual_iss_TBParametros_fiscais," & _
             "DFPercentual_irrf_TBParametros_fiscais," & _
             "DFPercentual_contribuicao_social_TBParametros_fiscais," & _
             "DFPercentual_cofins_TBParametros_fiscais," & _
             "DFPercentual_pis_TBParametros_fiscais " & _
             "FROM TBParametros_fiscais " & _
             "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    
    Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       If IsNull(rstAplicacao.Fields("DFPercentual_irrf_TBParametros_fiscais")) = False Then
          txtImposto_Renda.Text = Format(rstAplicacao.Fields("DFPercentual_irrf_TBParametros_fiscais"), "#,###0.00")
       Else
          txtImposto_Renda.Text = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFPercentual_iss_TBParametros_fiscais")) = False Then
          txtIss.Text = Format(rstAplicacao.Fields("DFPercentual_iss_TBParametros_fiscais"), "#,###0.00")
       Else
          txtIss.Text = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFPercentual_contribuicao_social_TBParametros_fiscais")) = False Then
          txtContribuicao_Social.Text = Format(rstAplicacao.Fields("DFPercentual_contribuicao_social_TBParametros_fiscais"), "#,###0.00")
       Else
          txtContribuicao_Social.Text = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFPercentual_cofins_TBParametros_fiscais")) = False Then
          txtCofins.Text = Format(rstAplicacao.Fields("DFPercentual_cofins_TBParametros_fiscais"), "#,###0.00")
       Else
          txtCofins.Text = "0,00"
       End If
       If IsNull(rstAplicacao.Fields("DFPercentual_pis_TBParametros_fiscais")) = False Then
          txtPis.Text = Format(rstAplicacao.Fields("DFPercentual_pis_TBParametros_fiscais"), "#,###0.00")
       Else
          txtPis.Text = "0,00"
       End If
    Else
       txtImposto_Renda.Text = "0,00"
       txtIss.Text = "0,00"
       txtCofins.Text = "0,00"
       txtPis.Text = "0,00"
       txtContribuicao_Social.Text = "0,00"
    End If
    
    Set rstAplicacao = Nothing

End Function

Private Function Grava_Corpo_Nota()
    Dim dblIndenizacao As Double
    Dim dblDesconto_Especial As Double
    Dim dblTotal_Pedido As Double
    Dim dblValor_Contrato As Double
    Dim intEmitente As Integer
    Dim intCodigo_Tabela_Vigente As Integer
    Dim intDia_Vencimento As Integer
    Dim dblImpostos As Double
    Dim strObservacao As String
    Dim datVencimento As Date
    
    frmAguarde.Show
    
    'BUSCANDO INFORMACOES PERTINENTES PARA GRAVAÇÃO
    
    'TABELA VIGENTE
    strsql = "SELECT DFNumero_tabela_vigente_TBParametros_venda " & _
             "FROM TBParametros_venda " & _
             "WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    
    Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       intCodigo_Tabela_Vigente = rstAplicacao.Fields("DFNumero_tabela_vigente_TBParametros_venda")
    End If
    
    Set rstAplicacao = Nothing
    
    'CFOP PARAMETRO FISCAL
    strsql = "SELECT PKID_TBCfop " & _
             "FROM TBParametros_fiscais " & _
             "INNER JOIN TBCFOP " & _
             "ON TBParametros_fiscais.DFProximo_cfop_venda_dentro_estado_TBParametros_fiscais = TBCFOP.DFCodigo_TBCfop " & _
             "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    
    Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       lngIDCfop = rstAplicacao.Fields("PKID_TBCfop")
    End If
    
    Set rstAplicacao = Nothing
    
    'INFORMACOES DO CLIENTE E DO CONTRATO
    strsql = "SELECT PKId_TBCliente,IXCodigo_TBCliente,FKId_TBPlano_pagamento," & _
             "DFDia_vencimento_TBCliente," & _
             "DFValor_TBContrato_cliente,DFDescricao_TBPlano_pagamento " & _
             "FROM TBCliente " & _
             "INNER JOIN TBContrato_cliente ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente  " & _
             "INNER JOIN TBPlano_Pagamento ON TBCliente.FKId_TBPlano_pagamento = TBPlano_pagamento.PKId_TBPlano_pagamento " & _
             "INNER JOIN TBCidade_Otica ON TBCliente.FKID_TBCidade_Otica = TBCidade_Otica.PKID_TBCidade_Otica " & _
             "WHERE TBCliente.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    
    If cbbEmissao.Text = "Individual" Then
       
       strsql = strsql + " AND IXCodigo_TBCliente = " & txtCliente.Text & ""
       
       Select_geral strsql, "BDRetaguarda", rstCliente, "Otica", Me
    
       'PLANO DE PAGAMENTO
       strsql = "SELECT PKId_TBPlano_pagamento,DFDigita_vencimento_TBPlano_pagamento " & _
                "FROM TBPlano_pagamento " & _
                "WHERE IXCodigo_TBPlano_pagamento = " & txtPlano_Pagamento.Text & " " & _
                "AND IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
       
       Select_geral strsql, "BDRetaguarda", rstPlano_Pagamento, "Otica", Me
       
       If rstPlano_Pagamento.RecordCount <> 0 Then
          intIDPlano = rstPlano_Pagamento.Fields("PKId_TBPlano_pagamento")
       End If

    Else
    
       If txtCliente.Text <> Empty Then
          strsql = strsql + " AND TBCliente.FKCodigo_TBRamo_atividade = " & txtCliente.Text & " "
       End If
       
       Select_geral strsql, "BDRetaguarda", rstCliente, "Otica", Me
       
       'ID DO PRODUTO PADRAO
       strsql = "SELECT FKId_contrato_TBProduto,DFUnidade_venda_TBProduto,DFCst1_TBProduto, " & _
                "DFCst2_TBProduto " & _
                "FROM TBParametros_servicos " & _
                "INNER JOIN TBProduto ON TBParametros_servicos.FKId_contrato_TBProduto = TBProduto.PKId_TBProduto " & _
                "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
       
       Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
        
       If rstAplicacao.RecordCount <> 0 Then
          lngIDProduto = rstAplicacao.Fields("FKId_contrato_TBProduto")
          strUnidade = rstAplicacao.Fields("DFUnidade_venda_TBProduto")
          strCST1 = rstAplicacao.Fields("DFCst1_TBProduto")
          strCST2 = rstAplicacao.Fields("DFCst2_TBProduto")
       Else
          Unload frmAguarde
          MsgBox "Produto Padrão não definido no cadastro de Parâmetros de Serviços. A operação está impossibilitada de continuar.", vbInformation, "Only Tech"
          Set rstAplicacao = Nothing
          Set rstPlano_Pagamento = Nothing
          Set rstCliente = Nothing
          Exit Function
       End If
        
       Set rstAplicacao = Nothing
    
    End If
    
    On Error GoTo Erro
    
    'ABRINDO CONEXAO
    cnGravacao.Initial_Catalog = "BDRetaguarda"
    cnGravacao.Abrir_conexao "Otica"
    cnGravacao.CNConexao.BeginTrans
    
    Do While rstCliente.EOF = False
         
       '''''''''''CAPTURANDO O DIA DE VENCIMENTO'''''''''''''''
       intDia_Vencimento = Format(rstCliente.Fields("DFDia_vencimento_TBCliente"), "00")
       
       If intDia_Vencimento = 0 Then
          intDia_Vencimento = 15
       End If
        
       If intDia_Vencimento <= Format(Now, "DD") Then
          datVencimento = intDia_Vencimento & "/" & Format(DateAdd("M", 1, Now), "MM/YYYY")
       Else
          datVencimento = intDia_Vencimento & "/" & Format(Now, "MM/YYYY")
       End If
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       If cbbEmissao.Text = "Individual" Then
       
          dblDesconto_Especial = lblDescontos_especiais.Caption
          
          If txtObservacao.Text = Empty Then
             strObservacao = "VENC. " & datVencimento
          Else
             strObservacao = txtObservacao.Text
          End If
          
          dblValor_Contrato = CDbl(lblTotal_Produtos.Caption)
          
          dblImpostos = Format(lblImpostos.Caption, "#,###0.00")
          
       Else
       
          'PLANO DE PAGAMENTO
          intIDPlano = rstCliente.Fields("FKId_TBPlano_pagamento")
          
          'MONTANDO OBSERVACAO
          strObservacao = "VENC. " & datVencimento
          
          dblDesconto_Especial = 0
          
          If IsNull(rstCliente.Fields("DFValor_TBContrato_cliente")) = False Then
             dblValor_Contrato = rstCliente.Fields("DFValor_TBContrato_cliente")
          Else
             dblValor_Contrato = 0
          End If
          
          dblImpostos = 0
          
          'Montando o total de Impostos
          If dblValor_Contrato > CDbl(strValor_Min_IR) Then
             dblImpostos = Format(CDbl(txtImposto_Renda.Text), "#,###0.00")
          End If
          
          If dblValor_Contrato > CDbl(strValor_Min_Contribuicao) Then
             dblImpostos = Format(dblImpostos + CDbl(txtCofins.Text) + CDbl(txtContribuicao_Social.Text) + CDbl(txtPis.Text), "#,###0.00")
          End If
          
          dblImpostos = Format(dblValor_Contrato * CDbl(dblImpostos) / 100, "#,###0.00")
          
       End If
       
       'CALCULANDO TOTAL DO PEDIDO
       dblTotal_Pedido = Format(dblValor_Contrato - dblImpostos - dblDesconto_Especial, "#,###0.00")
        
       'TIPO EMITENTE
       intEmitente = 0
       
       'GRAVANDO CORPO DO PEDIDO
       strsql = "INSERT INTO TBPedido(FKCodigo_TBEmpresa, " & _
                "FKCodigo_TBTabela_preco, " & _
                "FKId_TBVendedor," & _
                "FKId_TBPlano_pagamento," & _
                "FKCodigo_TBTransportadora," & _
                "DFTipo_operacao_TBPedido," & _
                "DFEmitente_TBPedido," & _
                "DFTotal_itens_TBPedido," & _
                "DFTotal_pedido_TBPedido," & _
                "DFTotal_pedido_tabelaTBPedido," & _
                "DFDesconto_especial_TBPedido," & _
                "DFDesconto_indenizacao_TBPedido," & _
                "DFData_Digitacao_TBPedido," & _
                "DFUsuario_TBPedido," & _
                "DFFaturado_TBPedido," & _
                "DFPrevisao_TBPedido," & _
                "DFValor_ipi_TBPedido," & _
                "DFBloqueado_TBPedido," & _
                "DFDespesas_acessorias_TBPedido,"
        
        strsql = strsql + "DFTotal_descontos_itens_TBPedido," & _
                "DFTotal_peso_liquido_TBPedido," & _
                "DFTotal_peso_bruto_TBPedido," & _
                "DFTipo_emitente_TBPedido, " & _
                "DFObservacao_TBPedido,DFBase_calculo_subst_tributaria_TBPedido," & _
                "DFValor_subst_tributaria_TBPedido,DFValor_Frete_TBPedido,DFTipo_Frete_TBPedido) " & _
                "VALUES (" & _
                " " & dtcEmpresa.BoundText & "," & _
                " " & intCodigo_Tabela_Vigente & "," & _
                " " & intIDVendedor & "," & _
                " " & intIDPlano & "," & _
                " " & intCodigo_Transportadora & "," & _
                " " & 1 & "," & _
                " " & rstCliente.Fields("IXCodigo_TBCliente") & " ," & _
                " " & Funcoes_Gerais.Grava_Moeda(dblValor_Contrato) & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(dblTotal_Pedido) & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(dblValor_Contrato) & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(dblDesconto_Especial) & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(dblImpostos) & "," & _
                " '" & Format(Now, "YYYYMMDD") & "'," & _
                " '" & MDIPrincipal.OCXUsuario.Nome & "'," & _
                " " & 0 & "," & _
                " " & 0 & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                " " & 0 & ","

        strsql = strsql + " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                " " & intEmitente & "," & _
                " '" & Funcoes_Gerais.Grava_String(strObservacao) & "'," & _
                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                " " & Funcoes_Gerais.Grava_Moeda(0) & ",0)"
        
        'Gravando o corpo do Pedido
        cnGravacao.CNConexao.Execute strsql
        
        Call Grava_Itens

        'Gravando o CFO na tabela CFO-PEDIDO
        strsql = "INSERT INTO TBCfop_pedido(FKId_TBCfop,FKId_TBPedido) " & _
                 "SELECT " & lngIDCfop & ",MAX(PKID_TBPedido) FROM TBPedido "
                 
        cnGravacao.CNConexao.Execute strsql
    
        rstCliente.MoveNext

    Loop
    
    cnGravacao.CNConexao.CommitTrans
    cnGravacao.Fechar_conexao
    
    Unload frmAguarde
    
    MsgBox "" & rstCliente.RecordCount & " Ordem(ns) gravada(s) corretamente."
    
    Set rstCliente = Nothing
    Set rstAplicacao = Nothing
    Set rstPlano_Pagamento = Nothing

    Call Cancelar
    
    Exit Function
Erro:
    'TRATAMENTO DE ERRO
    Unload frmAguarde
    cnGravacao.CNConexao.RollbackTrans
    cnGravacao.Fechar_conexao
    Call Erro.Erro(Me, "Otica", "Load")
End Function

Private Function Grava_Itens()

    Dim intRotina As Integer
    Dim dblValor_Item As Double
    Dim dblTotal_Item As Double
    Dim dblQuantidade_Item As Double
    
    If cbbEmissao.Text = "Individual" Then
       intRotina = hfgProduto.Rows - 1
    Else
       intRotina = 1
    End If
    
    intContador = 1
    
    Do While intContador <= intRotina
    
       If cbbEmissao.Text = "Individual" Then
          
          hfgProduto.Row = intContador
          hfgProduto.Col = 3
          dblQuantidade_Item = CDbl(hfgProduto.Text)
          hfgProduto.Col = 4
          strUnidade = hfgProduto.Text
          hfgProduto.Col = 5
          dblValor_Item = CDbl(hfgProduto.Text)
          hfgProduto.Col = 6
          dblTotal_Item = CDbl(hfgProduto.Text)
          
          hfgProduto.Col = 1
          strsql = "SELECT PKID_TBProduto " & _
                   "FROM TBProduto " & _
                   "WHERE IXCodigo_TBProduto = " & hfgProduto.Text & " " & _
                   "AND IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
          
          Select_geral strsql, "BDRetaguarda", rstAplicacao, "Otica", Me
          
          If rstAplicacao.RecordCount <> 0 Then
             lngIDProduto = rstAplicacao.Fields("PKID_TBProduto")
          End If
          
          Set rstAplicacao = Nothing
          
       Else
          dblValor_Item = rstCliente.Fields("DFValor_TBContrato_cliente")
          dblQuantidade_Item = 1
          dblTotal_Item = dblValor_Item
       End If
       
       strsql = "INSERT INTO TBItens_pedido(" & _
                "FKId_TBPedido," & _
                "FKId_TBProduto," & _
                "FKId_TBCfop," & _
                "DFCst1_TBItens_pedido," & _
                "DFCst2_TBItens_pedido," & _
                "DFQuantidade_TBItens_pedido," & _
                "DFTipo_preco_TBItens_pedido," & _
                "DFPreco_tabela_TBItens_pedido," & _
                "DFPercentual_desconto_TBItens_pedido," & _
                "DFPreco_praticado_TBItens_pedido," & _
                "DFValor_total_tabela_TBItens_pedido," & _
                "DFValor_total_praticado_TBItens_pedido," & _
                "DFPercentual_icms_TBItens_pedido," & _
                "DFValor_total_icms_TBItens_pedido," & _
                "DFUnidade_TBItens_pedido," & _
                "DFPeso_liquido_TBItens_pedido," & _
                "DFPeso_bruto_TBItens_pedido," & _
                "DFQuantidade_baixa_estoque_TBItens_pedido," & _
                "DFDivisor_baixa_estoque_TBItens_pedido," & _
                "FKId_TBVendedor,"
                 
        strsql = strsql + "DFValor_total_item_TBItens_pedido,DFBase_calculo_subst_tributaria_TBItens_pedido," & _
                          "DFValor_subst_tributaria_TBItens_pedido,DFValor_cotacao_dia_TBItens_pedido) " & _
                          "SELECT " & _
                          "MAX(PKID_TBPedido)," & _
                          "" & lngIDProduto & "," & _
                          "" & lngIDCfop & "," & _
                          "'" & strCST1 & "'," & _
                          "'" & strCST2 & "'," & _
                          "" & Funcoes_Gerais.Grava_Moeda(dblQuantidade_Item) & "," & _
                          "" & 1 & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(dblValor_Item) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(dblValor_Item) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Item) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Item) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                          "'" & strUnidade & "'," & _
                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(dblQuantidade_Item) & "," & _
                          "" & 1 & "," & _
                          "" & intIDVendedor & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Item) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(0) & ","
        
        strsql = strsql + "" & Funcoes_Gerais.Grava_Moeda(0) & "," & _
                          "" & Funcoes_Gerais.Grava_Moeda(0) & " " & _
                          "FROM TBPedido "
                           
        'Gravando o item do Pedido
        cnGravacao.CNConexao.Execute strsql
        
        intContador = intContador + 1
    Loop
    
End Function

Private Function Calcula_Resumos()

    Dim dblImpostos As Double
    Dim dblTotal_Produtos As Double
    
    hfgProduto.Col = 6
    intContador = 1
    Do While intContador <= hfgProduto.Rows - 1
       hfgProduto.Row = intContador
       If hfgProduto.Text <> Empty Then
          dblTotal_Produtos = dblTotal_Produtos + hfgProduto.Text
       End If
       intContador = intContador + 1
    Loop
    
    If dblTotal_Produtos > CDbl(strValor_Min_IR) Then
       dblImpostos = CDbl(txtImposto_Renda.Text)
    End If
    
    If dblTotal_Produtos > CDbl(strValor_Min_Contribuicao) Then
       dblImpostos = Format(dblImpostos + CDbl(txtCofins.Text) + CDbl(txtContribuicao_Social.Text) + CDbl(txtPis.Text), "#,###0.00")
    End If
    
    lblDescontos_especiais.Caption = txtDesconto_Especial.Text
    lblTotal_Produtos.Caption = Format(dblTotal_Produtos, "#,###0.00")
    
    If lblTotal_Produtos.Caption = Empty Then lblTotal_Produtos.Caption = "0,00"
    If lblDescontos_especiais.Caption = Empty Then lblDescontos_especiais.Caption = "0,00"
    If lblImpostos.Caption = Empty Then lblImpostos.Caption = "0,00"
    If lblTotal_Pedido.Caption = Empty Then lblTotal_Pedido.Caption = "0,00"
            
    lblImpostos.Caption = Format(CDbl(dblImpostos) * CDbl(lblTotal_Produtos.Caption) / 100, "#,###0.00")

    lblTotal_Pedido.Caption = Format(CDbl(lblTotal_Produtos.Caption) - CDbl(lblImpostos.Caption) - CDbl(lblDescontos_especiais.Caption), "#,###0.00")
    
End Function

