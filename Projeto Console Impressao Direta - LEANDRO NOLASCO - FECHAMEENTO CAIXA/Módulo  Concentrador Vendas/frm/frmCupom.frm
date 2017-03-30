VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCupom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cupom Fiscal"
   ClientHeight    =   7290
   ClientLeft      =   2790
   ClientTop       =   2355
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCupom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10125
   Begin VB.PictureBox OCXUsuario 
      Height          =   480
      Left            =   11010
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   68
      Top             =   1380
      Width           =   1200
   End
   Begin TabDlg.SSTab sstCupon 
      Height          =   6945
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   12250
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
      TabCaption(0)   =   "&Geral"
      TabPicture(0)   =   "frmCupom.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCliente_Fornecedor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblSituacao"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label20"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dtpHora"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dtpData"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dtcCodigo_ecf"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dtcEmitente"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cbbTipo_operacao"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "dtcVendedor"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "dtcEmpresa"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtVendedor"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtNumero_cupom"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtSerie"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCliente"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "freItens"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdObservacao"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdCancelamento"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame3"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtCodigo_PDV"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "dtcOperador"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtOperador"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtCod_COO"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmCupom.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRefresh"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdConsulta"
      Tab(1).Control(2)=   "txtConsulta"
      Tab(1).Control(3)=   "hfgCupom_fiscal"
      Tab(1).Control(4)=   "cbbCampos"
      Tab(1).Control(5)=   "dtpData_incial"
      Tab(1).Control(6)=   "dtpData_final"
      Tab(1).Control(7)=   "lblAte"
      Tab(1).Control(8)=   "Label6"
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtCod_COO 
         Enabled         =   0   'False
         Height          =   360
         Left            =   7020
         MaxLength       =   3
         TabIndex        =   13
         ToolTipText     =   "Código do Ponto de Venda"
         Top             =   1440
         Width           =   1830
      End
      Begin VB.TextBox txtOperador 
         Height          =   360
         Left            =   120
         MaxLength       =   3
         TabIndex        =   17
         ToolTipText     =   "Código do Operador"
         Top             =   2100
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtcOperador 
         Height          =   360
         Left            =   1500
         TabIndex        =   18
         Top             =   2100
         Width           =   3585
         _ExtentX        =   6324
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
      Begin VB.TextBox txtCodigo_PDV 
         Enabled         =   0   'False
         Height          =   360
         Left            =   5130
         MaxLength       =   3
         TabIndex        =   11
         ToolTipText     =   "Código do Ponto de Venda"
         Top             =   1440
         Width           =   1830
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   -65430
         Picture         =   "frmCupom.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   405
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   -65820
         Picture         =   "frmCupom.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resumo Financeiro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   120
         TabIndex        =   48
         Top             =   5640
         Width           =   8025
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Número de Itens..:"
            Height          =   240
            Left            =   4710
            TabIndex        =   55
            Top             =   225
            Width           =   1620
         End
         Begin VB.Label label23 
            AutoSize        =   -1  'True
            Caption         =   "Desconto Itens......:"
            Height          =   240
            Left            =   120
            TabIndex        =   51
            Top             =   465
            Width           =   1695
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Total Produtos......:"
            Height          =   240
            Left            =   120
            TabIndex        =   49
            Top             =   225
            Width           =   1680
         End
         Begin VB.Label lblDesconto_Itens 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1860
            TabIndex        =   52
            ToolTipText     =   "Total dos descontos por item"
            Top             =   465
            Width           =   2670
         End
         Begin VB.Label lblTotal_Produtos 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1860
            TabIndex        =   50
            ToolTipText     =   "Total bruto dos itens"
            Top             =   225
            Width           =   2670
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Total Cupom.........:"
            Height          =   240
            Left            =   120
            TabIndex        =   53
            Top             =   825
            Width           =   1710
         End
         Begin VB.Label lblTotal_Pedido 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1890
            TabIndex        =   54
            Top             =   825
            Width           =   2640
         End
         Begin VB.Line Line1 
            BorderStyle     =   6  'Inside Solid
            DrawMode        =   2  'Blackness
            X1              =   120
            X2              =   4560
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label lblTotal_Itens 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   6405
            TabIndex        =   56
            Top             =   225
            Width           =   1410
         End
      End
      Begin VB.CommandButton cmdCancelamento 
         Caption         =   "Cancelamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         TabIndex        =   57
         Top             =   5730
         Width           =   1725
      End
      Begin VB.CommandButton cmdObservacao 
         Caption         =   "Observação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         TabIndex        =   58
         Top             =   6330
         Width           =   1725
      End
      Begin VB.Frame freItens 
         Caption         =   "Itens do Cupom (F9)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         TabIndex        =   29
         Top             =   3210
         Width           =   9885
         Begin VB.CommandButton cmdConsulta_produto 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4140
            Picture         =   "frmCupom.frx":44F6
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Consulta detalhada do produto "
            Top             =   420
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
            Height          =   360
            Left            =   8610
            Picture         =   "frmCupom.frx":4880
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Adicionar item"
            Top             =   420
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
            Height          =   360
            Left            =   9390
            Picture         =   "frmCupom.frx":49CA
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Remove Item"
            Top             =   420
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
            Height          =   360
            Left            =   9000
            Picture         =   "frmCupom.frx":4F0C
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Cancelar"
            Top             =   420
            Width           =   375
         End
         Begin VB.TextBox txtProduto 
            Height          =   360
            Left            =   90
            MaxLength       =   6
            TabIndex        =   31
            ToolTipText     =   "Código do Produto"
            Top             =   450
            Width           =   765
         End
         Begin VB.TextBox txtQuantidade_produto 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   4560
            TabIndex        =   35
            ToolTipText     =   "Quantidade do Item"
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox txtPreco_unitario 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   6060
            TabIndex        =   39
            ToolTipText     =   "Preço unitário do item"
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox txtPercentual_desconto 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   6810
            TabIndex        =   41
            ToolTipText     =   "Percentual de desconto do item"
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox txtTotal_Item 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   360
            Left            =   7560
            TabIndex        =   43
            ToolTipText     =   "Total do item"
            Top             =   450
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dtcProduto 
            Height          =   360
            Left            =   900
            TabIndex        =   32
            ToolTipText     =   "Produto"
            Top             =   450
            Width           =   3195
            _ExtentX        =   5636
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgItem_cupom 
            Height          =   1485
            Left            =   90
            TabIndex        =   47
            Top             =   840
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   2619
            _Version        =   393216
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
            _Band(0).Cols   =   2
         End
         Begin MSDataListLib.DataCombo dtcUnidade_venda 
            Height          =   360
            Left            =   5310
            TabIndex        =   37
            Top             =   450
            Width           =   735
            _ExtentX        =   1296
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unid."
            Height          =   240
            Left            =   5310
            TabIndex        =   36
            Top             =   210
            Width           =   435
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto"
            Height          =   240
            Left            =   120
            TabIndex        =   30
            Top             =   210
            Width           =   660
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde"
            Height          =   240
            Left            =   4560
            TabIndex        =   34
            Top             =   210
            Width           =   405
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pr. Unit."
            Height          =   240
            Left            =   6060
            TabIndex        =   38
            Top             =   210
            Width           =   690
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%Desc."
            Height          =   240
            Left            =   6810
            TabIndex        =   40
            Top             =   210
            Width           =   645
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Item"
            Height          =   240
            Left            =   7560
            TabIndex        =   42
            Top             =   210
            Width           =   885
         End
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Left            =   120
         MaxLength       =   6
         TabIndex        =   23
         ToolTipText     =   "Código do Cliente"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtSerie 
         Height          =   360
         Left            =   4380
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1440
         Width           =   690
      End
      Begin VB.TextBox txtNumero_cupom 
         Height          =   360
         Left            =   2850
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1440
         Width           =   1470
      End
      Begin VB.TextBox txtVendedor 
         Height          =   360
         Left            =   5130
         MaxLength       =   4
         TabIndex        =   20
         ToolTipText     =   "Código do Vendedor"
         Top             =   2100
         Width           =   1335
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72120
         TabIndex        =   61
         Top             =   780
         Width           =   6195
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgCupom_fiscal 
         Height          =   5565
         Left            =   -74880
         TabIndex        =   64
         Top             =   1230
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   9816
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   -74880
         TabIndex        =   60
         Top             =   780
         Width           =   2715
         _ExtentX        =   4789
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
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   9885
         _ExtentX        =   17436
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
      Begin MSDataListLib.DataCombo dtcVendedor 
         Height          =   360
         Left            =   6510
         TabIndex        =   21
         Top             =   2100
         Width           =   3495
         _ExtentX        =   6165
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
      Begin AutoCompletar.CbCompleta cbbTipo_operacao 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2685
         _ExtentX        =   4736
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
      Begin MSDataListLib.DataCombo dtcEmitente 
         Height          =   360
         Left            =   1500
         TabIndex        =   24
         Top             =   2760
         Width           =   4995
         _ExtentX        =   8811
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
      Begin MSDataListLib.DataCombo dtcCodigo_ecf 
         Height          =   360
         Left            =   8910
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
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
      Begin MSComCtl2.DTPicker dtpData_incial 
         Height          =   360
         Left            =   -72120
         TabIndex        =   65
         Top             =   780
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   50266113
         CurrentDate     =   37949
      End
      Begin MSComCtl2.DTPicker dtpData_final 
         Height          =   360
         Left            =   -70230
         TabIndex        =   66
         Top             =   780
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   50266113
         CurrentDate     =   37949
      End
      Begin MSComCtl2.DTPicker dtpData 
         Height          =   360
         Left            =   6540
         TabIndex        =   26
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   50266113
         CurrentDate     =   37949
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   360
         Left            =   8310
         TabIndex        =   28
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   50266114
         CurrentDate     =   37858
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Código COO"
         Height          =   240
         Left            =   7020
         TabIndex        =   12
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Operador"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   1860
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   240
         Left            =   8310
         TabIndex        =   27
         Top             =   2520
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   240
         Left            =   6540
         TabIndex        =   25
         Top             =   2520
         Width           =   390
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Cód. PDV"
         Height          =   240
         Left            =   5130
         TabIndex        =   10
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblSituacao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   7650
         TabIndex        =   69
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label lblAte 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "até"
         Height          =   240
         Left            =   -70620
         TabIndex        =   67
         Top             =   930
         Width           =   270
      End
      Begin VB.Label lblCliente_Fornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Série"
         Height          =   240
         Left            =   4410
         TabIndex        =   8
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cupom"
         Height          =   240
         Left            =   2850
         TabIndex        =   6
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ECF"
         DragMode        =   1  'Automatic
         Height          =   240
         Left            =   8910
         TabIndex        =   14
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         Height          =   240
         Left            =   5130
         TabIndex        =   19
         Top             =   1860
         Width           =   825
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   59
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Operação"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1530
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10980
      Top             =   360
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
            Picture         =   "frmCupom.frx":5056
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCupom.frx":5370
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCupom.frx":568A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCupom.frx":5A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCupom.frx":5DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCupom.frx":60D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCupom.frx":63F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCupom.frx":7244
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
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Integração"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech Retaguarda                                           '
' Módulo.................: Concentrador de Vendas                                         '
' Objetivo...............: Cadastro Cupon Fiscal                                          '
' Data de Criação........: 30/04/2005                                                     '
' Equipe Responsável.....: Only Tech                                                      '
' Última Manutenção......: Leandro Nolasco Ferreira                                       '
' Data última manutenção.: 29/07/2006                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'na guia de listagem os principais campos de consulta são
'  - cupom
'  - cupom (nº da impressora)
'  - ecf
'  - pdv
'  - data
'  - tipo de operação


Public strUnidade_Venda As String
Public strSQL As String
Public strUF_Emitente As String
Dim I As Integer
Dim strTamanho As String
Dim strCampos As String
Dim strNomes As String
Dim strCampo_Produto As String
Dim strTamanho_Produto As String
Dim strCombo As String
Dim strConsulta As String
Dim booAlterar As Boolean
Dim intIDProduto As Long
Dim intNumero_Itens_NF As Integer
Dim intNumero_Itens_orca As Integer
Dim intContador As Integer
Dim booItem_Alterado As Boolean
Dim dblFator_compra As Double
Dim dblFator_venda As Double
Dim booLibera_produto_linha_pedido As Boolean
Dim booLibera_vendedor_cliente_linha_pedido As Boolean
Dim log As New DLLSystemManager.log
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean


' Declaração das variaveis da acessibilidade

Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim booItem_Alterar As Boolean
Dim booErro_No_Item As Boolean
Dim cnGravacao As New DLLConexao_Sistema.Conexao
Dim dblIdCupom As Double

' Estoque
Dim booBaixar_Estoque As Boolean
Dim booLibera_Estoque_Negativo As Boolean
Dim intManipula_Estoque As Integer

Option Explicit

Private Function Alterar_Itens_Pedido(Numero_Pedido As Long)

    Dim rstProduto As New ADODB.Recordset
    Dim dblQuantidade As Double
    Dim strUnidade_Item As String
    Dim dblPreco_Unitario_Item As Double
    Dim dblValor_Desconto_Item As Double
    Dim dblTotal_Item As Double
    Dim dblQuantidade_subtrair As Double
    
'    Dim booItemAlterado As Boolean
'    Dim booRegistros_Existe_Grid As Boolean
'    Dim booExcluir_Item_Banco As Boolean
'    Dim rstQuantidade_pedido As New ADODB.Recordset
'    Dim booManipular_estoque As Boolean
'    Dim intIDVendedor_Item As Integer
'    Dim intContador_2 As Long
    
    On Error GoTo Erro_Transacao
    
    intContador = 1
    
    ' Iniciando a Transação
    cnGravacao.CNconexao.BeginTrans
    
    ' Pegando os valores do Grid para gravação
    Do While intContador <= hfgItem_cupom.Rows - 2
    
        hfgItem_cupom.Row = intContador
                
        ' Posiscionando na segunda coluna (Código do Produto)
        hfgItem_cupom.Col = 1
       
        ' Passando o conteúdo do grid para a variavel de ID do Produto
        intIDProduto = CInt(hfgItem_cupom.Text)
                
        ' Query que retornará todas as inf. relacionadas a este produto
        strSQL = Empty
        strSQL = "SELECT * FROM TBProduto " & _
                "WHERE TBProduto.IXCodigo_TBProduto = " & intIDProduto & " " & _
                "AND TBProduto.IXCodigo_TBempresa = " & dtcEmpresa.BoundText & " "
                
        ' Montando a recordset que armazenará todas as informações do produto
        Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstProduto, "Otica", Me
       
        'ID,ID Prod,Cod.,Produto,Quant.,Emb.,Pr.Unitário,Vlr.Desc.,Tot.Item
        ' 1,      2,   3,      4,     5,   6,          7,        8,       9
        
        ' Quantidade
        hfgItem_cupom.Col = 3
        dblQuantidade = CDbl(hfgItem_cupom.Text)
        
        ' Unidade
        hfgItem_cupom.Col = 4
        strUnidade_Item = hfgItem_cupom.Text
        
        ' Preço Unitário
        hfgItem_cupom.Col = 5
        dblPreco_Unitario_Item = CDbl(hfgItem_cupom.Text)
        
        ' Valor Desconto Item
        hfgItem_cupom.Col = 6
        
        If hfgItem_cupom.Text = Empty Then
            dblValor_Desconto_Item = 0
        Else
            dblValor_Desconto_Item = CDbl(hfgItem_cupom.Text)
        End If
        
        ' Valor Total do Item
        hfgItem_cupom.Col = 7
       
        If hfgItem_cupom.Text = Empty Then
            dblTotal_Item = 0
        Else
            dblTotal_Item = CDbl(hfgItem_cupom.Text)
        End If
       
        If strUnidade_Item = rstProduto!DFUnidade_venda_TBProduto Then
            dblQuantidade_subtrair = dblQuantidade
        Else
            dblQuantidade_subtrair = dblQuantidade * rstProduto!DFFator_venda_TBProduto
        End If
                                 
        strSQL = Empty
        
        strSQL = "INSERT INTO TBItens_cupom ( " & _
                 "FKId_TBCupom, " & _
                 "DFCodigo_TBProduto, " & _
                 "DFCst1_TBItens_cupom, " & _
                 "DFCst2_TBItens_cupom, " & _
                 "DFQuantidade_TBItens_cupom, " & _
                 "DFTipo_preco_TBItens_cupom, " & _
                 "DFPreco_tabela_TBItens_cupom, " & _
                 "DFPercentual_desconto_TBItens_cupom, " & _
                 "DFPreco_praticado_TBItens_cupom, " & _
                 "DFValor_total_tabela_TBItens_cupom, " & _
                 "DFValor_total_praticado_TBItens_cupom, " & _
                 "DFPercentual_icms_TBItens_cupom, " & _
                 "DFValor_total_icms_TBItens_cupom, " & _
                 "DFUnidade_TBItens_cupom, " & _
                 "DFCusto_real_TBItens_cupom, " & _
                 "DFCusto_contabil_TBItens_cupom, " & _
                 "DFCusto_medio_TBItens_cupom, " & _
                 "DFPeso_liquido_TBItens_cupom, " & _
                 "DFPeso_bruto_TBItens_cupom, " & _
                 "DFQuantidade_baixa_estoque_TBItens_cupom, " & _
                 "DFValor_total_item_TBItens_cupom, " & _
                 "DFDivisor_baixa_estouqe_TBItens_cupom, " & _
                 "DFItens_cupom_Registrado_TBItens_cupom ) "
                 
        strSQL = strSQL & "VALUES (" & _
                 " " & Numero_Pedido & "," & _
                 " " & intIDProduto & "," & _
                 " '" & CStr(rstProduto!DFCst1_TBProduto) & "'," & _
                 " '" & CStr(rstProduto!DFCst2_TBProduto) & "'," & _
                 " " & Funcoes_Gerais.Grava_Moeda(dblQuantidade) & "," & _
                 " " & 1 & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda(dblPreco_Unitario_Item) & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda(dblValor_Desconto_Item) & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda(dblPreco_Unitario_Item) & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda(dblTotal_Item) & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda(dblTotal_Item) & "," & _
                 " " & 0 & "," & _
                 " " & 0 & "," & _
                 " '" & strUnidade_Item & "'," & _
                 " " & Funcoes_Gerais.Grava_Moeda(rstProduto!DFCusto_real_TBProduto) & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda(rstProduto!DFCusto_contabil_TBProduto) & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda(rstProduto!DFCusto_medio_TBProduto) & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda((rstProduto!DFPeso_liquido_TBProduto * dblQuantidade)) & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda((rstProduto!DFPeso_bruto_TBProduto * dblQuantidade)) & "," & _
                 " " & Funcoes_Gerais.Grava_Moeda(dblQuantidade_subtrair) & "," & _
                 " " & CDbl(lblTotal_Itens.Caption) & "," & _
                 " " & rstProduto!DFFator_venda_TBProduto & "," & _
                 " " & 1 & " ) "

        rstProduto.Close
        
        Set rstProduto = Nothing
    
        ' Passando comando para gravação
        cnGravacao.CNconexao.Execute strSQL
        
        ' Passando para o proximo item do Grid
        intContador = intContador + 1
        
    Loop
    
    ' Efetivando a Gravação dos Itens do Cupom
    cnGravacao.CNconexao.CommitTrans
    
    Exit Function
    
Erro_Transacao:

    ' Indicando q a alteração/inclusão do produto falhou. Assim sendo toda a operação deverá ser cancelada.
    booErro_No_Item = True
    
    ' Cancelando a Operação e fechando a conexão
    cnGravacao.CNconexao.RollbackTrans
    cnGravacao.Fechar_conexao

    Call Erro.Erro(Me, "Otica", "Alterar_Itens_Pedido")
    Exit Function
    
End Function

Private Function Consulta()
    Dim intContador As Integer
    Dim strTipo_Operacao As String
    
    If cbbCampos.Text <> "Todos" And txtConsulta.Text <> Empty Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
    
    If cbbCampos.Text = "Tipo de Operação" Then
       If txtConsulta.Text = "VENDA" Then
          strTipo_Operacao = 1
       ElseIf txtConsulta.Text = "TRANSFERÊNCIA" Or txtConsulta.Text = "TRANSFERENCIA" Then
          strTipo_Operacao = 2
       ElseIf txtConsulta.Text = "DEV.CLIENTE" Then
          strTipo_Operacao = 3
       ElseIf txtConsulta.Text = "DEV.FORNEC." Then
          strTipo_Operacao = 4
       Else
          strTipo_Operacao = 5
       End If
    End If

    ' Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    txtConsulta.Text = Funcoes_Gerais.Grava_String(txtConsulta.Text)
    
    strSQL = Empty
    
    strSQL = "SELECT PKId_TBCupom, " & _
             "IXCodigo_TBVendedor, DFNome_TBVendedor, " & _
             "DFNumero_ecf_TBPdv, DFTipo_operacao_TBCupom, " & _
             "DFNumero_TBCupom, DFSerie_TBCupom, DFEmitente_TBCupom, " & _
             "DFNome_TBCliente, " & _
             "DFTotal_itens_TBCupom, DFTotal_cupom_TBCupom, DFData_Saida_TBCupom," & _
             "DFHora_Saida,FKCodigo_TBOperadores_ecf,DFNome_TBOperadores_ecf," & _
             "DFCancelado_TBCupom,DFMotivo_cancelamento_TBCupom," & _
             "DFUsuario_cancelamento_TBCupom, DFObservacao_TBCupom,TBPdv.PKCodigo_TBPdv," & _
             "TBCupom.FKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa " & _
             "FROM TBCupom " & _
             "INNER JOIN TBEmpresa ON TBEmpresa.PKCodigo_TBEmpresa = TBCupom.FKCodigo_TBEmpresa " & _
             "INNER JOIN TBVendedor ON TBVendedor.PKId_TBVendedor = TBCupom.FKId_TBVendedor " & _
             "INNER JOIN TBPdv ON TBPdv.PKCodigo_TBPdv = TBCupom.PKCodigo_TBPdv " & _
             "LEFT JOIN TBCliente ON TBCliente.IXCodigo_TBCliente = TBCupom.DFEmitente_TBCupom " & _
             "INNER JOIN TBOperadores_ecf ON TBCupom.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf "
             
    If cbbCampos.Text <> "Todos" Then
        If cbbCampos.Text = "ECF" Then
            strSQL = strSQL & " WHERE convert(nvarchar,DFNumero_ecf_TBPdv) = " & txtConsulta.Text & " "
        ElseIf cbbCampos.Text = "Número do Cupom" Then
            strSQL = strSQL & " WHERE convert(nvarchar,DFNumero_TBCupom) = " & txtConsulta.Text & " "
        ElseIf cbbCampos.Text = "Série do Cupom" Then
            strSQL = strSQL & " WHERE convert(nvarchar,DFSerie_TBCupom) LIKE '%" & txtConsulta.Text & "%' "
        ElseIf cbbCampos.Text = "Data" Then
             strSQL = strSQL & " WHERE TBCupom.DFData_Saida_TBCupom >= '" & Format(dtpData_incial.Value, "YYYYMMDD") & "' " & _
                               " AND TBCupom.DFData_Saida_TBCupom <= '" & Format(dtpData_final.Value, "YYYYMMDD") & "' "
        ElseIf cbbCampos.Text = "Tipo de Operação" Then
            strSQL = strSQL & " WHERE convert(nvarchar,DFTipo_operacao_TBCupom) LIKE '" & strTipo_Operacao & "' "
        ElseIf cbbCampos.Text = "Cód. Vendedor" Then
            strSQL = strSQL & " WHERE convert(nvarchar,IXCodigo_TBVendedor) = " & txtConsulta.Text & " "
        ElseIf cbbCampos.Text = "Nome do Vendedor" Then
            strSQL = strSQL & " WHERE convert(nvarchar,DFNome_TBVendedor) LIKE '%" & txtConsulta.Text & "%' "
        ElseIf cbbCampos.Text = "Cód. Cliente" Then
            strSQL = strSQL & " WHERE convert(nvarchar,DFEmitente_TBCupom) = " & txtConsulta.Text & " "
        ElseIf cbbCampos.Text = "Nome do Cliente" Then
            strSQL = strSQL & " WHERE convert(nvarchar,DFNome_TBCliente) LIKE '%" & txtConsulta.Text & "%' "
        ElseIf cbbCampos.Text = "Cód. Operador" Then
            strSQL = strSQL & " WHERE convert(nvarchar,PKCodigo_TBOperadores_ecf) = '" & txtConsulta.Text & "' "
        ElseIf cbbCampos.Text = "Operador" Then
            strSQL = strSQL & " WHERE convert(nvarchar,DFNome_TBOperadores_ecf) LIKE '%" & txtConsulta.Text & "%' "
        ElseIf cbbCampos.Text = "Valor do Cupom" Then
            strSQL = strSQL & " WHERE DFTotal_cupom_TBCupom = " & txtConsulta.Text & " "
        ElseIf cbbCampos.Text = "Código PDV" Then
            strSQL = strSQL & " WHERE convert(nvarchar,PKCodigo_TBPdv) = '" & txtConsulta.Text & "' "
        ElseIf cbbCampos.Text = "Empresa" Then
            strSQL = strSQL & " WHERE convert(nvarchar,DFRazao_Social_TBEmpresa) LIKE '%" & txtConsulta.Text & "%' "
        End If
    End If
    
    frmAguarde.Show
    DoEvents
    
    strSQL = strSQL & " ORDER BY IXCodigo_TBVendedor "

    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgCupom_fiscal, strTamanho, strCampos, "BDRetaguarda", "OTICA", Me
    
    hfgCupom_fiscal.Col = 5
    intContador = 1
    Do While intContador <= hfgCupom_fiscal.Rows - 1
        hfgCupom_fiscal.Row = intContador
        hfgCupom_fiscal.Col = 5
        Select Case hfgCupom_fiscal.Text
            Case 1
                hfgCupom_fiscal.Text = "Venda"
            Case 2
                hfgCupom_fiscal.Text = "Transferência"
            Case 3
                hfgCupom_fiscal.Text = "Dev.Cliente"
            Case 4
                hfgCupom_fiscal.Text = "Dev.Fornec."
            Case 5
                hfgCupom_fiscal.Text = "Outras"
        End Select
        hfgCupom_fiscal.Col = 13
        hfgCupom_fiscal.Text = Format(hfgCupom_fiscal.Text, "HH:MM:SS")
        intContador = intContador + 1
    Loop
    
    hfgCupom_fiscal.Row = 1
    hfgCupom_fiscal.Col = 0
    
    If hfgCupom_fiscal.Text = Empty Then
       hfgCupom_fiscal.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgCupom_fiscal, strTamanho, strCampos, 22, "OTICA", Me
    End If
    
    hfgCupom_fiscal.Refresh
    
    Unload frmAguarde

End Function
Private Function Monta_Combo()

    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("ECF")
    cbbCampos.AddItem ("Número do Cupom")
    cbbCampos.AddItem ("Série do Cupom")
    cbbCampos.AddItem ("Data")
    cbbCampos.AddItem ("Tipo de Operação")
    cbbCampos.AddItem ("Cód. Vendedor")
    cbbCampos.AddItem ("Nome do Vendedor")
    cbbCampos.AddItem ("Cód. Cliente")
    cbbCampos.AddItem ("Nome do Cliente")
    cbbCampos.AddItem ("Cód. Operador")
    cbbCampos.AddItem ("Operador")
    cbbCampos.AddItem ("Valor do Cupom")
    cbbCampos.AddItem ("Código PDV")
    cbbCampos.AddItem ("Empresa")
    
    'Tipo de Operacao
    cbbTipo_operacao.Clear
    cbbTipo_operacao.AddItem ("Venda")
    cbbTipo_operacao.AddItem ("Transferência")
    cbbTipo_operacao.AddItem ("Dev.Cliente")
    cbbTipo_operacao.AddItem ("Dev.Fornec.")
    cbbTipo_operacao.AddItem ("Outras")

End Function

Private Function Monta_Resumos()

    Dim dblDesconto_Itens As Double
    Dim dblTotal_Item As Double
    
    'Resumo Financeiro
    'Calculando o Total do Item
    
    If txtTotal_item.Text = "" Then
        txtTotal_item.Text = 0
    End If
    
    If lblTotal_Produtos.Caption = "" Then
        lblTotal_Produtos.Caption = 0
    End If
    
    If txtPreco_unitario.Text = "" Then
        txtPreco_unitario.Text = 0
    End If
    
    dblTotal_Item = CDbl(txtQuantidade_produto.Text) * CDbl(txtPreco_unitario.Text)
    lblTotal_Produtos.Caption = lblTotal_Produtos + dblTotal_Item
    
    'Calculando o valor do desconto por Item
    If txtPercentual_desconto <> "" Then
        If lblDesconto_Itens.Caption = "" Then
            lblDesconto_Itens.Caption = 0
        End If
        dblDesconto_Itens = ((txtPreco_unitario.Text * txtQuantidade_produto.Text) * txtPercentual_desconto) / 100
        lblDesconto_Itens.Caption = lblDesconto_Itens + dblDesconto_Itens
    End If
     
    'Calculando o Total do Praticado
    If txtTotal_item.Text = "" Then
        txtTotal_item.Text = 0
    End If

    'Calculando o total do pedido
    If Me.lblDesconto_Itens.Caption = "" Then
        lblDesconto_Itens.Caption = 0
    End If
    
    lblTotal_Pedido.Caption = CDbl(lblTotal_Produtos.Caption) - CDbl(Me.lblDesconto_Itens.Caption)
    
    If lblTotal_Pedido.Caption < 0 Then
       lblTotal_Pedido.ForeColor = &HFF&
    Else
       lblTotal_Pedido.ForeColor = &HC00000
    End If
    
    lblTotal_Produtos.Caption = Format(lblTotal_Produtos.Caption, "#,###0.00")
    lblDesconto_Itens.Caption = Format(lblDesconto_Itens.Caption, "#,###0.00")
    lblTotal_Pedido.Caption = Format(lblTotal_Pedido.Caption, "#,###0.00")
    
End Function

Private Sub cbbCampos_Click()

    'Se Mudou de campo
    If cbbCampos.Text <> strCombo Then
       strCombo = cbbCampos.Text
    End If
            
    'se o campo "strCombo" for nulo
    If strCombo <> Empty Then
       cbbCampos.Text = strCombo
    End If
    
    'se o campo "cbbcampos" for nulo
    If cbbCampos.Text = Empty Then
       cbbCampos.SetFocus
       Exit Sub
    End If
        
    txtConsulta.Text = Empty
   
    If cbbCampos.Text = "Todos" Then
        txtConsulta.Visible = False
        lblAte.Visible = False
        dtpData_incial.Visible = False
        dtpData_final.Visible = False
        cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Data" Then
        txtConsulta.Visible = False
        lblAte.Visible = True
        dtpData_incial.Visible = True
        dtpData_final.Visible = True
        dtpData_incial.SetFocus
    Else
        txtConsulta.Visible = True
        lblAte.Visible = False
        dtpData_incial.Visible = False
        dtpData_final.Visible = False
        txtConsulta.SetFocus
    End If
        
End Sub

Private Sub cmdConsulta_Click()

    If cbbCampos.Text <> "Todos" Then
        If cbbCampos.Text <> "Data" Then
            If Trim(txtConsulta.Text) = Empty Then
                MsgBox "Data Final menor que Data Inicial. Verifique!", vbInformation, "Only Tech"
                txtConsulta.SetFocus
                Exit Sub
            End If
        Else
            If dtpData_incial.Value > dtpData_final.Value Then
                MsgBox "Data Final menor que Data Inicial. Verifique!", vbInformation, "Only Tech"
                dtpData_incial.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    Call Consulta
    
End Sub

Private Sub cmdIncluir_Item_Click()

    Dim dblAltera_Quantidade As Double
    Dim dblAltera_Preco_Unitario As Double
    Dim dblAltera_Percentual_Desconto As Double
    Dim dblAltera_Total_preco_tabela As Double
    Dim dblAltera_total_item_grid As Double
    Dim strEmbalagem As String
    Dim intRetorno As Integer
    Dim booItem_modificado As Boolean
            
    'Validação dos Inputs
    If txtProduto.Text = Empty Then
       MsgBox "Código do Item Inexistente! Redigite", vbInformation, "Only Tech"
       txtProduto.SetFocus
       Exit Sub
    ElseIf txtQuantidade_produto.Text = "" Then
       MsgBox "Quantidade do Item Inválida! Redigite", vbInformation, "Only Tech"
       txtQuantidade_produto.SetFocus
       Exit Sub
    ElseIf txtPreco_unitario.Text = "" Then
       MsgBox "Preço Unitário do Item Inválido! Redigite", vbInformation, "Only Tech"
       txtPreco_unitario.SetFocus
       Exit Sub
    ElseIf txtTotal_item.Text = "" Then
       MsgBox "Total do Item Inválido! Redigite", vbInformation, "Only Tech"
       txtTotal_item.SetFocus
       Exit Sub
    ElseIf txtVendedor.Text = "" Then
       MsgBox "Para incluir este item digite um vendedor válido para este pedido!", vbInformation, "Only Tech"
       txtVendedor.SetFocus
       Exit Sub
    End If
    
    'Validando informações do corpo do pedido
    'Verifica se é Dev.Fornecedor
    If cbbTipo_operacao.Text = "Dev.Fornec." Then
        'Fornecedor
        If txtCliente.Text = "" Then
           MsgBox "Para incluir este item digite um fornecedor válido para este pedido!", vbInformation, "Only Tech"
           txtCliente.SetFocus
           Exit Sub
        End If
    Else
        'Cliente
        If txtCliente.Text = "" Then
           MsgBox "Para incluir este item digite um cliente válido para este pedido!", vbInformation, "Only Tech"
           txtCliente.SetFocus
           Exit Sub
        End If
     
        On Error GoTo Erro
      
        '----------------------------------------------------------------------------------------------
        'Verifica se o produto possui ICMS na tabela de estado para ICMS
        Dim rstVerifica_Estado As New ADODB.Recordset
        
        'Pegando o ID do Produto
        intIDProduto = Funcoes_Gerais.Localiza_ID("PKId_TBProduto", "IXCodigo_TBProduto", " " & txtProduto.Text & " ", "TBProduto", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", " " & dtcEmpresa.BoundText & "")
        
        'Query para pegar ICMS do item
        strSQL = Empty
        strSQL = "SELECT " & _
                 "PKId_TBEstado_icms," & _
                 "FKId_TBProduto," & _
                 "DFUf_TBEstado_icms," & _
                 "DFPercentual_icms_saida_juridica_TBEstado_icms," & _
                 "DFPercentual_icms_saida_fisica_TBEstado_icms " & _
                 "FROM TBEstado_icms " & _
                 "WHERE FKId_TBProduto = " & intIDProduto & " " & _
                 "AND DFUf_TBEstado_icms = '" & strUF_Emitente & "'"
                 
        Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstVerifica_Estado, "Otica", Me
                
        'Esta modificação foi feita para corrigir um erro de cadastramento no programa
        'de Estados p/ ICMS. O mesmo estava permitindo o cadastramento do item com as respectivas
        'aliquotas gravadas com espaços ou seja " ". (Giordano)
        
        'jones
        If IsNull(rstVerifica_Estado.Fields("DFPercentual_icms_saida_fisica_TBEstado_icms")) Or _
           IsNull(rstVerifica_Estado.Fields("DFPercentual_icms_saida_juridica_TBEstado_icms")) Then
           Dim strMensagem_1 As String
           strMensagem_1 = Empty
           strMensagem_1 = strMensagem_1 & "Este item não possui aliquota de ICMS cadastrada em nenhum estado.Verifique!"
           MsgBox strMensagem_1, vbCritical, "Only Tech"
           Set rstVerifica_Estado = Nothing
           txtProduto.Text = Empty
           txtQuantidade_produto.Text = Empty
           txtPercentual_desconto.Text = Empty
           txtPreco_unitario.Text = Empty
           txtTotal_item.Text = Empty
           txtProduto.SetFocus
           Exit Sub
        End If
        
        'Verifica se existe icms para este estado nesse item
        If rstVerifica_Estado.BOF = True And rstVerifica_Estado.EOF = True Then
           'Mata a recordest
           Set rstVerifica_Estado = Nothing
           'Se não existir icms para este item no estado correspondente ele verifica novamente e testa com estado = **,
           '** foi uma forma encontrada para permitir que todos os itens estejam cadastrados em todos os estados,
           'não precisando espelhar todos os itens para todos os estados
           strSQL = Empty
           strSQL = "SELECT " & _
                    "PKId_TBEstado_icms," & _
                    "FKId_TBProduto," & _
                    "DFUf_TBEstado_icms," & _
                    "DFPercentual_icms_saida_juridica_TBEstado_icms," & _
                    "DFPercentual_icms_saida_fisica_TBEstado_icms " & _
                    "FROM TBEstado_icms " & _
                    "WHERE FKId_TBProduto = " & intIDProduto & " " & _
                    "AND DFUf_TBEstado_icms = '**'"
                    
           Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstVerifica_Estado, "Otica", Me

           'Ainda assim se não houver a presença do item com estado ** ele dá uma mensagem para o usuário
           If rstVerifica_Estado.BOF = True And rstVerifica_Estado.EOF = True Then
              Dim strMensagem As String
              strMensagem = Empty
              strMensagem = strMensagem & "Este item não possui aliquota de ICMS cadastrada em nenhum estado.Verifique!"
              MsgBox strMensagem, vbCritical, "Only Tech"
              Set rstVerifica_Estado = Nothing
              txtProduto.Text = Empty
              txtQuantidade_produto.Text = Empty
              txtPercentual_desconto.Text = Empty
              txtPreco_unitario.Text = Empty
              txtTotal_item.Text = Empty
              txtProduto.SetFocus
              Exit Sub
           End If
        End If
        
        'Verificar se o produto digitado possui estoque em seu cadastro
        Dim rstEstoque As New ADODB.Recordset
        
        strSQL = Empty
        strSQL = "SELECT DFEstoque_atual_TBProduto,DFUnidade_venda_TBProduto FROM TBProduto " & _
                 "WHERE IXCodigo_TBProduto = " & Me.txtProduto.Text & " " & _
                 "AND IXCodigo_TBEmpresa = " & Me.dtcEmpresa.BoundText & " "
                 
        Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstEstoque, "Otica", Me
        
        'Verifica se o estoque do produto já está negativo
        If CDbl(rstEstoque!DFEstoque_Atual_TBProduto) <= 0 And cbbTipo_operacao.Text <> "Dev.Cliente" Then
           'Verifica se pode liberar estoque pelo parâmetro de venda.
           intRetorno = MsgBox("Este Item está com o estoque negativo - Estoque Atual : " & CDbl(rstEstoque!DFEstoque_Atual_TBProduto) & " " & rstEstoque!DFUnidade_venda_TBProduto, vbYesNo, "Only Tech")
           If intRetorno = 7 Then
              Set rstEstoque = Nothing
              txtProduto.SetFocus
              Exit Sub
           End If
        End If
        
        'Verifica se o estoque do produto ficará negativo com a inclusão do item
        Dim dblQuantidade_verificacao As Double
        
'        If dtcUnidade_venda.text <> strUnidade_Venda Then
'           dblQuantidade_verificacao = CDbl(txtQuantidade_produto.Text) * dblFator_compra
'        Else
'           dblQuantidade_verificacao = CDbl(txtQuantidade_produto.Text)
'        End If
        
        If dblQuantidade_verificacao > CDbl(rstEstoque!DFEstoque_Atual_TBProduto) And cbbTipo_operacao.Text <> "Dev.Cliente" Then
           If booLibera_Estoque_Negativo = False Then
              MsgBox "O seu estoque atual é de " & rstEstoque!DFEstoque_Atual_TBProduto & " " & rstEstoque!DFUnidade_venda_TBProduto & "!Este Item deixará seu estoque negativo em: " & dblQuantidade_verificacao - CDbl(rstEstoque!DFEstoque_Atual_TBProduto) & " " & rstEstoque!DFUnidade_venda_TBProduto & ".O mesmo não pode ser incluido porque não há permissão no sistema para se vender um produto que não contenha estoque correspondente.", vbInformation, "Only Tech"
              txtProduto.Text = ""
              dtcProduto.Text = ""
              txtQuantidade_produto = ""
              txtPreco_unitario.Text = ""
              txtPercentual_desconto.Text = ""
              txtTotal_item.Text = ""
              txtProduto.SetFocus
              Exit Sub
           End If
           intRetorno = MsgBox("Este Item deixará seu estoque negativo em: " & dblQuantidade_verificacao - CDbl(rstEstoque!DFEstoque_Atual_TBProduto) & " " & rstEstoque!DFUnidade_venda_TBProduto, vbYesNo, "Only Tech")
           If intRetorno = 7 Then
              Set rstEstoque = Nothing
              txtProduto.SetFocus
              Exit Sub
           End If
           Set rstEstoque = Nothing
        End If
    End If
   
    'Verificar se o item está no grid de itens do pedido
    intContador = 1
    booItem_Alterado = False
    Do While intContador <= hfgItem_cupom.Rows - 2
        hfgItem_cupom.Row = intContador
        hfgItem_cupom.Col = 1
        If hfgItem_cupom.Text = txtProduto.Text And booItem_Alterar = False Then
           MsgBox "Este Item já está incluido neste Pedido.!", vbInformation, "Only Tech"
           'Limpando os campos dos Itens
           txtProduto.Text = ""
           dtcProduto.Text = ""
           txtQuantidade_produto = ""
           txtPreco_unitario.Text = ""
           txtPercentual_desconto.Text = ""
           txtTotal_item.Text = ""
           txtProduto.SetFocus
           Exit Sub
        End If
        '' INICIO ROTINA ALTERAÇÂO - ALTERA ITEM
        If hfgItem_cupom.Text = txtProduto.Text And booItem_Alterar = True Then
            booItem_Alterado = True
            booItem_modificado = False
            
            'Verifica se é rotina novo saio da function
            If booAlterar = False Then GoTo Nao_Modificado
            
            Me.hfgItem_cupom.Col = 3
            'Armazenado o conteúdo da celula do grid, antes que ele seja substituido pelo novo conteúdo
            'guardando para efetuar a subtração dos valores na montagem dos resumos abaixo.
            dblAltera_Quantidade = hfgItem_cupom.Text
            If txtQuantidade_produto.Text <> hfgItem_cupom.Text Then
               booItem_modificado = True
               hfgItem_cupom.Text = txtQuantidade_produto.Text
            End If
    
'            Me.hfgItem_cupom.Col = 4
'            'Armazenado o conteúdo da celula do grid, antes que ele seja substituido pelo novo conteúdo
'            'guardando para efetuar a subtração dos valores na montagem dos resumos abaixo.
'            strEmbalagem = hfgItem_cupom.Text
'            If dtcUnidade_venda.text <> hfgItem_cupom.Text Then
'               booItem_modificado = True
'               hfgItem_cupom.Text = dtcUnidade_venda.text
'            End If
            
'            Me.hfgItem_cupom.Col = 5
'            'Armazenado o conteúdo da celula do grid, antes que ele seja substituido pelo novo conteúdo
'            'guardando para efetuar a subtração dos valores na montagem dos resumos abaixo.
'            strTipo_Preco = hfgItem_cupom.Text
'            If cmbTipo_Preco.Text <> hfgItem_cupom.Text Then
'               booItem_modificado = True
'               hfgItem_cupom.Text = cmbTipo_Preco.Text
'            End If
    
            Me.hfgItem_cupom.Col = 6
            'Armazenado o conteúdo da celula do grid, antes que ele seja substituido pelo novo conteúdo
            'guardando para efetuar a subtração dos valores na montagem dos resumos abaixo.
            dblAltera_Preco_Unitario = hfgItem_cupom.Text
    
            If txtPreco_unitario.Text <> hfgItem_cupom.Text Then
               booItem_modificado = True
               hfgItem_cupom.Text = txtPreco_unitario.Text
            End If
    
            Me.hfgItem_cupom.Col = 7
            'Armazenado o conteúdo da celula do grid, antes que ele seja substituido pelo novo conteúdo
            'guardando para efetuar a subtração dos valores na montagem dos resumos abaixo.
            If hfgItem_cupom.Text = "" Then hfgItem_cupom.Text = 0
            If txtPercentual_desconto.Text = "" Then txtPercentual_desconto.Text = 0
            If txtPercentual_desconto.Text <> hfgItem_cupom.Text Then
               hfgItem_cupom.Text = txtPercentual_desconto.Text
               booItem_modificado = True
            End If
            dblAltera_Percentual_Desconto = hfgItem_cupom.Text
            
            Me.hfgItem_cupom.Col = 8
            
            dblAltera_total_item_grid = hfgItem_cupom.Text
            
            If txtTotal_item.Text <> hfgItem_cupom.Text Then
               hfgItem_cupom.Text = txtTotal_item.Text
               booItem_modificado = True
            End If
            
            If booItem_modificado = False Then GoTo Nao_Modificado
            
            '------------------------------------------------------------------------------------------------
            'Recalculando os resumos financeiros:
            'Subtraindo o valor inserido anteriormente para o item, para mais abaixo somar os novos valores
            'dos resumos
            '------------------------------------------------------------------------------------------------
    
'            Call Outras_Selecoes
'            Call Seleciona_Preco
    
            Dim dblAltera_Total_Item As Double
            Dim dblAltera_Desconto_Itens As Double
            Dim dblAltera_Total_peso_bruto As Double
            Dim dblAltera_Total_peso_liquido As Double
            Dim dblAltera_Quantidade_real As Double
            Dim dblAltera_Quantidade_trabalhar As Double
            Dim dblPreco_Unitario_Alterar As Double
            
            'Resumo Financeiro
            'Calculando o Total do Item
            'Verifica a unidade e faz a multiplicação quando as unidades(compra e venda) forem diferentes
            'pelo fator de venda
            
            If strEmbalagem <> strUnidade_Venda Then
               dblAltera_Quantidade = dblAltera_Quantidade * dblFator_compra
            End If
            
            If dtcUnidade_venda.Text <> strUnidade_Venda Then
               dblAltera_Quantidade_real = CDbl(txtQuantidade_produto.Text) * dblFator_compra
               dblPreco_Unitario_Alterar = CDbl(txtPreco_unitario.Text) / dblFator_compra
            Else
               dblAltera_Quantidade_real = CDbl(txtQuantidade_produto.Text)
               dblPreco_Unitario_Alterar = CDbl(txtPreco_unitario.Text)
            End If
            
            dblAltera_Quantidade_trabalhar = dblAltera_Quantidade_real - dblAltera_Quantidade
            
            '------------------------------------------------------------------------------------------
            'Este trecho verifica a diferença de arredondamento e faz o acerto
            '------------------------------------------------------------------------------------------
            
            Dim dblVerifica_Total As String
            Dim dblDiferenca_Total As Double
            Dim dblTotal_Atual As Double
            Dim a As Double
            
            dblVerifica_Total = dblPreco_Unitario_Alterar * CDbl(txtQuantidade_produto.Text)
            dblDiferenca_Total = 0
            
            dblTotal_Atual = CDbl(dblAltera_Preco_Unitario) * CDbl(txtQuantidade_produto.Text)
            
            dblDiferenca_Total = CDbl(dblVerifica_Total) - dblTotal_Atual
           
            If dblDiferenca_Total <> 0 Then
               dblDiferenca_Total = dblDiferenca_Total / dblFator_compra
               dblPreco_Unitario_Alterar = dblPreco_Unitario_Alterar - dblDiferenca_Total
            End If
            '------------------------------------------------------------------------------------------
            
            If dblAltera_Quantidade_trabalhar >= 0 Then
               lblTotal_Produtos.Caption = CDbl(lblTotal_Produtos) + (dblAltera_Quantidade_trabalhar * dblPreco_Unitario_Alterar)
            Else
               lblTotal_Produtos.Caption = CDbl(lblTotal_Produtos) - ((dblAltera_Quantidade_trabalhar * dblPreco_Unitario_Alterar) * (-1))
            End If
            
            'Acertando Descontos do item
            Dim dblAltera_Desconto_Itens_atual As String
            
            If ((dblAltera_Quantidade_real * dblPreco_Unitario_Alterar) * dblAltera_Percentual_Desconto) / 100 < 0 Then
               dblAltera_Desconto_Itens_atual = (((dblAltera_Quantidade_real * dblPreco_Unitario_Alterar) * dblAltera_Percentual_Desconto) / 100) * (-1)
            Else
               dblAltera_Desconto_Itens_atual = (((dblAltera_Quantidade_real * dblPreco_Unitario_Alterar) * dblAltera_Percentual_Desconto) / 100)
            End If
            
            'Calculando o valor do desconto por Item
            If dblAltera_Percentual_Desconto <> 0 Then
               dblAltera_Desconto_Itens = CDbl(lblDesconto_Itens) - dblAltera_Desconto_Itens_atual
               lblDesconto_Itens.Caption = lblDesconto_Itens - CDbl(dblAltera_Desconto_Itens)
            End If
    
            'Calculando o Total do Praticado
            'lblTotal_praticado.Caption = CDbl(lblTotal_Produtos)
    
            'Calculando o total do pedido
            If lblDesconto_Itens.Caption = "" Then lblDesconto_Itens.Caption = 0
            'If txtIpi.Text = "" Then txtIpi.Text = 0
            'If txtDespesas_acessorios.Text = "" Then txtDespesas_acessorios.Text = 0
            'If txtDesconto_especial = "" Then txtDesconto_especial.Text = 0
            'If txtIndenizacao.Text = "" Then txtIndenizacao.Text = 0
    
            lblTotal_Pedido.Caption = CDbl(lblTotal_Produtos.Caption) - CDbl(lblDesconto_Itens.Caption)
            
            If lblTotal_Pedido.Caption < 0 Then
               lblTotal_Pedido.ForeColor = &HFF&
            Else
               lblTotal_Pedido.ForeColor = &HC00000
            End If
    
            'Calculando o total de tabela
            'Selecionando o preço de tabela correspondente
            Dim rstPreco As New ADODB.Recordset
    
            intIDProduto = Funcoes_Gerais.Localiza_ID("PKID_TBProduto", "IXCodigo_TBProduto", txtProduto.Text, "TBProduto", "Otica", Me, "BDRetaguarda")
    
'            strSql = Empty
'            strSql = "SELECT * FROM TBItens_tabela_preco WHERE FKID_TBProduto = " & intIDProduto & " AND FKCodigo_TBTabela_preco = " & txtNumero_tabela.Text & " "
'
'            Movimentacoes.Select_geral strSql, "BDRetaguarda", rstPreco, "Otica", Me
'
'            If strTipo_Preco = "À Vista" Then dblAltera_Total_preco_tabela = rstPreco!DFPreco_avista_TBItens_tabela_preco
'            If strTipo_Preco = "Promoção" Then dblAltera_Total_preco_tabela = rstPreco!DFPreco_promocao_TBItens_tabela_preco
'            If strTipo_Preco = "Revenda" Then dblAltera_Total_preco_tabela = rstPreco!DFPreco_revenda_TBItens_tabela_preco
'            If strTipo_Preco = "Especial" Then dblAltera_Total_preco_tabela = rstPreco!DFPreco_especial_TBItens_tabela_preco
'
'            'Verifica seguindo o plano se a condição de pagamento força o acrésimo/desconto no preço ou se é normal
'            If strAcresimo_desconto_plano_pagamento <> "Normal" Then
'               dblPercentual_plano = (CDbl(dblAltera_Total_preco_tabela) * dblPercentual_plano_pagamento) / 100
'               If strAcresimo_desconto_plano_pagamento = "Acréscimo" Then
'                  dblAltera_Total_preco_tabela = CDbl(dblAltera_Total_preco_tabela) + dblPercentual_plano
'               End If
'               If strAcresimo_desconto_plano_pagamento = "Desconto" Then
'                  dblAltera_Total_preco_tabela = CDbl(dblAltera_Total_preco_tabela) - dblPercentual_plano
'               End If
'            End If
'
'            dblAltera_Total_preco_tabela = Format(dblAltera_Total_preco_tabela, "#,###0.00")
'            Set rstPreco = Nothing
            
            '------------------------------------------------------------------------------------------
            'Localizando o valor de tabela anterior para montar o resumo corretamente
'            Dim dblPreco_tabela_anterior As Double
'            Dim dblAltera_Quantidade_item_anterior As Double
'
'            dblPreco_tabela_anterior = dblAltera_Total_preco_tabela
'            dblAltera_Quantidade_item_anterior = dblAltera_Quantidade
'
'            If strEmbalagem = strUnidade_Venda Then
'               dblPreco_tabela_anterior = dblPreco_tabela_anterior / dblFator_compra
'            End If
'
'            If strEmbalagem <> strUnidade_Venda Then
'               dblAltera_Quantidade_item_anterior = dblAltera_Quantidade_item_anterior * dblFator_compra
'            End If
'
'            dblPreco_tabela_anterior = dblPreco_tabela_anterior * dblAltera_Quantidade_item_anterior
            
            '--------------------------------------------------------------------------------------------
            If dtcUnidade_venda.Text = strUnidade_Venda Then
               dblAltera_Total_preco_tabela = dblAltera_Total_preco_tabela / dblFator_compra
            End If
            
            If txtProduto.Text <> "" Then
            
               If dtcUnidade_venda.Text <> strUnidade_Venda Then
                  dblAltera_Total_preco_tabela = dblAltera_Total_preco_tabela / dblFator_compra
               End If
               
               'dblAltera_Total_preco_tabela = Format(dblAltera_Total_preco_tabela, "#,###0.00")
               
               'Calculando os total de tabela do Item a ser subtraido do resumo total de tabela
               dblAltera_Total_preco_tabela = dblAltera_Total_preco_tabela * dblAltera_Quantidade_trabalhar
               
'               If lblTotal_tabela.Caption = "" Then lblTotal_tabela.Caption = 0
            
'               If dblAltera_Quantidade_trabalhar >= 0 Then
'                  lblTotal_tabela.Caption = CDbl(lblTotal_tabela.Caption) + dblAltera_Total_preco_tabela
'               Else
'                  lblTotal_tabela.Caption = CDbl(lblTotal_tabela.Caption) - (dblAltera_Total_preco_tabela * (-1))
'               End If
'            End If
'
'            'Calculando a divergência
'            If lblTotal_tabela.Caption = "" Then lblTotal_tabela.Caption = 0
'            lblDivergencia.Caption = lblTotal_tabela.Caption - CDbl(lblTotal_praticado)
'            If lblDivergencia.Caption <> 0 Then
'               If CDbl(lblTotal_praticado.Caption) > CDbl(Me.lblTotal_tabela.Caption) Then
'                  lblDivergencia.Caption = lblDivergencia * (-1)
'                  lblDivergencia.ForeColor = &HC00000
'               Else
'                  lblDivergencia.ForeColor = &HFF&
'               End If
'            Else
'               lblDivergencia.ForeColor = &HC00000
'            End If
'
'            'Resumo Logistíco
'            'Calculando o Peso Liquído e bruto dos Itens
'            'Calculando o total do peso por Item
'            dblAltera_Total_peso_liquido = dblAltera_Quantidade_trabalhar * dblPeso_Liquido
'            dblAltera_Total_peso_bruto = dblAltera_Quantidade_trabalhar * dblPeso_Bruto
'
'            If lblPeso_Bruto = "" Then lblPeso_Bruto = 0
'            If lblPeso_Liquido = "" Then lblPeso_Liquido = 0
'
'            If dblAltera_Quantidade_trabalhar >= 0 Then
'               lblPeso_Liquido.Caption = lblPeso_Liquido + dblAltera_Total_peso_liquido
'               lblPeso_Bruto.Caption = lblPeso_Bruto + dblAltera_Total_peso_bruto
'            Else
'               lblPeso_Liquido.Caption = lblPeso_Liquido - (dblAltera_Total_peso_liquido * (-1))
'               lblPeso_Bruto.Caption = lblPeso_Bruto - (dblAltera_Total_peso_bruto * (-1))
'            End If
'
'            Call Formatar_Resumos
            End If
            Exit Do
        End If
        '' FIM ROTINA ALTERAÇÂO - ALTERA ITEM
        intContador = intContador + 1
    Loop
    
Nao_Modificado:

    If booItem_Alterado = False Then
    
        'Este if foi inserido aqui para corrigir um bug no componente flexgrid
        'BUG --> Diz respeito a contagem interna do indice de linhas do objeto
        If hfgItem_cupom.Rows - 2 = 0 Then hfgItem_cupom.Row = 0
        
        hfgItem_cupom.Row = hfgItem_cupom.TopRow
        If hfgItem_cupom.Text <> Empty Then
           hfgItem_cupom.Row = hfgItem_cupom.Rows
           hfgItem_cupom.Rows = hfgItem_cupom.Rows + 1
        Else
           hfgItem_cupom.Row = hfgItem_cupom.Rows - 1
        End If
 
        hfgItem_cupom.Col = 0
        hfgItem_cupom.ColWidth(0) = 500
        hfgItem_cupom.Font.Name = "Tahoma"
        hfgItem_cupom.CellFontSize = 7
        hfgItem_cupom.CellFontBold = False
        hfgItem_cupom.CellBackColor = &H80FFFF
        hfgItem_cupom.Text = hfgItem_cupom.Row
        
        hfgItem_cupom.Col = 1
        hfgItem_cupom.Text = txtProduto.Text
        
        hfgItem_cupom.Col = 2
        hfgItem_cupom.Text = dtcProduto.Text
        
        hfgItem_cupom.Col = 3
        hfgItem_cupom.Text = txtQuantidade_produto.Text
        
        hfgItem_cupom.Col = 4
        hfgItem_cupom.Text = dtcUnidade_venda.Text
        
        hfgItem_cupom.Col = 5
        hfgItem_cupom.Text = txtPreco_unitario.Text
        
        hfgItem_cupom.Col = 6
        hfgItem_cupom.Text = txtPercentual_desconto.Text
        
        hfgItem_cupom.Col = 7
        hfgItem_cupom.Text = txtTotal_item.Text
        
        hfgItem_cupom.Refresh
        
        txtProduto.SetFocus
        
        'Monta resumo do total de Itens no Pedido
        lblTotal_Itens.Caption = Me.hfgItem_cupom.Rows - 2
        
        'Montagem dos Resumos
        Call Monta_Resumos
    End If
        
    'Limpando os campos dos Itens
    txtProduto.Text = ""
    dtcProduto.Text = ""
    txtQuantidade_produto = ""
    dtcUnidade_venda.Text = ""

    txtPreco_unitario.Text = ""
    txtPercentual_desconto.Text = ""
    txtTotal_item.Text = ""
  
    Me.hfgItem_cupom.SetFocus
    Me.hfgItem_cupom.TopRow = Me.hfgItem_cupom.Rows - 2
    txtProduto.SetFocus
    
    Exit Sub
    
Erro:
    
    MsgBox "Ocorreu um erro n°: " & Err.Number & " -- Inclusão de Itens no Pedido -- " & Err.Description & "Última Query - " & strSQL, vbInformation, "Only Tech"
    Exit Sub

End Sub

Private Sub cmdLimpar_Click()

     'Limpando os campos dos Itens
     txtProduto.Text = ""
     dtcProduto.Text = ""
     txtQuantidade_produto = ""
     dtcUnidade_venda.Text = ""
     txtPreco_unitario.Text = ""
     txtPercentual_desconto.Text = ""
     txtTotal_item.Text = ""
     txtProduto.SetFocus
     
End Sub

Private Sub cmdRefresh_Click()

    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta

End Sub

Private Sub cmdRemover_Item_Click()

    Dim intIndice As Integer
    Dim intContador As Integer
    
    If hfgItem_cupom.Rows - 2 = 0 Then
       MsgBox "Não existem itens à serem removidos!", vbCritical, "Only Tech"
       txtProduto.SetFocus
       Exit Sub
    End If
    
    If txtProduto.Text = "" Then
       MsgBox "Não existe item a ser excluído.Verifique!", vbInformation, "Only Tech"
       txtProduto.SetFocus
       Exit Sub
    End If
    
    intContador = 1
    
    Do While intContador <= hfgItem_cupom.Rows - 2
        hfgItem_cupom.Col = 1
        hfgItem_cupom.Row = intContador
        If hfgItem_cupom.Text = txtProduto.Text Then
           intIndice = intContador
           Exit Do
        End If
        intContador = intContador + 1
    Loop
    
    hfgItem_cupom.RemoveItem (intIndice)
    
    'Recalcula os resumos, subtraindo
    'Call Subtrai_Resumos
   
    'Monta resumo do total de Itens no Pedido
    lblTotal_Itens.Caption = Me.hfgItem_cupom.Rows - 2 & " Item(s) no pedido"
    
    'Limpando os campos dos Itens
    txtProduto.Text = ""
    dtcProduto.Text = ""
    txtQuantidade_produto = ""
    dtcUnidade_venda.Text = ""
    txtPreco_unitario.Text = ""
    txtPercentual_desconto.Text = ""
    txtTotal_item.Text = ""
    txtProduto.SetFocus

End Sub


Private Sub dtcEmitente_GotFocus()
    If Me.txtCliente.Text = Empty Then
        Call Movimentacoes.Verifica_DataCombo(dtcEmitente.Text)
    End If
End Sub

Private Sub dtcEmitente_LostFocus()

    txtCliente.Text = dtcEmitente.BoundText
    
    If IsNumeric(txtCliente.Text) = False Or dtcEmitente.Text = Empty Then
        txtCliente.Text = Empty
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
    
    If IsNumeric(txtProduto.Text) = False Or dtcProduto.Text = Empty Then
        txtProduto.Text = Empty
    Else
        If txtProduto.Text <> Empty Then
            strSQL = Empty
            strSQL = "SELECT IXCodigo_TBProduto, DFUnidade_venda_TBProduto FROM TBProduto WHERE IXCodigo_TBProduto = " & txtProduto.Text & " "
            Movimenta_DataCombo "IXCodigo_TBProduto", "DFUnidade_venda_TBProduto", dtcUnidade_venda, strSQL, "BDRetaguarda", "Otica", Me
        End If
    End If

End Sub

Private Sub dtcVendedor_GotFocus()
    If Me.txtVendedor.Text = Empty Then
        Call Movimentacoes.Verifica_DataCombo(dtcVendedor.Text)
    End If
End Sub

Private Sub dtcVendedor_LostFocus()

    txtVendedor.Text = dtcVendedor.BoundText
    
    If IsNumeric(txtVendedor.Text) = False Or dtcVendedor.Text = Empty Then
        txtVendedor.Text = Empty
    End If
        
End Sub

Private Sub dtpData_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos no dataPicker pelo ENTER
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
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
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Cupom Fiscal"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstCupon, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando Cadastro de Cupom Fiscal"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    sstCupon.TabEnabled(0) = False
    sstCupon.Tab = 1
        
    Call Reposicao
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    
    lblAte.Visible = False
    dtpData_incial.Visible = False
    dtpData_final.Visible = False
    
    Exit Sub
    
Erro:

    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando Cadastro de Cupom Fiscal"
        
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

Private Sub hfgCupom_fiscal_Click()
    Dim intContador As String
    
    hfgItem_cupom.Clear
    hfgItem_cupom.ClearStructure
    
    If hfgCupom_fiscal.Col = 0 And hfgCupom_fiscal.Text <> Empty Then
    
        Call Objetos.Limpa_TXT(Me)
    
        ' Novo
        tlbBotoes.Buttons.Item(1).Enabled = False
        ' Gravar
        tlbBotoes.Buttons.Item(2).Enabled = booPrivilegio_Alterar
        ' Cancelar
        tlbBotoes.Buttons.Item(3).Enabled = True
        ' Excluir
        tlbBotoes.Buttons.Item(4).Enabled = booPrivilegio_Excluir
        ' Imprimir
        tlbBotoes.Buttons.Item(5).Enabled = False
        'Integração
        If booIntegra_Portal = True Then
           tlbBotoes.Buttons.Item(11).Enabled = True
        End If
        
        frmAguarde.Show
        DoEvents
        
        On Error Resume Next
        
        ' Preenchendo os campos do Corpo do Cupom Fiscal
        ' strTamanho = "0,1000,3500,1000,3500,500,1500,1500,500,1500,3500,2000,2000,1000,1000,2000,0,0,0,0"
        ' strCampos = "ID,Código,Empresa,Código,Vendedor,ECF,Tipo Operação,Número,Série,Código,Cliente,Total dos Itens,Total do Cupom,Data,Hora,Operador,Cancelado,Motivo,Usuário,Observação"

        dblIdCupom = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 1))
        txtVendedor.Text = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 2))
        dtcCodigo_ecf.Text = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 4))
        cbbTipo_operacao.Text = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 5))
        txtNumero_cupom.Text = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 6))
        txtSerie.Text = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 7))
        txtCliente.Text = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 8))
        lblTotal_Itens.Caption = Format(hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 10)), "#,###0.00")
        lblTotal_Pedido.Caption = Format(hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 11)), "#,###0.00")
        dtpData.Value = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 12))
        dtpHora.Value = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 13))
        txtOperador.Text = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 14))
        
        If hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 16)) = "Sim" Then
           lblSituacao.Caption = "CANCELADO"
        Else
           lblSituacao.Caption = ""
        End If
        
        txtCodigo_Pdv.Text = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 20))
        dtcEmpresa.BoundText = hfgCupom_fiscal.TextArray((hfgCupom_fiscal.Row * hfgCupom_fiscal.Cols + hfgCupom_fiscal.Col + 21))
       
       'Preenchendo o grid com os itens do cupom
       Dim rstItem_Cupom As New ADODB.Recordset
    
       strSQL = "SELECT TBItens_cupom.DFCodigo_TBProduto," & _
                "TBProduto.DFDescricao_TBProduto," & _
                "TBItens_cupom.DFQuantidade_TBItens_cupom," & _
                "TBItens_cupom.DFUnidade_TBItens_cupom," & _
                "TBItens_cupom.DFPreco_praticado_TBItens_cupom," & _
                "TBItens_cupom.DFPercentual_desconto_TBItens_cupom," & _
                "TBItens_cupom.DFValor_total_praticado_TBItens_cupom," & _
                "PKId_TBItens_cupom " & _
                "FROM TBItens_cupom " & _
                "INNER JOIN TBProduto ON TBProduto.IXCodigo_TBProduto = TBItens_cupom.DFCodigo_TBProduto " & _
                "WHERE FKId_TBCupom = " & dblIdCupom & ""
        
        Call Movimentacoes.Movimenta_HFlex_Grid(strSQL, hfgItem_cupom, strTamanho_Produto, strCampo_Produto, "BDRetaguarda", "Otica", Me)
        
        lblTotal_Produtos.Caption = 0
        lblTotal_Pedido.Caption = 0
        lblTotal_Itens.Caption = 0
        
        intContador = 1
        hfgItem_cupom.Col = 1
        hfgItem_cupom.Row = 1
        
        Dim strQuantidade As String
        Dim strPreco As String
        
        If hfgItem_cupom.Text = Empty Then
           hfgItem_cupom.Rows = 2
           Call Movimentacoes.Monta_HFlex_Grid(hfgItem_cupom, strTamanho_Produto, strCampo_Produto, 8, "Otica", Me)
           lblDesconto_Itens.Caption = "0,00"
           lblTotal_Itens.Caption = "0"
        ElseIf hfgItem_cupom.Text <> Empty Then
           Do While intContador <= hfgItem_cupom.Rows - 1

               hfgItem_cupom.Row = intContador
        
               hfgItem_cupom.Col = 3
               hfgItem_cupom.Text = Format(hfgItem_cupom.Text, "#,###0.00")
               
               strQuantidade = hfgItem_cupom.Text

               hfgItem_cupom.Col = 5
               hfgItem_cupom.Text = Format(hfgItem_cupom.Text, "#,###0.00")
                
               strPreco = hfgItem_cupom.Text
                
               hfgItem_cupom.Col = 6
               hfgItem_cupom.Text = Format(hfgItem_cupom.Text, "#,###0.00")

               hfgItem_cupom.Col = 7
               hfgItem_cupom.Text = Format(hfgItem_cupom.Text, "#,###0.00")
                
               lblTotal_Produtos.Caption = Format(CDbl(lblTotal_Produtos.Caption) + (CDbl(strQuantidade) * CDbl(strPreco)), "#,###0.00")
               lblTotal_Pedido.Caption = Format(CDbl(lblTotal_Pedido.Caption) + CDbl(hfgItem_cupom.Text), "#,###0.00")
               
               intContador = intContador + 1
            Loop
            lblDesconto_Itens.Caption = Format(CDbl(lblTotal_Produtos.Caption) - CDbl(lblTotal_Pedido.Caption), "#,###0.00")
            lblTotal_Itens.Caption = hfgItem_cupom.Rows - 1
        End If
        
        ' Habilitando componentes, setando foco e indo para a Guia Geral
        booAlterar = True
        txtConsulta.Text = Empty
        txtProduto.SetFocus
        
        sstCupon.TabEnabled(0) = True
        sstCupon.Tab = 0
    End If
    
    Unload frmAguarde

End Sub

Private Sub hfgCupom_fiscal_DblClick()
    hfgCupom_fiscal.Sort = 1
End Sub

Private Sub hfgCupom_fiscal_KeyPress(KeyAscii As Integer)
'   Retorna campos do grid com espaço
    If KeyAscii = 32 Then
       Call hfgCupom_fiscal_Click
    End If
End Sub

Private Sub hfgItem_cupom_Click()

    If hfgItem_cupom.Col = 0 Then
        On Error Resume Next
        txtProduto.Text = hfgItem_cupom.TextArray((hfgItem_cupom.Row * hfgItem_cupom.Cols + hfgItem_cupom.Col + 1))
        dtcProduto.BoundText = hfgItem_cupom.TextArray((hfgItem_cupom.Row * hfgItem_cupom.Cols + hfgItem_cupom.Col + 2))
        txtQuantidade_produto.Text = Format(hfgItem_cupom.TextArray((hfgItem_cupom.Row * hfgItem_cupom.Cols + hfgItem_cupom.Col + 3)), "#,###0.00")
        dtcUnidade_venda.Text = hfgItem_cupom.TextArray((hfgItem_cupom.Row * hfgItem_cupom.Cols + hfgItem_cupom.Col + 4))
        txtPreco_unitario.Text = Format(hfgItem_cupom.TextArray((hfgItem_cupom.Row * hfgItem_cupom.Cols + hfgItem_cupom.Col + 5)), "#,###0.00")
        txtPercentual_desconto.Text = Format(hfgItem_cupom.TextArray((hfgItem_cupom.Row * hfgItem_cupom.Cols + hfgItem_cupom.Col + 6)), "#,###0.00")
        txtTotal_item.Text = Format(hfgItem_cupom.TextArray((hfgItem_cupom.Row * hfgItem_cupom.Cols + hfgItem_cupom.Col + 7)), "#,###0.00")
      
        If booAlterar = True Then
            booItem_Alterar = True
        End If
        txtProduto_Change
        dtcProduto.SetFocus
        ' Desabilitando a opção de alteração de item
        cmdIncluir_Item.Enabled = False
    End If

End Sub

Private Sub sstCupon_Click(PreviousTab As Integer)
    If sstCupon.Tab = 0 Then
       txtVendedor.SetFocus
    ElseIf sstCupon.Tab = 1 Then
       If frmIntegracao.Visible = True Then
          Unload frmIntegracao
       End If
       If strCombo <> Empty And strCombo <> "Todos" Then
          cbbCampos.Text = strCombo
          txtConsulta.SetFocus
       ElseIf strCombo = "Todos" Then
          hfgCupom_fiscal.Row = 1
          hfgCupom_fiscal.Col = 0
          hfgCupom_fiscal.SetFocus
       End If
    End If
End Sub


Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
'           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 11: Call Integracao
        End Select
End Sub

Function Gravar()
    
    On Error GoTo Erro
    
    ' Validando os campos que não podem receber valores nulos
    
    If dtcCodigo_ecf.BoundText = Empty Then
        MsgBox "O campo PDV não pode ser nulo. Verifique!", vbInformation, "Only Tech"
        dtcCodigo_ecf.SetFocus
        Exit Function
    ElseIf cbbTipo_operacao.Text = Empty Then
        MsgBox "O campo Tipo de Operação não pode ser nulo. Verifique!", vbInformation, "Only Tech"
        cbbTipo_operacao.SetFocus
        Exit Function
    ElseIf txtNumero_cupom.Text = Empty Then
        MsgBox "O campo Número do Cupom não pode ser nulo. Verifique!", vbInformation, "Only Tech"
        txtNumero_cupom.SetFocus
        Exit Function
    ElseIf txtCliente.Text = Empty Then
        MsgBox "O campo Cliente não pode ser nulo. Verifique!", vbInformation, "Only Tech"
        txtCliente.SetFocus
        Exit Function
    ElseIf txtOperador.Text = Empty Then
        MsgBox "O campo Operador não pode ser nulo. Verifique!", vbInformation, "Only Tech"
        txtOperador.SetFocus
        Exit Function
    End If
    
    If txtVendedor.Text = Empty Then
        MsgBox "O campo vendedor não pode ser nulo. Verifique!", vbInformation, "Only Tech"
        txtVendedor.SetFocus
        Exit Function
    End If
    
    ' Verificando se há produtos cadastrados no Cupom
    
    hfgItem_cupom.Row = 1: hfgItem_cupom.Col = 1
    If hfgItem_cupom.Text = Empty Then
       MsgBox "Não foram digitado produtos para esse cupom. Verifique!", vbInformation, "Only Tech"
       txtProduto.SetFocus
       Exit Function
    End If
    
    Call Objetos.Retira_Espaco_Lateral(Me)
    Call Objetos.Maiusculo_TXT(Me)
    
    ' Pegando os Ids das tabelas para gravar na TBCupom
    
    Dim dblId_Vendedor As Double
    Dim intTipo_operacao As Integer

    Select Case cbbTipo_operacao.Text
        Case "Venda"
            intTipo_operacao = 1
        Case "Transferência"
            intTipo_operacao = 2
        Case "Dev.Cliente"
            intTipo_operacao = 3
        Case "Dev.Fornec."
            intTipo_operacao = 4
        Case "Outras"
            intTipo_operacao = 5
    End Select
    
    dblId_Vendedor = Funcoes_Gerais.Localiza_ID("PKId_TBVendedor", "IXCodigo_TBVendedor", txtVendedor.Text, "TBVendedor", "Otica", Me, "BDRetaguarda")
    
    ' Passando os campos e valores a serem gravados (ALTERACAO E INCLUSAO)
    
    Dim strCampo As String
    Dim strValores As String
    
    strCampo = "FKCodigo_TBEmpresa, FKId_TBVendedor, PKCodigo_TBPdv, DFTipo_operacao_TBCupom, DFNumero_TBCupom," & _
               "DFSerie_TBCupom, DFEmitente_TBCupom, DFTotal_itens_TBCupom," & _
               "DFTotal_cupom_TBCupom, DFTotal_cupom_tabela_TBCupom, DFData_Saida_TBCupom, DFHora_Saida," & _
               "DFCancelado_TBCupom, DFMotivo_cancelamento_TBCupom," & _
               "DFUsuario_cancelamento_TBCupom, DFIntegrado_fiscal_TBCupom, DFBase_calculo_subst_tributaria_TBCupom," & _
               "DFValor_subst_tributaria_TBCupom, DFObservacao_TBCupom, DFCupom_Registrado_TBCupom," & _
               "FKCodigo_TBOperadores_ecf,DFData_alteracao_TBCupom,DFIntegrado_filiais_TBCupom"
               
    If booIntegra_Portal = True Then
       strCampo = strCampo & ",DFIntegrado_portal_TBCupom"
    End If
               
    strValores = "" & dtcEmpresa.BoundText & "," & dblId_Vendedor & "," & dtcCodigo_ecf.BoundText & "," & _
                 "" & intTipo_operacao & "," & _
                 "" & txtNumero_cupom.Text & "," & txtSerie.Text & "," & 0 & "," & txtCliente.Text & "," & _
                 "" & Funcoes_Gerais.Grava_String(lblTotal_Itens.Caption) & "," & _
                 "" & Funcoes_Gerais.Grava_String(lblTotal_Pedido.Caption) & "," & _
                 "" & Funcoes_Gerais.Grava_String(lblTotal_Pedido.Caption) & "," & Format(dtpData.Value, "YYYYMMDD") & "," & _
                 "" & Format(dtpHora.Value, "hh:mm:ss") & "," & _
                 "" & 0 & "," & "Mot. Teste" & "," & "Us. Teste" & "," & _
                 "" & 0 & "," & 0 & "," & 0 & "," & "Obs. Teste" & "," & 0 & "," & _
                 "" & txtOperador.Text & ",'" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
     End If

    On Error GoTo Erro_Controle_Transacao
    
    ' Indicando o banco à conectar-se e estabelecendo a conexão com o banco
    
    cnGravacao.Initial_Catalog = "BDRetaguarda"
    cnGravacao.Abrir_conexao ("Otica")
    
    ' Verificando alteração ou inclusão
    
    strSQL = Empty
        
    If booAlterar = True Then
        strSQL = "UPDATE TBCupom " & _
                 "SET FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & "," & _
                 "FKId_TBVendedor = " & dblId_Vendedor & "," & _
                 "PKCodigo_TBPdv =  " & dtcCodigo_ecf.BoundText & "," & _
                 "DFTipo_operacao_TBCupom = " & intTipo_operacao & "," & _
                 "DFNumero_TBCupom = " & txtNumero_cupom.Text & "," & _
                 "DFSerie_TBCupom = '" & txtSerie.Text & "'," & _
                 "DFEmitente_TBCupom = " & txtCliente.Text & "," & _
                 "DFTotal_itens_TBCupom = " & Funcoes_Gerais.Grava_Moeda(lblTotal_Itens.Caption) & "," & _
                 "DFTotal_cupom_TBCupom = " & Funcoes_Gerais.Grava_Moeda(lblTotal_Pedido.Caption) & "," & _
                 "DFTotal_cupom_tabela_TBCupom  = " & Funcoes_Gerais.Grava_Moeda(lblTotal_Pedido.Caption) & "," & _
                 "DFData_Saida_TBCupom = '" & Format(dtpData.Value, "YYYYMMDD") & "'," & _
                 "DFHora_Saida = '" & Format(dtpHora.Value, "hh:mm:ss") & "'," & _
                 "DFCancelado_TBCupom = " & 0 & "," & _
                 "DFMotivo_cancelamento_TBCupom = 'Mot. Teste' ," & _
                 "DFUsuario_cancelamento_TBCupom = 'Us. Teste'," & _
                 "DFIntegrado_fiscal_TBCupom = " & 0 & "," & _
                 "DFBase_calculo_subst_tributaria_TBCupom = " & 0 & "," & _
                 "DFValor_subst_tributaria_TBCupom = " & 0 & "," & _
                 "DFObservacao_TBCupom = 'Obs. Teste'," & _
                 "DFCupom_Registrado_TBCupom =" & 0 & "," & _
                 "FKCodigo_TBOperadores_ecf = " & txtOperador.Text & " ," & _
                 "DFData_alteracao_TBCupom = '" & Format(Date, "YYYYMMDD") & "'," & _
                 "DFIntegrado_filiais_TBCupom = 0 "
                 
        If booIntegra_Portal = True Then
           strSQL = strSQL & ",DFIntegrado_portal_TBCupom = 0 "
        End If
                 
        strSQL = strSQL & "WHERE PKId_TBCupom = " & dblIdCupom & " "
                 
        log.Evento = "Alterar"
        ' Iniciando a Transação
        cnGravacao.CNconexao.BeginTrans
        ' Alterando o corpo do Pedido
        cnGravacao.CNconexao.Execute strSQL
        ' Executando a SQL de Alteração do corpo do cupon
        cnGravacao.CNconexao.Execute strSQL
        'log.Descricao = "Alterando o registro: " + txtCodigo.Text
        log.Tipo = 1
        log.Hora = Format(Now, "hh:mm:ss")
        log.Gravar_log "Otica", Me
        DoEvents
        ' Efetivando a gravação do registro
        cnGravacao.CNconexao.CommitTrans
    Else
    
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBCupom", strCampo, strValores, "OTICA", Me, "BDRetaguarda")
   
        Dim rstID_Cupom As New ADODB.Recordset
        strSQL = Empty
        strSQL = "SELECT MAX (PKID_TBCupom) AS Cupom FROM TBCupom "
        Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstID_Cupom, "Otica", Me
        ' Guarda o ID desse cupom para ser utilizado para gravar os Itens
        'lngID_Cupom = rstID_Cupom!Cupom
        'Call Alterar_Itens_Pedido(lngID_Cupom)
        ' Registrando a operação no Log
        'log.Descricao = "Gravando o registro: " + CStr(lngID_Cupom)
        log.Tipo = 1
        log.Hora = Format(Now, "hh:mm:ss")
        log.Gravar_log "Otica", Me
    End If
    
    'Integração
    tlbBotoes.Buttons.Item(11).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    sstCupon.Tab = 1
    sstCupon.TabEnabled(0) = False

    Exit Function

Erro_Controle_Transacao:
    
    'Em caso de erro a trasação sofre um rollback e a conexão é fechada.
    cnGravacao.CNconexao.RollbackTrans
    cnGravacao.Fechar_conexao
    
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
    
Erro:
    
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
    
End Function
Private Function Excluir()
'''    Dim intContador_Exclusao As String
'''
'''    ' Indicando o banco à conectar-se
'''    cnGravacao.Initial_Catalog = "BDRetaguarda"
'''
'''    ' Estabelecendo conexão com o banco
'''    cnGravacao.Abrir_conexao ("Otica")
'''
'''    ' Dando inicio a transação
'''    cnGravacao.CNconexao.BeginTrans
'''
'''    On Error GoTo Erro
'''
'''    Dim adrDeleta As New ADODB.Recordset
'''
'''    log.Evento = "Excluir"
'''    log.Descricao = "Exclusão do registro: " & dblIdCupom
'''    log.Tipo = 1
'''    log.Hora = Format(Now, "hh:mm:ss")
'''
'''    ' Montar rotina pra adicionar o estoque
'''
'''    intContador_Exclusao = 1
'''
'''    Do While intContador_Exclusao <= hfgItem_cupom.Rows - 2
'''       Dim rstEstoque_item_Exclusao As New ADODB.Recordset
'''       hfgItem_cupom.Col = 1
'''       hfgItem_cupom.Row = intContador_Exclusao
'''
'''       'Informações do Produto
'''       strSql = Empty
'''       strSql = "SELECT * FROM TBProduto Where IXCodigo_TBProduto = " & hfgItem_cupom.Text & " "
'''       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstEstoque_item_Exclusao, "Otica", Me
'''
'''       'Informações do Pedido
'''       strSql = "SELECT * FROM TBItens_pedido WHERE FKId_TBPedido = " & intPedido & ""
'''       Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstItens_Gravados_Pedido_Exclusao, "Otica", Me)
'''
'''       'Pegando o ID do Produto
'''       intIDProduto_Exclusao = Funcoes_Gerais.Localiza_ID("PKId_TBProduto", "IXCodigo_TBProduto", " " & Me.hfgItem_cupom.Text & " ", "TBProduto", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", " " & dtcEmpresa.BoundText & "")
'''       strSql = Empty
'''       strSql = "DELETE FROM TBItens_pedido WHERE FKId_TBpedido = " & intIDPedido & " AND FKId_TBProduto = " & intIDProduto_Exclusao & ""
'''       cnGravacao.CNconexao.Execute strSql
'''
'''       If booBaixar_Estoque = True Then
'''          If intManipula_Estoque = 1 Then
'''             'Ocorrência de Produto
'''             ocorrencia.Data_Movimento = Date
'''             ocorrencia.Estoque_Anterior = CDbl(rstEstoque_item_Exclusao!DFEstoque_atual_TBProduto)
'''             ocorrencia.Estoque_Atual = CDbl(rstEstoque_item_Exclusao!DFEstoque_atual_TBProduto) + CDbl(rstItens_Gravados_Pedido_Exclusao!DFQuantidade_baixa_estoque_TBItens_pedido)
'''             ocorrencia.Hora_Movimento = Format(Now, "hh:mm:ss")
'''             ocorrencia.ID_Produto = intIDProduto_Exclusao
'''             ocorrencia.Observacao = "Exclusão de Item no cadastro de Pedido Nº:" & intPedido & " - Adição no Estoque"
'''             ocorrencia.Programa = "Cadastro de Pedido"
'''             ocorrencia.Quantidade_Movimentada = CDbl(rstItens_Gravados_Pedido_Exclusao!DFQuantidade_baixa_estoque_TBItens_pedido)
'''             ocorrencia.Usuario = MDIPrincipal.OCXUsuario.Nome
'''             ocorrencia.Gravar "Otica", True, cnGravacao
'''             '--------------------------------------------------------------------------------------
'''             'Manipulação do Estoque.
'''             'Adicionando
'''             Estoque.ID_Produto = intIDProduto_Exclusao
'''             Estoque.Quantidade_Menor_Unidade_Item = CDbl(rstItens_Gravados_Pedido_Exclusao!DFQuantidade_baixa_estoque_TBItens_pedido)
'''             Estoque.Quantidade_Antes_Atualizar_Estoque = rstEstoque_item_Exclusao!DFEstoque_atual_TBProduto
'''             Estoque.Adicionar_Estoque "Otica", True, cnGravacao
'''          End If
'''
'''          If intManipula_Estoque = 2 Then
'''             ocorrencia.Data_Movimento = Date
'''             ocorrencia.Estoque_Anterior = CDbl(rstEstoque_item_Exclusao!DFEstoque_atual_TBProduto)
'''             ocorrencia.Estoque_Atual = CDbl(rstEstoque_item_Exclusao!DFEstoque_atual_TBProduto) - CDbl(rstItens_Gravados_Pedido_Exclusao!DFQuantidade_baixa_estoque_TBItens_pedido)
'''             ocorrencia.Hora_Movimento = Format(Now, "hh:mm:ss")
'''             ocorrencia.ID_Produto = intIDProduto_Exclusao
'''             ocorrencia.Observacao = "Exclusão de Item no cadastro de Pedido Nº:" & " & Numero_Pedido & " & "- Baixa de Estoque"
'''             ocorrencia.Programa = "Cadastro de Pedido"
'''             ocorrencia.Quantidade_Movimentada = CDbl(rstItens_Gravados_Pedido_Exclusao!DFQuantidade_baixa_estoque_TBItens_pedido)
'''             ocorrencia.Usuario = MDIPrincipal.OCXUsuario.Nome
'''             ocorrencia.Gravar "Otica", True, cnGravacao
'''             'Subtraindo o Item no Estoque , para depois subtrair a quantidade alterada
'''             Estoque.ID_Produto = intIDProduto_Exclusao
'''             Estoque.Quantidade_Menor_Unidade_Item = CDbl(rstItens_Gravados_Pedido_Exclusao!DFQuantidade_baixa_estoque_TBItens_pedido)
'''             Estoque.Quantidade_Antes_Atualizar_Estoque = rstEstoque_item_Exclusao!DFEstoque_atual_TBProduto
'''             Estoque.Subtrair_Estoque "Otica", True, cnGravacao
'''             'Adicionando ao estoque
'''             Estoque.ID_Produto = intIDProduto_Exclusao
'''             Estoque.Quantidade_Menor_Unidade_Item = CDbl(rstItens_Gravados_Pedido_Exclusao!DFQuantidade_baixa_estoque_TBItens_pedido)
'''             Estoque.Quantidade_Antes_Atualizar_Estoque = rstEstoque_item_Exclusao!DFEstoque_atual_TBProduto
'''             Estoque.Subtrair_Estoque "Otica", True, cnGravacao
'''          End If
'''       End If
'''
'''       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''       'Comitando e abrindo nova transação                                                              '
'''       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''       'Indica o sucesso da transação de gravação de corpo do pedido, dentro da função de gravação de   '
'''       'pedidos itens será feito outro controle de transação para que haja todo o controle transacional '
'''       'necessário.                                                                                     '
'''       'OBS: Esse COMMIT foi colocado aqui pois havia a necessidade de se obter o número do atual pedido'
'''       'para que pudesse ser gravado os itens do mesmo.                                                 '
'''       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''       cnGravacao.CNconexao.CommitTrans
'''       'Indica o inicio das transações restantes desse pedido
'''       cnGravacao.CNconexao.BeginTrans
'''
'''       intContador_Exclusao = intContador_Exclusao + 1
'''       Set rstEstoque_item_Exclusao = Nothing
'''       Set rstItens_Gravados_Pedido_Exclusao = Nothing
'''
'''    Loop
'''
'''    'Matando "Filhos" do registro na tabela de TBItens_Pedido
'''    cnGravacao.CNconexao.Execute "DELETE FROM TBItens_pedido WHERE FKId_TBpedido = " & intIDPedido & " "
'''
'''    'Matando o registro na tabela de TBPedido a tabela "PAI"
'''    cnGravacao.CNconexao.Execute "DELETE FROM TBPedido WHERE PKId_TBPedido = " & intIDPedido & ""
'''
'''    'Gravando log
'''    log.Gravar_log "Otica", Me
'''
'''    Call Limpar_Interface
'''    Call Limpa_Tudo
'''
'''    'Numero do Pedido
'''    lblNumero_pedido.Visible = False
'''    lblNumero_pedido.Caption = ""
'''
'''    'Novo
'''    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
'''    'Gravar
'''    tlbBotoes.Buttons.Item(2).Enabled = False
'''    'Cancelar
'''    tlbBotoes.Buttons.Item(3).Enabled = False
'''    'Excluir
'''    tlbBotoes.Buttons.Item(4).Enabled = False
'''    'Imprimir
'''    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
'''
'''    If booPrivilegio_Consultar = False Then
'''       adgPedido = False
'''    End If
'''
'''   'Confirmando o sucesso da transação
'''    cnGravacao.CNconexao.CommitTrans
'''
'''    'Matando a conexão
'''    cnGravacao.CNconexao.Close
'''    Set cnGravacao = Nothing
'''    sstPedido.TabEnabled(0) = False
'''    sstPedido.Tab = 1
'''    strObservacao = Empty
'''
'''    Exit Function
'''Erro:
'''    'Cancelando toda a transação em caso de falha
'''    cnGravacao.CNconexao.RollbackTrans
'''
'''    'Matando a conexão
'''    cnGravacao.CNconexao.Close
'''    Set cnGravacao = Nothing
'''
'''    Call Erro.Erro(Me, "Otica", "Excluir")
'''    Exit Function
'''
'''
'''    Exit Function
'''
'''Erro:
'''
'''     Call Erro.Erro(Me, "OTICA", "Excluir")
'''     Exit Function
     
End Function

Private Function Cancelar()

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
    tlbBotoes.Buttons.Item(11).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
'    If booPrivilegio_Consultar = False Then
'       hfgCupom_fiscal.Visible = False
'    End If
    
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Cadastro de Cupom Fiscal"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me

    sstCupon.TabEnabled(0) = False
    sstCupon.Tab = 1
    
    hfgItem_cupom.Clear
    hfgItem_cupom.ClearStructure
    
    Exit Function
    
Erro:

    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
    
End Function

Private Function Novo()

    On Error GoTo Erro
    
    Call Reposicao
    Call Objetos.Limpa_TXT(Me)
    
'    intContador = 2
'
'    hfgItem_cupom.ClearStructure
'
'    Do While intContador <= hfgItem_cupom.Rows
'        hfgItem_cupom.Row = intContador - 1
'        hfgItem_cupom.RemoveItem intContador - 1
'        intContador = intContador + 1
'    Loop

    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    Call Verifica_Parametro_Venda
    
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
    
    sstCupon.TabEnabled(0) = True
    sstCupon.Tab = 0
    
    booAlterar = False
    cmdCancelamento.Enabled = False
    lblSituacao.Visible = False
    
    dtcCodigo_ecf.SetFocus
    
    Exit Function
    
Erro:

    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
    
End Function

Private Function Reposicao()

    On Error GoTo Erro

    strCampos = "ID,Código,Vendedor," & _
    "ECF,Tipo Operação," & _
    "Número,Série," & _
    "Cod. Cliente," & _
    "Cliente,Total Itens," & _
    "Total Cupom,Data," & _
    "Hora," & _
    "Código,Operador,Cancelado," & _
    "Motivo," & _
    "Usuário,Observação,Código PDV," & _
    "Código,Empresa"

    strTamanho = "0,1000,3500," & _
    "800,2500," & _
    "1200,800," & _
    "1200," & _
    "3500,1500," & _
    "1500,1000," & _
    "1200," & _
    "1000,2500,1500," & _
    "0,0,0,1700,1000,2500"
    
    Movimentacoes.Monta_HFlex_Grid hfgCupom_fiscal, strTamanho, strCampos, 22, "OTICA", Me
    
    'Item_cupom
    strCampo_Produto = "Cód.Produto,Produto,Quantidade,Unidade,Preço Unitário,% Desconto,Total Item,IDItem"

    strTamanho_Produto = "1000,4400,950,750,1100,1000,1200,0"
    
    Movimentacoes.Monta_HFlex_Grid hfgItem_cupom, strTamanho_Produto, strCampo_Produto, 8, "OTICA", Me
    
    Call Monta_Combo
    Call Monta_DataCombo

    dtpData.Value = Date
    dtpHora.Value = Format(Now, "hh:mm:ss")
    
    Exit Function

Erro:

   Call Erro.Erro(Me, "OTICA", "Reposicao")
   Resume Next

End Function
Private Function Monta_DataCombo()

    ' Empresa
    strSQL = "SELECT TBEmpresa.PKCodigo_TBEmpresa, DFRazao_Social_TBEmpresa FROM TBEmpresa "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSQL, "BDRetaguarda", "Otica", Me

    ' Vendedor
    strSQL = "SELECT IXCodigo_TBVendedor,DFNome_TBVendedor FROM TBVendedor "
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBVendedor", "DFNome_TBVendedor", dtcVendedor, strSQL, "BDRetaguarda", "Otica", Me
    
    ' Codigo ECF
    strSQL = "SELECT PKCodigo_TBPdv,DFNumero_ecf_TBPdv FROM TBPdv "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBPdv", "DFNumero_ecf_TBPdv", dtcCodigo_ecf, strSQL, "BDRetaguarda", "Otica", Me
    
    ' Cliente
    strSQL = "SELECT IXCodigo_TBCliente, DFNome_TBCliente FROM TBCliente "
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcEmitente, strSQL, "BDRetaguarda", "Otica", Me
    
    ' Produto
    strSQL = "SELECT IXCodigo_TBProduto, DFDescricao_TBProduto FROM TBProduto "
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSQL, "BDRetaguarda", "Otica", Me
    
    'Operador
    strSQL = "SELECT PKCodigo_TBOperadores_ecf, DFNome_TBOperadores_ecf FROM TBOperadores_ecf "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSQL, "BDRetaguarda", "Otica", Me

End Function

Private Sub txtCliente_Change()
    dtcEmitente.BoundText = txtCliente.Text
End Sub

Private Sub txtCliente_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCliente_LostFocus()
    dtcEmitente.BoundText = txtCliente.Text
    Call Valida_Cliente(1)
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Sub txtNumero_cupom_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_cupom_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtOperador_Change()
    dtcOperador.BoundText = txtOperador.Text
End Sub

Private Sub txtOperador_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub dtcOperador_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_LostFocus()
    txtOperador.Text = UCase(txtOperador.Text)
End Sub

Private Sub txtPercentual_desconto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPercentual_desconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPercentual_desconto_LostFocus()

    Dim dblValor_total As Double
    
    If IsNumeric(txtPreco_unitario) Then
        dblValor_total = CDbl(txtQuantidade_produto.Text) * CDbl(txtPreco_unitario.Text)
    End If

    If IsNumeric(txtPercentual_desconto.Text) = False Then
        txtPercentual_desconto.Text = 0
    Else
        txtTotal_item.Text = dblValor_total - (CDbl(txtPercentual_desconto.Text) * dblValor_total / 100)
    End If
    
    txtPercentual_desconto.Text = Format(txtPercentual_desconto.Text, "#,###0.00")
    txtTotal_item.Text = Format(txtTotal_item.Text, "#,###0.00")
    
End Sub
Private Sub txtPreco_unitario_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPreco_unitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_Unitario_LostFocus()
    txtPreco_unitario.Text = Format(txtPreco_unitario.Text, "#,###0.00")
    If IsNumeric(txtPreco_unitario.Text) Then
        txtTotal_item.Text = Format(CDbl(txtQuantidade_produto.Text) * CDbl(txtPreco_unitario.Text), "#,###0.00")
    End If
End Sub

Private Sub txtProduto_Change()
    dtcProduto.BoundText = txtProduto.Text
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
    dtcProduto.BoundText = txtProduto.Text
    
    If txtProduto.Text <> Empty Then
        strSQL = Empty
        strSQL = "SELECT IXCodigo_TBProduto, DFUnidade_venda_TBProduto FROM TBProduto WHERE IXCodigo_TBProduto = " & txtProduto.Text & " "
        Movimenta_DataCombo "IXCodigo_TBProduto", "DFUnidade_venda_TBProduto", dtcUnidade_venda, strSQL, "BDRetaguarda", "Otica", Me
    End If
End Sub

Private Sub txtQuantidade_produto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtQuantidade_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtSerie_GotFocus()

    If txtSerie.Text = Empty Then
        txtSerie.Text = "CP"
    End If
    
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
    
End Sub

Private Sub txtSerie_LostFocus()
    txtSerie.Text = UCase(txtSerie.Text)
End Sub


Private Sub txtTotal_Item_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtVendedor_Change()
    dtcVendedor.BoundText = txtVendedor.Text
End Sub

Private Sub txtVendedor_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtVendedor_LostFocus()
    dtcVendedor.BoundText = txtVendedor.Text
End Sub

Private Function Valida_Cliente(Optional intVerifica As Integer)
    Dim rstValida As New ADODB.Recordset
    Dim rstEmpresa As New ADODB.Recordset
    Dim rstCFO As New ADODB.Recordset
    Dim rstCliente_Representante As New ADODB.Recordset
    
    If txtCliente.Text = "" Then Exit Function
    If dtcEmitente.Text = "" Then Exit Function
    If Me.cbbTipo_operacao.Text = "" Then Exit Function
    
    If txtVendedor.Text = "" Then
       MsgBox "Digite um vendedor válido para este pedido.", vbInformation, "Only Tech"
       txtVendedor.Text = Empty
       txtVendedor.SetFocus
       Exit Function
    End If
    
    On Error GoTo Erro
    
    'Verifica se este cliente é do dado representante
    strSQL = Empty
    strSQL = "SELECT " & _
             "TBVendedor.IXCodigo_TBVendedor " & _
             "FROM TBCLIENTE " & _
             "INNER JOIN TBVendedor_cliente " & _
             "ON TBCLIENTE.PKId_TBCliente = TBVendedor_cliente.FKId_TBCliente " & _
             "INNER JOIN TBVendedor " & _
             "ON TBVendedor_cliente.FKId_TBVendedor = TBVendedor.PKId_TBVendedor " & _
             "WHERE TBCLIENTE.IXCodigo_TBCliente = " & txtCliente.Text & " " & _
             "AND TBCLIENTE.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
             "AND TBVendedor.IXCodigo_TBVendedor = " & txtVendedor.Text & " "
             
    Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstCliente_Representante, "Otica", Me)
    
    Dim intRetorno As Integer
    
    If rstCliente_Representante.EOF = True And rstCliente_Representante.BOF = True And booAlterar = False Then
        'Verificar na tabela de parametros se libera ou não a digitação
        If booLibera_vendedor_cliente_linha_pedido = True Then
            intRetorno = MsgBox("Este cliente não é deste dado representante, prosseguir com a venda?", vbYesNo, "Only Tech")
            If intRetorno = 7 Then
               txtCliente.Text = Empty
               txtCliente.SetFocus
               Exit Function
            End If
        Else
            MsgBox "Este cliente não é deste dado representante.Redigite!", vbInformation, "Only Tech"
            txtCliente.Text = Empty
            txtCliente.SetFocus
            Exit Function
        End If
    End If
    
    'Descarregando a recordset
    Set rstCliente_Representante = Nothing
    
    'Verifica se o cliente está ativo/inativo - bloqueado/não bloqueado
    strSQL = Empty
    strSQL = "SELECT TBCliente.DFTipo_pessoa_TBCliente,TBCliente.DFInativo_TBCliente," & _
             "TBCliente.DFBloqueado_TBCliente,TBCliente.FKId_TBPlano_pagamento," & _
             "TBCidade_otica.DFUf_TBCidade_otica,TBCliente.DFTipo_entrega_TBCliente " & _
             "FROM TBCliente  " & _
             "INNER JOIN TBCidade_otica " & _
             "ON FKId_TBCidade_otica = PKId_TBCidade_otica " & _
             "WHERE IXCodigo_TBCliente = " & Me.txtCliente.Text & " " & _
             "AND IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
             
    Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstValida, "Otica", Me)
    
    strUF_Emitente = rstValida!DFUf_TBCidade_otica
    
    'Bloqueado/Não -Bloqueado
    If rstValida!DFBloqueado_TBCliente = True Then
       MsgBox "Este cliente está bloqueado!", vbQuestion = vbOKOnly, "Only Tech"
       txtCliente.Text = ""
       txtCliente.SetFocus
       Exit Function
    End If
  
'    If intVerifica <> 1 Then
'       If rstValida.Fields("DFTipo_entrega_TBCliente") = 1 Then
'          chkPrevisao.Value = 1
'       ElseIf rstValida.Fields("DFTipo_entrega_TBCliente") = 2 Then
'          chkPrevisao.Value = 0
'       End If
'    End If
    
    'Ativo/Inativo
    If rstValida!DFInativo_TBCliente = True And booAlterar = False Then
       Dim intResult As Integer
       intResult = MsgBox("Este cliente está momentaneamente inativo!Deseja continuar com a digitação do pedido?", vbYesNo, "Only Tech")
       If intResult = 7 Then
          txtCliente.Text = ""
          txtCliente.SetFocus
          Exit Function
       End If
    End If
    
'    'Montagem do CFO correspondente
'    'Localizando a UF da empresa
'
'    strSql = Empty
'    strSql = "SELECT TBCidade_otica.DFUf_TBCidade_otica FROM TBEmpresa " & _
'             "INNER JOIN TBCidade_otica " & _
'             "ON FKId_TBCidade_otica = PKId_TBCidade_otica " & _
'             "WHERE PKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstEmpresa, "Otica", Me)
'
'    'Comparando a UF da Empresa com a UF da Cidade do cliente e verificando se é de fora do estado
'    If rstEmpresa!DFUf_TBCidade_otica <> strUF_Emitente Then
'       'Fora do Estado
'       'Venda
'       If cmbTipo_operacao.Text = "Venda" Then
'          strSql = Empty
'          strSql = "SELECT DFProximo_cfop_venda_fora_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
'                   "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'       End If
'       'Transferência
'       If cmbTipo_operacao.Text = "Transferência" Then
'          strSql = Empty
'          strSql = "SELECT DFProximo_cfop_transferencia_fora_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
'                   "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'       End If
'       'Dev.Cliente
'       If cmbTipo_operacao.Text = "Dev.Cliente" Then
'          strSql = Empty
'          strSql = "SELECT DFProximo_cfop_devolucao_cliente_fora_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
'                   "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'       End If
'       'Dev.Fornec.
'       If cmbTipo_operacao.Text = "Dev.Fornec." Then
'          strSql = Empty
'          strSql = "SELECT DFProximo_cfop_devolucao_fornecedor_fora_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
'                   "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'       End If
'
'    End If
'
'    If booAlterar = False Then
'        'Comparando a UF da Empresa com a UF da Cidade do cliente e verificando se é de dentro do estado
'        If rstEmpresa!DFUf_TBCidade_otica = strUF_Emitente Then
'           'Dentro do Estado
'           'Venda
'           If cmbTipo_operacao.Text = "Venda" Then
'              strSql = Empty
'              strSql = "SELECT DFProximo_cfop_venda_dentro_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
'                       "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'           End If
'           'Transferência
'           If cmbTipo_operacao.Text = "Transferência" Then
'              strSql = Empty
'              strSql = "SELECT DFProximo_cfop_transferencia_dentro_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
'                       "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'           End If
'           'Dev.Cliente
'           If cmbTipo_operacao.Text = "Dev.Cliente" Then
'              strSql = Empty
'              strSql = "SELECT DFProximo_cfop_devolucao_cliente_dentro_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
'                       "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'           End If
'           'Dev.Fornec.
'           If cmbTipo_operacao.Text = "Dev.Fornec." Then
'              strSql = Empty
'              strSql = "SELECT DFProximo_cfop_devolucao_fornecedor_dentro_estado_TBParametros_fiscais AS CFO FROM TBParametros_fiscais " & _
'                       "WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
'           End If
'        End If
'
'        'Só vai abastecer o combo se for diferente de outras, dando abertura para digitação da CFO para o Operador
'        If cmbTipo_operacao.Text <> "Outras" Then
'            Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstCFO, "Otica", Me)
'            If rstCFO.BOF <> True And rstCFO.EOF <> True Then
'                txtCFO.Text = rstCFO!CFO
'            End If
'        End If
'    End If
'
'    Dim rstTitulos_Aberto_Cliente As New ADODB.Recordset
'    Dim dblID_Cliente As Double
'
'    dblID_Cliente = Funcoes_Gerais.Localiza_ID("PKId_TBCliente", "IXCodigo_TBCliente", Me.txtCliente.Text, "TBCliente", "Otica", Me, "BDRetaguarda")
'
'    'Em aberto e Vencidos
'    strSql = Empty
'    strSql = "SELECT COUNT(PKId_TBTitulo_receber) AS CONT," & _
'             "SUM(DFValor_TBTitulo_receber) TOTAL_DEBITO " & _
'             "FROM TBTitulo_receber " & _
'             "WHERE TBTitulo_receber.PKId_TBTitulo_receber " & _
'             "NOT IN (SELECT TBTitulo_recebido.FKId_TBTitulo_receber FROM TBTitulo_recebido) " & _
'             "AND TBTitulo_receber.FKID_TBCliente = " & dblID_Cliente & "" & _
'             "AND DFData_vencimento_TBTitulo_receber < '" & Format(Now, "YYYYMMDD") & "'"
'
'    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstTitulos_Aberto_Cliente, "Otica", Me
'
'    If rstTitulos_Aberto_Cliente.EOF = False And rstTitulos_Aberto_Cliente.BOF = False Then
'       If rstTitulos_Aberto_Cliente!TOTAL_DEBITO > 0 Then
'           MsgBox "Este cliente possui títulos em aberto ( " & rstTitulos_Aberto_Cliente!cont & " ), no valor de R$ " & Format(rstTitulos_Aberto_Cliente!TOTAL_DEBITO, "#,###0.00") & ""
'       End If
'    End If
    
    'Descarregando as recordests
    Set rstEmpresa = Nothing
    Set rstCFO = Nothing
    Set rstValida = Nothing
    
    Exit Function
    
Erro:
    
    MsgBox "Ocorreu um erro n°: " & Err.Number & " -- Valida Cliente -- " & Err.Description & "Última Query - " & strSQL, vbInformation, "Only Tech"
    Exit Function
    
End Function

Private Function Verifica_Parametro_Venda()

    Dim rstVerifica_parametro As New ADODB.Recordset
    
    On Error GoTo Erro
    
    strSQL = Empty
    strSQL = "SELECT DFLibera_produto_linha_venda_pedido_TBParametro_venda,DFLibera_vendedor_cliente_pedido_TBParametro_venda,DFNumero_itens_orcamento_TBParametros_venda,DFNumero_itens_nota_TBParametros_venda,DFEmite_banco_TBParametros_venda " & _
             "FROM TBParametros_venda " & _
             "WHERE TBParametros_venda.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
             
    Movimentacoes.Select_geral strSQL, "BDRetaguarda", rstVerifica_parametro, "Otica", Me
    
    If rstVerifica_parametro.BOF <> True And rstVerifica_parametro.EOF <> True Then
        'Parâmetros que a aplicação aramzena para uso futuro em situações particulares,tais como se vai ou não permitir
        'que um vendedor venda para cliente que não seja ligado a ele e se o vendedor pode vender um item que não seja de sua linha.
        booLibera_produto_linha_pedido = rstVerifica_parametro!DFLibera_produto_linha_venda_pedido_TBParametro_venda
        booLibera_vendedor_cliente_linha_pedido = rstVerifica_parametro!DFLibera_vendedor_cliente_pedido_TBParametro_venda
        'booLibera_prev_banco = rstVerifica_parametro!DFEmite_banco_TBParametros_venda
    Else
       MsgBox "Não existe tabela válida cadastrada no parâmetro de venda!Verifique antes de lançar um pedido!", vbInformation, "Only Tech"
       sstCupon.Tab = 1
       Set rstVerifica_parametro = Nothing
       Exit Function
    End If
    
    intNumero_Itens_NF = rstVerifica_parametro!DFNumero_itens_nota_TBParametros_venda
    'intNumero_Itens_orca = rstVerifica_parametro!DFNumero_itens_orcamento_TBParametros_venda
    
    Set rstVerifica_parametro = Nothing

    Exit Function
    
Erro:
    
    MsgBox "Ocorreu um erro n°: " & Err.Number & " -- Verificacao_Parametro_Venda -- " & Err.Description & "Última Query - " & strSQL, vbInformation, "Only Tech"
    Exit Function
    
End Function

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKId_TBCupom", CStr(dblIdCupom), "DFIntegrado_filiais_TBCupom", "TBCupom", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBCupom", Me.Top, Me.Left, Me.width, Me.Height, "Cupom Fiscal")
    
End Function
