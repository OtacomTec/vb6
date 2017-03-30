VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatorio_Fechamento_Diario_Vendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fechamento Diário de Vendas"
   ClientHeight    =   8190
   ClientLeft      =   1800
   ClientTop       =   1845
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelatorio_Fechamento_Diario_Vendas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   7560
   Begin VB.Frame Frame3 
      Caption         =   "Canceladas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   90
      TabIndex        =   15
      Top             =   5700
      Width           =   3675
      Begin VB.OptionButton optSim_Canceladas 
         Caption         =   "Sim"
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   330
         Width           =   735
      End
      Begin VB.OptionButton optNao_Canceladas 
         Caption         =   "Não"
         Height          =   240
         Left            =   1530
         TabIndex        =   26
         Top             =   330
         Width           =   735
      End
      Begin VB.OptionButton optTodos_Canceladas 
         Caption         =   "Todos"
         Height          =   240
         Left            =   2700
         TabIndex        =   27
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Cancel          =   -1  'True
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5010
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7680
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7680
      Width           =   1185
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   90
      TabIndex        =   14
      ToolTipText     =   "Tipo do Relatório"
      Top             =   4980
      Width           =   7365
      Begin VB.OptionButton optGrafico 
         Caption         =   "Gráfico"
         Height          =   240
         Left            =   6210
         TabIndex        =   24
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optSintetico 
         Caption         =   "Sintético"
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   330
         Width           =   1095
      End
      Begin VB.OptionButton optAnalitico 
         Caption         =   "Analítico"
         Height          =   240
         Left            =   3240
         TabIndex        =   23
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Frame freOrdenar 
      Caption         =   "Classificar"
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
      Left            =   90
      TabIndex        =   17
      ToolTipText     =   "Classificação/Agrupamento do Relatório"
      Top             =   6420
      Width           =   7365
      Begin VB.OptionButton optFinalizadora 
         Caption         =   "Finalizadora"
         Height          =   240
         Left            =   5880
         TabIndex        =   36
         Top             =   330
         Width           =   1365
      End
      Begin VB.OptionButton optPDV 
         Caption         =   "PDV"
         Height          =   240
         Left            =   4380
         TabIndex        =   33
         Top             =   660
         Width           =   675
      End
      Begin VB.OptionButton optFaixa_Horaria 
         Caption         =   "Faixa Hora"
         Height          =   240
         Left            =   4380
         TabIndex        =   30
         Top             =   330
         Width           =   1245
      End
      Begin VB.OptionButton optProduto 
         Caption         =   "Produto"
         Height          =   240
         Left            =   2970
         TabIndex        =   34
         Top             =   330
         Width           =   975
      End
      Begin VB.OptionButton optCupom 
         Caption         =   "Cupom"
         Enabled         =   0   'False
         Height          =   240
         Left            =   2970
         TabIndex        =   35
         Top             =   660
         Width           =   915
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Cliente"
         Height          =   240
         Left            =   1650
         TabIndex        =   32
         Top             =   660
         Width           =   885
      End
      Begin VB.OptionButton optVendedor 
         Caption         =   "Vendedor"
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   330
         Width           =   1125
      End
      Begin VB.OptionButton optData 
         Caption         =   "Data"
         Height          =   240
         Left            =   1650
         TabIndex        =   31
         Top             =   330
         Width           =   735
      End
      Begin VB.OptionButton optSecao 
         Caption         =   "Seção"
         Enabled         =   0   'False
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   660
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Previsão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3780
      TabIndex        =   16
      Top             =   5700
      Width           =   3675
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   2700
         TabIndex        =   39
         Top             =   330
         Width           =   855
      End
      Begin VB.OptionButton optNao 
         Caption         =   "Não"
         Height          =   240
         Left            =   1530
         TabIndex        =   38
         Top             =   330
         Width           =   735
      End
      Begin VB.OptionButton optSim 
         Caption         =   "Sim"
         Height          =   240
         Left            =   120
         TabIndex        =   37
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   90
      TabIndex        =   40
      Top             =   660
      Width           =   7365
      Begin VB.TextBox txtOperador 
         Height          =   360
         Left            =   1560
         TabIndex        =   12
         ToolTipText     =   "Código do Operador"
         Top             =   3750
         Width           =   1395
      End
      Begin VB.TextBox txtPdv 
         Height          =   360
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Código do Ponto de Venda"
         Top             =   3750
         Width           =   1395
      End
      Begin VB.TextBox txtFinalizadora 
         Height          =   360
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Código da Finalizadora"
         Top             =   3105
         Width           =   1395
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Código do Cliente"
         Top             =   2460
         Width           =   1395
      End
      Begin VB.TextBox txtProduto 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Código do Produto"
         Top             =   1830
         Width           =   1395
      End
      Begin VB.TextBox txtVendedor 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Código do Vendedor"
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox txtSecao 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Código da Seção"
         Top             =   1200
         Width           =   1395
      End
      Begin MSDataListLib.DataCombo dtcSecao 
         Height          =   360
         Left            =   1560
         TabIndex        =   4
         Top             =   1200
         Width           =   5655
         _ExtentX        =   9975
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
      Begin MSDataListLib.DataCombo dtcVendedor 
         Height          =   360
         Left            =   1560
         TabIndex        =   2
         Top             =   570
         Width           =   5655
         _ExtentX        =   9975
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
      Begin MSDataListLib.DataCombo dtcProduto 
         Height          =   360
         Left            =   1560
         TabIndex        =   6
         Top             =   1830
         Width           =   5655
         _ExtentX        =   9975
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
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   360
         Left            =   1560
         TabIndex        =   8
         Top             =   2460
         Width           =   5655
         _ExtentX        =   9975
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
      Begin MSDataListLib.DataCombo dtcFinalizadora 
         Height          =   360
         Left            =   1560
         TabIndex        =   10
         Top             =   3105
         Width           =   5655
         _ExtentX        =   9975
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
      Begin MSDataListLib.DataCombo dtcOperador 
         Height          =   360
         Left            =   3000
         TabIndex        =   13
         Top             =   3750
         Width           =   4215
         _ExtentX        =   7435
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Operador"
         Height          =   240
         Left            =   1560
         TabIndex        =   50
         Top             =   3510
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PDV"
         Height          =   240
         Left            =   120
         TabIndex        =   49
         Top             =   3510
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Finalizadora"
         Height          =   240
         Left            =   120
         TabIndex        =   48
         Top             =   2865
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   240
         Left            =   120
         TabIndex        =   47
         Top             =   2220
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         Height          =   240
         Left            =   120
         TabIndex        =   43
         Top             =   1590
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Seção"
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor"
         Height          =   240
         Left            =   120
         TabIndex        =   41
         Top             =   330
         Width           =   825
      End
   End
   Begin MSComCtl2.DTPicker dtpInicial 
      Height          =   360
      Left            =   90
      TabIndex        =   18
      Top             =   7710
      Width           =   1515
      _ExtentX        =   2672
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
      Format          =   54525953
      CurrentDate     =   37881
   End
   Begin MSComCtl2.DTPicker dtpFinal 
      Height          =   360
      Left            =   2040
      TabIndex        =   19
      Top             =   7710
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   54525953
      CurrentDate     =   37881
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "até"
      Height          =   240
      Left            =   1665
      TabIndex        =   46
      Top             =   7830
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   240
      Left            =   90
      TabIndex        =   45
      Top             =   7470
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      Height          =   240
      Left            =   90
      TabIndex        =   44
      Top             =   30
      Width           =   1290
   End
End
Attribute VB_Name = "frmRelatorio_Fechamento_Diario_Vendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Estatística de Resumo Diário de Vendas                         '
' Data de Criação........: 22/06/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public strSql As String
Dim log As New DLLSystemManager.log
'Declaração da variavel do intercomunicador de mensagens
Private Cliente_mensagem_exe As VetorDeMensagens.ClienteDeMensagens

Private Sub dtcCliente_GotFocus()
    If txtCliente.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCliente.Text)
    End If
End Sub

Private Sub dtcCliente_LostFocus()
    txtCliente.Text = dtcCliente.BoundText
    If IsNumeric(txtCliente.Text) = False Or dtcCliente.Text = Empty Then txtCliente.Text = Empty: Exit Sub
End Sub

Private Sub dtcEmpresa_Change()
    txtProduto.Text = Empty: txtOperador.Text = Empty: txtCliente.Text = Empty: txtVendedor.Text = Empty
End Sub

Private Sub dtcEmpresa_LostFocus()
    If Not IsNumeric(dtcEmpresa.BoundText) Then dtcEmpresa.Text = Empty
    If IsNumeric(dtcEmpresa.Text) Then dtcEmpresa.Text = Empty

    If dtcEmpresa.Text <> Empty Then
       strSql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me
       
       strSql = "SELECT PKCodigo_TBOperadores_ecf, DFNome_TBOperadores_ecf FROM TBOperadores_ecf WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
       Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSql, "BDRetaguarda", "Otica", Me
       
       strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
       
       strSql = "SELECT IXCodigo_TBVendedor,DFNome_TBVendedor FROM TBVendedor WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBVendedor", "DFNome_TBVendedor", dtcVendedor, strSql, "BDRetaguarda", "Otica", Me
    Else
       strSql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto"
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me
       
       strSql = "SELECT PKCodigo_TBOperadores_ecf, DFNome_TBOperadores_ecf FROM TBOperadores_ecf"
       Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSql, "BDRetaguarda", "Otica", Me
       
       strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente"
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
       
       strSql = "SELECT IXCodigo_TBVendedor,DFNome_TBVendedor FROM TBVendedor"
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBVendedor", "DFNome_TBVendedor", dtcVendedor, strSql, "BDRetaguarda", "Otica", Me
    End If
    
    dtcEmpresa.Enabled = False: txtVendedor.SetFocus
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
    If IsNumeric(txtOperador.Text) = False Or dtcOperador.Text = Empty Then txtOperador.Text = Empty: Exit Sub
End Sub

Private Sub dtpFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

Private Sub dtpInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = "113" Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
End Sub

Private Sub optAnalitico_Click()
    If optAnalitico.Value = True Then
       freOrdenar.Enabled = True
    End If
    optProduto.Value = True
    optFaixa_Horaria.Enabled = False
    optData.Enabled = False
    optCliente.Enabled = False
    optPDV.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    Call Cancelar
End Sub

Private Sub cmdImprimir_Click()
    If dtpInicial.Value > dtpFinal.Value Then
       MsgBox "Data Final menor que Data Inicial. Verifique!", vbInformation, "Only Tech"
       Exit Sub
    End If

    frmAguarde.Show
    DoEvents
    Call Impressao
    Unload frmAguarde
End Sub

Private Sub dtcProduto_GotFocus()
    If txtProduto.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcProduto.Text)
    End If
End Sub

Private Sub dtcProduto_LostFocus()
    txtProduto.Text = dtcProduto.BoundText
    If IsNumeric(txtProduto.Text) = False Or dtcProduto.Text = Empty Then txtProduto.Text = Empty: Exit Sub
End Sub

Private Sub dtcSecao_GotFocus()
    If txtSecao.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcSecao.Text)
    End If
End Sub

Private Sub dtcSecao_LostFocus()
    txtSecao.Text = dtcSecao.BoundText
    If IsNumeric(txtSecao.Text) = False Or dtcSecao.Text = Empty Then txtSecao.Text = Empty: Exit Sub
End Sub

Private Sub dtcVendedor_GotFocus()
    If txtVendedor.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcVendedor.Text)
    End If
End Sub

Private Sub dtcVendedor_LostFocus()
    txtVendedor.Text = dtcVendedor.BoundText
    If IsNumeric(txtVendedor.Text) = False Or dtcVendedor.Text = Empty Then txtVendedor.Text = Empty: Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
  
Private Sub Form_Load()
    On Error GoTo erro
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Relatorio de Fechamento Diário de Vendas"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
        Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Relatorio de Fechamento Diário de Vendas"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    'Montando os datacombo de tela
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me

    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    strSql = "SELECT IXCodigo_TBProduto,DFDescricao_TBProduto FROM TBProduto WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBProduto", "DFDescricao_TBProduto", dtcProduto, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
             
    strSql = "SELECT PKCodigo_TBSecao,DFDescricao_TBsecao FROM TBSecao "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBSecao", "DFDescricao_TBsecao", dtcSecao, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_TBVendedor,DFNome_TBVendedor FROM TBVendedor WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBVendedor", "DFNome_TBVendedor", dtcVendedor, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_TBFinalizadora, DFDescricao_TBFinalizadora FROM TBFinalizadora "
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBOperadores_ecf, DFNome_TBOperadores_ecf FROM TBOperadores_ecf WHERE FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSql, "BDRetaguarda", "Otica", Me
                   
    optNao_Canceladas.Value = True
    optVendedor.Value = True
    optTodos.Value = True
    
    optSintetico.Value = 1
    dtpInicial.Value = Date
    dtpFinal.Value = Date + 7
                        
    Exit Sub
erro:

    Call erro.erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Saindo de Relatorio de Fechamento Diário de Vendas"
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Exit Sub
erro:
    Call erro.erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Function Cancelar()
    On Error GoTo erro
    
    Call Objetos.Limpa_TXT(Me)
            
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Relatorio de Fechamento Diário de Vendas"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    dtpInicial.Value = Date
    dtpInicial.Day = 1
    dtpFinal = Date
    txtVendedor.SetFocus
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Function Impressao()
    
    Dim strID_Finalizadora As String
    
    If optProduto.Value = True And optGrafico.Value = False Then
       
       strSql = "SELECT TBCupom.PKCodigo_TBPdv AS PDV,CONVERT (INT, TBCupom.DFNumero_TBCupom) AS CUPOM," & _
                "TBCupom.DFSerie_TBCupom AS SERIE,TBCupom.DFTotal_cupom_TBCupom AS TOTAL_CUPOM," & _
                "TBCupom.DFData_Saida_TBCupom AS DATA,TBCupom.DFHora_Saida AS HORA,TBCupom.DFPrevisao_TBCupom AS PREVISAO," & _
                "TBCupom.FKCodigo_TBOperadores_ecf AS COD_OPERADOR,TBOperadores_ecf.DFNome_TBOperadores_ecf AS NOME_OPERADOR," & _
                "TBCupom.FKId_TBVendedor AS ID_VENDEDOR,TBVendedor.DFNome_TBVendedor AS VENDEDOR,TBVendedor.IXCodigo_TBVendedor AS COD_VENDEDOR," & _
                "TBCupom.DFEmitente_TBCupom AS COD_CLIENTE,TBCliente.DFNome_TBCliente AS CLIENTE,TBItens_cupom.DFCodigo_TBProduto," & _
                "TBProduto.DFDescricao_TBProduto,TBItens_cupom.DFQuantidade_TBItens_cupom," & _
                "TBItens_cupom.DFValor_total_item_TBItens_cupom,IXCodigo_TBFinalizadora,DFDescricao_TBFinalizadora," & _
                "DFPreco_praticado_TBItens_cupom,DFValor_total_praticado_TBItens_cupom," & _
                "DFCodigo_cupom_impressora_TBCupom " & _
                "FROM TBCupom " & _
                "INNER JOIN TBVendedor ON TBCupom.FKId_TBVendedor = TBVendedor.PKId_TBVendedor " & _
                "INNER JOIN TBOperadores_ecf ON TBCupom.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
                "INNER JOIN TBCliente ON TBCupom.DFEmitente_TBCupom = TBCliente.IXCodigo_TBCliente " & _
                "INNER JOIN TBItens_cupom ON TBCupom.PKId_TBCupom = TBItens_cupom.FKId_TBCupom " & _
                "INNER JOIN TBProduto ON TBItens_cupom.DFCodigo_TBProduto = TBProduto.IXCodigo_TBProduto " & _
                "LEFT JOIN TBOperacao_caixa ON TBCupom.PKId_TBCupom = TBOperacao_caixa.DFNumero_Cupom_TBOperacao_caixa " & _
                "LEFT JOIN TBFinalizadora ON TBOperacao_caixa.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                "WHERE (DFData_Saida_TBCupom BETWEEN '" & Format(dtpInicial.Value, "YYYYMMDD") & "'  " & _
                "AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "')  " & _
                "AND TBProduto.IXCodigo_TBEmpresa = TBCupom.FKCodigo_TBEmpresa "
                
    ElseIf optAnalitico.Value = True Then
       
       strSql = "SELECT TBCupom.PKCodigo_TBPdv AS PDV,CONVERT (INT, TBCupom.DFNumero_TBCupom) AS CUPOM," & _
                "TBCupom.DFSerie_TBCupom AS SERIE,TBCupom.DFTotal_cupom_TBCupom AS TOTAL_CUPOM," & _
                "TBCupom.DFData_Saida_TBCupom AS DATA,TBCupom.DFHora_Saida AS HORA,TBCupom.DFPrevisao_TBCupom AS PREVISAO," & _
                "TBCupom.FKCodigo_TBOperadores_ecf AS COD_OPERADOR,TBOperadores_ecf.DFNome_TBOperadores_ecf AS NOME_OPERADOR," & _
                "TBCupom.FKId_TBVendedor AS ID_VENDEDOR,TBVendedor.DFNome_TBVendedor AS VENDEDOR,TBVendedor.IXCodigo_TBVendedor AS COD_VENDEDOR," & _
                "TBCupom.DFEmitente_TBCupom AS COD_CLIENTE,TBCliente.DFNome_TBCliente AS CLIENTE,TBItens_cupom.DFCodigo_TBProduto," & _
                "TBProduto.DFDescricao_TBProduto,TBItens_cupom.DFValor_total_item_TBItens_cupom," & _
                "IXCodigo_TBFinalizadora,DFDescricao_TBFinalizadora,DFQuantidade_TBItens_cupom,DFPreco_praticado_TBItens_cupom," & _
                "DFCodigo_cupom_impressora_TBCupom " & _
                "FROM TBCupom " & _
                "INNER JOIN TBVendedor ON TBCupom.FKId_TBVendedor = TBVendedor.PKId_TBVendedor " & _
                "INNER JOIN TBOperadores_ecf ON TBCupom.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
                "INNER JOIN TBCliente ON TBCupom.DFEmitente_TBCupom = TBCliente.IXCodigo_TBCliente " & _
                "INNER JOIN TBItens_cupom ON TBCupom.PKId_TBCupom = TBItens_cupom.FKId_TBCupom " & _
                "INNER JOIN TBProduto ON TBItens_cupom.DFCodigo_TBProduto = TBProduto.IXCodigo_TBProduto " & _
                "LEFT JOIN TBOperacao_caixa ON TBCupom.PKId_TBCupom = TBOperacao_caixa.DFNumero_Cupom_TBOperacao_caixa " & _
                "LEFT JOIN TBFinalizadora ON TBOperacao_caixa.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                "WHERE (DFData_Saida_TBCupom BETWEEN '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                "AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "') " & _
                "AND TBProduto.IXCodigo_TBEmpresa = TBCupom.FKCodigo_TBEmpresa "
                
    ElseIf optSintetico.Value = True And optVendedor.Value = True Then
    
        strSql = "SELECT TBCupom.PKCodigo_TBPdv AS PDV,CONVERT (INT, TBCupom.DFNumero_TBCupom) AS CUPOM," & _
                "TBCupom.DFSerie_TBCupom AS SERIE,TBCupom.DFTotal_cupom_TBCupom AS TOTAL_CUPOM," & _
                "TBCupom.DFData_Saida_TBCupom AS DATA,TBCupom.DFHora_Saida AS HORA,TBCupom.DFPrevisao_TBCupom AS PREVISAO," & _
                "TBCupom.FKCodigo_TBOperadores_ecf AS COD_OPERADOR,TBOperadores_ecf.DFNome_TBOperadores_ecf AS NOME_OPERADOR," & _
                "TBCupom.FKId_TBVendedor AS ID_VENDEDOR,TBVendedor.DFNome_TBVendedor AS VENDEDOR,TBVendedor.IXCodigo_TBVendedor AS COD_VENDEDOR," & _
                "TBCupom.DFEmitente_TBCupom AS COD_CLIENTE,TBCliente.DFNome_TBCliente AS CLIENTE,TBItens_cupom.DFCodigo_TBProduto," & _
                "TBProduto.DFDescricao_TBProduto,TBItens_cupom.DFValor_total_item_TBItens_cupom," & _
                "IXCodigo_TBFinalizadora,DFDescricao_TBFinalizadora,DFQuantidade_TBItens_cupom,DFPreco_praticado_TBItens_cupom," & _
                "DFCodigo_cupom_impressora_TBCupom " & _
                "FROM TBCupom " & _
                "INNER JOIN TBVendedor ON TBCupom.FKId_TBVendedor = TBVendedor.PKId_TBVendedor " & _
                "INNER JOIN TBOperadores_ecf ON TBCupom.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
                "INNER JOIN TBCliente ON TBCupom.DFEmitente_TBCupom = TBCliente.IXCodigo_TBCliente " & _
                "INNER JOIN TBItens_cupom ON TBCupom.PKId_TBCupom = TBItens_cupom.FKId_TBCupom " & _
                "INNER JOIN TBProduto ON TBItens_cupom.DFCodigo_TBProduto = TBProduto.IXCodigo_TBProduto " & _
                "LEFT JOIN TBOperacao_caixa ON TBCupom.PKId_TBCupom = TBOperacao_caixa.DFNumero_Cupom_TBOperacao_caixa " & _
                "LEFT JOIN TBFinalizadora ON TBOperacao_caixa.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                "WHERE (DFData_Saida_TBCupom BETWEEN '" & Format(dtpInicial.Value, "YYYYMMDD") & "' " & _
                "AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "') " & _
                "AND TBProduto.IXCodigo_TBEmpresa = TBCupom.FKCodigo_TBEmpresa "

    Else
       
       strSql = "SELECT TBCupom.PKCodigo_TBPdv AS PDV,CONVERT (INT, TBCupom.DFNumero_TBCupom) AS CUPOM, " & _
                "TBCupom.DFSerie_TBCupom AS SERIE, TBCupom.DFTotal_cupom_TBCupom AS TOTAL_CUPOM, " & _
                "TBCupom.DFData_Saida_TBCupom AS DATA, TBCupom.DFHora_Saida AS HORA, " & _
                "TBCupom.DFPrevisao_TBCupom AS PREVISAO, TBCupom.FKCodigo_TBOperadores_ecf AS COD_OPERADOR, " & _
                "TBOperadores_ecf.DFNome_TBOperadores_ecf AS NOME_OPERADOR, TBCupom.FKId_TBVendedor AS ID_VENDEDOR, " & _
                "TBVendedor.DFNome_TBVendedor AS VENDEDOR, TBVendedor.IXCodigo_TBVendedor AS COD_VENDEDOR, " & _
                "TBCupom.DFEmitente_TBCupom AS COD_CLIENTE, TBCliente.DFNome_TBCliente AS CLIENTE," & _
                "IXCodigo_TBFinalizadora,DFDescricao_TBFinalizadora,TBItens_cupom.DFCodigo_TBProduto," & _
                "TBProduto.DFDescricao_TBProduto,DFQuantidade_TBItens_cupom,DFValor_total_item_TBItens_cupom,DFPreco_praticado_TBItens_cupom," & _
                "DFCodigo_cupom_impressora_TBCupom " & _
                "FROM TBCupom " & _
                "INNER JOIN TBVendedor ON TBCupom.FKId_TBVendedor = TBVendedor.pKId_TBVendedor " & _
                "INNER JOIN TBOperadores_ecf ON TBCupom.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
                "INNER JOIN TBCliente ON TBCupom.DFEmitente_TBCupom = TBCliente.IXCodigo_TBCliente " & _
                "INNER JOIN TBItens_cupom ON TBCupom.PKId_TBCupom = TBItens_cupom.FKId_TBCupom " & _
                "INNER JOIN TBProduto ON TBItens_cupom.DFCodigo_TBProduto = TBProduto.IXCodigo_TBProduto " & _
                "LEFT JOIN TBOperacao_caixa ON TBCupom.PKId_TBCupom = TBOperacao_caixa.DFNumero_Cupom_TBOperacao_caixa " & _
                "LEFT JOIN TBFinalizadora ON TBOperacao_caixa.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                "WHERE (DFData_Saida_TBCupom BETWEEN '" & Format(Me.dtpInicial.Value, "YYYYMMDD") & "' " & _
                "AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "' ) " & _
                "AND TBProduto.IXCodigo_TBEmpresa = TBCupom.FKCodigo_TBEmpresa "
    End If
   
    If dtcEmpresa.BoundText <> Empty Then
       strSql = strSql & " AND TBCupom.FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
                         " AND TBProduto.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
                         " AND TBCliente.IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
                         " AND TBOperadores_ecf.FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " "
    End If
   
    ' Notas canceladas
    If optSim_Canceladas.Value = True Then
       strSql = strSql & " AND DFCancelado_TBCupom = 1 "
    End If
    
    If Me.optNao_Canceladas.Value = True Then
       strSql = strSql & " AND DFCancelado_TBCupom = 0 "
    End If
    
    ' Previsao
    If optSim.Value = True Then
       strSql = strSql & " AND TBCupom.DFPrevisao_TBCupom = 1 "
    ElseIf optNao.Value = True Then
       strSql = strSql & " AND TBCupom.DFPrevisao_TBCupom = 0 "
    End If
    
    ' Vendedor
    If dtcVendedor.BoundText <> Empty Then
       strSql = strSql & " AND TBVendedor.IXCodigo_TBVendedor = " & dtcVendedor.BoundText & ""
    End If
    
    ' Secao
    If dtcSecao.BoundText <> "" Then
       strSql = strSql & " AND TBProduto.FKCodigo_TBSecao = " & dtcSecao.BoundText & ""
    End If
    
    ' Produto
    If dtcProduto.BoundText <> "" Then
       strSql = strSql & " AND TBItens_cupom.DFCodigo_TBProduto = " & dtcProduto.BoundText & ""
    End If
    
    ' Cliente
    If dtcCliente.BoundText <> "" Then
       strSql = strSql & " AND TBCupom.DFEmitente_TBCupom = " & dtcCliente.BoundText & ""
    End If

    ' Finalizadora
    If dtcFinalizadora.Text <> "" Then
       strSql = strSql & " AND TBFinalizadora.IXCodigo_TBFinalizadora = " & dtcFinalizadora.BoundText & " "
    End If
    
    ' PDV
    If txtPdv.Text <> "" Then
       strSql = strSql & " AND TBCupom.PKCodigo_TBPdv = " & txtPdv.Text & ""
    End If
    
    ' Operador
    If dtcOperador.BoundText <> "" Then
       strSql = strSql & " AND TBCupom.FKCodigo_TBOperadores_ecf = " & dtcOperador.BoundText & ""
    End If
    
    'If optSintetico.Value = True And optVendedor.Value = True Then
    '   strsql = strsql & "GROUP BY TBCupom.FKId_TBVendedor," & _
    '                     "TBVendedor.DFNome_TBVendedor," & _
    '                     "TBVendedor.IXCodigo_TBVendedor "
    'End If
    
    ' Ordenacao do relatorio
    If optVendedor.Value = True Then
        strSql = strSql & " ORDER BY TBVendedor.DFNome_TBVendedor "
    ElseIf optFaixa_Horaria.Value = True Then
        strSql = strSql & " ORDER BY TBCupom.DFHora_Saida "
    ElseIf optProduto.Value = True Then
        strSql = strSql & " ORDER BY TBProduto.IXCodigo_TBProduto "
    ElseIf optData.Value = True Then
        strSql = strSql & " ORDER BY TBCupom.DFData_Saida_TBCupom "
    ElseIf optCliente.Value = True Then
        strSql = strSql & " ORDER BY TBCliente.DFNome_TBCliente "
    ElseIf optPDV.Value = True Then
        strSql = strSql & " ORDER BY TBCupom.PKCodigo_TBPdv "
    ElseIf optCupom.Value = True Then
        strSql = strSql & " ORDER BY TBCupom.DFNumero_TBCupom "
    ElseIf optFinalizadora.Value = True Then
        strSql = strSql & " ORDER BY TBFinalizadora.IXCodigo_TBFinalizadora "
    End If
    
    Call frmConsole_Relatorio_Fechamento_Diario_Vendas.Show
    
End Function

Private Sub optGrafico_Click()
    optFaixa_Horaria.Enabled = True
    optData.Enabled = True
    optCliente.Enabled = True
    optPDV.Enabled = True
    optFinalizadora.Enabled = False
    optPDV.Enabled = False
End Sub

Private Sub optSintetico_Click()
    optFaixa_Horaria.Enabled = True
    optData.Enabled = True
    optCliente.Enabled = True
    optPDV.Enabled = True
    optFinalizadora.Enabled = True
End Sub

Private Sub txtCliente_Change()
    dtcCliente.BoundText = txtCliente.Text
    If IsNumeric(txtCliente.Text) = False Then txtCliente.Text = Empty: Exit Sub
End Sub

Private Sub txtCliente_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFinalizadora_Change()
    dtcFinalizadora.BoundText = txtFinalizadora.Text
    If IsNumeric(txtFinalizadora.Text) = False Then txtFinalizadora.Text = Empty: Exit Sub
End Sub

Private Sub txtFinalizadora_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_Change()
    dtcOperador.BoundText = txtOperador.Text
    If IsNumeric(txtOperador.Text) = False Then txtOperador.Text = Empty: Exit Sub
End Sub

Private Sub txtOperador_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_LostFocus()
    dtcOperador.BoundText = txtOperador.Text
End Sub

Private Sub txtPdv_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtProduto_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtSecao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtSecao_Change()
    dtcSecao.BoundText = txtSecao.Text
    If IsNumeric(txtSecao.Text) = False Then txtSecao.Text = Empty: Exit Sub
End Sub

Private Sub txtVendedor_Change()
    dtcVendedor.BoundText = txtVendedor.Text
    If IsNumeric(txtVendedor.Text) = False Then txtVendedor.Text = Empty: Exit Sub
End Sub

Private Sub txtProduto_Change()
    dtcProduto.BoundText = txtProduto.Text
    If IsNumeric(txtProduto.Text) = False Then txtProduto.Text = Empty: Exit Sub
End Sub

Private Sub txtVendedor_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub
