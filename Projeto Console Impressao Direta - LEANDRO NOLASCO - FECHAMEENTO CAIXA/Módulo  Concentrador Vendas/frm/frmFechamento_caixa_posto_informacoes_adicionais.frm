VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFechamento_caixa_posto_informacoes_adicionais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informações Adicionais de Venda"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFechamento_caixa_posto_informacoes_adicionais.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   9000
   Begin VB.TextBox txtData 
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
      Left            =   7500
      MaxLength       =   20
      TabIndex        =   1
      ToolTipText     =   "Código da Finalizadora"
      Top             =   225
      Width           =   1395
   End
   Begin VB.TextBox txtDescricao_Operador 
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
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   4
      ToolTipText     =   "Código da Finalizadora"
      Top             =   225
      Width           =   5925
   End
   Begin VB.TextBox txtOperador 
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
      Left            =   90
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Código da Finalizadora"
      Top             =   225
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgItens 
      Height          =   3495
      Left            =   90
      TabIndex        =   2
      Top             =   615
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   6165
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
   Begin VB.Label label15 
      AutoSize        =   -1  'True
      Caption         =   "Data"
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
      Left            =   7500
      TabIndex        =   5
      Top             =   30
      Width           =   345
   End
   Begin VB.Label label15 
      AutoSize        =   -1  'True
      Caption         =   "Operador"
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
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   690
   End
End
Attribute VB_Name = "frmFechamento_caixa_posto_informacoes_adicionais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador de Vendas                                                       '
' Objetivo...............: Informaçoes Adicionais de Venda                                                '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Peixoto                                                  '
' Data de Criação........: 04/03/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim rstAplicacao As New ADODB.Recordset

Private Sub Form_Load()

    'Abastecendo os itens
    strSql = "SELECT IXCodigo_TBProduto," & _
             "DFDescricao_TBProduto,SUM(DFQuantidade_TBItens_cupom)," & _
             "DFPreco_praticado_TBItens_cupom,SUM(DFValor_total_item_TBItens_cupom) " & _
             "FROM TBItens_cupom " & _
             "INNER JOIN TBCupom ON TBItens_cupom.FKId_TBCupom = TBCupom.PKId_TBCupom " & _
             "INNER JOIN TBProduto ON TBItens_cupom.DFCodigo_TBProduto = TBProduto.IXCodigo_TBProduto " & _
             "WHERE FKCodigo_TBOperadores_ecf = " & frmFechamento_caixa_posto.txtOperador.Text & " " & _
             "AND DFData_Saida_TBCupom = '" & Format(frmFechamento_caixa_posto.dtpFechamento.Value, "YYYYMMDD") & "' " & _
             "AND TBCupom.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
             "AND TBProduto.IXCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
             "AND TBCupom.DFCancelado_TBCupom = 0 " & _
             "AND FKCodigo_TBSecao = " & frmFechamento_caixa_posto.hfgSecao.Text & " " & _
             "GROUP BY IXCodigo_TBProduto,DFDescricao_TBProduto,DFPreco_praticado_TBItens_cupom "

     Movimentacoes.Movimenta_HFlex_Grid strSql, hfgItens, "750,4000,1300,750,1250", "Código,Descrição,Quantidade,Preço,Total", "BDRetaguarda", "Otica", Me, "N"
     
     hfgItens.Col = 0
     hfgItens.Row = 1
     If hfgItens.Text = Empty Then
        hfgItens.Rows = 2
        Movimentacoes.Monta_HFlex_Grid hfgItens, "750,4000,1300,1000,1300", "Código,Descrição,Quantidade,Preço,Total", 5, "OTICA", Me
     End If

    hfgItens.Col = 0
    hfgItens.Row = 1

    txtOperador.Text = frmFechamento_caixa_posto.txtOperador.Text
    txtDescricao_Operador.Text = frmFechamento_caixa_posto.dtcOperador.Text
    txtData.Text = frmFechamento_caixa_posto.dtpFechamento.Value
End Sub
