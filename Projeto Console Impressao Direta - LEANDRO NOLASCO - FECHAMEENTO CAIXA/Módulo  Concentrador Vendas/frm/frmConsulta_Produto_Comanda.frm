VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConsulta_Produto_Comanda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Produtos"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmConsulta_Produto_Comanda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8070
   Begin VB.CommandButton cmdConsulta 
      Height          =   360
      Left            =   7650
      Picture         =   "frmConsulta_Produto_Comanda.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Consultar"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtConsulta 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      MaxLength       =   50
      TabIndex        =   0
      Top             =   450
      Width           =   7515
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgProduto 
      Height          =   4305
      Left            =   60
      TabIndex        =   2
      Top             =   900
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   7594
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   210
      Width           =   660
   End
End
Attribute VB_Name = "frmConsulta_Produto_Comanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Consulta de produtos                                           '
' Data de Criação........: 04/02/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log

Private Sub cmdConsulta_Click()

     If Me.txtConsulta.Text = "" Then
        MsgBox "Favor selecione o produto", vbInformation, "Only Tech"
        Exit Sub
     End If
     
     Dim rstProduto As New ADODB.Recordset
        
     frmAguarde.Show
     DoEvents
     
     strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBITENS_TABELA_PRECO.DFPreco_varejo_TBItens_tabela_preco " & _
              "FROM TBPRODUTO " & _
              "INNER JOIN TBITENS_TABELA_PRECO " & _
              "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
              "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA) " & _
              "AND convert(nvarchar,TBProduto.DFDescricao_TBProduto) LIKE '%" & txtConsulta.Text & "%' "
     
     Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1500", "Código,Descrição,Preço($)", "BDRetaguarda", "Otica", Me
     
     Unload frmAguarde
      
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub hfgProduto_Click()

    If hfgProduto.Col = 0 Then
       frmComanda.txtCodigo_Produto.Text = hfgProduto.TextArray((hfgProduto.Row * hfgProduto.Cols + hfgProduto.Col + 1))
       frmComanda.txtDescricao_Produto.Text = hfgProduto.TextArray((hfgProduto.Row * hfgProduto.Cols + hfgProduto.Col + 2))
       Unload Me
       frmComanda.txtDescricao_Produto.SetFocus
    End If
   
End Sub

Private Sub hfgProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgProduto_Click
    End If
End Sub
