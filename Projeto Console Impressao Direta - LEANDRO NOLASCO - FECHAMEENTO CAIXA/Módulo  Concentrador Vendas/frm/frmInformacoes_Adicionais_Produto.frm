VERSION 5.00
Object = "{40BD39E3-6F1E-11D1-B2DF-444553540000}#1.0#0"; "OCXShape.ocx"
Begin VB.Form frmInformacoes_Adicionais_Produto 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmInformacoes_Adicionais_Produto.frx":0000
   ScaleHeight     =   3900
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin FormShape.FormShape FormShape1 
      Left            =   5430
      Top             =   3450
      ShapeType       =   1
      MaskColor       =   16777215
      AutoScale       =   -1  'True
      ScaleX          =   1
      ScaleY          =   1
      ShapeString     =   ""
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque Rua"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   30
      Top             =   2820
      Width           =   1455
   End
   Begin VB.Label lblEstoque_Rua 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   29
      Top             =   3000
      Width           =   1485
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque Prédio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2430
      TabIndex        =   28
      Top             =   2820
      Width           =   1455
   End
   Begin VB.Label lblEstoque_Predio 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2430
      TabIndex        =   27
      Top             =   3000
      Width           =   1485
   End
   Begin VB.Label lblEstoque_Apt 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   3000
      Width           =   1485
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque Apt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   25
      Top             =   2820
      Width           =   1455
   End
   Begin VB.Label lblPreco3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2430
      TabIndex        =   24
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label lblCabecalho_preco3 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2430
      TabIndex        =   23
      Top             =   3270
      Width           =   1155
   End
   Begin VB.Label lblPreco2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1260
      TabIndex        =   22
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label lblCabecalho_preco2 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1260
      TabIndex        =   21
      Top             =   3270
      Width           =   1155
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque Atual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   20
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblEstoque_atual 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   2580
      Width           =   1485
   End
   Begin VB.Label lblDescricao 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   18
      Top             =   750
      Width           =   5745
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
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
      Left            =   180
      TabIndex        =   17
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label lblPreco5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4710
      TabIndex        =   16
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label lblCabecalho_preco5 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço5"
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
      Left            =   4710
      TabIndex        =   15
      Top             =   3270
      Width           =   1155
   End
   Begin VB.Label lblFabricante 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   2130
      Width           =   5745
   End
   Begin VB.Label lblEstoque_max 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2430
      TabIndex        =   13
      Top             =   2580
      Width           =   1485
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque Máximo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2430
      TabIndex        =   12
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblEstoque_min 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   2580
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque Mínimo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblCabecalho_preco4 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço4"
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
      TabIndex        =   9
      Top             =   3270
      Width           =   1155
   End
   Begin VB.Label lblPreco4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3540
      TabIndex        =   8
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label lblCabecalho_preco1 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   7
      Top             =   3270
      Width           =   1155
   End
   Begin VB.Label lblPreco1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricante"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Seção"
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
      Left            =   180
      TabIndex        =   4
      Top             =   1470
      Width           =   1455
   End
   Begin VB.Label lblSecao 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1680
      Width           =   5745
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Categoria"
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
      Left            =   180
      TabIndex        =   2
      Top             =   990
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Informações Adicionais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   5535
   End
   Begin VB.Label lblCategoria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   1200
      Width           =   5745
   End
End
Attribute VB_Name = "frmInformacoes_Adicionais_Produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Faturamento                                                    '
' Objetivo...............: Informações do Produto                                         '
' Data de Criação........: 26/06/2006                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........: Rodrigo Santos                                                 '
' Data última manutenção.:                                                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Info_Produto(Codigo_produto As Integer, Empresa As Integer, Aplicacao As String, Banco As String, Optional Top As Integer, Optional Left As Integer, Optional Width As Integer, Optional Height As Integer)
        
    Dim strSql As String
    Dim rstProduto As New ADODB.Recordset
    Dim rstCabecalho As New ADODB.Recordset
    Dim conexao_Integracao As New DLLConexao_Sistema.conexao
           
    'INDICANDO O BANCO A CONECTAR-SE
    conexao_Integracao.Initial_Catalog = Banco
    
    'ABRINDO CONEXAO COM BANCO
    conexao_Integracao.Abrir_conexao (Aplicacao)
    
    DoEvents
    
    rstProduto.CursorLocation = adUseClient
    
    'STRING QUE COLETA DADOS RELATIVOS AO PRODUTO
    strSql = "SELECT TBProduto.IXCodigo_TBProduto," & _
             "TBProduto.DFDescricao_TBProduto," & _
             "TBProduto.DFEstoque_minimo_TBProduto," & _
             "TBProduto.DFEstoque_maximo_TBProduto," & _
             "TBProduto.DFEstoque_atual_TBProduto," & _
             "TBCategoria.PKId_TBCategoria,TBCategoria.DFDescricao_TBCategoria," & _
             "TBCategoria.DFSigla_TBCategoria," & _
             "DFLocalizacao_estoque_rua_TBProduto," & _
             "DFLocalizacao_estoque_predio_TBProduto," & _
             "DFLocalizacao_estoque_apartamento_TBProduto," & _
             "TBFabricante.PKCodigo_TBFabricante,TBFabricante.DFNome_TBFabricante," & _
             "TBFabricante.DFSigla_TBFabricante,"
             
    strSql = strSql & "TBSecao.PKCodigo_TBSecao,TBSecao.DFDescricao_TBsecao," & _
                      "TBItens_tabela_preco.DFPreco_avista_TBItens_tabela_preco," & _
                      "TBItens_tabela_preco.DFPreco_promocao_TBItens_tabela_preco," & _
                      "TBItens_tabela_preco.DFPreco_revenda_TBItens_tabela_preco," & _
                      "TBItens_tabela_preco.DFPreco_especial_TBItens_tabela_preco," & _
                      "TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco " & _
                      "FROM TBProduto " & _
                      "LEFT JOIN TBFabricante ON " & _
                      "TBProduto.FKCodigo_TBFabricante = TBFabricante.PKCodigo_TBFabricante " & _
                      "LEFT JOIN TBSecao ON " & _
                      "TBProduto.FKCodigo_TBSecao = TBSecao.PKCodigo_TBSecao " & _
                      "LEFT JOIN TBCategoria ON " & _
                      "TBProduto.FKId_TBCategoria = TBCategoria.PKId_TBCategoria " & _
                      "LEFT JOIN TBItens_tabela_preco ON " & _
                      "TBProduto.PKId_TBProduto = TBItens_tabela_preco.FKId_TBProduto "
             
    strSql = strSql & "WHERE IXCodigo_TBEmpresa = '" & Empresa & "' AND TBProduto.IXCodigo_TBProduto = '" & Codigo_produto & "' " & _
                      "GROUP BY TBProduto.IXCodigo_TBProduto," & _
                      "TBProduto.DFDescricao_TBProduto," & _
                      "TBProduto.DFEstoque_minimo_TBProduto," & _
                      "TBProduto.DFEstoque_maximo_TBProduto," & _
                      "TBProduto.DFEstoque_atual_TBProduto," & _
                      "TBCategoria.PKId_TBCategoria,TBCategoria.DFDescricao_TBCategoria," & _
                      "TBCategoria.DFSigla_TBCategoria," & _
                      "DFLocalizacao_estoque_rua_TBProduto," & _
                      "DFLocalizacao_estoque_predio_TBProduto," & _
                      "DFLocalizacao_estoque_apartamento_TBProduto," & _
                      "TBFabricante.PKCodigo_TBFabricante,TBFabricante.DFNome_TBFabricante," & _
                      "TBFabricante.DFSigla_TBFabricante," & _
                      "TBSecao.PKCodigo_TBSecao,TBSecao.DFDescricao_TBsecao," & _
                      "TBItens_tabela_preco.DFPreco_avista_TBItens_tabela_preco," & _
                      "TBItens_tabela_preco.DFPreco_promocao_TBItens_tabela_preco," & _
                      "TBItens_tabela_preco.DFPreco_revenda_TBItens_tabela_preco," & _
                      "TBItens_tabela_preco.DFPreco_especial_TBItens_tabela_preco," & _
                      "TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco "
             
    rstProduto.Open strSql, conexao_Integracao.CNConexao, adOpenStatic, adLockReadOnly
    
    rstProduto.MoveFirst
    Me.Show
    
    'PREENCHENDO LABELS
    If rstProduto.BOF = False Then
       If IsNull(rstProduto!DFDescricao_TBProduto) = False Then
          lblDescricao.Caption = rstProduto!DFDescricao_TBProduto
       End If
       If IsNull(rstProduto!PKId_TBCategoria) = False Then
          lblCategoria.Caption = rstProduto!PKId_TBCategoria & " - " & rstProduto!DFDescricao_TBCategoria & "   " & rstProduto!DFSigla_TBCategoria
       End If
       If IsNull(rstProduto!PKCodigo_TBSecao) = False Then
          lblSecao.Caption = rstProduto!PKCodigo_TBSecao & " - " & rstProduto!DFDescricao_TBsecao
       End If
       If IsNull(rstProduto!PKCodigo_TBFabricante) = False Then
          lblFabricante.Caption = rstProduto!PKCodigo_TBFabricante & " - " & rstProduto!DFNome_TBFabricante & "   " & rstProduto!DFSigla_TBFabricante
       End If
       If IsNull(rstProduto!DFEstoque_minimo_TBProduto) = False Then
          lblEstoque_min.Caption = rstProduto!DFEstoque_minimo_TBProduto
       End If
       If IsNull(rstProduto!DFEstoque_maximo_TBProduto) = False Then
          lblEstoque_max.Caption = rstProduto!DFEstoque_maximo_TBProduto
       End If
       If IsNull(rstProduto!DFEstoque_atual_TBProduto) = False Then
          lblEstoque_atual.Caption = rstProduto!DFEstoque_atual_TBProduto
       End If
       If IsNull(rstProduto!DFPreco_avista_TBItens_tabela_preco) = False Then
          lblPreco1.Caption = rstProduto!DFPreco_avista_TBItens_tabela_preco
       End If
       If IsNull(rstProduto!DFPreco_promocao_TBItens_tabela_preco) = False Then
          lblPreco2.Caption = rstProduto!DFPreco_promocao_TBItens_tabela_preco
       End If
       If IsNull(rstProduto!DFPreco_revenda_TBItens_tabela_preco) = False Then
          lblPreco3.Caption = rstProduto!DFPreco_revenda_TBItens_tabela_preco
       End If
       If IsNull(rstProduto!DFPreco_especial_TBItens_tabela_preco) = False Then
          lblPreco4.Caption = rstProduto!DFPreco_especial_TBItens_tabela_preco
       End If
       If IsNull(rstProduto!DFPreco_varejo_TBItens_tabela_preco) = False Then
          lblPreco5.Caption = rstProduto!DFPreco_varejo_TBItens_tabela_preco
       End If
       If IsNull(rstProduto!DFLocalizacao_estoque_rua_TBProduto) = False Then
          lblEstoque_Rua.Caption = rstProduto!DFLocalizacao_estoque_rua_TBProduto
       End If
       If IsNull(rstProduto!DFLocalizacao_estoque_predio_TBProduto) = False Then
          lblEstoque_Predio.Caption = rstProduto!DFLocalizacao_estoque_predio_TBProduto
       End If
       If IsNull(rstProduto!DFLocalizacao_estoque_apartamento_TBProduto) = False Then
          lblEstoque_Apt.Caption = rstProduto!DFLocalizacao_estoque_apartamento_TBProduto
       End If
    End If
      
    Set rstProduto = Nothing
    
    conexao_Integracao.Fechar_conexao
    
    'PREENCHENDO CABEÇALHO PREÇO
    strSql = "SELECT DFNome_Preco_avista_TBTipo_preco," & _
             "DFNome_Preco_promocao_TBTipo_preco," & _
             "DFNome_Preco_revenda_TBTipo_preco," & _
             "DFNome_Preco_especial_TBTipo_preco," & _
             "DFNome_Preco_varejo_TBTipo_preco FROM TBTipo_preco "
                       
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstCabecalho, "Otica", Me)
    
    rstCabecalho.MoveFirst
    
    If rstCabecalho.EOF = False Then
       If IsNull(rstCabecalho!DFNome_Preco_avista_TBTipo_preco) = False Then
          lblCabecalho_preco1.Caption = rstCabecalho!DFNome_Preco_avista_TBTipo_preco
       End If
       If IsNull(rstCabecalho!DFNome_Preco_promocao_TBTipo_preco) = False Then
          lblCabecalho_preco2.Caption = rstCabecalho!DFNome_Preco_promocao_TBTipo_preco
       End If
       If IsNull(rstCabecalho!DFNome_Preco_revenda_TBTipo_preco) = False Then
          lblCabecalho_preco3.Caption = rstCabecalho!DFNome_Preco_revenda_TBTipo_preco
       End If
       If IsNull(rstCabecalho!DFNome_Preco_especial_TBTipo_preco) = False Then
          lblCabecalho_preco4.Caption = rstCabecalho!DFNome_Preco_especial_TBTipo_preco
       End If
       If IsNull(rstCabecalho!DFNome_Preco_varejo_TBTipo_preco) = False Then
          lblCabecalho_preco5.Caption = rstCabecalho!DFNome_Preco_varejo_TBTipo_preco
       End If
    End If
    
    DoEvents
    
    'POSICIONA FORM
    Me.Left = Left + (Width / 2) - 3345
    Me.Top = (Top + (Height / 2)) - 350
    
End Function

Private Sub Form_Click()
    Unload frmInformacoes_Adicionais_Produto
End Sub

Private Sub Form_Load()
    FormShape1.hWnd = frmInformacoes_Adicionais_Produto.hWnd
    FormShape1.ShapePicture = frmInformacoes_Adicionais_Produto.Picture
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub Label10_Click()
    Unload Me
End Sub

Private Sub Label11_Click()
    Unload Me
End Sub

Private Sub Label12_Click()
    Unload Me
End Sub

Private Sub Label14_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Label3_Click()
    Unload Me
End Sub

Private Sub Label4_Click()
    Unload Me
End Sub

Private Sub Label5_Click()
    Unload Me
End Sub

Private Sub Label6_Click()
    Unload Me
End Sub

Private Sub Label9_Click()
    Unload Me
End Sub

Private Sub lblCabecalho_preco1_Click()
    Unload Me
End Sub

Private Sub lblCabecalho_preco2_Click()
    Unload Me
End Sub

Private Sub lblCabecalho_preco3_Click()
    Unload Me
End Sub

Private Sub lblCabecalho_preco4_Click()
    Unload Me
End Sub

Private Sub lblCabecalho_preco5_Click()
    Unload Me
End Sub

Private Sub lblCategoria_Click()
    Unload Me
End Sub

Private Sub lblDescricao_Click()
    Unload Me
End Sub

Private Sub lblEstoque_Apt_Click()
    Unload Me
End Sub

Private Sub lblEstoque_atual_Click()
    Unload Me
End Sub

Private Sub lblEstoque_max_Click()
    Unload Me
End Sub

Private Sub lblEstoque_min_Click()
    Unload Me
End Sub

Private Sub lblEstoque_Predio_Click()
    Unload Me
End Sub
    
Private Sub lblEstoque_Rua_Click()
    Unload Me
End Sub

Private Sub lblFabricante_Click()
    Unload Me
End Sub

Private Sub lblPreco1_Click()
    Unload Me
End Sub

Private Sub lblPreco2_Click()
    Unload Me
End Sub

Private Sub lblPreco3_Click()
    Unload Me
End Sub

Private Sub lblPreco4_Click()
    Unload Me
End Sub

Private Sub lblPreco5_Click()
    Unload Me
End Sub

Private Sub lblSecao_Click()
    Unload Me
End Sub
