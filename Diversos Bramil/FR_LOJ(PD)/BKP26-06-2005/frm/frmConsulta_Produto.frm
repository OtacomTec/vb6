VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConsulta_Produto 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Produtos"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsulta_Produto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgProduto 
      DragMode        =   1  'Automatic
      Height          =   2685
      Left            =   300
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2010
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   4736
      _Version        =   393216
      BackColor       =   8454143
      BackColorFixed  =   8454143
      BackColorBkg    =   8454143
      BackColorUnpopulated=   8454143
      GridColorFixed  =   8454143
      GridColorUnpopulated=   8454143
      AllowBigSelection=   0   'False
      HighLight       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   8280
      ScaleHeight     =   2625
      ScaleWidth      =   405
      TabIndex        =   8
      Top             =   2070
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1635
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1635
   End
   Begin VB.OptionButton optDescricao 
      BackColor       =   &H0080FFFF&
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.OptionButton optCodigo_interno 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cod. Interno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   690
      Width           =   1815
   End
   Begin VB.OptionButton optCod_ean 
      BackColor       =   &H0080FFFF&
      Caption         =   "Código barra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   1815
   End
   Begin VB.TextBox txtCodigo_Produto 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   2550
      MaxLength       =   14
      TabIndex        =   3
      Top             =   810
      Width           =   6135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2985
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   1890
      Width           =   8655
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2925
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   2070
      Width           =   8475
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   675
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   5220
      Width           =   1665
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   675
      Left            =   7050
      Shape           =   4  'Rounded Rectangle
      Top             =   5220
      Width           =   1665
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   345
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   1290
      Width           =   1845
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   345
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   1845
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   345
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   270
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   525
      Left            =   2430
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   6315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2430
      TabIndex        =   7
      Top             =   330
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   525
      Left            =   2340
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   6315
   End
End
Attribute VB_Name = "frmConsulta_Produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Faturamento                                                    '
' Objetivo...............: Consulta de produtos no ECF                                    '
' Data de Criação........: 13/01/2005                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim conexao As New DLLConexao_Sistema.conexao
Dim acesso As New DLLSystemManager.Acessibilidade
Dim booIntegracao_Retaguarda As Boolean
Dim booPreco_online As Boolean
Dim Log As New DLLSystemManager.Log

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If hfgProduto.Col = 0 Then
       frmTela_Venda.txtCodigo_Produto = hfgProduto.TextArray((hfgProduto.Row * hfgProduto.Cols + hfgProduto.Col + 1))
       Unload Me
       frmTela_Venda.txtDescricao_Produto.SetFocus
    End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
       booIntegracao_Retaguarda = frmTela_Venda.booIntegracao_Retaguarda
       booPreco_online = frmTela_Venda.booPreco_online
End Sub

Private Sub hfgProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       'Call hfgProduto_click
    End If
End Sub

Private Sub txtCodigo_Produto_LostFocus()
    
     Dim rstProduto As New ADODB.Recordset
     
     If Me.txtCodigo_Produto.Text <> "" Then
        'Descrição
        If optDescricao.Value = True Then
        
           If booIntegracao_Retaguarda = True And booPreco_online = True Then
              strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto,TBITENS_TABELA_PRECO.DFPreco_avista_TBItens_tabela_preco " & _
                       "FROM TBPRODUTO " & _
                       "INNER JOIN TBITENS_TABELA_PRECO " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                       "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA) " & _
                       "AND convert(nvarchar,TBProduto.DFDescricao_TBProduto) LIKE '% " & Me.txtCodigo_Produto.Text & " %' "
              
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDRetaguarda", "Otica", Me
           Else
              strSql = Empty
              strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPreco_venda_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                       "FROM TBProduto " & _
                       "AND convert(nvarchar,TBProduto.DFDescricao_TBProduto) LIKE '%" & Me.txtCodigo_Produto.Text & "%' "
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDPDV", "PDV", Me
           End If
           
        End If
        
        'Cod. Interno
        If optCodigo_interno.Value = True Then
        
           If booIntegracao_Retaguarda = True And booPreco_online = True Then
              strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto,TBITENS_TABELA_PRECO.DFPreco_avista_TBItens_tabela_preco " & _
                       "FROM TBPRODUTO " & _
                       "INNER JOIN TBITENS_TABELA_PRECO " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                       "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA) " & _
                       "AND TBProduto.IXCodigo_TBProduto = " & Me.txtCodigo_Produto.Text & ""
              
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDRetaguarda", "Otica", Me
           Else
              strSql = Empty
              strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPreco_venda_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                       "FROM TBProduto " & _
                       "WHERE TBProduto.IXCodigo_TBProduto = " & Me.txtCodigo_Produto.Text & ""
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDPDV", "PDV", Me
           End If
           
        End If
        
        'Código de barra
        If Me.optCod_ean.Value = True Then
        
           If booIntegracao_Retaguarda = True And booPreco_online = True Then
              strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBITENS_TABELA_PRECO.DFPreco_varejo_TBItens_tabela_preco,DFPath_imagem_TBProduto " & _
                       "FROM TBPRODUTO " & _
                       "INNER JOIN TBITENS_TABELA_PRECO " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                       "INNER JOIN TBCodigo_barras " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBCodigo_barras.FKID_TBProduto " & _
                       "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA) " & _
                       "AND TBCodigo_barras.IXCodigo_TBCodigo_barras = " & txtCodigo_Produto.Text & " "
              
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDRetaguarda", "Otica", Me
           Else
              strSql = Empty
              strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPreco_venda_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                       "FROM TBProduto " & _
                       "INNER JOIN TBCodigo_barras " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBCodigo_barras.FKID_TBProduto " & _
                       "WHERE TBCodigo_barras.IXCodigo_TBCodigo_barras = " & txtCodigo_Produto.Text & " "
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDPDV", "PDV", Me
           End If
           
        End If
     End If
     
End Sub
