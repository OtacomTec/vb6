VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConsulta_Produto 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Consulta de Produtos"
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   0
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
   ScaleHeight     =   6870
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   8460
      Picture         =   "frmConsulta_Produto.frx":1782
      ScaleHeight     =   615
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   30
      Width           =   435
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgProduto 
      Height          =   2685
      Left            =   300
      TabIndex        =   4
      Top             =   2820
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   4736
      _Version        =   393216
      BackColor       =   8454143
      BackColorFixed  =   8454143
      BackColorBkg    =   8454143
      BackColorUnpopulated=   8454143
      GridColorFixed  =   8454143
      GridColorUnpopulated=   8454143
      AllowBigSelection=   0   'False
      FocusRect       =   2
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   8280
      ScaleHeight     =   2625
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   2880
      Width           =   315
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
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5970
      Width           =   1635
   End
   Begin VB.OptionButton optDescricao 
      BackColor       =   &H0080FFFF&
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1785
   End
   Begin VB.OptionButton optCodigo_interno 
      BackColor       =   &H0080FFFF&
      Caption         =   "código Interno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   1530
      Width           =   1785
   End
   Begin VB.OptionButton optCod_ean 
      BackColor       =   &H0080FFFF&
      Caption         =   "código Barra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   1020
      Width           =   1785
   End
   Begin VB.TextBox txtCodigo_Produto 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   2550
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1650
      Width           =   6105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consulta Detalhada de Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   750
      TabIndex        =   10
      Top             =   210
      Width           =   5565
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   3600
      X2              =   60
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8910
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblAguarde 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aguarde......."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      TabIndex        =   8
      Top             =   6090
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Line Line3 
      X1              =   8910
      X2              =   8910
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8910
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2985
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   2700
      Width           =   8565
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2925
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   8385
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   675
      Left            =   6900
      Shape           =   4  'Rounded Rectangle
      Top             =   6030
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
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   345
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   1590
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   345
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   825
      Left            =   2430
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   6285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2430
      TabIndex        =   6
      Top             =   1110
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   825
      Left            =   2340
      Shape           =   4  'Rounded Rectangle
      Top             =   1650
      Width           =   6285
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
Dim log As New DLLSystemManager.log


Private Sub cmdOk_Click()
    
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
    
    'Habilita a saida com ESC
    If KeyAscii = 27 Then
        Unload Me
    End If
    
    'Verifica se foi prescionado o CTRL + A
    If KeyAscii = 4 Then
        Me.optDescricao.SetFocus
    End If
    
    'Verifica se foi prescionado o CTRL + I
    If KeyAscii = 9 Then
        Me.optCodigo_interno.SetFocus
    End If
    
    'Verifica se foi prescionado o CTRL + B
    If KeyAscii = 2 Then
        Me.optCod_ean.SetFocus
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

Private Sub optCod_ean_Click()
    txtCodigo_Produto.Text = Empty
    Me.txtCodigo_Produto.SetFocus
End Sub

Private Sub optCodigo_interno_Click()
    txtCodigo_Produto.Text = Empty
    Me.txtCodigo_Produto.SetFocus
End Sub

Private Sub optDescricao_Click()
    txtCodigo_Produto.Text = Empty
    Me.txtCodigo_Produto.SetFocus
End Sub

Private Sub txtCodigo_Produto_KeyPress(KeyAscii As Integer)
    If Me.optCod_ean.Value = True Or Me.optCodigo_interno.Value = True Then
       If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub txtCodigo_Produto_LostFocus()
    
     Dim rstProduto As New ADODB.Recordset
     
     Me.lblAguarde.Visible = True
     
     If Me.txtCodigo_Produto.Text <> "" Then
        'Descrição
        If optDescricao.Value = True Then
        
           If booIntegracao_Retaguarda = True And booPreco_online = True Then
              strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto,TBITENS_TABELA_PRECO.DFPreco_VAREJO_TBItens_tabela_preco " & _
                       "FROM TBPRODUTO " & _
                       "INNER JOIN TBITENS_TABELA_PRECO " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                       "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA WHERE IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ") " & _
                       "AND convert(nvarchar,TBProduto.DFDescricao_TBProduto) LIKE '%" & Me.txtCodigo_Produto.Text & "%' " & _
                       "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
              
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDRetaguarda", "Otica", Me
           Else
              strSql = Empty
              strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPreco_venda_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                       "FROM TBProduto " & _
                       "WHERE convert(nvarchar,TBProduto.DFDescricao_TBProduto) LIKE '%" & Me.txtCodigo_Produto.Text & "%' " & _
                       "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDPDV", "PDV", Me
           End If
           
        End If
        
        'Cod. Interno
        If optCodigo_interno.Value = True Then
        
           If booIntegracao_Retaguarda = True And booPreco_online = True Then
              strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto,TBITENS_TABELA_PRECO.DFPreco_VAREJO_TBItens_tabela_preco " & _
                       "FROM TBPRODUTO " & _
                       "INNER JOIN TBITENS_TABELA_PRECO " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                       "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA WHERE IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ")" & _
                       "AND TBProduto.IXCodigo_TBProduto = " & Me.txtCodigo_Produto.Text & " " & _
                       "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDRetaguarda", "Otica", Me
           Else
              strSql = Empty
              strSql = "SELECT TBProduto.PKId_TBProduto,TBProduto.DFCst1_TBProduto,TBProduto.DFCst2_TBProduto,TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_resumida_TBProduto,TBProduto.DFPreco_venda_TBProduto,TBProduto.DFPath_imagem_TBProduto " & _
                       "FROM TBProduto " & _
                       "WHERE TBProduto.IXCodigo_TBProduto = " & Me.txtCodigo_Produto.Text & " " & _
                       "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDPDV", "PDV", Me
           End If
           
        End If
        
        'Código de barra
        If Me.optCod_ean.Value = True Then
           If booIntegracao_Retaguarda = True And booPreco_online = True Then
              strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto,TBITENS_TABELA_PRECO.DFPreco_varejo_TBItens_tabela_preco " & _
                       "FROM TBPRODUTO " & _
                       "INNER JOIN TBITENS_TABELA_PRECO " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                       "INNER JOIN TBCodigo_barras " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBCodigo_barras.FKID_TBProduto " & _
                       "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA WHERE IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ") " & _
                       "AND TBCodigo_barras.IXCodigo_TBCodigo_barras = " & txtCodigo_Produto.Text & " " & _
                       "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDRetaguarda", "Otica", Me
           Else
              strSql = Empty
              strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto,TBProduto.DFPreco_varejo_TBItens_tabela_preco " & _
                       "FROM TBProduto " & _
                       "INNER JOIN TBCodigo_barras " & _
                       "ON TBPRODUTO.PKID_TBProduto = TBCodigo_barras.FKID_TBProduto " & _
                       "WHERE TBCodigo_barras.IXCodigo_TBCodigo_barras = " & txtCodigo_Produto.Text & " " & _
                       "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgProduto, "1000,4700,1000", "Código,Descrição,Preço($)", "BDPDV", "PDV", Me
           End If
        End If
     End If
     
     Me.lblAguarde.Visible = False
     
     txtCodigo_Produto.Text = Empty
     
     Me.hfgProduto.SetFocus
     
End Sub
