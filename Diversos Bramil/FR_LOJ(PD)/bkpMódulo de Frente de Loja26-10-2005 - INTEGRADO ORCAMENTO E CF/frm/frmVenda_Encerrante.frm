VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmVenda_Encerrante 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Vendedor"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   Icon            =   "frmVenda_Encerrante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5910
      Picture         =   "frmVenda_Encerrante.frx":1782
      ScaleHeight     =   615
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   30
      Width           =   435
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   8400
      Picture         =   "frmVenda_Encerrante.frx":4B90
      ScaleHeight     =   615
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   0
      Width           =   435
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5880
      ScaleHeight     =   615
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   1740
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   270
      ScaleHeight     =   75
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   1680
      Width           =   5835
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   270
      ScaleHeight     =   675
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   1710
      Width           =   15
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   270
      ScaleHeight     =   75
      ScaleWidth      =   5835
      TabIndex        =   2
      Top             =   2280
      Width           =   5835
   End
   Begin MSDataListLib.DataCombo dtcEncerrante 
      Height          =   570
      Left            =   270
      TabIndex        =   4
      ToolTipText     =   "Finalizadora"
      Top             =   1740
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   1005
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   8454143
      ForeColor       =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2820
      Width           =   1635
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   900
      TabIndex        =   13
      Top             =   3060
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblDescricao_produto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   1110
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Label lblProduto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   180
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblDescricao 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Produto associado a este bico:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Top             =   2700
      Visible         =   0   'False
      Width           =   3720
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   3060
      X2              =   60
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bico na bomba"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   210
      Width           =   2640
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   3630
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3630
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6360
      Y1              =   3630
      Y2              =   3630
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   645
      Left            =   4470
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   1665
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   1590
      Width           =   6045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   1140
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   915
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   1650
      Width           =   6015
   End
End
Attribute VB_Name = "frmVenda_Encerrante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Dim intTipo_Preco As Integer

Private Sub cmdOk_Click()

    If dtcEncerrante.Text = "" Then
       dtcEncerrante.SetFocus
    Else
       frmTela_Venda.txtCodigo_Produto.Text = CLng(Me.lblProduto.Caption)
       frmTela_Venda.txtQuantidade_Produto.TabStop = True
       
       Dim rstInf_Produtos As New ADODB.Recordset
       
       strSql = Empty
       strSql = "SELECT DFPreco_avista_TBItens_tabela_preco,DFPreco_promocao_TBItens_tabela_preco,DFPreco_revenda_TBItens_tabela_preco,DFPreco_especial_TBItens_tabela_preco,DFPreco_varejo_TBItens_tabela_preco " & _
                "FROM TBProduto " & _
                "INNER JOIN TBITENS_TABELA_PRECO " & _
                "ON TBPRODUTO.PKID_TBProduto = TBITENS_TABELA_PRECO.FKID_TBProduto " & _
                "WHERE TBITENS_TABELA_PRECO.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA WHERE IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ")" & _
                "AND TBProduto.IXCodigo_TBProduto = " & CLng(Me.lblProduto.Caption) & " " & _
                "AND TBPRODUTO.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""

       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstInf_Produtos, "Otica", Me
      
       If rstInf_Produtos.BOF = True And rstInf_Produtos.EOF = True Then
          MsgBox "Código Interno não cadastrado.Verifique!", vbCritical, "Only Tech"
          frmTela_Venda.txtCodigo_Produto.Text = Empty
          frmTela_Venda.txtQuantidade_Produto.TabStop = False
          Set rstInf_Produtos = Nothing
          Exit Sub
       End If
    
       If rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco <= 0 Or IsNull(rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco) Then
          MsgBox "Preço do Item não cadastrado.Verifique!", vbCritical, "Only Tech"
          frmTela_Venda.txtCodigo_Produto.Text = Empty
          frmTela_Venda.txtQuantidade_Produto.TabStop = False
          Set rstInf_Produtos = Nothing
          Exit Sub
       End If
       
       'Preço 1
       If intTipo_Preco = 1 Then
          frmTela_Venda.txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_avista_TBItens_tabela_preco
       End If
       'Preço 2
       If intTipo_Preco = 2 Then
          frmTela_Venda.txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_promocao_TBItens_tabela_preco
       End If
       'Preço 3
       If intTipo_Preco = 3 Then
          frmTela_Venda.txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_revenda_TBItens_tabela_preco
       End If
       'Preço 4
       If intTipo_Preco = 4 Then
          frmTela_Venda.txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_especial_TBItens_tabela_preco
       End If
       'Preço 5
       If intTipo_Preco = 5 Then
          frmTela_Venda.txtPreco_Unitario.Text = rstInf_Produtos!DFPreco_varejo_TBItens_tabela_preco
       End If
              
       Set rstInf_Produtos = Nothing
       
       frmTela_Venda.booConsulta = False
       
       Unload Me
       
       Call frmTela_Venda.txtCodigo_Produto_LostFocus
       
       
    End If
    
    frmTela_Venda.booConsulta = False
    
End Sub

Private Sub dtcEncerrante_LostFocus()

    Dim rstProduto_bico As New ADODB.Recordset
    If Me.dtcEncerrante.Text <> "" Then
         strSql = Empty
         strSql = "SELECT TBProduto.IXCodigo_TBProduto,TBProduto.DFDescricao_TBProduto,DFTipo_preco_TBBomba_bico " & _
                  "FROM TBBOMBA_BICO " & _
                  "INNER JOIN TBProduto " & _
                  "ON TBBOMBA_BICO.FKId_TBProduto = TBProduto.PKId_TBProduto " & _
                  "WHERE TBBOMBA_BICO.IXCodigo_TBBomba_bico = " & Me.dtcEncerrante.Text & " " & _
                  "AND TBProduto.IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
                  
        Movimentacoes.Select_geral strSql, "BDRetaguarda", rstProduto_bico, "Otica", Me
        
        If rstProduto_bico.BOF = True And rstProduto_bico.EOF = True Then
           MsgBox "Produto não associado no bico!Verifique.", vbCritical, "Only Tech"
           Set rstProduto_bico = Nothing
           Exit Sub
        End If
        
        Me.lblDescricao_produto.Visible = True
        Me.lblProduto.Visible = True
        Me.Label3.Visible = True
        Me.lblDescricao.Visible = True
        
        If rstProduto_bico.BOF = False And rstProduto_bico.EOF = False Then
            Me.lblProduto.Caption = rstProduto_bico!IXCodigo_TBProduto
            Me.lblDescricao_produto.Caption = rstProduto_bico!DFDescricao_TBProduto
            intTipo_Preco = rstProduto_bico!DFTipo_preco_TBBomba_bico
        End If
        
        Set rstProduto_bico = Nothing
        
    End If
End Sub

Private Sub Form_Activate()
  
  Call Movimentacoes.Verifica_DataCombo(Me.dtcEncerrante)
  Me.dtcEncerrante.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    'Habilita a saida com ESC
    If KeyAscii = 27 Then
        frmTela_Venda.booConsulta = False
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

  frmTela_Venda.booConsulta = True
  frmTela_Venda.booEncerrante = True
  
  strSql = Empty
  strSql = "SELECT PKId_TBBomba_bico,IXCodigo_TBBomba_bico FROM TBBOMBA_BICO"
  Call Movimentacoes.Movimenta_DataCombo("PKId_TBBomba_bico", "IXCodigo_TBBomba_bico", dtcEncerrante, strSql, "BDRetaguarda", "Otica", Me)

End Sub
