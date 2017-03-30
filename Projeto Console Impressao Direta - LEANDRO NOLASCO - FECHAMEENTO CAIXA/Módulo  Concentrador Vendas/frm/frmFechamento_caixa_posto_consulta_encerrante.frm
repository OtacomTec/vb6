VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFechamento_caixa_posto_consulta_encerrante 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Encerrante"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFechamento_caixa_posto_consulta_encerrante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11865
   StartUpPosition =   1  'CenterOwner
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
      Height          =   1935
      Left            =   90
      TabIndex        =   6
      Top             =   6300
      Width           =   11655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgGrupo_Produto 
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   1720
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Agrupado"
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
         Left            =   120
         TabIndex        =   12
         Top             =   570
         Width           =   1320
      End
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
         Height          =   195
         Left            =   8850
         TabIndex        =   10
         ToolTipText     =   "Total de Valor pego pelo Cliente"
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Vendas:"
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
         Left            =   7500
         TabIndex        =   9
         ToolTipText     =   "Total de IPI  + Total de despesas  acessórios"
         Top             =   270
         Width           =   1200
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
         Height          =   195
         Left            =   1875
         TabIndex        =   8
         ToolTipText     =   "Total de Títulos pegos pelo Cliente"
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Total Quantidade.:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   1605
      End
   End
   Begin VB.TextBox txtOperador 
      Enabled         =   0   'False
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
      Left            =   90
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Código da Finalizadora"
      Top             =   285
      Width           =   1455
   End
   Begin VB.TextBox txtDescricao_Operador 
      Enabled         =   0   'False
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
      Left            =   1590
      MaxLength       =   20
      TabIndex        =   1
      ToolTipText     =   "Código da Finalizadora"
      Top             =   285
      Width           =   8325
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
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
      Left            =   9960
      MaxLength       =   20
      TabIndex        =   2
      ToolTipText     =   "Código da Finalizadora"
      Top             =   285
      Width           =   1785
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgProduto 
      Height          =   5535
      Left            =   90
      TabIndex        =   3
      Top             =   705
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   9763
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
      Caption         =   "Operador"
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
      Index           =   3
      Left            =   90
      TabIndex        =   5
      Top             =   30
      Width           =   810
   End
   Begin VB.Label label15 
      AutoSize        =   -1  'True
      Caption         =   "Data"
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
      Index           =   0
      Left            =   9960
      TabIndex        =   4
      Top             =   30
      Width           =   390
   End
End
Attribute VB_Name = "frmFechamento_caixa_posto_consulta_encerrante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador de Vendas                                         '
' Objetivo...............: Consulta de Encerrante Fechamento Caixa                        '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Peixoto                                                  '
' Data de Criação........: 04/02/2006                                                     '
' Desenvolvedor..........: Leandro Nolasco Ferreira                                       '
' Data última manutenção.: 22/08/2006                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const c_strMaxEncerrante = "999999999.999"

Dim strSQL As String
Dim rstAplicacao As New ADODB.Recordset
Dim booNovo As Boolean
Dim arrAux() As String
Dim dblMaxEncerrante As Double

Public Sub setParametros(ByRef arrEncerrante() As String, Optional ByVal Novo As Boolean = True)

    booNovo = Novo
    arrAux = arrEncerrante

End Sub

Private Sub Form_Load()
    
    Dim rstAux As ADODB.Recordset
    
    Dim intIdx As Integer
    Dim intJdx As Integer
    
    lblTotal_Quantidade.Caption = "0,000"
    lblTotal_Vendas.Caption = "0,00"
    
    hfgProduto.Rows = 2
    Movimentacoes.Monta_HFlex_Grid hfgProduto, "0,0,600,0,1500,1500,1500,900,1550,900,950,1500,0,0", "ID_Bomba_Bico,Bomba,Bico,Cod_Produto,Combustível,Inicial,Final,Aferição,Vendas(L),Pr. Varejo,Custo,Venda ($),Cod_Secao,Desc_Secao", 14, "OTICA", Me
    
    hfgGrupo_Produto.Rows = 2
    Movimentacoes.Monta_HFlex_Grid hfgGrupo_Produto, "1000,4000,2550,1000,2000", "Cod.Produto,Combustível,Venda(L),Pr. Varejo,Venda ($)", 5, "OTICA", Me
    
    'Busca encerrante máximo
    Set rstAux = New ADODB.Recordset
    Call Movimentacoes.Select_geral("SELECT TOP 1 DFNumero_maximo_encerrante_TBBomba_bico FROM TBBomba_Bico", "BDRetaguarda", rstAux, "Otica", Me)
    If rstAux.RecordCount > 0 Then
        rstAux.MoveFirst
        dblMaxEncerrante = CDbl(rstAux(0))
    Else
        dblMaxEncerrante = CDbl(c_strMaxEncerrante)
    End If
    If Not rstAux Is Nothing Then
        If rstAux.State = adStateClosed Then
            rstAux.Close
        End If
        Set rstAux = Nothing
    End If
    
    If booNovo Then
    
        Set rstAux = New ADODB.Recordset
        
        strSQL = Monta_Query(True)
        Call Movimentacoes.Select_geral(strSQL, "BDRetaguarda", rstAux, "Otica", Me)
        If rstAux.RecordCount = 0 Then
            strSQL = Monta_Query
        End If
        
        If Not rstAux Is Nothing Then
            If Not rstAux.State = adStateClosed Then
                rstAux.Close
            End If
            Set rstAux = Nothing
        End If
        
        'Abastecendo os itens
        'Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgProduto, "0,600,600,0,1600,1250,1250,950,1100,900,950,1100,0,0", "ID_Bomba_Bico,Bomba,Bico,Cod_Produto,Combustível,Inicial,Final,Aferição,Vendas(L),Pr. Varejo,Custo,Venda ($),Cod_Secao,Desc_Secao", "BDRetaguarda", "Otica", Me, "N", 3
        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgProduto, "0,0,600,0,1500,1500,1500,900,1550,900,950,1500,0,0", "ID_Bomba_Bico,Bomba,Bico,Cod_Produto,Combustível,Inicial,Final,Aferição,Vendas(L),Pr. Varejo,Custo,Venda ($),Cod_Secao,Desc_Secao", "BDRetaguarda", "Otica", Me, "N", 3
    
    Else
 
        hfgProduto.Rows = UBound(arrAux, 2) + 2
 
        For intIdx = 0 To UBound(arrAux, 2)
        
            For intJdx = 0 To UBound(arrAux, 1) - 1
            
                hfgProduto.TextMatrix(intIdx + 1, intJdx) = arrAux(intJdx, intIdx)
                If intJdx = 0 Then
                    
                    hfgProduto.Row = intIdx + 1
                    hfgProduto.ColWidth(0) = 500
                    hfgProduto.Font.Name = "Tahoma"
                    hfgProduto.CellFontSize = 7
                    hfgProduto.CellFontBold = False
                    hfgProduto.CellBackColor = &H80FFFF
                    
                End If
            
            Next intJdx
        
        Next intIdx
 
 
    End If
    
    If hfgProduto.TextMatrix(1, 1) <> Empty Then
        Call Recalcula_Totais
        Call Recalcula_Totais_Grupo
    End If
       
    txtOperador.Text = frmFechamento_caixa_posto.txtOperador.Text
    txtDescricao_Operador.Text = frmFechamento_caixa_posto.dtcOperador.Text
    txtData.Text = frmFechamento_caixa_posto.dtpFechamento.Value
    
End Sub

Private Function Monta_Query(Optional ByVal Filtra_Ultimo_Encerrante As Boolean = False) As String

    Dim strSQL As String
    
    strSQL = "SELECT PKId_TBbomba_bico, " & _
                    "IXCodigo_Bomba, " & _
                    "IXCodigo_TBBomba_bico, " & _
                    "TBproduto.IXCodigo_TBproduto, " & _
                    "TBproduto.DFdescricao_TBproduto, " & _
                    "ISNULL(DFEncerrante_final_TBEncerrante_caixa_posto, 0.00) AS DFencerrante_inicial, " & _
                    "CONVERT(MONEY, 0.00) AS DFencerrante_final, " & _
                    "CONVERT(MONEY,0.00) AS DFafericao, " & _
                    "CONVERT(MONEY, 0.00) AS DFqtde_venda, " & _
                    "ISNULL(DFValor_unitario_DFEncerrante_inicial_TBEncerrante_caixa_posto, CONVERT(MONEY, ISNULL(DFPreco_varejo_TBItens_tabela_preco, 0.00))) AS DFPreco_varejo_TBItens_tabela_preco, " & _
                    "CONVERT(MONEY, ISNULL(TBproduto.DFcusto_real_TBproduto, 0.00)) AS DFcusto_real_TBproduto, " & _
                    "CONVERT(MONEY, 0.00) AS DFvalor_total_venda, " & _
                    "PKCodigo_TBsecao, " & _
                    "DFDescricao_TBsecao "
    strSQL = strSQL & _
               "FROM TBbomba_bico " & _
         "INNER JOIN TBbomba " & _
                 "ON TBbomba.PKId_TBbomba = TBbomba_bico.FKId_TBbomba " & _
         "LEFT  JOIN TBEncerrante_caixa_posto " & _
                 "ON PKId_TBbomba_bico = FKId_TBbomba_bico " & _
         "INNER JOIN TBProduto " & _
                 "ON TBbomba_bico.FKId_TBproduto = TBProduto.PKId_TBProduto " & _
                "AND TBProduto.IXCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' " & _
         "INNER JOIN TBitens_tabela_preco " & _
                 "ON PKId_TBProduto = TBitens_tabela_preco.FKId_TBproduto " & _
                "AND FKCodigo_TBTabela_preco IN ( SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBparametros_venda WHERE IXCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' ) " & _
         "LEFT  JOIN TBsecao " & _
                 "ON TBproduto.FKCodigo_TBsecao = TBsecao.PKCodigo_TBsecao "

    If Filtra_Ultimo_Encerrante Then
        strSQL = strSQL & _
              "WHERE PKId_TBEncerrante_caixa_posto IN " & _
                    "( SELECT MAX(PKId_TBEncerrante_caixa_posto) FROM TBEncerrante_caixa_posto GROUP BY FKId_TBBomba_bico )"
    End If
    
    Monta_Query = strSQL
    
End Function

Private Sub Form_Unload(Cancel As Integer)

    Cancel = Abs(CInt(Not frmFechamento_caixa_posto.setLista_Encerrante_Bico(hfgProduto)))

End Sub

Private Sub hfgProduto_KeyPress(KeyAscii As Integer)
    If hfgProduto.Rows >= 2 Then
        If hfgProduto.TextMatrix(hfgProduto.Row, 1) <> Empty Then
            If (hfgProduto.Col >= 6 And hfgProduto.Col < 11 And hfgProduto.Col <> 9) And hfgProduto.Row > 0 And hfgProduto.ColWidth(hfgProduto.Col) > 0 Then
                EscreveNaGrid hfgProduto, hfgProduto.Col, hfgProduto.Row, KeyAscii, True, False
            End If
        End If
    End If
End Sub

Private Sub hfgProduto_EnterCell()
    If (hfgProduto.Col < 6 Or hfgProduto.Col > 11) Then
        Exit Sub
    End If
    hfgProduto.CellFontBold = True
    'hfgProduto.CellForeColor = &H808080
End Sub

Private Sub hfgProduto_LeaveCell()
    
    With hfgProduto
        If (.Col < 6 And .Col > 11) Then
            Exit Sub
        End If
        
        If (.Col >= 6 And .Col < 12) Then
            .Text = Format(.Text, "##,###0.000")
        ElseIf .Col = 12 Then
            .Text = Format(.Text, "##,##0.00")
        End If
        
        .CellFontBold = False
        .CellForeColor = &H0
        
        If IsNumeric(.TextMatrix(.Row, .Col)) Then
            If CDbl(.TextMatrix(.Row, 7)) < CDbl(.TextMatrix(.Row, 6)) Then
                .TextMatrix(.Row, 9) = Format((Round(dblMaxEncerrante - CDbl(.TextMatrix(.Row, 6)), 3) + CDbl(.TextMatrix(.Row, 7))) - CDbl(.TextMatrix(.Row, 8)), "##,###0.000")
            Else
                .TextMatrix(.Row, 9) = Format(CDbl(.TextMatrix(.Row, 7)) - CDbl(.TextMatrix(.Row, 6)) - CDbl(.TextMatrix(.Row, 8)), "##,###0.000")
            End If
            .TextMatrix(.Row, 12) = Format(CDbl(.TextMatrix(.Row, 9)) * CDbl(.TextMatrix(.Row, 10)), "##,##0.00")
        End If
        
        Call Recalcula_Totais(True)
        Call Recalcula_Totais_Grupo
    End With
    
End Sub

Private Sub hfgProduto_GotFocus()
    If hfgProduto.Col >= 6 Or hfgProduto.Col < 11 Then
        Call hfgProduto_EnterCell
    End If
End Sub

Private Sub hfgProduto_LostFocus()
    Call hfgProduto_LeaveCell
    Call Recalcula_Totais
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

Private Sub Recalcula_Totais(Optional ByVal Digitacao As Boolean = False)

    Dim I As Integer
    Dim J As Integer
    Dim dblValor As Double
    Dim dblQtde As Double
    
    For I = 1 To hfgProduto.Rows - 1
        dblQtde = dblQtde + CDbl(IIf(hfgProduto.TextMatrix(I, 9) = Empty, 0, hfgProduto.TextMatrix(I, 9)))
        dblValor = dblValor + CDbl(IIf(hfgProduto.TextMatrix(I, 12) = Empty, 0, hfgProduto.TextMatrix(I, 12)))
    Next I

    lblTotal_Quantidade.Caption = Format(dblQtde, "##,###0.000")
    lblTotal_Vendas.Caption = Format(dblValor, "##,##0.00")

End Sub

Private Sub Recalcula_Totais_Grupo()

    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    Dim l As Integer
    
    Dim booGrupo_Existente As Boolean
    
    Dim strCod_Novo As String
    Dim strDesc_Novo As String
    
    Dim dblPreco_Praticado As Double
    Dim dblTotal_Valor As Double
    Dim dblTotal_Qtde As Double
    
    'Limpando grid totalizador
    hfgGrupo_Produto.Clear
    hfgGrupo_Produto.Rows = 2
    Movimentacoes.Monta_HFlex_Grid hfgGrupo_Produto, "1000,4000,2550,1000,2000", "Cod.Produto,Combustível,Venda(L),Pr. Varejo,Venda ($)", 5, "OTICA", Me
    
    With hfgProduto
    
        For I = 1 To .Rows - 1
        
            'Seleciona um grupo - excluindo já inseridos
            booGrupo_Existente = False
            
            'Verifica se o produto e preço selecionado já foi inserido
            For l = 1 To hfgGrupo_Produto.Rows - 1
                If hfgGrupo_Produto.TextMatrix(l, 1) = .TextMatrix(I, 4) And hfgGrupo_Produto.TextMatrix(l, 4) = .TextMatrix(I, 10) Then
                    booGrupo_Existente = True
                    Exit For
                End If
            Next l
            
            If Not booGrupo_Existente Then
            
               'Seleciona novos produtos
               strCod_Novo = .TextMatrix(I, 4)
               strDesc_Novo = .TextMatrix(I, 5)
               dblTotal_Qtde = .TextMatrix(I, 9)
               dblPreco_Praticado = .TextMatrix(I, 10)
               dblTotal_Valor = .TextMatrix(I, 12)
    
               'Totaliza grupo
               'Se ordenar por código do produto e preço unitário, J deve começar de I e não de 1, para aumentar o desempenho do algoritimo
               For J = 1 To .Rows - 1
                    If J <> I Then
                        If .TextMatrix(J, 4) = strCod_Novo And .TextMatrix(J, 10) = dblPreco_Praticado Then
                            dblTotal_Qtde = dblTotal_Qtde + CDbl(.TextMatrix(J, 9))
                            dblTotal_Valor = dblTotal_Valor + CDbl(.TextMatrix(J, 12))
                        End If
                    End If
               Next J
    
               'Insere o grupo
               If hfgGrupo_Produto.Rows >= 2 And hfgGrupo_Produto.TextMatrix(1, 1) <> Empty Then
                   hfgGrupo_Produto.AddItem ""
               End If
               K = hfgGrupo_Produto.Rows - 1
               hfgGrupo_Produto.TextMatrix(K, 0) = K
               
               'formatando grid
               hfgGrupo_Produto.Row = K
               hfgGrupo_Produto.Col = 0
               hfgGrupo_Produto.ColWidth(0) = 500
               hfgGrupo_Produto.Font.Name = "Tahoma"
               hfgGrupo_Produto.CellFontSize = 7
               hfgGrupo_Produto.CellFontBold = False
               hfgGrupo_Produto.CellBackColor = &H80FFFF
               
               'inserindo valores
               hfgGrupo_Produto.TextMatrix(K, 1) = strCod_Novo
               hfgGrupo_Produto.TextMatrix(K, 2) = strDesc_Novo
               hfgGrupo_Produto.TextMatrix(K, 3) = Format(dblTotal_Qtde, "##,###0.000")
               hfgGrupo_Produto.TextMatrix(K, 4) = Format(dblPreco_Praticado, "##,###0.000")
               hfgGrupo_Produto.TextMatrix(K, 5) = Format(dblTotal_Valor, "##,##0.00")
               
            End If
            
        Next I
    
    End With
    
    
End Sub
