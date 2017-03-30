VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFechamento_caixa_posto_conferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conferência de Finalizadora"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFechamento_caixa_posto_conferencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   10200
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCliente 
      Enabled         =   0   'False
      Height          =   360
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   4
      Top             =   900
      Width           =   1395
   End
   Begin VB.TextBox txtNome_Cliente 
      Enabled         =   0   'False
      Height          =   360
      Left            =   3030
      MaxLength       =   20
      TabIndex        =   5
      Top             =   900
      Width           =   4635
   End
   Begin VB.CommandButton cmdAlterar 
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
      Left            =   9690
      Picture         =   "frmFechamento_caixa_posto_conferencia.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Alterar"
      Top             =   900
      Width           =   405
   End
   Begin VB.TextBox txtValor 
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7740
      MaxLength       =   20
      TabIndex        =   6
      Top             =   900
      Width           =   1875
   End
   Begin VB.TextBox txtCupom 
      Enabled         =   0   'False
      Height          =   360
      Left            =   90
      TabIndex        =   3
      Top             =   900
      Width           =   1395
   End
   Begin VB.CommandButton cmdSalvar 
      Cancel          =   -1  'True
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
      Left            =   9690
      Picture         =   "frmFechamento_caixa_posto_conferencia.frx":3B54
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salvar"
      Top             =   3690
      Width           =   405
   End
   Begin VB.TextBox txtOperador 
      Enabled         =   0   'False
      Height          =   360
      Left            =   90
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Código do Operador"
      Top             =   270
      Width           =   1395
   End
   Begin VB.TextBox txtData 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7740
      MaxLength       =   20
      TabIndex        =   2
      Top             =   270
      Width           =   2325
   End
   Begin VB.TextBox txtDescricao_Operador 
      Enabled         =   0   'False
      Height          =   360
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   1
      Top             =   270
      Width           =   6105
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgConferencia 
      Height          =   2295
      Left            =   90
      TabIndex        =   9
      Top             =   1320
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   240
      Left            =   1560
      TabIndex        =   16
      Top             =   660
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   240
      Left            =   7740
      TabIndex        =   15
      Top             =   660
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Cupom"
      Height          =   240
      Left            =   90
      TabIndex        =   14
      Top             =   660
      Width           =   600
   End
   Begin VB.Label Label47 
      Caption         =   "Total :"
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
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   885
   End
   Begin VB.Label lblTotal 
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
      Left            =   1110
      TabIndex        =   10
      Top             =   3720
      Width           =   4995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   240
      Left            =   7740
      TabIndex        =   12
      Top             =   30
      Width           =   390
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Operador"
      Height          =   240
      Left            =   90
      TabIndex        =   11
      Top             =   30
      Width           =   810
   End
End
Attribute VB_Name = "frmFechamento_caixa_posto_conferencia"
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
' Data última manutenção.: 04/08/2006                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String
Dim rstAplicacao As New ADODB.Recordset
Dim intContador As Integer
Dim intLinha As Integer

Private Type tParametros
    Data_Fechamento As Date
    Cod_Operador As String
    Nome_Operador As String
    Cod_Finalizadora As String
    Lista_Existe As Boolean
End Type

Dim objParam As tParametros

Public Function setParametros(ByVal datData_Fechamento As Date, ByVal strCod_Operador As String, ByVal strNome_Operador As String, ByVal strCod_Finalizadora As String, Optional ByVal Lista_Alterada As Boolean = False) As Boolean

    setParametros = False
    
    With objParam
        .Cod_Finalizadora = strCod_Finalizadora
        .Cod_Operador = strCod_Operador
        .Nome_Operador = strNome_Operador
        .Data_Fechamento = datData_Fechamento
        .Lista_Existe = Lista_Alterada
    End With
    
    setParametros = True

End Function

Private Sub cmdAlterar_Click()
    
    If txtCupom.Text = Empty Then
        MsgBox "Cupom inválido. Verifique.", vbInformation, "Only Tech"
        Exit Sub
    ElseIf txtValor.Text = Empty Then
       MsgBox "Valor inválido. Verifique.", vbInformation, "Only Tech"
       txtValor.SetFocus
       Exit Sub
    End If
    
    lblTotal.Caption = Format(CDbl(txtValor.Text) - CDbl(hfgConferencia.TextMatrix(intLinha, 5)) + CDbl(lblTotal.Caption), "##,##0.00")
    hfgConferencia.TextMatrix(intLinha, 5) = txtValor.Text
    
    hfgConferencia.Col = 1
    hfgConferencia.CellForeColor = &H8000&
    hfgConferencia.CellFontBold = True
    hfgConferencia.Text = "X"
          
    txtCupom.Text = Empty
    txtCliente.Text = Empty
    txtNome_Cliente.Text = Empty
    txtValor.Text = Empty
    
End Sub

Private Sub cmdSalvar_Click()
    Dim Conexao As New DLLConexao_Sistema.Conexao
    Dim strValor_Alterado As String
    Dim strIDOperacao As String
    
    On Error GoTo Erro
    
'    Conexao.Initial_Catalog = "BDRetaguarda"
'    Conexao.Abrir_conexao "Otica"
'    Conexao.CNconexao.BeginTrans
'
'    intContador = 1
'    Do While intContador <= hfgConferencia.Rows - 1
'
'       hfgConferencia.Row = intContador
'       hfgConferencia.Col = 11
'       strIDOperacao = hfgConferencia.Text
'
'       hfgConferencia.Col = 1
'       'Alteracao
'       If hfgConferencia.CellForeColor = &H8000& Then
'          hfgConferencia.Col = 5
'          strValor_Alterado = hfgConferencia.Text
'          hfgConferencia.Col = 6
'          If CDbl(strValor_Alterado) <> CDbl(hfgConferencia.Text) Then
'             strSql = "UPDATE TBOperacao_caixa " & _
'                      "SET DFValor_TBOperacao_caixa = " & Funcoes_Gerais.Grava_Moeda(strValor_Alterado) & " " & _
'                      "WHERE PKId_TBOperacao_caixa = " & strIDOperacao & ""
'
'             Conexao.CNconexao.Execute strSql
'          End If
'       'Exclusao
'       ElseIf hfgConferencia.CellForeColor = &HC0& Then
'          strSql = "DELETE FROM TBOperacao_caixa " & _
'                   "WHERE PKId_TBOperacao_caixa = " & strIDOperacao & ""
'
'          Conexao.CNconexao.Execute strSql
'       End If
'       intContador = intContador + 1
'    Loop
'
'    Conexao.CNconexao.CommitTrans
'    Conexao.Fechar_conexao
'
'    frmFechamento_caixa_posto.hfgFinalizadora.Col = 1
'    intContador = 1
'    Do While intContador <= frmFechamento_caixa_posto.hfgFinalizadora.Rows - 1
'       frmFechamento_caixa_posto.hfgFinalizadora.Row = intContador
'       If frmFechamento_caixa_posto.hfgFinalizadora.Text = frmFechamento_caixa_posto.strCodigo_Finalizadora Then
'          frmFechamento_caixa_posto.hfgFinalizadora.Col = 3
'          'Recalculando o subtotal
'          frmFechamento_caixa_posto.txtSubTotal.Text = Format(CDbl(lblTotal.Caption) - CDbl(frmFechamento_caixa_posto.hfgFinalizadora.Text) + CDbl(frmFechamento_caixa_posto.txtSubTotal.Text), "#,###0.000")
'          frmFechamento_caixa_posto.hfgFinalizadora.Text = lblTotal.Caption
'          frmFechamento_caixa_posto.txtResultado_Caixa.Text = Format(CDbl(frmFechamento_caixa_posto.txtSubTotal.Text) - CDbl(frmFechamento_caixa_posto.txtTotal_Vendas.Text), "#,###0.00")
'          Exit Do
'       End If
'       intContador = intContador + 1
'    Loop
'
'    Unload Me
    
'    If frmFechamento_caixa_posto.setLista_Itens_Finalizadora(hfgConferencia) Then
'        Call frmFechamento_caixa_posto.setValor_Finalizadora_Atualizada(lblTotal.Caption)
'        MsgBox "Alteração dos cupons concluída", vbOKOnly + vbInformation, "Only Tech"
'        Unload Me
'    End If
    
    Exit Sub
    
Erro:
    'Conexao.CNconexao.RollbackTrans
    'Conexao.Fechar_conexao
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim intContador As Integer
    
    With objParam
    
        txtOperador.Text = .Cod_Operador
        txtDescricao_Operador.Text = .Nome_Operador
        txtData.Text = CStr(objParam.Data_Fechamento)
    
    End With
    
    lblTotal.Caption = "0,00"
    
    If Not objParam.Lista_Existe Then
    
        'Abastecendo os itens
        strSQL = "SELECT '',IXCodigo_TBFinalizadora," & _
                        "DFDescricao_TBFinalizadora,TBCupom.DFNumero_TBCupom," & _
                        "DFValor_TBOperacao_caixa,DFValor_TBOperacao_caixa,convert(char,TBOperacao_caixa.DFHora_TBOperacao_caixa,108)," & _
                        "IXCodigo_TBCliente,DFNome_TBCliente,DFDebito_credito_TBFinalizadora," & _
                        "PKId_TBOperacao_caixa " & _
                   "FROM TBOperacao_caixa " & _
             "INNER JOIN TBFinalizadora " & _
                     "ON TBOperacao_caixa.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
             "LEFT  JOIN TBCupom " & _
                     "ON TBOperacao_caixa.DFNumero_Cupom_TBOperacao_caixa = TBCupom.PKId_TBCupom " & _
             "LEFT  JOIN TBCliente " & _
                     "ON TBCupom.DFEmitente_TBCupom = TBCliente.IXCodigo_TBCliente " & _
                  "WHERE TBOperacao_caixa.FKCodigo_TBOperadores_ecf = " & txtOperador.Text & " " & _
                    "AND DFData_TBOperacao_caixa = '" & Format(txtData.Text, "YYYYMMDD") & "' " & _
                    "AND TBOperacao_caixa.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & " " & _
                    "AND (TBCliente.IXCodigo_TBEmpresa = TBOperacao_caixa.FKCodigo_TBEmpresa OR TBCliente.IXCodigo_TBEmpresa IS NULL) " & _
                    "AND TBFinalizadora.IXCodigo_TBFinalizadora = '" & objParam.Cod_Finalizadora & "' " & _
               "ORDER BY TBOperacao_caixa.DFData_TBOperacao_caixa, TBOperacao_caixa.DFHora_TBOperacao_caixa "
    
        Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgConferencia, "300,700,1700,1100,1350,0,1050,700,2340,0,0", " ,Código,Finalizadora,Cupom,Valor,Valor,Hora Saída,Código,Cliente,Debito_credito,IDCupom", "BDRetaguarda", "Otica", Me, "N"
    
        If hfgConferencia.TextMatrix(1, 0) = Empty Then
            hfgConferencia.Rows = 2
            Movimentacoes.Monta_HFlex_Grid hfgConferencia, "300,700,1700,1100,1350,0,1050,700,2340,0,0", " ,Código,Finalizadora,Cupom,Valor,Valor,Hora Saída,Código,Cliente,Debito_credito,IDCupom", 11, "OTICA", Me
        Else
            For intContador = 1 To hfgConferencia.Rows - 1
                hfgConferencia.Row = intContador
                hfgConferencia.TextMatrix(intContador, 5) = Format(hfgConferencia.TextMatrix(intContador, 5), "##,##0.00")
                lblTotal.Caption = Format(CDbl(lblTotal.Caption) + CDbl(hfgConferencia.TextMatrix(intContador, 5)), "##,##0.00")
            Next
        End If
        
    Else
        
        hfgConferencia.Rows = 2
        Movimentacoes.Monta_HFlex_Grid hfgConferencia, "300,700,1700,1100,1350,0,1050,700,2340,0,0", " ,Código,Finalizadora,Cupom,Valor,Valor,Hora Saída,Código,Cliente,Debito_credito,IDCupom", 11, "OTICA", Me
    
'        'Interface de fechamento preenche o grid...
'        Call frmFechamento_caixa_posto.Preenche_Lista_Itens_Finalizadora_Existente(hfgConferencia)
        
        For intContador = 1 To hfgConferencia.Rows - 1
            hfgConferencia.Row = intContador
            hfgConferencia.Col = 0
            hfgConferencia.CellBackColor = &H80FFFF
            hfgConferencia.Col = 1
            hfgConferencia.CellForeColor = &H8000&
            hfgConferencia.CellFontBold = True
            
            hfgConferencia.TextMatrix(intContador, 5) = Format(hfgConferencia.TextMatrix(intContador, 5), "##,##0.00")
            lblTotal.Caption = Format(CDbl(lblTotal.Caption) + CDbl(hfgConferencia.TextMatrix(intContador, 5)), "##,##0.00")
        Next
    
    End If
    
    hfgConferencia.ColAlignment(1) = 5
    hfgConferencia.ColWidth(0) = 480
    hfgConferencia.Col = 0
    hfgConferencia.Row = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'SETANDO O NÚMERO DE LINHAS PARA DOIS
'    hfgConferencia.Rows = 2
End Sub

Private Sub hfgConferencia_Click()
    If hfgConferencia.Col = 1 Then
       If hfgConferencia.Text = Empty Then
          hfgConferencia.CellForeColor = &H8000&
          hfgConferencia.CellFontBold = True
          hfgConferencia.Text = "X"
       ElseIf hfgConferencia.CellForeColor = &HC0& Then
          hfgConferencia.CellForeColor = &H8000&
          hfgConferencia.CellFontBold = True
          hfgConferencia.Text = "X"
          hfgConferencia.Col = 5
          lblTotal.Caption = Format(CDbl(lblTotal.Caption) + CDbl(hfgConferencia.Text), "#,###0.000")
          hfgConferencia.Col = 1
       Else
          hfgConferencia.CellForeColor = &HC0&
          hfgConferencia.CellFontBold = True
          hfgConferencia.Text = "X"
          hfgConferencia.Col = 5
          lblTotal.Caption = Format(CDbl(lblTotal.Caption) - CDbl(hfgConferencia.Text), "#,###0.000")
          hfgConferencia.Col = 1
       End If
    ElseIf hfgConferencia.Col <> 0 Then
       txtCupom.Text = Empty
       txtCliente.Text = Empty
       txtNome_Cliente.Text = Empty
       txtValor.Text = Empty
    End If
End Sub

Private Sub hfgConferencia_DblClick()
    If hfgConferencia.Col = 0 Then
       hfgConferencia.Col = 4
       txtCupom.Text = hfgConferencia.Text
       hfgConferencia.Col = 8
       txtCliente.Text = hfgConferencia.Text
       hfgConferencia.Col = 9
       txtNome_Cliente.Text = hfgConferencia.Text
       hfgConferencia.Col = 5
       txtValor.Text = hfgConferencia.Text
       intLinha = hfgConferencia.Row
       hfgConferencia.Col = 0
       txtValor.SetFocus
    Else
       txtCupom.Text = Empty
       txtCliente.Text = Empty
       txtNome_Cliente.Text = Empty
       txtValor.Text = Empty
    End If
End Sub

Private Sub hfgConferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgConferencia_Click
    End If
End Sub

Private Sub txtValor_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_LostFocus()
    If IsNumeric(txtValor.Text) = False Then txtValor.Text = Empty
    txtValor.Text = Format(txtValor.Text, "##,##0.00")
End Sub
