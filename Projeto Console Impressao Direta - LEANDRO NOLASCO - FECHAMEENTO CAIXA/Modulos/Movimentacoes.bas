Attribute VB_Name = "Movimentacoes"
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim strID_menu As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim multempresa As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log

Function Estoque_Grade_Altera(Estoque_Atual_Grade As String, Estoque_Temporario_Grade As String, ID_PRODUTO As String, ID_Grade_Categoria As String, Form As Object, Valor_Campo_Codigo_Empresa As Integer, Adiciona_Subtrai As String, cnConexao_Aberta As Object, Optional booErro As Boolean)

    On Error GoTo Erro_Transacao
    
    Dim strSql As String
    
    booErro = False
       
    strSql = Empty
    strSql = "UPDATE TBEstoque_grade "

    If Adiciona_Subtrai = "Subtrai" Then
       strSql = strSql & "SET DFEstoque_TBEstoque_grade = " & (Funcoes_Gerais.Grava_Moeda(CDbl(Estoque_Atual_Grade) - CDbl(Estoque_Temporario_Grade))) & " "
    Else
       strSql = strSql & "SET DFEstoque_TBEstoque_grade = " & (Funcoes_Gerais.Grava_Moeda(CDbl(Estoque_Atual_Grade) + CDbl(Estoque_Temporario_Grade))) & " "
    End If

    strSql = strSql & "WHERE FKId_TBProduto = '" & ID_PRODUTO & "' " & _
                      "AND FKId_TBGrade_categoria = '" & ID_Grade_Categoria & "' " & _
                      "AND FKCodigo_TBEmpresa = '" & Valor_Campo_Codigo_Empresa & "' "

    'ALTERANDO REGISTRO NA TBGRADE_CATEGORIA
    cnConexao_Aberta.CNconexao.Execute strSql
       
    DoEvents
    
    Exit Function
    
Erro_Transacao:
    
    'MARCANDO ERRO NA TRANSAÇÃO
    Form.booErro = True
    
    'REALIZANDO ROLLBACK NA TRANSAÇÃO
    cnConexao_Aberta.CNconexao.RollbackTrans
       
    Err.Raise vbObjectError + 1313, "Estoque", "Erro ao manipular Estoque - Adição Grade Alterar"
    
    Exit Function
End Function

Function Estoque_Grade_Insere(Estoque_Temporario_Grade As String, ID_PRODUTO As String, ID_Grade_Categoria As String, Form As Object, Valor_Campo_Codigo_Empresa As Integer, Adiciona_Subtrai As String, cnConexao_Aberta As Object, Optional booErro As Boolean)
    On Error GoTo Erro_Transacao
    
    Dim strSql  As String
    
    booErro = False
        
    strSql = Empty
    strSql = "INSERT INTO TBEstoque_grade " & _
             "(FKId_TBProduto," & _
             "FKId_TBGrade_categoria," & _
             "DFEstoque_TBEstoque_grade," & _
             "FKCodigo_TBEmpresa) " & _
             "VALUES('" & ID_PRODUTO & "'," & _
             "'" & ID_Grade_Categoria & "',"
    
    If Adiciona_Subtrai = "Subtrai" Then
       strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda((CDbl(Estoque_Temporario_Grade) * (-1))) & ","
    Else
       strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda(CDbl(Estoque_Temporario_Grade)) & ","
    End If
    
    strSql = strSql & "'" & Valor_Campo_Codigo_Empresa & "')"

    'ALTERANDO REGISTRO NA TBGRADE_CATEGORIA
    cnConexao_Aberta.CNconexao.Execute strSql
    
    DoEvents
    
    Exit Function
    
Erro_Transacao:

    'MARCANDO ERRO NA TRANSAÇÃO
    Form.booErro = True

    'REALIZANDO ROLLBACK NA TRANSAÇÃO
    cnConexao_Aberta.CNconexao.RollbackTrans
    
    Err.Raise vbObjectError + 1313, "Estoque", "Erro ao manipular Estoque - Adição Grade Inserir"
    
    Exit Function
End Function

Function Movimenta_Data_Grid(strSql As String, datagrid As Object, tamanho_colunas As String, caption_campos As String, Banco As String, Aplicacao As String, Form As Object)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: PDV Only Tech                                                  '
' Equipe Responsável.....: Onlytech                                                       '
' Data da criação........: 24/02/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Public Property Let --> Recebe um valor para a propriedade     '
'                          (em run time)                                                  '
'                          Public Property Get --> Retorna para a aplicação um valor      '
'                                                                                         '
'                         *Esta Classe movimenta o controle datagrid. Cuidados Especiais  '
'                          Apos o uso não esquecer de descarregar a                       '
'                                                                                         '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim rstgrid As New ADODB.Recordset
    Dim vdatagrid As Object
    Dim Contador As Integer
    Dim Colunas As Integer
    Dim matriz_tamanho() As String
    Dim matriz_caption() As String
    Dim conexao_grid As New DLLConexao_Sistema.Conexao
    
    'On Error GoTo Erro
    
    Set vdatagrid = datagrid
    
    'Indicando o banco à conectar-se
    conexao_grid.Initial_Catalog = Banco
    
    conexao_grid.Abrir_conexao (Aplicacao)
       
    rstgrid.CursorLocation = adUseClient
    rstgrid.Open strSql, conexao_grid.CNconexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set vdatagrid = datagrid
    Set vdatagrid.DataSource = rstgrid
    
    'Montando a matriz
    matriz_tamanho = Split(tamanho_colunas, ",")
    matriz_caption = Split(caption_campos, ",")
    
    Contador = 0
    Colunas = 0
    
    'Montando as características do grid - Caption e Tamanho
    Do While matriz_caption(Contador) <> Empty
        vdatagrid.Columns(Colunas).Caption = matriz_caption(Contador)
        vdatagrid.Columns(Colunas).width = Val(matriz_tamanho(Contador))
        Contador = Contador + 1
        Colunas = Colunas + 1
        If Contador > UBound(matriz_caption) Then Exit Do
    Loop
   
   ' Exit Function
    
'Erro:
'    Call Erro.Erro(Form, Aplicacao, "Movimenta_Data_Grid")
    
End Function

Function Movimenta_HFlex_Grid(strSql As String, HFlexGrid As Object, tamanho_colunas As String, caption_campos As String, Banco As String, Aplicacao As String, Form As Object, Optional Controle_focu As String, Optional ByVal Casas_Decimais As Integer = 3)

    On Error GoTo Erro

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: PDV Only Tech                                                  '
' Equipe Responsável.....: Only Tech                                                      '
' Data da criação........: 05/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............:                                                                '
'                                                                                         '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim rstHFlexgrid As New ADODB.Recordset
    Dim Contador As Integer
    Dim Colunas As Integer
    Dim matriz_tamanho() As String
    Dim matriz_caption() As String
    Dim contador_colunas As Integer
    Dim I As Integer
    Dim conexao_grid As New DLLConexao_Sistema.Conexao
    
 '   On Error GoTo Erro_HFlexgrid
        
    'Indicando o banco à conectar-se
    conexao_grid.Initial_Catalog = Banco
    
    conexao_grid.Abrir_conexao (Aplicacao)
       
    rstHFlexgrid.CursorLocation = adUseClient
    rstHFlexgrid.Open strSql, conexao_grid.CNconexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    'INSERIDO AQUI PARA SOLUCIONAR UM PROBLEMA DE ÚLTIMO ITEM EXCLUIDO E CONTINUA A APARECER NO GRID 16/06/2004
    If rstHFlexgrid.BOF = True And rstHFlexgrid.EOF = True Then HFlexGrid.Clear: Exit Function
    
    'Marreta para acertar bug do Hflex grid
    'Não tirar o IF abaixo dessa posição. Se tirar vão ocorrer erros no FlexGrid. (Giordano e Marcos).
    
    If rstHFlexgrid.RecordCount = 0 Then
       Set rstHFlexgrid = Nothing
       Exit Function
    End If
    
    Set HFlexGrid.DataSource = rstHFlexgrid
   
    HFlexGrid.Clear
         
    'Montando a matriz
    matriz_tamanho = Split(tamanho_colunas, ",")
    matriz_caption = Split(caption_campos, ",")
    
    Contador = 0
    Colunas = 1
    
    HFlexGrid.Cols = rstHFlexgrid.Fields.Count + 1
    
    HFlexGrid.ColWidth(0) = 480
    
    'Montando as características do cabeçalho do MSHFlexgrid - Caption e Tamanho
    Do While matriz_caption(Contador) <> Empty
        DoEvents
        HFlexGrid.Row = 0
        HFlexGrid.Col = Colunas
        HFlexGrid.FixedAlignment(Colunas) = 2
        HFlexGrid.Font.Name = "Tahoma"
        HFlexGrid.Text = matriz_caption(Contador)
        HFlexGrid.ColWidth(Colunas) = Val(matriz_tamanho(Contador))
        Contador = Contador + 1
        Colunas = Colunas + 1
        If Contador > UBound(matriz_caption) Then Exit Do
    Loop
    
    HFlexGrid.Rows = rstHFlexgrid.RecordCount + 1
    
    'Essa Mudança não pode ser retirada, senão o produto nao funcionara
    If rstHFlexgrid.EOF <> True And rstHFlexgrid.BOF <> True Then
       rstHFlexgrid.MoveFirst
    End If
    
    contador_colunas = 1
    Linhas = 1
    I = 0
    
    Do While Linhas <= rstHFlexgrid.RecordCount
       DoEvents
       HFlexGrid.Row = Linhas
       HFlexGrid.Col = 0
       HFlexGrid.CellBackColor = &H80FFFF
       HFlexGrid.CellFontBold = False
       HFlexGrid.CellFontSize = 7
       HFlexGrid.Text = Linhas
       Do While contador_colunas <= rstHFlexgrid.Fields.Count
          HFlexGrid.Col = contador_colunas
          'Essse if abaixo foi incluido para identificar os campos "Booleanos" e
          'mover "Sim" ou "Não" para o Grid ao inves de "True" ou "False".(Giordano Vilela)
          If rstHFlexgrid.Fields.Item(I).Type = adBoolean Then
             If rstHFlexgrid.Fields(I).Value = False Then
                HFlexGrid.Text = "Não"
             Else
                HFlexGrid.Text = "Sim"
             End If
          Else
              'if inserido para atribyir ao campo NULL espaço para não dar problema no HFlexgrid
              If IsNull(rstHFlexgrid.Fields(I).Value) Then
                 HFlexGrid.Text = " "
              Else
                
                '3 casas sempre foi o padrão da função...
                If Casas_Decimais <> 3 Then
                
                    If rstHFlexgrid.Fields.Item(I).Type = adCurrency Then
                       HFlexGrid.Text = Format(rstHFlexgrid.Fields(I).Value, "##," & String(Casas_Decimais, "#") & "0." & String(Casas_Decimais, "0"))
                    Else
                       HFlexGrid.Text = rstHFlexgrid.Fields(I).Value
                    End If
                
                Else
                
                    If rstHFlexgrid.Fields.Item(I).Type = adCurrency Then
                       HFlexGrid.Text = Format(rstHFlexgrid.Fields(I).Value, "#,###0.000")
                    Else
                       HFlexGrid.Text = rstHFlexgrid.Fields(I).Value
                    End If
                    
                End If
              End If
          End If
          contador_colunas = contador_colunas + 1
          I = I + 1
       Loop
       rstHFlexgrid.MoveNext
       I = 0
       contador_colunas = 1
       Linhas = Linhas + 1
    Loop
      
    Set rstHFlexgrid = Nothing
    
    HFlexGrid.Row = 1
    HFlexGrid.Col = 0
    
    If Controle_focu = "" Then Controle_focu = "S"
    If Controle_focu = "S" Then HFlexGrid.SetFocus
    
''''INSERIDO AQUI POR RAFAEL PARA RESOLVER O PROBLEMA DE ALINHAMENTO DO GRID 18/10/2005''''''''''''''''''''
    I = HFlexGrid.Cols - 1
    
    Do While I <> 0
       HFlexGrid.Col = I
       HFlexGrid.Row = 1
          
       If IsNumeric(HFlexGrid.Text) Then
          HFlexGrid.ColAlignment(I) = 7  'ALINHAMENTO A DIREITA - 678
       Else
          HFlexGrid.ColAlignment(I) = 1  'ALINHAMENTO A ESQUERDA - 123
       End If
       
       I = I - 1
    Loop
    
    HFlexGrid.ColAlignment(0) = 7
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  '  Exit Function
    
'Erro_HFlexgrid:

 '   Call Erro.Erro(Form, "Otica", "Movimenta_HFlex_Grid")
 
    Exit Function

Erro:
    MsgBox Err.Number & " - " & Err.Description, vbCritical + vbOKOnly, "Only Tech"
    Exit Function
    Resume
    
End Function

Function Movimenta_HFlex_GridII(strSql As String, HFlexGrid As Object, tamanho_colunas As String, caption_campos As String, Banco As String, Aplicacao As String, Form As Object)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: PDV Only Tech                                                  '
' Equipe Responsável.....: Only Tech                                                      '
' Data da criação........: 05/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............:                                                                '
'                                                                                         '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim rstHFlexgrid As New ADODB.Recordset
    Dim Contador As Integer
    Dim Colunas As Integer
    Dim matriz_tamanho() As String
    Dim matriz_caption() As String
    Dim contador_colunas As Integer
    Dim I As Integer
    Dim conexao_grid As New DLLConexao_Sistema.Conexao
    Dim C As Integer

  '  On Error GoTo Erro_HFlexgrid

    'Indicando o banco à conectar-se
    conexao_grid.Initial_Catalog = Banco

    conexao_grid.Abrir_conexao (Aplicacao)

    rstHFlexgrid.CursorLocation = adUseClient
    rstHFlexgrid.Open strSql, conexao_grid.CNconexao, adOpenKeyset, adLockOptimistic, adCmdText

    Set HFlexGrid.DataSource = rstHFlexgrid
    HFlexGrid.Clear

    'Montando a matriz
    matriz_tamanho = Split(tamanho_colunas, ",")
    matriz_caption = Split(caption_campos, ",")

    Contador = 0
    Colunas = 1

    HFlexGrid.Cols = rstHFlexgrid.Fields.Count + 1

    HFlexGrid.ColWidth(0) = 300

    'Montando as características do cabeçalho do MSHFlexgrid - Caption e Tamanho
    Do While matriz_caption(Contador) <> Empty
        HFlexGrid.Row = 0
        HFlexGrid.Col = Colunas
        HFlexGrid.FixedAlignment(Colunas) = 2
        HFlexGrid.Font.Name = "Tahoma"
        HFlexGrid.Text = matriz_caption(Contador)
        HFlexGrid.ColWidth(Colunas) = Val(matriz_tamanho(Contador))
        Contador = Contador + 1
        Colunas = Colunas + 1
        If Contador > UBound(matriz_caption) Then Exit Do
    Loop

    HFlexGrid.Rows = rstHFlexgrid.RecordCount + 1

    Linhas = 1
    
    I = 0
    C = 0
    Dim Teste As String
    
    'Essa Mudança não pode ser retirada, senão o produto nao funcionara
    If rstHFlexgrid.EOF <> True And rstHFlexgrid.BOF <> True Then
       rstHFlexgrid.MoveFirst
    End If
    
    Dim strLinha As String
    
    Do While Linhas <= rstHFlexgrid.RecordCount
       DoEvents
       Do While C < rstHFlexgrid.Fields.Count
          'strRecebe = rstHFlexgrid.Fields.Item(C).Value
          'vetNome(C) = strRecebe & "ø"
          If IsNull(rstHFlexgrid.Fields.Item(C).Value) Then
              rstHFlexgrid.Fields.Item(C).Value = " "
          End If
          If C = rstHFlexgrid.Fields.Count - 1 Then
            strLinha = strLinha & rstHFlexgrid.Fields.Item(C).Value
          Else
            strLinha = strLinha & rstHFlexgrid.Fields.Item(C).Value & " + Chr$(9) + "
          End If
          C = C + 1
       Loop
       Teste = Linhas
       strLinha = strLinha
       HFlexGrid.AddItem Teste + Chr(9) + strLinha, Teste
       HFlexGrid.Refresh
       HFlexGrid.Row = Linhas
       HFlexGrid.Col = 0
       HFlexGrid.CellBackColor = &H80FFFF
       HFlexGrid.CellFontBold = False
       HFlexGrid.CellFontSize = 7
       HFlexGrid.Text = Linhas
       strLinha = Empty
       Linhas = Linhas + 1
       C = 0
       If rstHFlexgrid.EOF <> True Then rstHFlexgrid.MoveNext
    Loop

    Set rstHFlexgrid = Nothing

   ' Exit Function

'Erro_HFlexgrid:

 '   Call Erro.Erro(Form, "Otica", "Movimenta_HFlex_GridII")
    
End Function

Public Function Movimenta_DataCombo(Nome_Campo_Codigo As String, Nome_Campo_Descricao As String, DataCombo As Object, String_Sql As String, Banco As String, Aplicacao As String, Form As Object, Optional Nome_Campo_Ordenacao As String) As String

    Dim rstCombo As New ADODB.Recordset
    Dim vDatacombo As Object
    Dim conexao_combo As New DLLConexao_Sistema.Conexao
    
  '  On Error GoTo Erro
    
    'Indicando o banco à conectar-se
    conexao_combo.Initial_Catalog = Banco
    
    conexao_combo.Abrir_conexao (Aplicacao)
    
    DoEvents
    rstCombo.CursorLocation = adUseClient
    
    If Nome_Campo_Ordenacao <> "" Then
       String_Sql = String_Sql & " ORDER BY " & Nome_Campo_Ordenacao
    Else
       String_Sql = String_Sql & " ORDER BY " & Nome_Campo_Descricao
    End If
    
    rstCombo.Open String_Sql, conexao_combo.CNconexao, adOpenStatic, adLockReadOnly
                
    Set vDatacombo = DataCombo
    Set vDatacombo.DataSource = rstCombo
    
    Set vDatacombo.RowSource = rstCombo
        vDatacombo.ListField = Nome_Campo_Descricao
        'BoundColumn -> sendo usado para retornar no TextBox o valor pedido, neste caso
        '               Nome_Campo_Codigo
        vDatacombo.BoundColumn = Nome_Campo_Codigo
    
    Set rstCombo.ActiveConnection = Nothing
    
    conexao_combo.Fechar_conexao
    
    String_Sql = Empty
    
   ' Exit Function
    
'Erro:

 '   Call Erro.Erro(Form, Aplicacao, "Movimenta Combo")
    
End Function

Public Function Verifica_Numero(Nome_Campo As String, Nome_Tabela As String, Nome_textbox As Object, Aplicacao As String, Form As Object, Optional Nome_Campo_Codigo_Empresa As String, Optional Valor_Campo_Codigo_Empresa As Integer, Optional Controle_focu As String) As Boolean

    Dim conexao_verifica As New DLLConexao_Sistema.Conexao
    Dim rstVerificacao As New ADODB.Recordset
    
    If Nome_textbox.Text = Empty Then
       Exit Function
    End If
    
    Dim strSql As String
    
    On Error GoTo Erro
    
    strSql = Empty
    strSql = "SELECT " & Nome_Campo & " " & _
             "FROM " & Nome_Tabela & " " & _
             "WHERE " & Nome_Campo & " = '" & Nome_textbox.Text & "' "
             
    If Nome_Campo_Codigo_Empresa <> "" Then
       strSql = strSql & "AND " & Nome_Campo_Codigo_Empresa & " = " & Valor_Campo_Codigo_Empresa & ""
    End If

    conexao_verifica.Abrir_conexao (Aplicacao)
    
    rstVerificacao.CursorLocation = adUseClient
    rstVerificacao.Open strSql, conexao_verifica.CNconexao, adOpenStatic, adLockReadOnly
    If rstVerificacao.EOF = True And rstVerificacao.BOF = True Then
        Verifica_Numero = False
    Else
      If Val(rstVerificacao(Nome_Campo)) = Val(Nome_textbox.Text) Then
          MsgBox "Registro já existente nesta empresa.", vbCritical, "Only Tech"
          If Controle_focu = "S" Or Controle_focu = "" Then
             Nome_textbox.Text = Empty
             Nome_textbox.SetFocus
          End If
          Verifica_Numero = True
      Else
          Verifica_Numero = False
      End If
      
    End If
    Exit Function
    
Erro:

    If Err.Number = 3021 Then
        Verifica_Numero = False
        Exit Function
    Else
        Call Erro.Erro(Form, Aplicacao, "Verifica Registro")
    End If
    Nome_textbox.SetFocus
    
End Function

Public Function Select_geral(String_Sql As String, Banco As String, recordset_aplicacao As ADODB.Recordset, Aplicacao As String, Form As Object, Optional Usuario As String, Optional Senha As String)

    Dim conexao_select As New DLLConexao_Sistema.Conexao
    
  '  On Error GoTo Erro
    
    'Trecho inserido aqui porque é necessario que esta classe tenha a possibilidade de se conectar
    'a mais de um banco no SQL Server
    conexao_select.Initial_Catalog = Banco
    
    If Usuario <> "" Then
       conexao_select.User_ID = Usuario
    End If
    
    If Senha <> "" Then
       conexao_select.Password = Senha
    End If
        
    'Estabelecendo conexão com o banco
    Call conexao_select.Abrir_conexao(Aplicacao)
    
    DoEvents
    recordset_aplicacao.CursorLocation = adUseClient
    
    recordset_aplicacao.Open String_Sql, conexao_select.CNconexao, adOpenStatic, adLockBatchOptimistic
     
    'Desconectando Recordset
    recordset_aplicacao.ActiveConnection = Nothing
    
    'Fecha a conexão com o banco
    conexao_select.Fechar_conexao

   ' Exit Function

'Erro:

    'Fecha a conexão com o banco
 '   conexao_select.Fechar_conexao

  '  Call Erro.Erro(Form, Aplicacao, "Movimentações")
    
End Function

Function Monta_HFlex_Grid(HFlexGrid As Object, tamanho_colunas As String, caption_campos As String, Quant_Campos As Integer, Aplicacao As String, Form As Object)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Otica/PDV Only Tech                                            '
' Equipe Responsável.....: Only Tech                                                      '
' Data da criação........: 18/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............:                                                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Contador As Integer
    Dim Colunas As Integer
    Dim matriz_tamanho() As String
    Dim matriz_caption() As String
    Dim contador_colunas As Integer
    Dim I As Integer
        
   ' On Error GoTo Erro_grid
    
    HFlexGrid.Clear
 
    'Montando a matriz
    matriz_tamanho = Split(tamanho_colunas, ",")
    matriz_caption = Split(caption_campos, ",")
    
    Contador = 0
    Colunas = 1
    
    HFlexGrid.Cols = Quant_Campos + 1
    HFlexGrid.ColWidth(0) = 200
    
    'Montando as características do cabeçalho do MSHFlexgrid - Caption e Tamanho
    Do While matriz_caption(Contador) <> Empty
        DoEvents
        HFlexGrid.Row = 0
        HFlexGrid.Col = Colunas
        HFlexGrid.Font.Name = "Tahoma"
        HFlexGrid.Text = matriz_caption(Contador)
        HFlexGrid.ColWidth(Colunas) = Val(matriz_tamanho(Contador))
        Contador = Contador + 1
        Colunas = Colunas + 1
        If Contador > UBound(matriz_caption) Then Exit Do
    Loop
        
    'Exit Function
    
'Erro_grid:

 '   Call Erro.Erro(Form, "Otica", "Monta_HFlex_Grid")
        
End Function
Public Function Acessibilidade_inicio(Caption_form As String, botao_consultar As Object, botao_refresh As Object, SSTab_form As Object, Variavel_Incluir As Boolean, Variavel_Alterar As Boolean, Variavel_Excluir As Boolean, Variavel_Consultar As Boolean, Codigo_Usuario As Long, barra_ferramentas As Object, Form As Object, Aplicacao As String, Banco As String)

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Only Tech                                                                               '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Sistema................: Retaguarda                                                     '
        ' Equipe Responsável.....: Only Tech                                                      '
        ' Data da criação........: 10/07/2004                                                     '
        ' Data última manutenção.:                                                                '
        ' Observação.............: Função desenvolvida para trabalhar em conjunto com o módulo    '
        '                          movimentacoes.acessibilidade_inicio                            '
        '                          da DLLSystem_manager.Acessibilidade e será usada no Activate do'
        '                          form, será o primeiro teste de acessibilidade do form          '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        'Verificando acessibilidade
        'Tenho a descricao(caption do form) do programa e indico pra conseguir o ID
        acesso.Consulta_ID_programa Caption_form, Aplicacao, Banco, strID_Acessibilidade
        
        'Indicar o usuário a verificar sua acessibilidade
        acesso.Codigo_Usuario = Codigo_Usuario
    
        Set rstAplicacao = acesso.Verifica_acessibilidade(strID_Acessibilidade, Aplicacao, Banco)
        
        'Abastecendo a OCX do usuário para ser usada pelo restante da execução em Run Time
        Variavel_Alterar = rstAplicacao!DFAlteracao_TBAcessibilidade
        Variavel_Consultar = rstAplicacao!DFConsulta_TBAcessibilidade
        Variavel_Excluir = rstAplicacao!DFExclusao_TBAcessibilidade
        Variavel_Incluir = rstAplicacao!DFInclusao_TBAcessibilidade
         
        If Not rstAplicacao.EOF And Not rstAplicacao.BOF Then
            SSTab_form = 1
            'Habilita somente as operações que o usuário tem permissão
            'Novo
            barra_ferramentas.Buttons.Item(1).Enabled = Variavel_Incluir
            'Gravar
            barra_ferramentas.Buttons.Item(2).Enabled = False
            'Cancelar
            barra_ferramentas.Buttons.Item(3).Enabled = False
            'Excluir
            barra_ferramentas.Buttons.Item(4).Enabled = False
            'Imprimir
            barra_ferramentas.Buttons.Item(5).Enabled = Variavel_Consultar
            
            If Variavel_Consultar = False Then
               botao_consultar.Enabled = False
               botao_refresh.Enabled = False
            End If
         Else
            botao_consultar.Enabled = False
            botao_refresh.Enabled = False
            MsgBox "Usuário não tem privilégio para esta operação Verifique com o administrador do sistema!", vbInformation, "Only Tech"
            log.Descricao = "Operação cancelada por falta de privilégio do usuário - " & Caption_form & " "
            'Gravando o log
            log.Gravar_log "Otica", Form
            Set rstAplicacao = Nothing
            'Finaliza o executável
           ' End
         End If
         Set rstAplicacao = Nothing
    
End Function
Public Function Acessibilidade_inicio_relatorios(Caption_form As String, OCXUsuario_form As Object, Form As Object, Aplicacao As String, Banco As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Otica/Only Tech                                                '
' Equipe Responsável.....: Only Tech                                                      '
' Data da criação........: 30/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o método da '
'                          da DLLSystem_manager.Acessibilidade e será usada no Activate do'
'                          form, será o primeiro teste de acessibilidade do form          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '
   '     On Error GoTo Erro_acessibilidade
    
        'Verificando acessibilidade
        'Tenho a descricao(caption do form) do programa e indico pra conseguir o ID
        acesso.Consulta_ID_programa Caption_form, Aplicacao, Banco, strID_Acessibilidade
        'Indicar o usuário a verificar sua acessibilidade
        acesso.Codigo_Usuario = OCXUsuario_form.Codigo
    
        Set rstAplicacao = acesso.Verifica_acessibilidade(strID_Acessibilidade, Aplicacao, Banco)
        
        'Abastecendo a OCX do usuário para ser usada pelo restante da execução em Run Time
        OCXUsuario_form.PrivilégioAcessar = rstAplicacao!DFAcesso_TBAcessibilidade
        OCXUsuario_form.PrivilégioAlterar = rstAplicacao!DFAlteracao_TBAcessibilidade
        OCXUsuario_form.PrivilégioConsultar = rstAplicacao!DFConsulta_TBAcessibilidade
        OCXUsuario_form.PrivilégioExcluir = rstAplicacao!DFExclusao_TBAcessibilidade
        OCXUsuario_form.PrivilégioIncluir = rstAplicacao!DFInclusao_TBAcessibilidade
        
        If rstAplicacao.EOF And rstAplicacao.BOF And OCXUsuario_form.PrivilégioAcessar = True Then
           MsgBox "Usuário não tem privilégio para esta operação Verifique com o administrador do sistema!", vbInformation, "Only Tech"
           log.Descricao = "Operação cancelada por falta de privilégio do usuário - " & Caption_form & " "
           'Gravando o log
           log.Gravar_log "Otica", Form
           Set rstAplicacao = Nothing
           'Finaliza o executável
           'End
        End If
        Set rstAplicacao = Nothing
    
    Exit Function
    
'Erro_acessibilidade:

    'Call Erro.Erro(Form, "Otica", "Acessibilidade_inicio_relatorios")
      
End Function

Public Function Refresh_Grid(FlexGrid_Form As Object, Campos_Query As String, Nome_Tabela As String, FlexGrid_Linha_Atualizar As Integer, FlexGrid_Colunas_Atualizar As String, Indice_Fields As String, Contador As Integer, Aplicacao As String, Banco As String, Form As Object, Optional Clausula_WHERE As String, Optional Clausula_INNER_JOIN As String)
    
    Dim Conexao_Refresh As New DLLConexao_Sistema.Conexao
    Dim rstRefresh As New ADODB.Recordset
    Dim Vetor_Colunas() As String
    Dim Vetor_Indice_Fields() As String
    Dim I As Integer
    Dim strSql As String
    
    Vetor_Colunas = Split(FlexGrid_Colunas_Atualizar, ",")
    Vetor_Indice_Fields = Split(Indice_Fields, ",")
              
    Conexao_Refresh.Initial_Catalog = Banco
    Conexao_Refresh.Abrir_conexao (Aplicacao)
    DoEvents
                 
    strSql = ""
      
    strSql = "SELECT " & Campos_Query & " " & _
             "FROM " & Nome_Tabela & " "
             
    If Clausula_INNER_JOIN <> Empty Then
       strSql = strSql & " INNER JOIN " & Clausula_INNER_JOIN & " "
    End If
    
    If Clausula_WHERE <> Empty Then
       strSql = strSql & " WHERE " & Clausula_WHERE & " "
    End If
        
    rstRefresh.CursorLocation = adUseClient
    rstRefresh.Open strSql, Conexao_Refresh.CNconexao, adOpenStatic, adLockBatchOptimistic
    
    If rstRefresh.RecordCount <> 0 Then
       FlexGrid_Form.Row = FlexGrid_Linha_Atualizar
       For I = 0 To Contador - 1
          FlexGrid_Form.Col = Vetor_Colunas(I)
          If rstRefresh.Fields.Item(Val(Vetor_Indice_Fields(I))).Type = adCurrency Then
             FlexGrid_Form.Text = Format(rstRefresh.Fields.Item(Val(Vetor_Indice_Fields(I))).Value, "#,###0.00")
          Else
             FlexGrid_Form.Text = rstRefresh.Fields.Item(Val(Vetor_Indice_Fields(I))).Value
          End If
       Next I
    End If
              
    FlexGrid_Form.Row = FlexGrid_Linha_Atualizar
    FlexGrid_Form.Col = 0
    FlexGrid_Form.SetFocus
    
    FlexGrid_Form.Refresh
        
    rstRefresh.ActiveConnection = Nothing
    Conexao_Refresh.Fechar_conexao
    
    Exit Function
    
End Function

Public Function Verifica_DataCombo(Valor_Data_Combo As String)
        
    If Valor_Data_Combo = Empty Then
       SendKeys "{F4}"
    Else
       SendKeys "{TAB}"
    End If

End Function

Public Function Digito_Nosso_Numero(Codigo_Banco As Long, Agencia As String, CC As String, Carteira As String, Nosso_Numero As String) As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda Only Tech                                           '
' Equipe Responsável.....: Only Tech                                                      '
' Data da criação........: 24/02/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função criada para geração automática do gig. do nosso número, '
'                          variando, de banco para banco                                  '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim strNumero_Concatenado As String
    Dim intTamanho_string_concatenado As Integer
    Dim intTamanho_string_resultado_parcial As Integer
    Dim intContador_string As Integer
    Dim dblResultado_parcial As Double
    Dim dblResultado_final As Double
    Dim dblCalculo_Digito  As Double
    Dim dblDigito  As Double
    
    '--------------------------------------------------------------------------------------------------
    'Itaú
    'Regra:
    '1 - Concatenar Agência,CC,Carteira e Nosso Número;
    '2 - Multiplicar os digitos um a um alternados por 1 e 2 da esquerda para a direita;
    '3 - Somar os produtos acima de 9 entre eles ex: 10 ---> 1 + 0;
    '4 - Dividir o resultado por 10;
    '5 - Subtrair Divisor por 10.
    '--------------------------------------------------------------------------------------------------
    
    If Codigo_Banco = 341 Then 'Itau
    
       strNumero_Concatenado = Trim(Agencia) & Trim(CC) & Trim(Carteira) & Trim(Nosso_Numero)
       intTamanho_string_concatenado = Len(strNumero_Concatenado)
       
       intContador_string = 1
       
       Do While intContador_string <= intTamanho_string_concatenado
          dblResultado_parcial = CDbl(Mid(strNumero_Concatenado, intContador_string, 1))
          'Multiplica os impares por 1
          If intContador_string Mod 2 <> 0 Then
             dblResultado_parcial = dblResultado_parcial * 1
          Else
             dblResultado_parcial = dblResultado_parcial * 2
          End If
          
          intTamanho_string_resultado_parcial = Len(CStr(dblResultado_parcial))
          'Soma os digitos dos produtos acima de 2
          If intTamanho_string_resultado_parcial > 1 Then
             dblResultado_parcial = CDbl(Mid(dblResultado_parcial, 1, 1)) + CDbl(Mid(dblResultado_parcial, 2, 1))
          End If
          
          dblResultado_final = dblResultado_final + dblResultado_parcial
                    
          intContador_string = intContador_string + 1
       Loop
       
       dblCalculo_Digito = dblResultado_final Mod 10
    
       dblDigito = 10 - dblCalculo_Digito
       
       If dblDigito = 10 Then dblDigito = 0
       
    End If
    
    Digito_Nosso_Numero = dblDigito
    
End Function
Public Function Acessibilidade_Menu(Caption_Menu As String, Aplicacao As String, Banco As String, Usuario As Long) As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                         '
' Equipe Responsável.....: Marcos Baião                                                   '
' Data da criação........: 30/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o método da '
'                          da DLLSystem_manager.Acessibilidade e será usada no Admin      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Verificando acessibilidade
    'Tenho a descricao(caption do form) do menu e indico pra conseguir o ID
    acesso.Consulta_ID_menu Caption_Menu, Aplicacao, Banco, strID_Acessibilidade
    'Indicar o usuário a verificar sua acessibilidade
    acesso.Codigo_Usuario = Usuario

    Set rstAplicacao = acesso.Verifica_acessibilidade_menu(strID_Acessibilidade, Aplicacao, Banco)
    
    If strID_Acessibilidade = 0 Then Acessibilidade_Menu = False: Exit Function
    
    If Not rstAplicacao.EOF And Not rstAplicacao.BOF Then
       If rstAplicacao!DFAcesso_TBAcessibilidade_Menu = True Then
          Acessibilidade_Menu = True
       Else
          Acessibilidade_Menu = False
       End If
    Else
       Acessibilidade_Menu = False
    End If
    
    Set rstAplicacao = Nothing
    
End Function

Public Function Acessibilidade_Item_Menu(Caption_Menu As String, Aplicacao As String, Banco As String, Usuario As Long, Item_Menu_Componente As Object)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                         '
' Equipe Responsável.....: Marcos Baião                                                   '
' Data da criação........: 05/07/2004                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o método da '
'                          da DLLSystem_manager.Acessibilidade e será usada no MDI        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Verificando acessibilidade
    'Tenho a descricao(caption do form) do menu e indico pra conseguir o ID
    acesso.Consulta_ID_programa Caption_Menu, Aplicacao, Banco, strID_Acessibilidade
    'Indicar o usuário a verificar sua acessibilidade
    acesso.Codigo_Usuario = Usuario
    
    If strID_Acessibilidade = 0 Then Item_Menu_Componente.Enabled = False: Exit Function
    
    Set rstAplicacao = acesso.Verifica_acessibilidade(strID_Acessibilidade, Aplicacao, Banco)
    
    If strID_Acessibilidade = 0 Then Item_Menu_Componente.Enabled = False: Exit Function
     
    If Not rstAplicacao.EOF And Not rstAplicacao.BOF Then
       If rstAplicacao!DFAcesso_TBAcessibilidade = True Then
          Item_Menu_Componente.Enabled = True
       Else
          Item_Menu_Componente.Enabled = False
       End If
    Else
       Item_Menu_Componente.Enabled = False
    End If
    
    Set rstAplicacao = Nothing
    
End Function

Private Function Montando_Acessibilidade_Programas(Caption_form As String, Variavel_Incluir As Boolean, Variavel_Alterar As Boolean, Variavel_Excluir As Boolean, Variavel_Consultar As Boolean, Aplicacao As String, Banco As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                     '
' Equipe Responsável.....: Only Tech                                                      '
' Data da criação........: 10/07/2004                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o módulo    '
'                          movimentacoes.acessibilidade_inicio                            '
'                          da DLLSystem_manager.Acessibilidade e será usada no Activate do'
'                          form, será o primeiro teste de acessibilidade do form          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Verificando acessibilidade
    'Tenho a descricao(caption do form) do programa e indico pra conseguir o ID
    acesso.Consulta_ID_programa Caption_form, Aplicacao, Banco, strID_Acessibilidade
    
    'Indicar o usuário a verificar sua acessibilidade
    acesso.Codigo_Usuario = OCXUsuario_form.Codigo

    Set rstAplicacao = acesso.Verifica_acessibilidade(strID_Acessibilidade, Aplicacao, Banco)
    
    'Abastecendo a OCX do usuário para ser usada pelo restante da execução em Run Time
    Variavel_Alterar = rstAplicacao!DFAlteracao_TBAcessibilidade
    Variavel_Consultar = rstAplicacao!DFConsulta_TBAcessibilidade
    Variavel_Excluir = rstAplicacao!DFExclusao_TBAcessibilidade
    Variavel_Incluir = rstAplicacao!DFInclusao_TBAcessibilidade
        
End Function
Public Function Consulta_Contingencia_Acessibilidade(Aplicacao As String) As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                      '
' Equipe Responsável.....: Only Tech                                                       '
' Data da criação........: 20/07/2004                                                      '
' Data última manutenção.:                                                                 '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o todos os   '
'                          módulos nos MDI's ela será útil, quando o programa não conseguir'                            '
'                          acessar o endereço de memória para colher as informações dos us_'
'                          uários logados no admin                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Registro As New DLLSystemManager.Registro
    Dim Nome_Usuario As String, Codigo_Usuario As String, Empresa_Usuario As String, Estacao_Usuario As String

    Empresa_Usuario = CStr(Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "EMPRESA"))
    Codigo_Usuario = CStr(Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "CODIGO"))
    Nome_Usuario = CStr(Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "NOME"))
    Estacao_Usuario = CStr(Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "ESTACAO"))
    
    '--------------------------------------------------------------------------------------------------
    'Monta a mensagem na mesma ordem da Intercomunicador
    '--------------------------------------------------------------------------------------------------
    Consulta_Contingencia_Acessibilidade = Estacao_Usuario & "¤" & _
                                           Nome_Usuario & "¤" & _
                                           "********" & "¤" & _
                                           Nome_Usuario & "¤" & _
                                           Codigo_Usuario & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           Empresa_Usuario
End Function
Public Function Grava_Contingencia_Acessibilidade(Estacao_Usuario As String, Nome_Usuario As String, Codigo_Usuario As Integer, Empresa_Usuario As String, Aplicacao As String) As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                      '
' Equipe Responsável.....: Only Tech                                                       '
' Data da criação........: 20/07/2004                                                      '
' Data última manutenção.:                                                                 '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o todos os   '
'                          módulos nos MDI's ela será útil, quando o programa não conseguir'                            '
'                          acessar o endereço de memória para colher as informações dos us_'
'                          uários logados no admin                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Registro As New DLLSystemManager.Registro

    'Criando chave no registro
    Registro.WinRegCriarChave "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER"
    
    'Limpando as inf. do usuário
    Call Limpa_Contingencia_Acessibilidade(Aplicacao)
    
    'Gravando as inf. do usuário
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "ESTACAO", Estacao_Usuario
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "EMPRESA", Empresa_Usuario
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "CODIGO", CInt(Codigo_Usuario)
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "NOME", Nome_Usuario
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "DATA", Format(Now, "DD/MM/YYYY")
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "HORA", Format(Now, "HH:MM:SS")
    
End Function

Public Function Limpa_Contingencia_Acessibilidade(Aplicacao As String) As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                      '
' Equipe Responsável.....: Only Tech                                                       '
' Data da criação........: 20/07/2004                                                      '
' Data última manutenção.:                                                                 '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o todos os   '
'                          módulos nos MDI's ela será útil, quando o programa não conseguir'                            '
'                          acessar o endereço de memória para colher as informações dos us_'
'                          uários logados no admin                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Registro As New DLLSystemManager.Registro

    'Limpando as inf. do usuário
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "EMPRESA", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "ESTACAO", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "CODIGO", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "NOME", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "DATA", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "HORA", ""
   
End Function

Public Function Verifica_Acesso_Usuario(DataCombo_Empresa As Object, Aplicacao As String, Banco As String, Form As Object, Optional Controle_focu As String)
        
    Dim strSql As String
    Dim rstcomparacao As New ADODB.Recordset
    Dim conexao_combo As New DLLConexao_Sistema.Conexao
           
    'INDICANDO O BANCO A CONECTAR-SE
    conexao_combo.Initial_Catalog = Banco
    
    'ABRINDO CONEXAO COM BANCO
    conexao_combo.Abrir_conexao (Aplicacao)
    
    DoEvents
    
    rstcomparacao.CursorLocation = adUseClient
    
    'STRING QUE COLETA DADOS RELATIVOS A ACESSIBILIDADE DO USUARIO
    strSql = "SELECT  DFNivel_TBUsuario FROM TBUsuario " & _
             "WHERE DFNome_TBUsuario = '" & MDIPrincipal.OCXUsuario.Nome & "'"
             
    rstcomparacao.Open strSql, conexao_combo.CNconexao, adOpenStatic, adLockReadOnly
    
    rstcomparacao.MoveFirst

    
    'VERIFICANDO NA ACESSIBILIDADE SE USUARIO PODE NAVEGAR ENTRE AS EMPRESAS
    If rstcomparacao!DFNivel_TBUsuario >= 5 Then
       DataCombo_Empresa.Enabled = True
       If Controle_focu = "S" Or Controle_focu = "" Then
          DataCombo_Empresa.SetFocus
       End If
    End If
      
    Set rstcomparacao = Nothing
    
    conexao_combo.Fechar_conexao
    
    DoEvents
       
    Exit Function
    
End Function

Public Function Gera_Caixa(Banco As String, Aplicacao As String, intCodigo_empresa As Integer, lngID_Historico As Long, datData_lancamento As Date, strComplemento As String, dblValor As Double, Optional Conexao As Object, Optional controle_transacional As String)

    Dim strSql As String
    
    If controle_transacional = "" Or controle_transacional = "N" Then
       Dim conexao_caixa As New DLLConexao_Sistema.Conexao
       'INDICANDO O BANCO A CONECTAR-SE
       conexao_caixa.Initial_Catalog = Banco
       'ABRINDO CONEXAO COM BANCO
       conexao_caixa.Abrir_conexao (Aplicacao)
    End If
        
    'MONTANDO A STRING DE CONEXÃO PARA GRAVAR O CAIXA
    strSql = Empty
    strSql = "INSERT INTO TBCaixa(FKCodigo_TBEmpresa,FKId_TBHistorico_padrao,DFData_lancamento_TBCaixa,DFComplemento_TBCaixa,DFValor_TBCaixa) " & _
             "VALUES(" & intCodigo_empresa & "," & lngID_Historico & ",'" & Format(datData_lancamento, "YYYYMMDD") & "','" & strComplemento & "'," & Funcoes_Gerais.Grava_Moeda(dblValor) & ")"
             
    If controle_transacional = "" Or controle_transacional = "N" Then
       conexao_caixa.CNconexao.Execute strSql
    Else
       Conexao.CNconexao.Execute strSql
    End If

End Function

Public Function Acessibilidade_multempresa(Caption_form As String, Caption_modulo As String, Codigo_Usuario As Long, Aplicacao As String, Banco As String, Empresa As Long) As Boolean

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Only Tech                                                                               '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Sistema................: Retaguarda                                                     '
        ' Equipe Responsável.....: Only Tech                                                      '
        ' Data da criação........: 23/12/2005                                                     '
        ' Data última manutenção.:                                                                '
        ' Observação.............: Função desenvolvida para trabalhar em conjunto com o módulo    '
        '                          movimentacoes.acessibilidade_inicio                            '
        '                          da DLLSystem_manager.Acessibilidade e será usada no Activate do'
        '                          form, será o primeiro teste de acessibilidade do Mult Empresa  '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        'Verificando acessibilidade
        'Tenho a descricao(caption do form) do programa e indico pra conseguir o ID
        multempresa.Consulta_ID_programa Caption_form, Aplicacao, Banco, strID_Acessibilidade
        
        'Indicar o usuário a verificar sua acessibilidade
        multempresa.Codigo_Usuario = Codigo_Usuario
        
        'Indicar o usuário a verificar sua acessibilidade
        multempresa.Codigo_empresa = Empresa
        
        'Indicar o menu a verificar sua acessibilidade
        multempresa.Consulta_ID_menu "Cadastros Base", "Otica", "BDRetaguarda", strID_menu
       
        Set rstAplicacao = multempresa.Verifica_acessibilidade_multempresa(strID_Acessibilidade, strID_menu, Aplicacao, Banco)
         
        If rstAplicacao.EOF = True And rstAplicacao.BOF = True Then
           Acessibilidade_multempresa = False
        Else
           Acessibilidade_multempresa = True
        End If
        
        Set rstAplicacao = Nothing
    
End Function
Public Function Acessibilidade_nivel_usuario(Caption_form As Object, Codigo_Usuario As Long, Aplicacao As String, Banco As String, Empresa As Long) As Boolean
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Only Tech                                                                               '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Sistema................: Retaguarda                                                     '
    ' Equipe Responsável.....: Only Tech                                                      '
    ' Data da criação........: 12/05/2006                                                     '
    ' Data última manutenção.:                                                                '
    ' Observação.............: Função desenvolvida para trabalhar em conjunto com os módulos  '
    '                          que utilizam a integração                                      '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim rstIntegracao As New ADODB.Recordset
    Dim strSql As String
    
    strSql = Empty
    strSql = "SELECT  DFNivel_TBUsuario FROM TBUsuario " & _
             "WHERE PKCodigo_TBUsuario = '" & Codigo_Usuario & "'" & _
             "AND FKCodigo_TBEmpresa = " & Empresa & ""
    Movimentacoes.Select_geral strSql, Banco, rstIntegracao, Aplicacao, Caption_form
    
    If rstIntegracao.BOF = False Then
       If rstIntegracao!DFNivel_TBUsuario < 5 Then
          Acessibilidade_nivel_usuario = False
       Else
          Acessibilidade_nivel_usuario = True
       End If
    End If
          
    Set rstIntegracao = Nothing

End Function

Public Function Gera_TX(Recordset As ADODB.Recordset, nome_arquivo As String, Optional Nome_campo_empresa As String, Optional cod_empresa As Long, Optional Complemento_linha As String, Optional campo_codigo_item_cii As String, Optional valor_codigo_item_cii As String, Optional campo_codigo_chave As String, Optional valor_codigo_chave As String, Optional Tabela_pai As String, Optional campo_id_atualizado_filial As String, Optional Form As Object, Optional Quantidade_FKS As Long, Optional Campos_FK_tabela_filha As String, Optional Tabela_pai_FK As String, Optional Id_Tabela_pai_FK As String, Optional CII_Tabela_Pai_FK As String)

    Dim NumArq As Integer
    Dim intContador As Long
    Dim strCampos As String
    
    NumArq = FreeFile
    Open nome_arquivo For Append As #NumArq
    
    Recordset.MoveFirst
        
    'Varrendo os registros - - - LINHAS
    Do While Recordset.EOF = False
        If campo_codigo_item_cii <> "" Then
           'GRAVAR O CII
           If Recordset.Fields(campo_codigo_item_cii) = "" Or IsNull(Recordset.Fields(campo_codigo_item_cii)) = True Then
              If Nome_campo_empresa = "" Then
                 funcoes_banco.Alterar Tabela_pai, "SET " & campo_codigo_item_cii & " = " & Recordset.Fields(valor_codigo_item_cii) & "", campo_codigo_chave, Recordset.Fields(valor_codigo_chave), "Otica", Form, "BDRetaguarda"
              Else
                 funcoes_banco.Alterar Tabela_pai, "SET " & campo_codigo_item_cii & " = " & Recordset.Fields(valor_codigo_item_cii) & "", campo_codigo_chave, Recordset.Fields(valor_codigo_chave), "Otica", Form, "BDRetaguarda", Nome_campo_empresa, Recordset.Fields(Nome_campo_empresa)
              End If
           End If
        End If
        
        'Varrendo os Campos - - - COLUNAS
        intColunas = 0
        
        Do While Recordset.Fields.Count > intColunas
        
           intContador = 0
           Debug.Print Recordset.Fields(intColunas).Name
           'Verificando se é FK e puxando FK's pelo CII
           If Campos_FK_tabela_filha <> "" Then
                Dim rstFKS As New ADODB.Recordset
                Dim strTamanho_string As String
                Dim strTamanho_string_CII As String
                Dim strTamanho_string_Tabela As String
                Dim strSql As String
                Dim strIni As Long
                Dim strCII_FK As String
                Dim strTabela_FK As String
                Dim strCampos_CII_FK As String
                Dim strCampos_Tabela_Fk As String
                Dim strCampos_ID_FK As String
                
                strTamanho_string = 1
                strTamanho_string_CII = 1
                strTamanho_string_Tabela = 1
                strTamanho_string_ID = 1
                                
                Do While Quantidade_FKS > intContador
                
                   strString = Campos_FK_tabela_filha
                   strString_CII_FK = CII_Tabela_Pai_FK
                   strString_Tabela_FK = Tabela_pai_FK
                   strString_Id_Tabela_FK = Id_Tabela_pai_FK
                    
                   'strString = Campos_FK_tabela_filha
                   strTamanho_string = InStr(1, strString, ",")
                   strTamanho_string_CII = InStr(1, strString_CII_FK, ",")
                   strTamanho_string_Tabela = InStr(1, strString_Tabela_FK, ",")
                   strTamanho_string_ID = InStr(1, strString_Id_Tabela_FK, ",")
                   
                   If strTamanho_string > 0 Then
                      'Campos Filhas
                      strCampos = Mid(strString, 1, strTamanho_string - 1)
                      strString = Mid(Campos_FK_tabela_filha, CInt(strTamanho_string_total) + CInt(strTamanho_string) + 1, Len(Campos_FK_tabela_filha))
                      'CII
                      strCampos_CII_FK = Mid(strString_CII_FK, 1, strTamanho_string_CII - 1)
                      strString_CII_FK = Mid(CII_Tabela_Pai_FK, CInt(strTamanho_string_total_CII) + CInt(strTamanho_string_CII) + 1, Len(CII_Tabela_Pai_FK))
                      'Tabela
                      strCampos_Tabela_Fk = Mid(strString_Tabela_FK, 1, strTamanho_string_Tabela - 1)
                      strString_Tabela_FK = Mid(Tabela_pai_FK, CInt(strTamanho_string_total_tabela) + CInt(strTamanho_string_Tabela) + 1, Len(Tabela_pai_FK))
                      'ID
                      strCampos_ID_FK = Mid(strString_Id_Tabela_FK, 1, strTamanho_string_ID - 1)
                      strString_Id_Tabela_FK = Mid(Id_Tabela_pai_FK, CInt(strTamanho_string_total_ID) + CInt(strTamanho_string_ID) + 1, Len(Id_Tabela_pai_FK))
                   Else
                      strCampos = strString
                      strCampos_CII_FK = strString_CII_FK
                      strCampos_Tabela_Fk = strString_Tabela_FK
                      strCampos_ID_FK = strString_Id_Tabela_FK
                   End If
                   
                   If strCampos = Recordset.Fields.Item(intColunas).Name Then
                      'Localizar o CII do ID
                      strSql = "SELECT * FROM " & strCampos_Tabela_Fk & " " & _
                               "WHERE  " & strCampos_ID_FK & " = " & Recordset.Fields.Item(intColunas).Value & " "
                      Movimentacoes.Select_geral strSql, "BDRetaguarda", rstFKS, "Otica", Form
                        
                      'Tratar CII NULL
                      strLinha = strLinha & "|" & CStr(rstFKS.Fields(strCampos_CII_FK))
                        
                      Set rstFKS = Nothing
                        
                      GoTo PROXIMA_COLUNA
                   End If
                   
                   strTamanho_string_total = CInt(strTamanho_string_total) + CInt(strTamanho_string)
                   strTamanho_string_total_CII = CInt(strTamanho_string_total_CII) + CInt(strTamanho_string_CII)
                   strTamanho_string_total_tabela = CInt(strTamanho_string_total_tabela) + CInt(strTamanho_string_Tabela)
                   
                   intContador = intContador + 1
                   'Campos
                   strTam_campos = InStr(1, Campos_FK_tabela_filha, ",")
                   Campos_FK_tabela_filha = Mid(Campos_FK_tabela_filha, strTam_campos + 1, Len(Campos_FK_tabela_filha))
                   'ID
                   strTam_campos_ID = InStr(1, Id_Tabela_pai_FK, ",")
                   Id_Tabela_pai_FK = Mid(Id_Tabela_pai_FK, strTam_campos_ID + 1, Len(Id_Tabela_pai_FK))
                   'Tabela
                   strTam_campos_Tabela = InStr(1, Tabela_pai_FK, ",")
                   Tabela_pai_FK = Mid(Tabela_pai_FK, strTam_campos_Tabela + 1, Len(Tabela_pai_FK))
                   'CII
                   strTam_campos_CII = InStr(1, CII_Tabela_Pai_FK, ",")
                   CII_Tabela_Pai_FK = Mid(CII_Tabela_Pai_FK, strTam_campos_CII + 1, Len(CII_Tabela_Pai_FK))
                Loop
           End If
        
           If intColunas = 0 Then
              If Nome_campo_empresa <> "" Then
                 If Nome_campo_empresa = Recordset.Fields.Item(intColunas).Name Then
                    strLinha = strLinha & "|" & cod_empresa
                 Else
                    strLinha = CStr(Recordset.Fields(intColunas))
                 End If
              Else
                 strLinha = CStr(Recordset.Fields(intColunas))
              End If
           Else
              If IsNull(Recordset.Fields(intColunas)) Then
                 strLinha = strLinha & "| "
              Else
                 If Nome_campo_empresa <> "" Then
                    If Nome_campo_empresa = Recordset.Fields.Item(intColunas).Name Then
                       strLinha = strLinha & "|" & cod_empresa
                    Else
                        Select Case Recordset.Fields.Item(intColunas).Type
                               Case adBoolean: strLinha = strLinha & "|" & CInt(Recordset.Fields(intColunas))
                               Case 135: strLinha = strLinha & "|" & CStr(Format(Recordset.Fields(intColunas), "YYYYMMDD hh:mm:ss"))
                               Case 130: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                               Case 202: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                               Case adTypeText: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                               Case 3: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                               Case 6: strLinha = strLinha & "|" & Funcoes_Gerais.Grava_Moeda(Recordset.Fields(intColunas))
                               Case 129: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                        End Select
                    End If
                 Else
                    Select Case Recordset.Fields.Item(intColunas).Type
                           Case adBoolean: strLinha = strLinha & "|" & CInt(Recordset.Fields(intColunas))
                           Case 135: strLinha = strLinha & "|" & CStr(Format(Recordset.Fields(intColunas), "YYYYMMDD"))
                           Case 130: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                           Case 202: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                           Case adTypeText: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                           Case 3: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                           Case 6: strLinha = strLinha & "|" & Funcoes_Gerais.Grava_Moeda(Recordset.Fields(intColunas))
                           Case 129: strLinha = strLinha & "|" & CStr(Recordset.Fields(intColunas))
                    End Select
                 End If
              End If
           End If
           
PROXIMA_COLUNA:

           intColunas = intColunas + 1
           
        Loop
        
        If Complemento_linha <> "" Then
           strLinha = strLinha & "|" & Complemento_linha & "|"
        Else
           strLinha = strLinha & "|" & Recordset.Fields(campo_codigo_chave)
        End If
        
        Print #NumArq, strLinha
        
        'GRAVAR ATUALIZADO SIM
        If campo_id_atualizado_filial <> "" Then
           If Nome_campo_empresa = "" Then
              funcoes_banco.Alterar Tabela_pai, "SET " & campo_id_atualizado_filial & " = 1 ", campo_codigo_chave, Recordset.Fields(campo_codigo_chave), "Otica", Form, "BDRetaguarda"
           Else
              funcoes_banco.Alterar Tabela_pai, "SET " & campo_id_atualizado_filial & " = 1 ", campo_codigo_chave, Recordset.Fields(campo_codigo_chave), "Otica", Form, "BDRetaguarda", Nome_campo_empresa, Recordset.Fields(Nome_campo_empresa)
           End If
        End If
        
        Recordset.MoveNext
    Loop
    
    Close #NumArq

End Function

Public Function Processa_TX(Recordset As ADODB.Recordset, nome_arquivo As String, Caminho_arquivo As String, Tabela_a_ser_atualizada As String, Conexao As Object, Form As Object, Field_CII As String, Field_ID_na_Tabela As String, Optional Campo_Empresa As String, Optional Valor_Campo_Empresa As String, Optional Pula_campo_chave As Boolean, Optional Presente_complemento_na_string As Boolean, Optional Campo_a_ser_preenchido_pelo_complemento As String, Optional ID_Tabela_pai As String, Optional Tabela_pai As String, Optional Field_CII_pai As String, Optional Sobrepor_dados As Boolean, Optional Id_Tabela_pai_FK As String, Optional Tabela_pai_FK As String, Optional Field_CII_pai_FK As String, Optional Quantidade_FKS As String, Optional Campos_FK_tabela_filha As String)
                                                                                                                                     
    Dim strLinha  As String
    Dim strVetLinha() As String
    Dim strValues As String
    Dim strFields As String
    Dim intColunas As Long
    Dim strSql As String
    Dim rstVerifica_reg_pai As New ADODB.Recordset
    
    ''On Error GoTo Erro
    
    Open Caminho_arquivo For Input As #1
    
    'Cabeçalho dos pedidos
    Do While Not EOF(1)
       Line Input #1, strLinha
       
       strVetLinha = Split(strLinha, "|")
       
       'Varrendo os Campos - COLUNAS
       intColunas = 0
       strFields = Empty
       strValues = Empty
       
       Do While intColunas < Recordset.Fields.Count
          If intColunas = 0 Then
             If Pula_campo_chave = True Then
                intColunas = intColunas + 1
             End If
             If Presente_complemento_na_string = True Then
                If Campo_a_ser_preenchido_pelo_complemento = Recordset.Fields.Item(intColunas).Name Then
                   intColunas = intColunas + 1
                End If
             End If
             
             Debug.Print Recordset.Fields.Item(intColunas).Name
             
             strFields = Recordset.Fields.Item(intColunas).Name
             strValues = strVetLinha(intColunas)
             
             Select Case Recordset.Fields.Item(intColunas).Type
                    Case adBoolean
                         If strVetLinha(intColunas) = " " Then
                            strVetLinha(intColunas) = 0
                         End If
                         strValues = strVetLinha(intColunas)
                    Case 135: strValues = "'" & strVetLinha(intColunas) & "'"
                    Case 130: strValues = "'" & strVetLinha(intColunas) & "'"
                    Case 202: strValues = "'" & strVetLinha(intColunas) & "'"
                    Case adTypeText: strValues = "'" & strVetLinha(intColunas) & "'"
                    Case 3
                         If strVetLinha(intColunas) = " " Then
                            strVetLinha(intColunas) = 0
                         End If
                         strValues = strVetLinha(intColunas)
                    Case 6
                         If strVetLinha(intColunas) = " " Then
                            strVetLinha(intColunas) = 0
                         End If
                         strValues = "" & strVetLinha(intColunas) & ""
                   Case 129: strValues = "'" & CStr(strVetLinha(intColunas)) & "'"
             End Select
             
             booAlterar = False
             
             Dim rstVerifica_reg As New ADODB.Recordset
             
             If Recordset.EOF <> True And Recordset.BOF <> True Then
                If IsNull(Recordset.Fields(Field_ID_na_Tabela)) = False Then
                   Recordset.MoveFirst
                   Recordset.Find (Field_ID_na_Tabela & " = " & Recordset.Fields(Field_ID_na_Tabela) & "")
                Else
                   GoTo FIM_VERIFICA_REG
                End If
             End If
             
             If Recordset.EOF = True Then
                GoTo FIM_VERIFICA_REG
             End If
             
             If IsNull(Recordset.Fields(Field_CII)) = False Then
                'Verificando se o registro já esta cadastrado
                strSql = Empty
                strSql = "SELECT * FROM " & Tabela_a_ser_atualizada & " WHERE " & Field_CII & " = " & Recordset.Fields(Field_CII) & " "
             
                If Campo_Empresa <> "" Then
                   strSql = strSql & "AND " & Campo_Empresa & " = " & Valor_Campo_Empresa & ""
                End If
             
                Movimentacoes.Select_geral strSql, "BDRetaguarda", rstVerifica_reg, "Otica", Form
             
                'Caso seja cadastrado o mesmo será excluido.
                If rstVerifica_reg.RecordCount > 0 Then
                   booAlterar = True
                End If
               
               Set rstVerifica_reg = Nothing
             Else
               booAlterar = False
             End If
FIM_VERIFICA_REG:

          Else
             If Presente_complemento_na_string = True Then
                If Campo_a_ser_preenchido_pelo_complemento = Recordset.Fields.Item(intColunas).Name Then
                   intColunas = intColunas + 1
                End If
             End If
             
             strFields = strFields & "," & Recordset.Fields.Item(intColunas).Name
             
             intContador_FK = 0
             Debug.Print Recordset.Fields(intColunas).Name
             
             'Verificando se é FK e puxando FK's pelo CII
             If Campos_FK_tabela_filha <> "" Then
                Dim rstFKS As New ADODB.Recordset
                Dim strTamanho_string As String
                Dim strTamanho_string_CII As String
                Dim strTamanho_string_Tabela As String
                Dim strIni As Long
                Dim strCII_FK As String
                Dim strTabela_FK As String
                Dim strCampos_CII_FK As String
                Dim strCampos_Tabela_Fk As String
                Dim strCampos_ID_FK As String
                
                strTamanho_string = 1
                strTamanho_string_CII = 1
                strTamanho_string_Tabela = 1
                strTamanho_string_ID = 1
                                
                Do While Quantidade_FKS > intContador_FK
                
                   strString = Campos_FK_tabela_filha
                   'strString_CII_FK = CII_Tabela_Pai_FK
                   strString_Tabela_FK = Tabela_pai_FK
                   strString_Id_Tabela_FK = Id_Tabela_pai_FK
                    
                   'strString = Campos_FK_tabela_filha
                   strTamanho_string = InStr(1, strString, ",")
                   'strTamanho_string_CII = InStr(1, strString_CII_FK, ",")
                   strTamanho_string_Tabela = InStr(1, strString_Tabela_FK, ",")
                   strTamanho_string_ID = InStr(1, strString_Id_Tabela_FK, ",")
                   
                   If strTamanho_string > 0 Then
                      'Campos Filhas
                      strCampos = Mid(strString, 1, strTamanho_string - 1)
                      strString = Mid(Campos_FK_tabela_filha, CInt(strTamanho_string_total) + CInt(strTamanho_string) + 1, Len(Campos_FK_tabela_filha))
                      'CII
                      strCampos_CII_FK = Mid(strString_CII_FK, 1, strTamanho_string_CII - 1)
                      strString_CII_FK = Mid(CII_Tabela_Pai_FK, CInt(strTamanho_string_total_CII) + CInt(strTamanho_string_CII) + 1, Len(CII_Tabela_Pai_FK))
                      'Tabela
                      strCampos_Tabela_Fk = Mid(strString_Tabela_FK, 1, strTamanho_string_Tabela - 1)
                      strString_Tabela_FK = Mid(Tabela_pai_FK, CInt(strTamanho_string_total_tabela) + CInt(strTamanho_string_Tabela) + 1, Len(Tabela_pai_FK))
                      'ID
                      strCampos_ID_FK = Mid(strString_Id_Tabela_FK, 1, strTamanho_string_ID - 1)
                      strString_Id_Tabela_FK = Mid(Id_Tabela_pai_FK, CInt(strTamanho_string_total_ID) + CInt(strTamanho_string_ID) + 1, Len(Id_Tabela_pai_FK))
                   Else
                      strCampos = strString
                      strCampos_CII_FK = strString_CII_FK
                      strCampos_Tabela_Fk = strString_Tabela_FK
                      strCampos_ID_FK = strString_Id_Tabela_FK
                   End If
                   
                   If strCampos = Recordset.Fields.Item(intColunas).Name Then
                      Dim rstID As New ADODB.Recordset
                      
                      'Pegando pelo CII o ID certo
                      strSql = Empty
                      strSql = "SELECT " & Id_Tabela_pai_FK & " FROM " & Tabela_pai_FK & " WHERE " & Field_CII_pai_FK & " = " & strVetLinha(intColunas) & " "
                      
                      Movimentacoes.Select_geral strSql, "BDRetaguarda", rstID, "Otica", Form
                      
                      strValues = strValues & "," & rstID.Fields.Item(Id_Tabela_pai_FK).Value
                      
                      Set rstID = Nothing
                      
                      GoTo PROXIMA_COLUNA:
                   End If
                   
                   strTamanho_string_total = CInt(strTamanho_string_total) + CInt(strTamanho_string)
                   strTamanho_string_total_CII = CInt(strTamanho_string_total_CII) + CInt(strTamanho_string_CII)
                   strTamanho_string_total_tabela = CInt(strTamanho_string_total_tabela) + CInt(strTamanho_string_Tabela)
                   
                   intContador_FK = intContador_FK + 1
                   'Campos
                   strTam_campos = InStr(1, Campos_FK_tabela_filha, ",")
                   Campos_FK_tabela_filha = Mid(Campos_FK_tabela_filha, strTam_campos + 1, Len(Campos_FK_tabela_filha))
                   'ID
                   strTam_campos_ID = InStr(1, Id_Tabela_pai_FK, ",")
                   Id_Tabela_pai_FK = Mid(Id_Tabela_pai_FK, strTam_campos + 1, Len(Id_Tabela_pai_FK))
                   'Tabela
                   strTam_campos_Tabela = InStr(1, Tabela_pai_FK, ",")
                   Tabela_pai_FK = Mid(Tabela_pai_FK, strTam_campos + 1, Len(Tabela_pai_FK))
                   'CII
                   strTam_campos_CII = InStr(1, CII_Tabela_Pai_FK, ",")
                   CII_Tabela_Pai_FK = Mid(CII_Tabela_Pai_FK, strTam_campos + 1, Len(CII_Tabela_Pai_FK))
                Loop
                   
             End If
             
             Select Case Recordset.Fields.Item(intColunas).Type
                    Case adBoolean
                         If strVetLinha(intColunas) = " " Then
                            strVetLinha(intColunas) = 0
                         End If
                         strValues = strValues & "," & strVetLinha(intColunas)
                    Case 135: strValues = strValues & "," & "'" & Funcoes_Gerais.Grava_String(strVetLinha(intColunas)) & "'"
                    Case 130: strValues = strValues & "," & "'" & Funcoes_Gerais.Grava_String(strVetLinha(intColunas)) & "'"
                    Case 202: strValues = strValues & "," & "'" & Funcoes_Gerais.Grava_String(strVetLinha(intColunas)) & "'"
                    Case adTypeText: strValues = strValues & "," & "'" & Funcoes_Gerais.Grava_String(strVetLinha(intColunas)) & "'"
                    Case 3
                         If strVetLinha(intColunas) = " " Then
                            strVetLinha(intColunas) = 0
                         End If
                         strValues = strValues & "," & strVetLinha(intColunas)
                    Case 6:
                         If strVetLinha(intColunas) = " " Then
                            strVetLinha(intColunas) = 0
                         End If
                         strValues = strValues & "," & "" & strVetLinha(intColunas) & ""
                   Case 129: strValues = strValues & "," & "'" & Funcoes_Gerais.Grava_String(strVetLinha(intColunas)) & "'"
             End Select
                         
          End If
          
PROXIMA_COLUNA:

          intColunas = intColunas + 1

       Loop
    
       If Presente_complemento_na_string = True Then
       
          'Verificando se o registro já esta cadastrado
          strSql = Empty
          strSql = "SELECT " & ID_Tabela_pai & " FROM " & Tabela_pai & " WHERE " & Field_CII_pai & " = " & strVetLinha(intColunas) & " "
           
          If Campo_Empresa <> "" Then
             strSql = strSql & "AND " & Campo_Empresa & " = " & Valor_Campo_Empresa & ""
          End If
          
          Movimentacoes.Select_geral strSql, "BDRetaguarda", rstVerifica_reg_pai, "Otica", Form
          
          If rstVerifica_reg_pai.EOF = True And rstVerifica_reg_pai.BOF = True Then
             GoTo Fim_complemento
          End If
          
          id_pai = rstVerifica_reg_pai.Fields(ID_Tabela_pai)
          
          strFields = strFields & "," & Campo_a_ser_preenchido_pelo_complemento
          strValues = strValues & "," & id_pai
          
Fim_complemento:

          Set rstVerifica_reg_pai = Nothing
          
       End If
       
       If booAlterar = False Then
          If strFields <> "" Then
             strSql = Empty
             strSql = "INSERT INTO " & Tabela_a_ser_atualizada & " ( " & strFields & " ) VALUES ( " & strValues & ") "
             Conexao.CNconexao.BeginTrans
             Conexao.CNconexao.Execute strSql
             Conexao.CNconexao.CommitTrans
          End If
       Else
          If Sobrepor_dados = True Then
             Dim intTam_string As Integer
             Dim intTam_string2 As Integer
             Dim strField_alterar As String
             Dim strValues2 As String
             Dim strFields2 As String
             
             strFields2 = strFields
             strValues2 = strValues
             
             intContador = 1

             Do While InStr(1, strFields2, ",") > 0
               
                intTam_string = InStr(1, strFields2, ",")
                intTam_string2 = InStr(1, strValues2, ",")
                
                If intContador = 1 Then
                   strSql = "UPDATE " & Tabela_a_ser_atualizada & " SET "
                   strField_alterar = Mid(strFields2, 1, intTam_string - 1) & " = " & Mid(strValues2, 1, intTam_string2 - 1)
                Else
                   strField_alterar = strField_alterar & "," & Mid(strFields2, 1, intTam_string - 1) & " = " & Mid(strValues2, 1, intTam_string2 - 1)
                End If
                
                strFields2 = Mid(strFields2, intTam_string + 1, Len(strFields2))
                
                strValues2 = Mid(strValues2, intTam_string2 + 1, Len(strValues2))
                
                intContador = intContador + 1
                
             Loop
             
             Dim rstVerificaUpdate As New ADODB.Recordset
             Dim strSql2 As String
             
             'Verificando se o registro já esta cadastrado
             strSql2 = Empty
             strSql2 = "SELECT " & ID_Tabela_pai & " FROM " & Tabela_pai & " WHERE " & Field_CII_pai & " = " & strVetLinha(intColunas - 1) & " "
           
             If Campo_Empresa <> "" Then
                strSql2 = strSql & "AND " & Campo_Empresa & " = " & Valor_Campo_Empresa & ""
             End If
          
             Movimentacoes.Select_geral strSql2, "BDRetaguarda", rstVerificaUpdate, "Otica", Form
          
             If rstVerificaUpdate.EOF = True And rstVerificaUpdate.BOF = True Then
                GoTo Fim_UPDATE
             End If
             
             lngID_pai = rstVerificaUpdate.Fields(ID_Tabela_pai)
             
             strSql = strSql & strField_alterar
             
             If ID_Tabela_pai <> "" Then
                strSql = strSql & " WHERE " & ID_Tabela_pai & " = " & lngID_pai & ""
                If Campo_Empresa <> "" Then
                   strSql = strSql & "AND " & Campo_Empresa & " = " & Valor_Campo_Empresa & ""
                End If
             Else
                If Campo_Empresa <> "" Then
                   strSql = strSql & "AND " & Campo_Empresa & " = " & Valor_Campo_Empresa & ""
                End If
             End If
             
             Conexao.CNconexao.BeginTrans
             Conexao.CNconexao.Execute strSql
             Conexao.CNconexao.CommitTrans
             
Fim_UPDATE:

             Set rstVerificaUpdate = Nothing
             
          End If
       End If
    Loop
    
    Close #1
    
    Exit Function
Erro:

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TRATAMENTO DE ERRO PARA ARQUIVO INEXISTENTE NO DIRETORIO DO VENDEDOR'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Err.Number = "53" Then
       Exit Function
       Err.Clear
    End If
    
    Call Erro.Erro(Form, "Otica")
    
    Exit Function
    
End Function

