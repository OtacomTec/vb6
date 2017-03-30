Attribute VB_Name = "Integracoes"
'---------------------------------------------------------------------
' Este módulo contém funçoes específicas para integração entre os
' módulos do director
'---------------------------------------------------------------------
Option Explicit

Dim strCod_Empresa As String

Dim adrFuncionario As ADODB.Recordset
Dim adrAfastamento As ADODB.Recordset
Dim adrAdmissao As ADODB.Recordset

Dim vetDia_Sem_Vale(1 To 7) As Integer
Dim vetDia_Sem_Afas(1 To 7) As Integer

Dim strJornada As String

Enum eRetono
    Retornar_ID_Deposito
    Retornar_Deposito_Item_Unidade
End Enum
Enum eTipo_Movimento
    Entrada
    Saida
End Enum

Public Sub Movto_Estoque_Saida(Cod_Empresa, Cod_Cliente, Cod_Tipo_Documento, Num_Documento, Data_Movto, Cod_Deposito, Cod_Item, Cod_Unidade_Armazenagem, Qtde_Saida, Optional ID_Movto)
'******************************************************************************
'Sistema...........................: Director
'Módulo............................: Integracao
'Procedimento/Função...............: Movto_Estoque_Saida
'Objetivo:.........................: Gerar um movimento de saída e executando
'                                    todos os tratamentos necessários para manter
'                                    a integridade do estoque.
'Desenvolvimento...................: Fernando Souza
'Data de criação...................: ../03/2001
'Data da última manutenção.........:
'Manutenção executada por..........:
'Observaçãoes......................:
'
'    Parâmetros:
'    • Cod_Empresa: Código da Empresa
'    • Cod_Cliente: Código do Cliente
'    • Cod_Tipo_Documento: Código do Tipo de Documento
'    • Num_Documento: Código do Número do Documento
'    • Data_Movto: Data do Movimento
'    • Cod_Deposito: Código do Depósito
'    • Cod_Item: Código ddo Item de Estoque
'    • Cod_Unidade_Armazenagem: Código da Unidade de Armazenagem
'    • Qtde_Saida: Quantidade de Saida do item de estoque
'    • ID_Movto: Se esse parâmetro for passado com o id do movimento usado pela
'                interface então será interpretado como uma alteração a esse movimento,
'                senão, se esse parâmetro for passado com 0 será interpretado que o
'                usuário quer que o id do movimento seja retornado a esse parâmetro,
'                senão, a função irá trabalhar sozinha sem retornar valor algum.
'
'******************************************************************************
    On Error GoTo Erro
    Dim strSQL As String

    If Not IsNumeric(ID_Movto) Or IsMissing(ID_Movto) Then
        ID_Movto = 0
    End If

    ' declarações
    strSQL = _
        "DECLARE " & _
            "/* variaveis para movimento */ " & _
            "@ID_Movto INT, " & _
            "@Cod_Empresa INT, " & _
            "@Cod_Cliente INT, " & _
            "@Cod_Tipo_Docto INT, " & _
            "@Numero_Docto NVARCHAR(20), " & _
            "@Data_Movto SMALLDATETIME, " & _
            "" & _
            "/* variaveis para itens de movimento */ " & _
            "@ID_Deposito_Item INT, " & _
            "@Cod_Deposito INT, " & _
            "@Cod_Item INT, " & _
            "@Transito_Direto BIT, " & _
            "@Cod_Unidade_Armazenagem INT, " & _
            "@Qtde_Saida DECIMAL(18,4) "


    ' atribuição de valores
    strSQL = strSQL & _
 _
        "SET @ID_Movto = " & ID_Movto & " " & _
        "SET @Cod_Empresa = " & Cod_Empresa & " " & _
        "SET @Cod_Cliente = " & Cod_Cliente & " " & _
        "SET @Cod_Tipo_Docto = " & Cod_Tipo_Documento & " " & _
        "SET @Numero_Docto = '" & Num_Documento & "'" & _
        "SET @Data_Movto = '" & Format(Data_Movto, "yyyyMMdd") & "' " & _
        "" & _
        "SET @Cod_Deposito = " & Cod_Deposito & " " & _
        "SET @Cod_Item = " & Cod_Item & " " & _
        "SET @Cod_Unidade_Armazenagem = " & Cod_Unidade_Armazenagem & " " & _
        "SET @Qtde_Saida = " & Funcoes_Gerais.Grava_Moeda(CStr(Qtde_Saida)) & " " & _
        "" & _
        "SET @Transito_Direto = (SELECT DFtransito_direto FROM TBitem_estoque WHERE DFcod_item_estoque = @Cod_Item) "


    ' deposito_item
    strSQL = strSQL & _
 _
        "/* id depósito */ " & _
        "SET @ID_Deposito_Item = (SELECT DFid_deposito_item " & _
                                 "FROM TBdeposito_item " & _
                                 "WHERE DFcod_deposito = @Cod_Deposito " & _
                                   "AND DFcod_item_estoque = @Cod_Item " & _
                                   "AND DFcod_unidade_armazenagem = @Cod_Unidade_Armazenagem) " & _
        "IF ISNULL(@ID_Deposito_Item,-1) = -1 " & _
            "RETURN /* cancela a execução desta sql */"



    ' atualizar movimento de estoque
    strSQL = strSQL & _
 _
        "/* movimento_estoque */ " & _
        "IF @ID_Movto = 0 " & _
         "BEGIN " & _
            "SET @ID_Movto = (SELECT DFid_movimento_estoque FROM TBmovimento_estoque WHERE DFcod_empresa = @Cod_Empresa AND DFcod_tipo_documento = @Cod_Tipo_Docto AND DFnumero_docto = @Numero_Docto AND DFtipo_movto_estoque = 'S') " & _
            "IF ISNULL(@ID_Movto,-1) = -1 " & _
             "BEGIN " & _
                "INSERT INTO TBmovimento_estoque (DFcod_empresa, DFcod_tipo_documento, DFnumero_docto, DFtipo_movto_estoque, DFdata_movimento) " & _
                "SELECT @Cod_Empresa, @Cod_Tipo_Docto, @Numero_Docto, 'S', @Data_Movto " & _
                "" & _
                "SET @ID_Movto = (SELECT MAX(DFid_movimento_estoque) FROM TBmovimento_estoque) " & _
                "" & _
                "/* inserindo agregacao entre movto_estoque e cliente */ " & _
                "INSERT INTO TBmovimento_estoque_cliente (DFid_movimento_estoque, DFcod_cliente) " & _
                "SELECT @ID_Movto, @Cod_Cliente " & _
             "END "



    ' incluir movimento de estoque
    strSQL = strSQL & _
 _
            "ELSE " & _
             "BEGIN " & _
                "UPDATE TBmovimento_estoque " & _
                "SET DFcod_empresa = @Cod_Empresa, DFcod_tipo_documento = @Cod_Tipo_Docto, DFnumero_docto = @Numero_Docto, DFtipo_movto_estoque = 'S', DFdata_movimento = @Data_Movto " & _
                "WHERE DFid_movimento_estoque = @ID_Movto " & _
                "" & _
                "/* atualizando agregacao entre movto_estoque e cliente */ " & _
                "UPDATE TBmovimento_estoque_cliente " & _
                "SET DFcod_cliente = @Cod_Cliente " & _
                "WHERE DFid_movimento_estoque = @ID_Movto " & _
             "END " & _
         "END " & _
        "Else " & _
         "BEGIN " & _
             "Update TBmovimento_estoque " & _
             "SET DFcod_empresa = @Cod_Empresa, DFcod_tipo_documento = @Cod_Tipo_Docto, DFnumero_docto = @Numero_Docto, DFtipo_movto_estoque = 'S', DFdata_movimento = @Data_Movto " & _
             "WHERE DFid_movimento_estoque = @ID_Movto " & _
             "" & _
             "/* atualizando agregacao entre movto_estoque e cliente */ " & _
             "Update TBmovimento_estoque_cliente " & _
             "SET DFcod_cliente = @Cod_Cliente " & _
             "WHERE DFid_movimento_estoque = @ID_Movto " & _
         "End "

    ' incluir movimento de estoque
    strSQL = strSQL & _
 _
    "/* item movimento */ " & _
    "IF EXISTS (SELECT DFid_item_movto_estoque FROM TBitem_movto_estoque WHERE DFid_movimento_estoque = @ID_Movto AND DFid_deposito_item = @ID_Deposito_Item) " & _
     "BEGIN " & _
        "DECLARE @Diferenca DECIMAL(18,4) " & _
        "" & _
        "/* para atualizar a quantidade do estoque é necessário subtrair da qtde alterada pela interface a qtde em estoque */ " & _
        "SET @Diferenca = (SELECT @Qtde_Saida - DFqtde FROM TBitem_movto_estoque WHERE DFid_movimento_estoque = @ID_Movto AND DFid_deposito_item = @ID_Deposito_Item) " & _
        "" & _
        "UPDATE TBitem_movto_estoque " & _
        "SET DFqtde = @Qtde_Saida , DFvalor = 0 " & _
        "WHERE DFid_movimento_estoque = @ID_Movto AND DFid_deposito_item = @ID_Deposito_Item " & _
        "" & _
        "IF @Transito_Direto = 0 " & _
            "Update TBdeposito_item " & _
            "SET DFquantidade_estoque = ISNULL(DFquantidade_estoque,0) - @Diferenca " & _
            "WHERE TBdeposito_item.DFid_deposito_item = @ID_Deposito_Item " & _
     "End "



    ' incluir movimento de estoque
    strSQL = strSQL & _
 _
    "Else " & _
     "BEGIN " & _
        "INSERT INTO TBitem_movto_estoque (DFid_movimento_estoque, DFid_deposito_item, DFqtde, DFvalor) " & _
        "SELECT @ID_Movto, @ID_Deposito_Item, @Qtde_Saida, 0 " & _
        "" & _
        "IF @Transito_Direto = 0 " & _
            "Update TBdeposito_item " & _
            "SET DFquantidade_estoque = ISNULL(DFquantidade_estoque,0) - @Qtde_Saida " & _
            "WHERE TBdeposito_item.DFid_deposito_item = @ID_Deposito_Item " & _
     "End "



    ' ultima data de saída
    strSQL = strSQL & _
 _
    "/* última data saida */ " & _
    "Update TBitem_estoque " & _
    "SET DFdata_ultima_saida = @Data_Movto " & _
    "WHERE DFcod_item_estoque = @Cod_Item AND @Data_Movto > DFdata_ultima_saida "



    ' movendo o id do movimento para uma view temporária
    strSQL = strSQL & "SELECT @ID_Movto AS DFid_movto INTO TBtemp_movto"



    ' Executando a Instrução em SQL --------------------------------------------
    CNConexao.Execute strSQL


    ' Retornando o Id do Movimento ---------------------------------------------
    strSQL = "SELECT DFid_movto FROM TBtemp_movto"
    ID_Movto = CNConexao.Execute(strSQL).Fields("DFid_movto")


    ' Excluindo a Tablea Temporária --------------------------------------------
    CNConexao.Execute "DROP TABLE TBtemp_movto"


    Exit Sub
Erro:
    Call Erro.Erro
End Sub

Public Sub Movto_Estoque_Entrada(Cod_Empresa, Cod_Fornecedor, Cod_Tipo_Documento, Num_Documento, Data_Movto, Cod_Deposito, Cod_Item, Cod_Unidade_Armazenagem, Qtde_Entrada, Optional ID_Movto)
'******************************************************************************
'Sistema...........................: Director
'Módulo............................: Integracao
'Procedimento/Função...............: Movto_Estoque_Entrada
'Objetivo:.........................: Gerar um movimento de entrada e executando
'                                    todos os tratamentos necessários para manter
'                                    a integridade do estoque.
'Desenvolvimento...................: Fernando Souza
'Data de criação...................: ../03/2001
'Data da última manutenção.........:
'Manutenção executada por..........:
'Observaçãoes......................:
'
'    Parâmetros:
'    • Cod_Empresa: Código da Empresa
'    • Cod_Fornecedor: Código do Fornecedor
'    • Cod_Tipo_Documento: Código do Tipo de Documento
'    • Num_Documento: Código do Número do Documento
'    • Data_Movto: Data do Movimento
'    • Cod_Deposito: Código do Depósito
'    • Cod_Item: Código ddo Item de Estoque
'    • Cod_Unidade_Armazenagem: Código da Unidade de Armazenagem
'    • Qtde_Entrada: Quantidade de entrada do item de estoque
'    • ID_Movto: Se esse parâmetro for passado com o id do movimento usado pela
'                interface então será interpretado como uma alteração a esse movimento,
'                senão, se esse parâmetro for passado com 0 será interpretado que o
'                usuário quer que o id do movimento seja retornado a esse parâmetro,
'                senão, a função irá trabalhar sozinha sem retornar valor algum.
'
'******************************************************************************
    Dim strSQL As String

    If Not IsNumeric(ID_Movto) Or IsMissing(ID_Movto) Then
        ID_Movto = 0
    End If

    strSQL = _
 _
        "DECLARE " & _
            "/* variaveis para movimento */ " & _
            "@ID_Movto INT, " & _
            "@Cod_Empresa INT, " & _
            "@Cod_Fornecedor INT, " & _
            "@Cod_Tipo_Docto INT, " & _
            "@Numero_Docto NVARCHAR(20), " & _
            "@Data_Movto SMALLDATETIME, " & _
            "" & _
            "/* variaveis para itens de movimento */ " & _
            "@ID_Deposito_Item INT, " & _
            "@Cod_Deposito INT, " & _
            "@Cod_Item INT, " & _
            "@Transito_Direto BIT, " & _
            "@Cod_Unidade_Armazenagem INT, " & _
            "@Qtde_Entrada DECIMAL(18,4), " & _
            "@Valor MONEY "


    strSQL = strSQL & _
 _
        "SET @ID_Movto = " & ID_Movto & " " & _
        "SET @Cod_Empresa = " & Cod_Empresa & " " & _
        "SET @Cod_Fornecedor = " & Cod_Fornecedor & " " & _
        "SET @Cod_Tipo_Docto = " & Cod_Tipo_Documento & " " & _
        "SET @Numero_Docto = '" & Num_Documento & "' " & _
        "SET @Data_Movto = '" & Format(Data_Movto, "yyyyMMdd") & "' " & _
        "" & _
        "SET @Cod_Deposito = " & Cod_Deposito & " " & _
        "SET @Cod_Item = " & Cod_Item & " " & _
        "SET @Cod_Unidade_Armazenagem = " & Cod_Unidade_Armazenagem & " " & _
        "SET @Qtde_Entrada = " & Funcoes_Gerais.Grava_Moeda(CStr(Qtde_Entrada)) & " " & _
        "" & _
        "SET @Transito_Direto = (SELECT DFtransito_direto FROM TBitem_estoque WHERE DFcod_item_estoque = @Cod_Item) "


    strSQL = strSQL & _
 _
        "/* VALOR DO ITEM */ " & _
        "SET @Valor = (SELECT DFpreco_custo FROM TBitem_estoque WHERE DFcod_item_estoque = @Cod_Item) " & _
        "" & _
        "/* DEPÓSITO */ " & _
        "IF NOT EXISTS (SELECT DFid_deposito_item FROM TBdeposito_item WHERE DFcod_deposito = @Cod_Deposito AND DFcod_item_estoque = @Cod_Item AND DFcod_unidade_armazenagem = @Cod_Unidade_Armazenagem) " & _
            "INSERT INTO TBdeposito_item (DFcod_deposito, DFcod_item_estoque, DFcod_unidade_armazenagem, DFquantidade_estoque, DFquantidade_estoque_reservado, DFlocalizacao) " & _
            "SELECT @Cod_Deposito, @Cod_Item, @Cod_Unidade_Armazenagem, 0, 0, '' " & _
            "SET @ID_Deposito_Item = (SELECT DFid_deposito_item FROM TBdeposito_item WHERE DFcod_deposito = @Cod_Deposito AND DFcod_item_estoque = @Cod_Item AND DFcod_unidade_armazenagem = @Cod_Unidade_Armazenagem) "


    strSQL = strSQL & _
 _
        "/* MOVIMENTO */ " & _
        "IF @ID_Movto = 0 " & _
         "BEGIN " & _
            "SET @ID_Movto = (SELECT DFid_movimento_estoque FROM TBmovimento_estoque WHERE DFcod_empresa = @Cod_Empresa AND DFcod_tipo_documento = @Cod_Tipo_Docto AND DFnumero_docto = @Numero_Docto AND DFtipo_movto_estoque = 'E') " & _
            "IF ISNULL(@ID_Movto,-1) = -1 " & _
             "BEGIN " & _
                "INSERT INTO TBmovimento_estoque (DFcod_empresa, DFcod_tipo_documento, DFnumero_docto, DFtipo_movto_estoque, DFdata_movimento)" & _
                "SELECT @Cod_Empresa, @Cod_Tipo_Docto, @Numero_Docto, 'E', @Data_Movto " & _
                "" & _
                "SET @ID_Movto = (SELECT MAX(DFid_movimento_estoque) FROM TBmovimento_estoque) " & _
                "" & _
                "INSERT INTO TBmovimento_estoque_fornecedor (DFid_movimento_estoque, DFcod_fornecedor) " & _
                "SELECT @ID_Movto, @Cod_Fornecedor " & _
             "End "

    strSQL = strSQL & _
 _
            "Else " & _
             "BEGIN " & _
                "Update TBmovimento_estoque " & _
                "SET DFcod_empresa = @Cod_Empresa, DFcod_tipo_documento = @Cod_Tipo_Docto, DFnumero_docto = @Numero_Docto, DFtipo_movto_estoque = 'E', DFdata_movimento = @Data_Movto " & _
                "WHERE DFid_movimento_estoque = @ID_Movto " & _
                "" & _
                "Update TBmovimento_estoque_fornecedor " & _
                "SET DFcod_fornecedor = @Cod_Fornecedor " & _
                "WHERE DFid_movimento_estoque = @ID_Movto " & _
             "End " & _
         "END " & _
        "Else " & _
         "BEGIN " & _
            "Update TBmovimento_estoque " & _
            "SET DFcod_empresa = @Cod_Empresa, DFcod_tipo_documento = @Cod_Tipo_Docto, DFnumero_docto = @Numero_Docto, DFtipo_movto_estoque = 'E', DFdata_movimento = @Data_Movto " & _
            "WHERE DFid_movimento_estoque = @ID_Movto " & _
         "" & _
            "Update TBmovimento_estoque_fornecedor " & _
            "SET DFcod_fornecedor = @Cod_Fornecedor " & _
            "WHERE DFid_movimento_estoque = @ID_Movto " & _
         "End "


    strSQL = strSQL & _
 _
        "/* ITEM DE MOVIMENTO */ " & _
        "DECLARE " & _
            "@Qtde_Estoque_Em_Unidade_Compra DECIMAL(18,4), " & _
            "@Qtde_Entrada_Em_Unidade_Compra DECIMAL(18,4), " & _
            "@Fator_Conversao DECIMAL(18,4), " & _
            "@Diferenca DECIMAL(18,4) " & _
        "" & _
        "/* CARREGANDO VARIÁVEIS PARA O CÁLCULO DO CUSTO MÉDIO " & _
            "para calcular o custo médio de um item é necessário converter a qtde deste item em todos os depósitos " & _
            "da empresa para a mesma unidade de compra deste item e somar estas qtdes encontradas, e também, " & _
            "converter a qtde entrada deste item para a unidade de compra. " & _
            "" & _
            "a conversão de uma qtde em uma determinada unidade para a unidade de compra é feita da seguinte forma: " & _
            "fator de conversão da unidade dividido pelo fator de conversão da unidade " & _
            "de compra multiplicado pela qtde. */ "



    strSQL = strSQL & _
 _
        "SET @Qtde_Estoque_Em_Unidade_Compra = (SELECT SUM((CONVERT(DECIMAL,TBunidade.DFfator_conversao) / CONVERT(DECIMAL,TBunidade_compra.DFfator_conversao)) * TBdeposito_item.DFquantidade_estoque) " & _
                                               "From TBunidade " & _
                                                   "INNER JOIN TBunidade AS TBunidade_compra " & _
                                                   "INNER JOIN TBitem_estoque " & _
                                                   "INNER JOIN TBdeposito_item " & _
                                                       "ON TBdeposito_item.DFcod_item_estoque = TBitem_estoque.DFcod_item_estoque " & _
                                                       "ON TBunidade_compra.DFcod_unidade = TBitem_estoque.DFcod_unidade_compra " & _
                                                       "ON TBunidade.DFcod_unidade = TBdeposito_item.DFcod_unidade_armazenagem " & _
                                               "WHERE TBdeposito_item.DFcod_item_estoque = @Cod_Item) " & _
                                               "" & _
        "SET @Fator_Conversao =(SELECT CONVERT(DECIMAL,TBunidade.DFfator_conversao) / CONVERT(DECIMAL,TBunidade_compra.DFfator_conversao) " & _
                               "FROM TBunidade AS TBunidade_compra " & _
                                  "INNER JOIN TBunidade " & _
                                  "INNER JOIN TBunidade_venda " & _
                                  "INNER JOIN TBitem_estoque " & _
                                      "ON TBitem_estoque.DFcod_item_estoque = TBunidade_venda.DFcod_item_estoque " & _
                                      "ON TBunidade_venda.DFcod_unidade = TBunidade.DFcod_unidade " & _
                                      "ON TBunidade_compra.DFcod_unidade = TBitem_estoque.DFcod_unidade_compra " & _
                               "WHERE TBitem_estoque.DFcod_item_estoque = @Cod_Item " & _
                                 "AND TBunidade_venda.DFcod_unidade = @Cod_Unidade_Armazenagem) "


    strSQL = strSQL & _
 _
        "IF EXISTS (SELECT DFid_item_movto_estoque FROM TBitem_movto_estoque WHERE DFid_movimento_estoque = @ID_Movto AND DFid_deposito_item = @ID_Deposito_Item) " & _
         "BEGIN /* alteracao de item de movimento de estoque */ " & _
            "" & _
            "/* a qtde de entrada na alteração do estoque é a diferença entre a qtde entrada na interface e a qtde estocada */ " & _
            "SET @Diferenca = (SELECT @Qtde_Entrada - DFqtde FROM TBitem_movto_estoque WHERE DFid_movimento_estoque = @ID_Movto AND DFid_deposito_item = @ID_Deposito_Item) " & _
            "" & _
            "/* inclusão de item de movimento de estoque */ " & _
            "Update TBitem_movto_estoque " & _
            "SET DFid_movimento_estoque = @ID_Movto, " & _
                "DFid_deposito_item = @ID_Deposito_Item, " & _
                "DFqtde = @Qtde_Entrada, " & _
                "DFvalor = @Valor " & _
            "WHERE DFid_movimento_estoque = @ID_Movto " & _
              "AND DFid_deposito_item = @ID_Deposito_Item "

    strSQL = strSQL & _
 _
            "/* CALCULO CUSTO MEDIO */ " & _
            "IF @Transito_Direto = 0 " & _
             "BEGIN " & _
                "" & _
                "IF @Diferenca > 0 " & _
                 "BEGIN " & _
                    "SET @Qtde_Entrada_Em_Unidade_Compra = (@Fator_Conversao * @Diferenca) " & _
                    "" & _
                    "Update TBitem_estoque " & _
                    "SET DFcusto_medio = (IsNull((DFcusto_medio * @Qtde_Estoque_Em_Unidade_Compra), 0) + (DFpreco_custo * @Qtde_Entrada_Em_Unidade_Compra)) / (@Qtde_Estoque_Em_Unidade_Compra  + @Qtde_Entrada_Em_Unidade_Compra) " & _
                    "WHERE DFcod_item_estoque = @Cod_Item " & _
                 "End " & _
                "/* FIM DO CÁLCULO DO CUSTO MÉDIO */ " & _
                "" & _
                "Update TBdeposito_item " & _
                "SET DFquantidade_estoque = ISNULL(DFquantidade_estoque,0) + @Diferenca " & _
                "WHERE TBdeposito_item.DFid_deposito_item = @ID_Deposito_Item " & _
             "END " & _
         "End "


    strSQL = strSQL & _
 _
        "Else " & _
         "BEGIN " & _
            "/* inclusão de item de movimento de estoque */ " & _
            "INSERT INTO TBitem_movto_estoque (DFid_movimento_estoque, DFid_deposito_item, DFqtde, DFvalor) " & _
            "SELECT @ID_Movto, @ID_Deposito_Item, @Qtde_Entrada, @Valor " & _
            "" & _
            "IF @Transito_Direto = 0 " & _
             "BEGIN " & _
                "IF @Qtde_Entrada > 0 " & _
                 "BEGIN " & _
                    "SET @Qtde_Entrada_Em_Unidade_Compra = (@Fator_Conversao * @Qtde_Entrada) " & _
                    "" & _
                    "Update TBitem_estoque " & _
                    "SET DFcusto_medio = (IsNull((DFcusto_medio * @Qtde_Estoque_Em_Unidade_Compra), 0) + (DFpreco_custo * @Qtde_Entrada_Em_Unidade_Compra)) / (@Qtde_Estoque_Em_Unidade_Compra + @Qtde_Entrada_Em_Unidade_Compra) " & _
                    "WHERE DFcod_item_estoque = @Cod_Item " & _
                 "End " & _
                 "" & _
                "Update TBdeposito_item " & _
                "SET DFquantidade_estoque = ISNULL(DFquantidade_estoque,0) + @Qtde_Entrada " & _
                "WHERE TBdeposito_item.DFid_deposito_item = @ID_Deposito_Item " & _
             "END " & _
         "End "



    strSQL = strSQL & _
 _
        "/* ÚLTIMA DATA DE ENTRADA */ " & _
        "Update TBitem_estoque " & _
        "SET DFdata_ultima_entrada = @Data_Movto " & _
        "WHERE DFcod_item_estoque = @Cod_Item AND @Data_Movto > DFdata_ultima_entrada "


    ' movendo o id do movimento para uma view temporária
    strSQL = strSQL & "SELECT @ID_Movto AS DFid_movto INTO TBtemp_movto"

    ' Executando a Instrução em SQL --------------------------------------------
    CNConexao.Execute strSQL

    ' Retornando o Id do Movimento ---------------------------------------------
    strSQL = "SELECT DFid_movto FROM TBtemp_movto"
    ID_Movto = CNConexao.Execute(strSQL).Fields("DFid_movto")

    ' Excluindo a Tablea Temporária --------------------------------------------
    CNConexao.Execute "DROP TABLE TBtemp_movto"


End Sub


Public Sub Excluir_Item_Movto_Estoque(Grid As MSFlexGrid, Tipo_Movimento As eTipo_Movimento, ID_Movto, Col_Deposito, Col_Item, Col_Unidade)
'******************************************************************************
'Sistema...........................: Director
'Módulo............................: Integracao
'Procedimento/Função...............: Deposito_Item
'Objetivo:.........................: Excluir todos os itens da TBmovimento_estoque
'                                    que forma excluídas dos MSFlexGrid pelo usuário
'Desenvolvimento...................: Fernando Souza
'Data de criação...................: ../03/2001
'Data da última manutenção.........:
'Manutenção executada por..........:
'Observaçãoes......................:
'    Parâmetros:
'    • Grid: nome do grid que sofre as alterações do usuário
'    • Tipo_Movimento: valores do tipo eTipo_Movimento definido do Declarations deste módulo
'                     (vale: 0 ou 1 / Entrada ou Saída)
'    • ID_Movto: ID do movimento referente ao grid
'    • Col_Deposito: A coluna do grid referente ao código do depósito
'    • Col_Item: A coluna do grid referente ao código do item de estoque
'    • Col_Unidade: A coluna do grid referente ao código da unidade
'******************************************************************************
    Dim adrItem_Movto As ADODB.Recordset
    Dim strSQL As String

    Dim Indice As Integer
    Dim booEncontrado As Boolean


    strSQL = _
        "SELECT TBitem_movto_estoque.DFid_item_movto_estoque, " & _
               "TBitem_movto_estoque.DFqtde, " & _
               "TBitem_estoque.DFtransito_direto, " & _
               "TBdeposito_item.DFid_deposito_item, " & _
               "TBdeposito_item.DFcod_deposito, " & _
               "TBdeposito_item.DFcod_item_estoque, " & _
               "TBdeposito_item.DFcod_unidade_armazenagem, " & _
               "CONVERT(DECIMAL,TBunidade.DFfator_conversao) / CONVERT(DECIMAL,TBunidade_padrao.DFfator_conversao) AS DFfator_conversao " & _
        "FROM TBunidade " & _
            "INNER JOIN TBunidade AS TBunidade_padrao " & _
            "INNER JOIN TBitem_estoque " & _
            "INNER JOIN TBdeposito_item " & _
            "INNER JOIN TBitem_movto_estoque " & _
                "ON TBitem_movto_estoque.DFid_deposito_item = TBdeposito_item.DFid_deposito_item " & _
                "ON TBdeposito_item.DFcod_item_estoque = TBitem_estoque.DFcod_item_estoque " & _
                "ON TBitem_estoque.DFcod_unidade_compra = TBunidade_padrao.DFcod_unidade " & _
                "ON TBdeposito_item.DFcod_unidade_armazenagem = TBunidade.DFcod_unidade " & _
        "WHERE TBitem_movto_estoque.DFid_movimento_estoque = " & ID_Movto
    Call Banco_Dados.SQLgeral(strSQL, adrItem_Movto)


    Do Until adrItem_Movto.EOF
        booEncontrado = False
        For Indice = 1 To (Grid.Rows - 1)
            If adrItem_Movto("DFcod_item_estoque") = Grid.TextMatrix(Indice, Col_Item) And _
               adrItem_Movto("DFcod_deposito") = Grid.TextMatrix(Indice, Col_Deposito) And _
               adrItem_Movto("DFcod_unidade_armazenagem") = Grid.TextMatrix(Indice, Col_Unidade) _
            Then

                booEncontrado = True
                Exit For
            End If
        Next Indice

        If Not booEncontrado Then
            If Tipo_Movimento = Entrada And adrItem_Movto("DFtransito_direto") = False Then
                strSQL = _
                    "DECLARE " & _
                        "@Qtde_Excluida DECIMAL(18,4), " & _
                        "@Cod_Item INT, " & _
                        "@Transito_Direto BIT, " & _
                        "@ID_Deposio_Item INT, " & _
                        "@Cod_Unidade_Armazenagem INT, " & _
                        "@Fator_Conversao DECIMAL(18,4), " & _
                        "@Qtde_Estoque DECIMAL(18,4) " & _
                        "" & _
                    "SET @Qtde_Excluida = " & Funcoes_Gerais.Grava_Moeda(adrItem_Movto("DFqtde")) & "" & _
                    "SET @Cod_Item = " & adrItem_Movto("DFcod_item_estoque") & "" & _
                    "SET @ID_Deposio_Item = " & adrItem_Movto("DFid_deposito_item") & "" & _
                    "SET @Cod_Unidade_Armazenagem = " & adrItem_Movto("DFcod_unidade_armazenagem") & "" & _
                    "SET @Fator_Conversao = " & Funcoes_Gerais.Grava_Moeda(adrItem_Movto("DFfator_conversao")) & " " & _
                    "SET @Transito_Direto = (SELECT DFtransito_direto FROM TBitem_estoque WHERE DFcod_item_estoque = @Cod_Item) "


                strSQL = strSQL & _
                    "IF (SELECT DFquantidade_estoque - @Qtde_Excluida FROM TBdeposito_item WHERE DFid_deposito_item = @ID_Deposio_Item) > 0 " & _
                     "BEGIN " & _
                        "SET @Qtde_Estoque = ( " & _
                            "SELECT SUM((CONVERT(DECIMAL,TBunidade.DFfator_conversao) / CONVERT(DECIMAL,TBunidade_compra.DFfator_conversao)) * TBdeposito_item.DFquantidade_estoque) " & _
                            "FROM TBunidade " & _
                                "INNER JOIN TBunidade AS TBunidade_compra " & _
                                "INNER JOIN TBitem_estoque " & _
                                "INNER JOIN TBdeposito_item " & _
                                    "ON TBdeposito_item.DFcod_item_estoque = TBitem_estoque.DFcod_item_estoque " & _
                                    "ON TBunidade_compra.DFcod_unidade = TBitem_estoque.DFcod_unidade_compra " & _
                                    "ON TBunidade.DFcod_unidade = TBdeposito_item.DFcod_unidade_armazenagem " & _
                            "WHERE TBdeposito_item.DFcod_item_estoque = @Cod_Item) " & _
                        "UPDATE TBitem_estoque " & _
                        "SET DFcusto_medio = (@Qtde_Estoque * ISNULL(DFcusto_medio,0)) - ((@Qtde_Excluida * @Fator_Conversao) * ISNULL(DFpreco_custo,0)) / (@Qtde_Estoque - (@Qtde_Excluida * @Fator_Conversao)) " & _
                        "WHERE DFcod_item_estoque = @Cod_Item " & _
                     "END "

                strSQL = strSQL & _
                    "ELSE " & _
                        "UPDATE TBitem_estoque " & _
                        "SET DFcusto_medio = DFpreco_custo " & _
                        "WHERE DFcod_item_estoque = @Cod_Item " & _
                    "" & _
                    "UPDATE TBdeposito_item " & _
                    "SET DFquantidade_estoque = DFquantidade_estoque - @Qtde_Excluida " & _
                    "WHERE TBdeposito_item.DFid_deposito_item = @ID_Deposio_Item "
                CNConexao.Execute strSQL

            ElseIf Tipo_Movimento = Saida And adrItem_Movto("DFtransito_direto") = False Then
                strSQL = _
                    "UPDATE TBdeposito_item " & _
                    "SET DFquantidade_estoque = DFquantidade_estoque + " & adrItem_Movto("DFqtde") & _
                    "WHERE DFid_deposito_item = " & adrItem_Movto("DFid_deposito_item")
                CNConexao.Execute strSQL
            End If

            strSQL = _
                "DELETE FROM TBitem_movto_estoque " & _
                "WHERE DFid_item_movto_estoque = " & adrItem_Movto("DFid_item_movto_estoque")
            CNConexao.Execute strSQL
        End If

        booEncontrado = False
        adrItem_Movto.MoveNext
    Loop
End Sub

Public Function Calcular_Vencimento_Cliente(Cod_Cliente As Integer, Data_Lancamento As Date, Data_Emissao As Date) As Date
'********************************************************************************
'Sistema...........................: Director
'Módulo............................: Integracoes
'Procedimento/Função...............: Calcular_Vencimento_Cliente
'Objetivo:.........................: Calcular a data de vencimento para documentos
'                                    segundo os dados cadastrais do cliente.
'Desenvolvimento...................: Fernando Souza
'Data de criação...................: 28/05/2001
'Data da última manutenção.........:
'Manutenção executada por..........:
'Observaçãoes......................:
'
' Parâmetros:
'    Cod_Cliente......: O código do cliente
'    Data_Lancamento..: A data da entrada do documento no sistema
'    Data_Emissao.....: A data de emissão (impressão) do documento
'
' Obs.: Existem casos em que a data de vencimento é calculada em cima da data de
'       emissão, outras, em cima da data de lançamento no sistema.
'
'********************************************************************************
    Dim adrCliente As ADODB.Recordset
    Dim strSQL As String

    strSQL = _
        "SELECT DFdia_vencimento, " & _
               "DFconsidera_sab_dom_fer, " & _
               "DFvencimento_sob, " & _
               "DFnum_dias_vencimento " & _
        "FROM TBcliente " & _
        "WHERE DFcod_cliente = " & Cod_Cliente
    Call Banco_Dados.SQLgeral(strSQL, adrCliente)


    ' calcula data por dia fixo
    If Not IsNull(adrCliente("DFdia_vencimento")) Then
        Calcular_Vencimento_Cliente = DateSerial(Year(Data_Emissao), Month(Data_Emissao) + 1, adrCliente("DFdia_vencimento"))


    ' calcula data por prazo validando ou não dias úteis
    Else
        Dim adrFeriado As ADODB.Recordset

        Dim datData As Date
        Dim datDia As Date

        Dim I As Integer
        Dim booFeriado As Boolean

        strSQL = _
            "SELECT DFdata_feriado " & _
            "FROM TBferiado " & _
            "WHERE DFdata_feriado > '" & Format(datData, "yyyyMMdd") & "'"
        Call Banco_Dados.SQLgeral(strSQL, adrFeriado)



        If adrCliente("DFvencimento_sob") = "E" Then
            datData = Data_Lancamento
        Else
            datData = Data_Emissao
        End If



        datDia = datData
        For I = 1 To adrCliente("DFnum_dias_vencimento")

            Do
                datDia = datDia + 1

                If Weekday(datDia) <> 1 And Weekday(datDia) <> 7 Then

                    adrFeriado.MoveFirst
                    booFeriado = False
                    Do While Not adrFeriado.EOF
                        If datDia = adrFeriado("DFdata_feriado") Then
                            booFeriado = True
                            Exit Do
                        End If
                        adrFeriado.MoveNext
                    Loop

                    If booFeriado = False Then
                        Exit Do
                    End If

                End If
            Loop

        Next I

        Calcular_Vencimento_Cliente = datDia

    End If
End Function

Public Sub Lancar_Imposto_Automatico(Tabela_Aplicacao As String, Id_Documento As Variant, Cod_Cliente As Variant, ByVal Valor As Variant)
'******************************************************************************
'Sistema...........................: Director
'Módulo............................: Integracao
'Procedimento/Função...............: Lancar_Imposto_Automatico
'Objetivo:.........................: Lancar INSS, ISS e IR automaticamente caso estas opçõs estejam
'                                    marcadas na TBopcoes
'Desenvolvimento...................: Fernando Souza
'Data de criação...................: 27/06/2001
'Data da última manutenção.........:
'Manutenção executada por..........:
'Observaçãoes......................:
'
'Parâmetros
'    • Tabela_Aplicacao..:  O Nome da tabela de agregação do documento com o imposto
'    • Id_Documento......:  O id (string ou integer) do documento em que será aplicado o imposto
'    • Cod_Cliente.......:  O código do cliente
'    • Valor.............:  O valor do documento (string, integer, single ou currency)
'
'    Exemplo da chamada da função, o INSS, ISS e IR serão lançados sobre o título a receber de id 288
'     e de valor R$1.254,45  (a tabela de agregação entre o imposto e o titulo a receber
'     é a TBimposto_aplicado_titulo_receber)
'
'    Call Lancar_INSS_Automatico( "TBimposto_aplicado_titulo_receber", 288, 1254.45 )
'
'******************************************************************************
    On Error GoTo Erro
    Dim strSQL As String
    Dim curImposto_Aplicado As Currency
    Dim adrImposto As ADODB.Recordset


    If IsNumeric(Valor) Then
        Valor = CCur(Valor)
    Else
        Exit Sub
    End If



    '--------------------------------------------------------------------
    ' INSS
    '--------------------------------------------------------------------
    strSQL = _
        "SELECT TBimposto.DFcod_imposto, " & _
               "TBimposto.DFpercentual, " & _
               "TBopcoes_b.DFvalor AS DFpiso_cobranca_inss, " & _
               "TBopcoes_c.DFvalor AS DFdestacar_descontar " & _
        "FROM TBopcoes TBopcoes_c " & _
            "INNER JOIN TBopcoes TBopcoes_b " & _
            "INNER JOIN TBopcoes TBopcoes_a " & _
            "INNER JOIN TBimposto " & _
                "ON CONVERT( NVARCHAR( 50 ), TBimposto.DFcod_imposto ) = TBopcoes_a.DFvalor " & _
                "ON TBopcoes_b.DFdescricao = 'Piso para cobrança automática de INSS' " & _
                "ON TBopcoes_c.DFdescricao = 'INSS Destacar/Calcular' " & _
        "WHERE TBopcoes_a.DFdescricao = 'Código do imposto referente ao INSS' "
    Call Banco_Dados.SQLgeral(strSQL, adrImposto)



    If adrImposto.RecordCount > 0 Then
        If Valor >= CCur(adrImposto("DFpiso_cobranca_inss")) Then

            curImposto_Aplicado = (Valor * adrImposto("DFpercentual")) / 100


            ' gravando INSS
            strSQL = _
                "IF NOT EXISTS ( " & _
                        "SELECT 1 FROM " & Tabela_Aplicacao & " " & _
                        "WHERE DFid_" & Replace(Tabela_Aplicacao, "TBimposto_aplicado_", "") & " = " & Id_Documento & " " & _
                          "AND DFcod_imposto = " & adrImposto("DFcod_imposto") & ") " & _
                   "INSERT INTO " & Tabela_Aplicacao & " " & _
                       "( DFid_" & Replace(Tabela_Aplicacao, "TBimposto_aplicado_", "") & ", DFcod_imposto, DFvalor, DFdestacar_calcular ) " & _
                   "SELECT " & Id_Documento & ", " & adrImposto("DFcod_imposto") & ", " & Funcoes_Gerais.Grava_Moeda(CStr(curImposto_Aplicado)) & ", '" & Left(adrImposto("DFdestacar_descontar"), 1) & "' "
            CNConexao.Execute strSQL

        End If
    End If
    '--------------------------------------------------------------------


    '--------------------------------------------------------------------
    ' ISS
    '--------------------------------------------------------------------
    strSQL = _
        "SELECT TBcliente.DFcod_cliente, " & _
               "TBcliente.DFaplicar_iss, " & _
               "TBimposto.DFcod_imposto, " & _
               "TBimposto.DFpercentual AS DFpercentual_iss, " & _
               "TBopcoes_b.DFvalor AS DFdestacar_descontar " & _
        "FROM TBcliente, TBopcoes AS TBopcoes_b, " & _
             "TBimposto " & _
                "INNER JOIN TBopcoes AS TBopcoes_a " & _
                    "ON CONVERT( NVARCHAR(50), TBimposto.DFcod_imposto ) = TBopcoes_a.DFvalor " & _
        "WHERE TBopcoes_a.DFdescricao = 'Código do imposto referente ao ISS' " & _
          "AND TBcliente.DFcod_cliente = " & Cod_Cliente & " " & _
          "AND TBcliente.DFaplicar_iss = 1 " & _
          "AND TBopcoes_b.DFdescricao = 'ISS Destacar/Calcular' "
    Call Banco_Dados.SQLgeral(strSQL, adrImposto)

    If adrImposto.RecordCount > 0 Then

            curImposto_Aplicado = (Valor * adrImposto("DFpercentual_iss")) / 100

            ' gravanco ISS
            strSQL = _
                "IF NOT EXISTS ( " & _
                        "SELECT 1 FROM " & Tabela_Aplicacao & " " & _
                        "WHERE DFid_" & Replace(Tabela_Aplicacao, "TBimposto_aplicado_", "") & " = " & Id_Documento & " " & _
                          "AND DFcod_imposto = " & adrImposto("DFcod_imposto") & ") " & _
                   "INSERT INTO " & Tabela_Aplicacao & " " & _
                       "( DFid_" & Replace(Tabela_Aplicacao, "TBimposto_aplicado_", "") & ", DFcod_imposto, DFvalor, DFdestacar_calcular ) " & _
                   "SELECT " & Id_Documento & ", " & adrImposto("DFcod_imposto") & ", " & Funcoes_Gerais.Grava_Moeda(CStr(curImposto_Aplicado)) & ", '" & Left(adrImposto("DFdestacar_descontar"), 1) & "' "
            CNConexao.Execute strSQL

    End If
    '--------------------------------------------------------------------



    '--------------------------------------------------------------------
    ' IR
    '--------------------------------------------------------------------
    strSQL = _
        "SELECT TBcliente.DFcod_cliente, " & _
               "TBcliente.DFaplicar_ir, " & _
               "TBimposto.DFcod_imposto, " & _
               "TBimposto.DFpercentual AS DFpercentual_ir, " & _
               "TBopcoes_b.DFvalor AS DFdestacar_descontar " & _
        "FROM TBcliente, TBopcoes AS TBopcoes_b, " & _
             "TBimposto As TBimposto " & _
                "INNER JOIN TBopcoes AS TBopcoes " & _
                    "ON CONVERT( NVARCHAR(50), TBimposto.DFcod_imposto ) = TBopcoes.DFvalor " & _
        "WHERE TBopcoes.DFdescricao = 'Código do imposto referente ao IR' " & _
          "AND TBcliente.DFcod_cliente = " & Cod_Cliente & " " & _
          "AND TBcliente.DFaplicar_ir = 1 " & _
          "AND TBopcoes_b.DFdescricao = 'IR Destacar/Calcular' "
    Call Banco_Dados.SQLgeral(strSQL, adrImposto)

    If adrImposto.RecordCount > 0 Then

            curImposto_Aplicado = (Valor * adrImposto("DFpercentual_ir")) / 100

            ' gravanco IR
            strSQL = _
                "IF NOT EXISTS ( " & _
                        "SELECT 1 FROM " & Tabela_Aplicacao & " " & _
                        "WHERE DFid_" & Replace(Tabela_Aplicacao, "TBimposto_aplicado_", "") & " = " & Id_Documento & " " & _
                          "AND DFcod_imposto = " & adrImposto("DFcod_imposto") & ") " & _
                   "INSERT INTO " & Tabela_Aplicacao & " " & _
                       "( DFid_" & Replace(Tabela_Aplicacao, "TBimposto_aplicado_", "") & ", DFcod_imposto, DFvalor, DFdestacar_calcular ) " & _
                   "SELECT " & Id_Documento & ", " & adrImposto("DFcod_imposto") & ", " & Funcoes_Gerais.Grava_Moeda(CStr(curImposto_Aplicado)) & ", '" & Left(adrImposto("DFdestacar_descontar"), 1) & "' "
            CNConexao.Execute strSQL

    End If
    '--------------------------------------------------------------------


    Exit Sub
Erro:
    Call Erro.Erro
End Sub

Public Sub Montar_SQL_Vale_Transporte(Empresa As String, Optional Num_Registros As Integer)
    Dim adrFuncionario As ADODB.Recordset
    Dim strSQL As String
    
    CNConexao.Execute "If Exists(SELECT * FROM sysObjects WHERE id = Object_id('dbo.TBtemp_vale_transporte'))Begin DROP TABLE TBtemp_vale_transporte End "

    strSQL = "SELECT TBfuncionario.DFmatricula, " & _
                    "TBfuncionario.DFnome, " & _
                    "TBempresa.DFnome_fantasia, " & _
                    "TBfuncionario.DFinicio_jornada, " & _
                    "TBfuncionario.DFtermino_jornada, " & _
                    "TBtipo_vale_transporte.DFdescricao AS DFtipo_vale, " & _
                    "TBtipo_vale_transporte_funcionario.DFqtde, " & _
                    "CONVERT(INT,0) AS DFqtde_mes, " & _
                    "TBtipo_vale_transporte.DFvalor, " & _
                    "CONVERT(Money, 0) As DFvalor_mes " & _
             "INTO TBtemp_vale_transporte " & _
             "FROM TBfuncionario " & _
                  "INNER JOIN TBtipo_vale_transporte_funcionario " & _
                          "ON TBtipo_vale_transporte_funcionario.DFmatricula = TBfuncionario.DFmatricula " & _
                  "INNER JOIN TBtipo_vale_transporte " & _
                          "ON TBtipo_vale_transporte.DFid_tipo_vale_transporte = TBtipo_vale_transporte_funcionario.DFid_tipo_vale_transporte " & _
                  "INNER JOIN TBempresa " & _
                          "ON TBfuncionario.DFcod_empresa = TBempresa.DFcod_empresa " & _
             "WHERE TBfuncionario.DFvale_transporte = 1 " & _
               "AND TBfuncionario.DFcod_afastamento IS NULL "

    If Empresa <> Empty Then
        strSQL = strSQL & _
            "AND TBfuncionario.DFcod_empresa IN (" & Empresa & ") "
    End If
    
    strSQL = strSQL & _
        "ORDER BY TBfuncionario.DFnome "
    
    CNConexao.Execute strSQL
    
    strSQL = "SELECT * FROM TBtemp_vale_transporte "
    
    Call Banco_Dados.SQLgeral(strSQL, adrFuncionario)
    
    If adrFuncionario.RecordCount = 0 Then
        Num_Registros = 0
    Else
        Num_Registros = adrFuncionario.RecordCount
    End If
End Sub

Public Sub Calcular_Vales_Funcionario(Empresa As String, Data_Emissao As Date, Ultimo_Dia_Mes As Integer, Optional Progress_Bar As ProgressBar)
    On Error GoTo Erro
    '---------------------------------------------------------------------------------------
    Dim adrFuncionario As ADODB.Recordset
    Dim adrAdmissao As ADODB.Recordset
    Dim strSQL As String
    
    Dim booAchou As Boolean
    Dim strValores As String
    Dim intInicio_Mes As Integer, intFim_Mes As Integer
    Dim intDias_Trab As Integer
    '---------------------------------------------------------------------------------------
    strSQL = "SELECT * FROM TBtemp_vale_transporte "
    
    Call Banco_Dados.SQLgeral(strSQL, adrFuncionario)
    '---------------------------------------------------------------------------------------
    If adrFuncionario.RecordCount <> 0 Then
    '---------------------------------------------------------------------------------------
        CNConexao.BeginTrans
        '-----------------------------------------------------------------------------------
        adrFuncionario.MoveFirst
        '-----------------------------------------------------------------------------------
        Do While Not adrFuncionario.EOF
        '-----------------------------------------------------------------------------------
            intDias_Trab = 0
            '-------------------------------------------------------------------------------
            
            ' Verificar Admissao do Funcionario
            strSQL = "SELECT DFdata_admissao " & _
                     "FROM TBfuncionario " & _
                     "WHERE (CONVERT(CHAR(4),YEAR(DFdata_admissao)))+ " & _
                           "(CONVERT(CHAR(2),MONTH(DFdata_admissao))) = " & _
                           "'" & Format(Data_Emissao, "yyyyM") & "' " & _
                       "AND DFmatricula = " & adrFuncionario("DFmatricula")
                    
            Call Banco_Dados.SQLgeral(strSQL, adrAdmissao)
            
            '-------------------------------------------------------------------------------
            If adrAdmissao.RecordCount = 0 Then
            '-------------------------------------------------------------------------------
                Call Verificar_Movimentacoes_Funcionario(Data_Emissao, adrFuncionario("DFmatricula"), Ultimo_Dia_Mes, strValores)
                '---------------------------------------------------------------------------
                If strValores <> Empty Then
                    Do While strValores <> Empty
                        intInicio_Mes = Mid(strValores, 1, 2)
                        intFim_Mes = Mid(strValores, 4, 2)
                        
                        Call Calcular_Dias_Trabalhados(Data_Emissao, intInicio_Mes, intFim_Mes, adrFuncionario("DFinicio_jornada"), adrFuncionario("DFtermino_jornada"), intDias_Trab)
                        strValores = IIf(Len(strValores) > 5, Mid(strValores, 7), Empty)
                    Loop
                Else
                    Call Calcular_Dias_Trabalhados(Data_Emissao, 1, Ultimo_Dia_Mes, adrFuncionario("DFinicio_jornada"), adrFuncionario("DFtermino_jornada"), intDias_Trab)
                End If
            '-------------------------------------------------------------------------------
            Else
                Call Calcular_Dias_Trabalhados(Data_Emissao, Day(adrAdmissao("DFdata_admissao")), Ultimo_Dia_Mes, adrFuncionario("DFinicio_jornada"), adrFuncionario("DFtermino_jornada"), intDias_Trab)
            End If
            '-------------------------------------------------------------------------------
            strSQL = _
                "UPDATE TBtemp_vale_transporte " & _
                   "SET DFqtde_mes = DFqtde * " & intDias_Trab & ", " & _
                       "DFvalor_mes = (DFqtde * " & intDias_Trab & ") * DFvalor " & _
                "WHERE DFmatricula = " & adrFuncionario("DFmatricula")
            '-------------------------------------------------------------------------------
            CNConexao.Execute strSQL
            '-------------------------------------------------------------------------------
            
            
            If Not Progress_Bar Is Nothing Then
                Progress_Bar = Progress_Bar + 1
            End If
            '-------------------------------------------------------------------------------
            adrFuncionario.MoveNext
        '-----------------------------------------------------------------------------------
        Loop
        '-----------------------------------------------------------------------------------
        Dim strEmpresa As String
        
        If Empresa <> Empty Then
            strEmpresa = Empresa
        Else
            strEmpresa = "TODAS"
        End If
        '-----------------------------------------------------------------------------------
        strSQL = _
            "UPDATE TBsistema " & _
               "SET DFcampo1 = '" & strEmpresa & "', " & _
                   "DFcampo2 = '" & Data_Emissao & "' " & _
             "WHERE DFcod_registro = 2 "
        
        CNConexao.Execute strSQL
        '-----------------------------------------------------------------------------------
        CNConexao.CommitTrans
    '---------------------------------------------------------------------------------------
    End If
    '---------------------------------------------------------------------------------------
    Exit Sub
    '---------------------------------------------------------------------------------------

Erro:
    CNConexao.RollbackTrans
    Call Erro.Erro("Calcular_Vales_Funcionario")
End Sub

Public Sub Calcular_Dias_Trabalhados(Data_Emissao As Date, Primeiro_Dia_Mes As Integer, Ultimo_Dia_Mes As Integer, Inicio_Jornada As Integer, Termino_Jornada As Integer, Dias_Trab As Integer)
    Dim Pri_Dia_Mes As Integer
    Dim Ult_Dia_Mes As Integer
    Dim Pri_Dia_Sem As Integer
    
    Dim strJornada As String
    
    Dim intInicio_Sem As Integer
    Dim intFim_Sem As Integer
    Dim Qtde_Dias_Sem As Integer
            
    Dim intNum_Sem As Integer
    
    Dim Cont As Integer
    Dim Indice As Integer
    '---------------------------------------------------------------------------------------
    Pri_Dia_Mes = Primeiro_Dia_Mes
    Ult_Dia_Mes = Ultimo_Dia_Mes

    Pri_Dia_Sem = Pri_Dia_Mes
    
    intInicio_Sem = Weekday(CDate(Format(Pri_Dia_Sem, "00") & "/" & Format(Data_Emissao, "MM/yyyy")))
    Qtde_Dias_Sem = (7 - intInicio_Sem)
    intFim_Sem = Weekday(CDate(Format((Pri_Dia_Sem + Qtde_Dias_Sem), "00") & "/" & Format(Data_Emissao, "MM/yyyy")))
    '---------------------------------------------------------------------------------------
    
    ' Montar string de jornada do funcionario
    Indice = Inicio_Jornada
    strJornada = Indice
    Do While Indice <> Termino_Jornada
        Indice = Indice + 1
        Indice = IIf(Indice = 8, 1, Indice)
        strJornada = strJornada & "," & Indice
    Loop
    
    '---------------------------------------------------------------------------------------
    intNum_Sem = (DateDiff("ww", CDate(Format(Pri_Dia_Sem, "00") & "/" & Format(Data_Emissao, "MM/yyyy")), CDate(Format(Ult_Dia_Mes, "00") & "/" & Format(Data_Emissao, "MM/yyyy"))) + 1)
    '---------------------------------------------------------------------------------------
    For Cont = 1 To intNum_Sem
        If (Pri_Dia_Sem + (7 - intInicio_Sem)) <= Ult_Dia_Mes Then
            intInicio_Sem = Weekday(CDate(Format(Pri_Dia_Sem, "00") & "/" & Format(Data_Emissao, "MM/yyyy")))
            Qtde_Dias_Sem = (7 - intInicio_Sem)
            intFim_Sem = Weekday(CDate(Format((Pri_Dia_Sem + Qtde_Dias_Sem), "00") & "/" & Format(Data_Emissao, "MM/yyyy")))
        Else
            intInicio_Sem = Weekday(CDate(Format(Pri_Dia_Sem, "00") & "/" & Format(Data_Emissao, "MM/yyyy")))
            intFim_Sem = Weekday(CDate(Format(Ult_Dia_Mes, "00") & "/" & Format(Data_Emissao, "MM/yyyy")))
        End If
        '-----------------------------------------------------------------------------------
        For Indice = intInicio_Sem To intFim_Sem
            If InStr(1, strJornada, Indice) <> 0 Then
                Dias_Trab = (Dias_Trab + 1)
            End If
        Next Indice
        '-----------------------------------------------------------------------------------
        Pri_Dia_Sem = (Pri_Dia_Sem + Qtde_Dias_Sem) + 1
        '-----------------------------------------------------------------------------------
    Next Cont
    '---------------------------------------------------------------------------------------
    Call Verificar_Feriado(Data_Emissao, Pri_Dia_Mes, Ult_Dia_Mes, strJornada, Dias_Trab)
    '---------------------------------------------------------------------------------------
End Sub

Public Function Retornar_Ultimo_Dia_Mes(Data As Date) As Integer
    Select Case Month(Data)
        Case 1, 3, 5, 7, 8, 10, 12
            Retornar_Ultimo_Dia_Mes = 31
        Case 4, 6, 9, 11
            Retornar_Ultimo_Dia_Mes = 30
        Case 2
            If Year(Data) Mod 4 Then
                Retornar_Ultimo_Dia_Mes = 28
            Else
                Retornar_Ultimo_Dia_Mes = 29
            End If
    End Select
End Function

Public Sub Verificar_Feriado(Data_Emissao As Date, Inicio_Mes As Integer, Fim_Mes As Integer, Jornada As String, Dias_Trab As Integer)
    Dim strSQL As String
    Dim adrVerifica_Feriado As ADODB.Recordset

    strSQL = "SELECT DFdata_feriado FROM TBferiado " & _
             "WHERE (CONVERT(CHAR(4),YEAR(DFdata_feriado)))+ " & _
                   "(CONVERT(CHAR(2),MONTH(DFdata_feriado))) = " & Format(Data_Emissao, "yyyyM") & " " & _
               "AND (CONVERT(CHAR(2),DAY(DFdata_feriado))) BETWEEN " & Inicio_Mes & " AND " & Fim_Mes
    
    Call Banco_Dados.SQLgeral(strSQL, adrVerifica_Feriado)

    If adrVerifica_Feriado.RecordCount <> 0 Then
        adrVerifica_Feriado.MoveFirst
        Do While Not adrVerifica_Feriado.EOF
            If InStr(1, Jornada, Weekday(CDate(adrVerifica_Feriado("DFdata_feriado")))) <> 0 Then
                Dias_Trab = (Dias_Trab - 1)
            End If
            adrVerifica_Feriado.MoveNext
        Loop
    End If
End Sub

Public Sub Verificar_Movimentacoes_Funcionario(Data_Emissao As Date, Matricula As String, Ultimo_Dia_Mes As Integer, Valores As String)
    Dim adrMovimentacoes As ADODB.Recordset
    Dim strSQL As String

    strSQL = _
        "SELECT TBafastamento_funcionario.DFdata_inicio, TBafastamento.DFtipo " & _
        "FROM TBfuncionario " & _
             "INNER JOIN TBafastamento_funcionario " & _
                     "ON TBfuncionario.DFmatricula = TBafastamento_funcionario.DFmatricula " & _
             "INNER JOIN TBafastamento " & _
                     "ON TBafastamento_funcionario.DFcod_movimentacao = TBafastamento.DFcod_afastamento " & _
        "WHERE (CONVERT(CHAR(4),YEAR(DFdata_inicio)))+ " & _
              "(CONVERT(CHAR(2),MONTH(DFdata_inicio))) = " & _
              "'" & Format(Data_Emissao, "yyyyM") & "' " & _
          "AND TBafastamento_funcionario.DFmatricula = " & Matricula
    
    Call Banco_Dados.SQLgeral(strSQL, adrMovimentacoes)
        
    If adrMovimentacoes.RecordCount <> 0 Then
        Do While Not adrMovimentacoes.EOF
            If Day(adrMovimentacoes("DFdata_inicio")) < 1 Then
                Valores = Valores & _
                          IIf(Valores <> Empty, ",", Empty) & _
                          Format(Day(IIf(adrMovimentacoes("DFtipo") = "A", _
                                        (adrMovimentacoes("DFdata_inicio") - 1), _
                                         adrMovimentacoes("DFdata_inicio"))), "00")
            Else
                Valores = Valores & _
                          IIf(Valores <> Empty, ",", Empty) & _
                          Format(Day(IIf(adrMovimentacoes("DFtipo") = "A", _
                                        (adrMovimentacoes("DFdata_inicio")), _
                                         adrMovimentacoes("DFdata_inicio"))), "00")
            End If
            adrMovimentacoes.MoveNext
        Loop
        adrMovimentacoes.MoveFirst
        
        If adrMovimentacoes("DFtipo") = "A" And (adrMovimentacoes.RecordCount Mod 2) = 0 Then
            Valores = "01," & Valores & "," & Ultimo_Dia_Mes
        ElseIf adrMovimentacoes("DFtipo") = "A" And (adrMovimentacoes.RecordCount Mod 2) <> 0 Then
            Valores = "01," & Valores
        ElseIf adrMovimentacoes("DFtipo") = "R" And (adrMovimentacoes.RecordCount Mod 2) <> 0 Then
            Valores = Valores & "," & Ultimo_Dia_Mes
        End If
        
    Else
        Valores = Empty
    End If
    
End Sub
