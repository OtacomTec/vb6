Attribute VB_Name = "Banco_Dados"
 '*******************************************************************************************
'
'Sistema...........................: Director
'Módulo............................: Nenhum
'Conexão...........................: Nenhuma
'Formulário........................: Banco_Dados
'Objetivo do formulário............: Banco de Dados
'Análise...........................: Eugênio Gomes
'Programação.......................: Pablo Souza, Eduardo Cruz
'Data..............................: 07/04/2000
'Data da última manutenção.........: 21/02/2001
'Manutenção executada por..........: Vagner Vilela
'
'*******************************************************************************************
'Declaracao das variaveis de Conexao e uma Recordset
Public CNConexao As ADODB.Connection 'Propriedade da classe BDados
Public CNconexaoII As ADODB.Connection
Public Conexao As ADODB.Connection
Public TBrecordset As ADODB.Recordset
Public Conexao_Aberta As Boolean
'Variáveis para receberem os parâmetros de conexão do banco de dados
Public strNome_Servidor As String 'Propriedade da classe BDados
Public strSenha_Admin As String 'Propriedade da classe BDados
Public strUser_Admin As String 'Propriedade da classe BDados

Public Function Abre_Conexao(Conexao As ADODB.Connection) As String
    '******************************************************************************
    'Módulo............................: Banco_Dados
    'Procedimento/Função...............: Abre_Conexao
    'Objeto/classe correspondente......: PIBDados.Abre_Conexao
    'Objetivo:.........................: Estabelece uma conexão com o banco de dados
    'Desenvolvimento...................: Pablo Souza, Eduardo Cruz
    'Data de criação...................: 07/04/2000
    'Utilização........................: Call Banco_Dados.Abre_Conexao(v1)
    'Parâmetros de entrada.............: Variável v1 do tipo ADODB.Connection
    '                                    Utiliza-se no Director uma variável (CNConexao) declarada
    '                                    pública no módulo de código Banco_Dados e utilizada
    '                                    em todo o projeto como conexão com o banco.
    'Saída.............................: -
    'Data da última manutenção.........:
    'Manutenção executada por..........:
    'Observações.......................: esta função conecta-se somente ao DBDirector, uma vez
    '                                    que tal informação é uma constante dentro da mesma.
    '                                    A própria função passa a senha, nome do servidor e usuário
    '                                    obtidos no registro do Windows.
    '******************************************************************************
    On Error GoTo ErroConexao
    Static intCont As Integer
    
    strNome_Servidor = GetSetting("Director", "Parâmetros", "Servidor", "Caminho não encontrado")
    strSenha_Admin = GetSetting("Director", "Parâmetros", "Senha", "Senha não encontrada")
    strUser_Admin = GetSetting("Director", "Parâmetros", "Usuário", "Usuário não encontrado")

    Set Conexao = New ADODB.Connection
        Conexao.ConnectionString = "Provider=SQLOLEDB.1;Password=" & strSenha_Admin & ";Persist Security Info=True;User ID=" & strUser_Admin & ";Initial Catalog=dbdirector;Data Source=" & strNome_Servidor & ";Connect Timeout=5"
        Conexao.Open
        
    
    intCont = 0
    Exit Function
    
ErroConexao:
    If Err.Number = -2147467259 Then
        If intCont < 3 Then
            Dim msgResposta As VbMsgBoxResult
            
            msgResposta = MsgBox("Não foi possível encontrar o Servidor de Dados. " & _
                                 "O nome dele pode estar errado ou ele pode estar desligado ou o serviço desconectado." & vbNewLine & _
                                 "Deseja tentar a conexao novamente.", vbCritical + vbSystemModal + vbYesNo, "Director")
            
            If msgResposta = vbYes Then
                intCont = intCont + 1
                Call Banco_Dados.Abre_Conexao(Conexao)
            Else
                End
            End If
        Else
            MsgBox "Não foi possível encontrar o Servidor de Dados. " & _
                   "Ele pode estar desligado ou o serviço desconectado." & vbNewLine & _
                   "Contate o Administrador da Rede.", vbCritical + vbSystemModal, "Director"
            End
        End If
    
    Else
        Call erro.erro("Abre_Conexao")
        End
    End If
End Function

Public Function Fecha_RecordSet(Recordset_Memoria As ADODB.Recordset) As String
    'Objeto/classe correspondente......: PIBDados.Fecha_Recordset
    On Error Resume Next
        'Fecha a RecordSet
        Recordset_Memoria.Close
        'Destroi a variavel de memoria
    Set Recordset_Memoria = Nothing
End Function


Public Function Condicao_String(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Condicao As String, Operador As String, strString As String, Recordset_Memoria As ADODB.Recordset) As String
    'Objeto/classe correspondente......: PIBDados.SQL_Geral
    Dim strSQL As String
    
    On Error GoTo erro
    'Sua SQL de pesquisa
    strSQL = ""
    strSQL = strSQL & "SELECT " & Nome_Campos & " "
    strSQL = strSQL & "FROM " & Nome_Tabela & " "
    strSQL = strSQL & "WHERE " & Nome_Campo_Condicao & " "
    strSQL = strSQL & "" & Operador & "'" & strString & "'"
                
    'Passando a copia da Tabela para a RecordSet indicada
    Set Recordset_Memoria = New ADODB.Recordset
        Recordset_Memoria.CursorLocation = adUseClient
        Recordset_Memoria.Open strSQL, CNConexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    Exit Function
    
erro:
    Call erro.erro("Condicao_String")
    Resume Next
End Function


Public Function Condicao_String_CMB(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Condicao As String, Operador As String, Combo As ComboBox, Recordset_Memoria As ADODB.Recordset) As String
    '******************************************************************************
    'Módulo............................: Banco_Dados
    'Procedimento/Função...............: Condicao_String_CMB
    'Objeto/classe correspondente......: -
    'Objetivo:.........................: Executa uma SQL em CNConexao realizando um
    '                                    filtro de acordo com o conteúdo de uma combo-box
    '                                    e gera um recordset
    'Desenvolvimento...................: Pablo Souza, Eduardo Cruz
    'Data de criação...................:
    'Utilização........................: Call Banco_Dados.Abre_Conexao(v1,v3,v3,v4, v5, v6)
    'Parâmetros de entrada.............: Variável v1, v2, v3 e v4 do tipo string, v5 do tipo Combo e
    '                                    v6 do tipo ADODB.Recordset
    'Saída.............................: Retorna um recordset na variável v6, fruto da SQL passada
    '                                    para a função
    'Data da última manutenção.........:
    'Manutenção executada por..........:
    'Observações.......................: Análoga a Condicao_String_ADB
    '
    '******************************************************************************

    Dim strSQL As String
    
    On Error GoTo erro
    strSQL = ""
    strSQL = strSQL & "SELECT " & Nome_Campos & " "
    strSQL = strSQL & "FROM " & Nome_Tabela & " "
    strSQL = strSQL & "WHERE " & Nome_Campo_Condicao & " "
    strSQL = strSQL & "" & Operador & " '" & Combo & "'"
    
    Set Recordset_Memoria = New ADODB.Recordset
        Recordset_Memoria.CursorLocation = adUseClient
        Recordset_Memoria.Open strSQL, CNConexao, adOpenKeyset, adLockOptimistic, adCmdText
        
    Exit Function
    
erro:
    Call erro.erro("Condicao_String_CMB")
    Resume Next
End Function

Public Function Condicao_Integer(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Condicao As String, Operador As String, Variavel As Integer, Recordset_Memoria As ADODB.Recordset) As String
    '******************************************************************************
    'Módulo............................: Banco_Dados
    'Procedimento/Função...............: Condicao_Integer
    'Objeto/classe correspondente......: PIBDados.SQL_Geral
    'Objetivo:.........................: Executa uma SQL contra uma Conexao e gera um recordset
    'Desenvolvimento...................: Pablo Souza, Eduardo Cruz
    'Data de criação...................: 07/04/2000
    'Utilização........................: Call Banco_Dados.Condicao_Integer(v1, v2, v3, v4, v5, v6)
    'Parâmetros de entrada.............: Variável v1, v2, v3, v4 do tipo string, v5 do tipo integer e v6 ADODB.Recordset
    'Saída.............................: Retorna um recordset na variável v6, fruto da SQL passada
    '                                    para a função
    'Data da última manutenção.........:
    'Manutenção executada por..........:
    'Observações.......................: A SQL executada tem formato fixo do tipo
    '                                    SELECT v1 FROM v2 WHERE v3 <v4> v5, onde v3 é uma condição
    '                                    Veja também Banco_Dados.SQL_Geral/Condicao_Integer/
    '                                    Veja também:
    '                                    Condicao_Numerico/Condicao_String/Condicao_String_ADB/
    '                                    Condicao_String_CMB
    '                                    Pode ser substituída por SQL_Geral
    '******************************************************************************
    
    Dim strSQL As String
    
    On Error GoTo erro
    strSQL = ""
    strSQL = strSQL & "SELECT " & Nome_Campos & " "
    strSQL = strSQL & "FROM " & Nome_Tabela & " "
    strSQL = strSQL & "WHERE " & Nome_Campo_Condicao & " "
    strSQL = strSQL & "" & Operador & " " & Variavel & ""
        
    Set Recordset_Memoria = New ADODB.Recordset
        Recordset_Memoria.CursorLocation = adUseClient
        Recordset_Memoria.Open strSQL, CNConexao, adOpenKeyset, adLockOptimistic, adCmdText
    Exit Function
    
erro:
    Call erro.erro("Condicao_Integer")
    Resume Next
End Function

Public Function Fecha_Conexao(Nome_Conexao As ADODB.Connection) As String
    'Objeto/classe correspondente......: PIBDados.Fecha_Conexao
    On Error Resume Next
        Nome_Conexao.Close
    Set Nome_Conexao = Nothing
End Function

Public Function SQLgeral(strSQL_comando As String, Recordset_Memoria As ADODB.Recordset)
    On Error GoTo erro
    Set Recordset_Memoria = New ADODB.Recordset
        Recordset_Memoria.CursorLocation = adUseClient
        Recordset_Memoria.Open strSQL_comando, CNConexao, adOpenKeyset, adLockOptimistic, adCmdText
    Exit Function
erro:
    Call erro.erro("SQLgeral")
    
    Resume Next
End Function

Public Function Acessibilidade(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Ordem As String, Recordset_Memoria As ADODB.Recordset, Conexao As ADODB.Connection) As String
    '******************************************************************************
    'Módulo............................: Banco_Dados
    'Procedimento/Função...............: Acessibilidade
    'Objeto/classe correspondente......: PIBDados.SQL_Geral
    'Objetivo:.........................: Executa uma SQL em CNConexao e gera um recordset
    'Desenvolvimento...................: Pablo Souza, Eduardo Cruz
    'Data de criação...................: 07/04/2000
    'Utilização........................: Call Banco_Dados.Abre_Conexao(v1,v3,v3,v4)
    'Parâmetros de entrada.............: Variável v1, v2 e v3 do tipo string, v4 do tipo ADODB.Recordset
    'Saída.............................: Retorna um recordset na variável v4, fruto da SQL passada
    '                                    para a função
    'Data da última manutenção.........:
    'Manutenção executada por..........:
    'Observações.......................: A SQL executada tem formato fixo do tipo
    '                                    SELECT v1 FROM v2 ORDER BY v3
    '                                    Veja também Banco_Dados.SQL_Geral/Condicao_Integer/
    '                                    Condicao_Numerico/Condicao_String/Condicao_String_ADB/
    '                                    Condicao_String_CMB
    '                                    Pode ser substituída por SQL_Geral
    '******************************************************************************
    Dim strSQL As String
    
    On Error GoTo erro
    strSQL = ""
    strSQL = strSQL & "SELECT " & Nome_Campos & " "
    strSQL = strSQL & "FROM " & Nome_Tabela & " "
    strSQL = strSQL & "ORDER BY " & Nome_Campo_Ordem & ""
    
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open strSQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
    Set Recordset_Memoria = TBrecordset
    
    Set TBrecordset = Nothing
    Exit Function
erro:
    Call erro.erro("Acessibilidade")
    End
End Function

Public Function Integridade(Deletar_SN As Boolean, Optional Nome_Stored_Procedure As String, Optional Parametro_Stored_Procedure As String) As Boolean
    On Error GoTo erro
    
    If Deletar_SN = True Then

        Dim adrIntegridade As ADODB.Recordset
        Dim strResp As String

        strResp = MsgBox("Serão excluidos todos os relacionamentos existentes. Confirma a operação?", vbQuestion + vbYesNo, "Director")

        If strResp = vbYes Then
            CNConexao.BeginTrans
            On Error GoTo Erro_Conexao
            DoEvents
            
            Set adrIntegridade = New ADODB.Recordset
                adrIntegridade.CursorLocation = adUseClient
                adrIntegridade.Open "{ Call " & Nome_Stored_Procedure & " } '" & Parametro_Stored_Procedure & "' ", CNConexao, adOpenKeyset, adLockOptimistic
            
            'Comentario colocado por
            'Eduardo Cruz - 24/04/2001
            If booExcluir = True Then
                CNConexao.CommitTrans
                Integridade = True
            Else
                CNConexao.RollbackTrans
                Integridade = False
            End If
        Else
            Integridade = False
        End If
    Else
        MsgBox "Deverão ser excluidos todos os itens existentes relacionados a esta.", vbExclamation, "Director"
        Integridade = False
    End If

    Exit Function
    
Erro_Conexao:
    CNConexao.RollbackTrans
erro:
    If Err.Number = -2147217900 Then
        MsgBox "Exclusão não permitida. Registro relacionado a outras tabelas.", vbCritical, "Director"
        Integridade = False
        Exit Function
    Else
        Call erro.erro("Integridade")
    End If
    Resume Next
End Function

Sub Movimentar_Item(Cod_Empresa As Integer, Cod_Cliente_Fornecedor As Integer, Cod_Tipo_Docto As Integer, Numero_Docto As String, Tipo_Movto As String, Data_Movto As Date, Cod_Deposito As Integer, Cod_Item As Integer, Cod_Unidade_Armazenagem As Integer, Qtde_Entrada_Saida As Single, Optional Valor As String = "0,00", Optional ID_Movto As Integer = 0)
    '******************************************************************************
    'Sistema...........................: Director
    'Módulo............................: Estoque
    'Procedimento/Função...............: Movimentar_Item
    'Objetivo:.........................: Efetuar todas as transações necessárias para
    '                                    a entrada ou saída de um item no estoque
    '                                    (cálculo de custo médio, atualização do total
    '                                    do estoque, inclusão ou alteração de movimento
    '                                    de estoque, inclusão ou alteração de item de
    '                                    movimento de estoque e atualização de última
    '                                    data de entrada e saída)
    'Desenvolvimento...................: Fernando Souza
    'Data de criação...................: 13/02/2001
    'Data da última manutenção.........: 07/03/2001
    'Manutenção executada por..........: FernandoS
    'Observaçãoes......................:
    '******************************************************************************
    Dim strSQL As String
    
    strSQL = _
        "DECLARE " & _
            "/* variaveis para movimento */ " & _
            "@ID_Movto INT, " & _
            "@Cod_Empresa INT, " & _
            "@Cod_Tipo_Docto INT, " & _
            "@Cod_Cliente_Fornecedor INT, " & _
            "@Numero_Docto NVARCHAR(50), " & _
            "@Tipo_Movto VARCHAR(1), " & _
            "@Data_Movto SMALLDATETIME, " & _
        "" & _
            "/* variaveis para itens de movimento */ " & _
            "@ID_Deposito_Item INT, " & _
            "@Cod_Deposito INT, " & _
            "@Cod_Item INT, " & _
            "@Cod_Unidade_Armazenagem INT, " & _
            "@Qtde_Entrada_Saida DECIMAL(18,4), " & _
            "@Valor MONEY, " & _
        "" & _
            "/* variaveis para contole de estoque */ " & _
            "@Qtde_Estoque DECIMAL(18,4), " & _
            "@Diferenca DECIMAL(18,4) "

    strSQL = strSQL & _
        "SET @ID_Movto = " & ID_Movto & " " & _
        "SET @Cod_Empresa = " & Cod_Empresa & " " & _
        "SET @Cod_Tipo_Docto = " & Cod_Tipo_Docto & " " & _
        "SET @Cod_Cliente_Fornecedor = " & Cod_Cliente_Fornecedor & " " & _
        "SET @Numero_Docto = " & Numero_Docto & " " & _
        "SET @Tipo_Movto = '" & Tipo_Movto & "' " & _
        "SET @Data_Movto = '" & Format(Data_Movto, "yyyyMMdd") & "' " & _
        "SET @Cod_Deposito = " & Cod_Deposito & " " & _
        "SET @Cod_Item = " & Cod_Item & " " & _
        "SET @Cod_Unidade_Armazenagem = " & Cod_Unidade_Armazenagem & " " & _
        "SET @Qtde_Entrada_Saida = " & Funcoes_Gerais.Grava_Moeda(CStr(Qtde_Entrada_Saida)) & " " & _
        "SET @Valor = " & Funcoes_Gerais.Grava_Moeda(Valor) & " "

    strSQL = strSQL & _
        "/* id depósito */ " & _
        "IF NOT EXISTS (SELECT DFid_deposito_item FROM TBdeposito_item WHERE DFcod_deposito = @Cod_Deposito AND DFcod_item_estoque = @Cod_Item AND DFcod_unidade_armazenagem = @Cod_Unidade_Armazenagem) " & _
            "INSERT INTO TBdeposito_item (DFcod_deposito, DFcod_item_estoque, DFcod_unidade_armazenagem, DFquantidade_estoque, DFquantidade_estoque_reservado, DFlocalizacao) " & _
            "SELECT @Cod_Deposito, @Cod_Item, @Cod_Unidade_Armazenagem, 0, 0, '' " & _
        "SET @ID_Deposito_Item = (SELECT DFid_deposito_item FROM TBdeposito_item WHERE DFcod_deposito = @Cod_Deposito AND DFcod_item_estoque = @Cod_Item AND DFcod_unidade_armazenagem = @Cod_Unidade_Armazenagem) " & _
        "" & _
        "/* qtde total em deposito */ " & _
        "SET @Qtde_Estoque = (SELECT SUM((CONVERT(DECIMAL,TBunidade.DFfator_conversao) / CONVERT(DECIMAL,TBunidade_padrao.DFfator_conversao)) * TBdeposito_item.DFquantidade_estoque) AS DFtotal_estoque " & _
                             "FROM TBunidade " & _
                                "INNER JOIN TBunidade AS TBunidade_padrao " & _
                                "INNER JOIN TBdeposito_item " & _
                                "INNER JOIN TBitem_estoque " & _
                                    "ON TBitem_estoque.DFcod_item_estoque = TBdeposito_item.DFcod_item_estoque " & _
                                    "ON TBitem_estoque.DFcod_unidade_compra = TBunidade_padrao.DFcod_unidade " & _
                                    "ON TBdeposito_item.DFcod_unidade_armazenagem = TBunidade.DFcod_unidade " & _
                             "WHERE TBitem_estoque.DFcod_item_estoque = @Cod_Item) "

    strSQL = strSQL & _
        "/* movimento */ " & _
        "IF @ID_Movto = 0 " & _
         "BEGIN " & _
           "SET @ID_Movto = ISNULL((SELECT DFid_movimento_estoque FROM TBmovimento_estoque WHERE DFcod_empresa = @Cod_Empresa AND DFcod_tipo_documento = @Cod_Tipo_Docto AND DFnumero_docto = @Numero_Docto AND DFtipo_movto_estoque = @Tipo_Movto),0) " & _
           "IF @ID_Movto = 0 " & _
            "BEGIN " & _
               "INSERT INTO TBmovimento_estoque (DFcod_empresa, DFcod_tipo_documento, DFnumero_docto, DFtipo_movto_estoque, DFdata_movimento) " & _
               "SELECT @Cod_Empresa, @Cod_Tipo_Docto, @Numero_Docto, @Tipo_Movto, @Data_Movto " & _
               "SET @ID_Movto = (SELECT MAX(DFid_movimento_estoque) FROM TBmovimento_estoque) " & _
               "IF @Tipo_Movto = 'E' " & _
                   "INSERT INTO TBmovimento_estoque_fornecedor (DFid_movimento_estoque, DFcod_fornecedor) " & _
                   "SELECT @ID_Movto, @Cod_Cliente_Fornecedor " & _
               "ELSE " & _
                   "INSERT INTO TBmovimento_estoque_cliente (DFid_movimento_estoque, DFcod_cliente) " & _
                   "SELECT @ID_Movto, @Cod_Cliente_Fornecedor " & _
            "END " & _
         "END "

    strSQL = strSQL & _
        "ELSE " & _
         "BEGIN " & _
            "UPDATE TBmovimento_estoque " & _
            "SET DFcod_empresa = @Cod_Empresa, DFcod_tipo_documento = @Cod_Tipo_Docto, DFnumero_docto = @Numero_Docto, DFtipo_movto_estoque = @Tipo_Movto, DFdata_movimento = @Data_Movto " & _
            "WHERE DFid_movimento_estoque = @ID_Movto " & _
            "IF @Tipo_Movto = 'E' " & _
                "UPDATE TBmovimento_estoque_fornecedor " & _
                "SET DFcod_fornecedor = @Cod_Cliente_Fornecedor " & _
                "WHERE DFid_movimento_estoque = @ID_Movto " & _
            "ELSE " & _
                "UPDATE TBmovimento_estoque_cliente " & _
                "SET DFcod_cliente = @Cod_Cliente_Fornecedor " & _
                "WHERE DFid_movimento_estoque = @ID_Movto " & _
         "END "

    strSQL = strSQL & _
        "/* item movimento */ " & _
        "IF EXISTS (SELECT DFid_item_movto_estoque FROM TBitem_movto_estoque WHERE DFid_movimento_estoque = @ID_Movto AND DFid_deposito_item = @ID_Deposito_Item) " & _
         "BEGIN " & _
            "SET @Diferenca = (SELECT @Qtde_Entrada_Saida - DFqtde FROM TBitem_movto_estoque WHERE DFid_movimento_estoque = @ID_Movto AND DFid_deposito_item = @ID_Deposito_Item) " & _
            "IF @Tipo_Movto = 'E' " & _
             "BEGIN " & _
                "/* custo medio */ " & _
                "IF @Diferenca > 0 " & _
                    "UPDATE TBitem_estoque " & _
                    "SET DFcusto_medio = (IsNull((DFcusto_medio * @Qtde_Estoque), 0) + (DFpreco_custo * @Diferenca)) / (@Diferenca  + @Qtde_Estoque) " & _
                    "WHERE DFcod_item_estoque = @Cod_Item " & _
                "UPDATE TBdeposito_item " & _
                "SET DFquantidade_estoque = ISNULL(DFquantidade_estoque,0) + @Diferenca " & _
                "WHERE TBdeposito_item.DFid_deposito_item = @ID_Deposito_Item " & _
             "END " & _
            "ELSE " & _
                "UPDATE TBdeposito_item " & _
                "SET DFquantidade_estoque = ISNULL(DFquantidade_estoque,0) - @Diferenca " & _
                "WHERE TBdeposito_item.DFid_deposito_item = @ID_Deposito_Item " & _
            "UPDATE TBitem_movto_estoque " & _
            "SET DFqtde = @Qtde_Entrada_Saida , DFvalor = @Valor " & _
            "WHERE DFid_movimento_estoque = @ID_Movto AND DFid_deposito_item = @ID_Deposito_Item " & _
         "END "

    strSQL = strSQL & _
        "ELSE " & _
         "BEGIN " & _
            "INSERT INTO TBitem_movto_estoque (DFid_movimento_estoque, DFid_deposito_item, DFqtde, DFvalor) " & _
            "SELECT @ID_Movto, @ID_Deposito_Item, @Qtde_Entrada_Saida, @Valor " & _
            "IF @Tipo_Movto = 'E' " & _
             "BEGIN " & _
                "/* custo medio */ " & _
                "Update TBitem_estoque " & _
                "SET DFcusto_medio = (IsNull((DFcusto_medio * @Qtde_Estoque), 0) + (DFpreco_custo * @Qtde_Entrada_Saida)) / (@Qtde_Entrada_Saida  + @Qtde_Estoque) " & _
                "WHERE DFcod_item_estoque = @Cod_Item " & _
                "UPDATE TBdeposito_item " & _
                "SET DFquantidade_estoque = ISNULL(DFquantidade_estoque,0) + @Qtde_Entrada_Saida " & _
                "WHERE TBdeposito_item.DFid_deposito_item = @ID_Deposito_Item " & _
             "END " & _
            "ELSE " & _
                "UPDATE TBdeposito_item " & _
                "SET DFquantidade_estoque = ISNULL(DFquantidade_estoque,0) - @Qtde_Entrada_Saida " & _
                "WHERE TBdeposito_item.DFid_deposito_item = @ID_Deposito_Item " & _
         "END "

    strSQL = strSQL & _
        "/* última data entrada/saida*/ " & _
        "UPDATE TBitem_estoque " & _
        "SET DFdata_ultima_saida = @Data_Movto " & _
        "WHERE DFcod_item_estoque = @Cod_Item AND @Data_Movto > DFdata_ultima_saida " & _
        "" & _
        "UPDATE TBitem_estoque " & _
        "SET DFdata_ultima_entrada = @Data_Movto " & _
        "WHERE DFcod_item_estoque = @Cod_Item AND @Data_Movto > DFdata_ultima_entrada " & _
        "" & _
        "/* movendo o id do movimento para uma tabela temporária */ " & _
        "SELECT @ID_Movto AS DFid_movto INTO TBtemp_movto "
    CNConexao.Execute strSQL

    If ID_Movto = 0 Then
        ID_Movto = CNConexao.Execute("SELECT DFid_movto FROM TBtemp_movto").Fields("DFid_movto")
    End If
    CNConexao.Execute "DROP TABLE TBtemp_movto"

End Sub

Public Sub Executa_Stored_Procedure(Nome_Stored_Procedure As String, Optional Parametro_Stored_Procedure)
    On Error GoTo erro
    Dim adrStored_Procedure As ADODB.Recordset
    
    Set adrStored_Procedure = New ADODB.Recordset
        adrStored_Procedure.CursorLocation = adUseClient
        adrStored_Procedure.Open "{ Call " & Nome_Stored_Procedure & " } '" & Parametro_Stored_Procedure & "' ", CNConexao, adOpenKeyset, adLockOptimistic
    
        adrStored_Procedure.Close
    Set adrStored_Procedure = Nothing
    
    Exit Sub
erro:
    Call erro.erro("Executa_Stored_Procedure")
End Sub
