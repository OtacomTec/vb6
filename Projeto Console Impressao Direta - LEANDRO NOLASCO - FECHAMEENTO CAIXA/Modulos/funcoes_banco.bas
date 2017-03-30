Attribute VB_Name = "funcoes_banco"
Option Explicit
Dim conexao_banco As New DLLConexao_Sistema.conexao
Dim conexao_banco_retaguarda As New DLLConexao_Sistema.conexao
Dim strSQL As String

Public Sub Gravar(Tabela As String, Campos_tabela As String, valores_campos As String, Aplicacao As String, Form As Object, Optional Banco As String)

    'Indicando o banco à conectar-se
    conexao_banco.Initial_Catalog = Banco
    
    'Estabelecendo conexão com o banco
    conexao_banco.Abrir_conexao (Aplicacao)
    
    'Indica o inicio da transação junto o banco
    conexao_banco.CNConexao.BeginTrans
    
    On Error GoTo Erro
    
    strSQL = "INSERT INTO " & Tabela & " (" & Campos_tabela & ") " & _
             "SELECT " & valores_campos & " "
                                     
    conexao_banco.CNConexao.Execute strSQL
                                  
    'Indica o sucesso da transação do banco
    conexao_banco.CNConexao.CommitTrans
    
    'Fecha a conexão com o banco
    conexao_banco.Fechar_conexao
    
    DoEvents
       
    Exit Sub
    
Erro:
        
    Call Erro.Erro(Form, Aplicacao, "Gravar")
    
    'Indica o fracasso da transação do banco
    conexao_banco.CNConexao.RollbackTrans
    
    'Fecha a conexão com o banco
    conexao_banco.Fechar_conexao
    
End Sub

Public Sub Excluir(Tabela As String, campo_identificador As String, valor_campo_identificador As String, Aplicacao As String, Form As Object, Banco As String, Optional Nome_Campo_Codigo_Empresa As String, Optional Valor_Campo_Codigo_Empresa As String)

    On Error GoTo Erro
  
    'Indicando o banco à conectar-se
    conexao_banco.Initial_Catalog = Banco
      
    'Estabelecendo conexão com o banco
    conexao_banco.Abrir_conexao (Aplicacao)
        
    'Indica o inicio da transação junto o banco
    conexao_banco.CNConexao.BeginTrans
    
    strSQL = "DELETE FROM " & Tabela & " " & _
             "WHERE " & campo_identificador & " = '" & valor_campo_identificador & "' "
    
    If Nome_Campo_Codigo_Empresa <> Empty Then
       strSQL = strSQL + "AND " & Nome_Campo_Codigo_Empresa & " = " & Valor_Campo_Codigo_Empresa & " "
    End If
    
    conexao_banco.CNConexao.Execute strSQL
                                  
    'Indica o sucesso da transação do banco
    conexao_banco.CNConexao.CommitTrans
    
    'Fecha a conexão com o banco
    conexao_banco.Fechar_conexao
    
    DoEvents
       
    Exit Sub
    
Erro:
    
    Call Erro.Erro(Form, "Otica", "Excluir")
    
    'Indica o fracasso da transação do banco
    conexao_banco.CNConexao.RollbackTrans
    
    'Fecha a conexão com o banco
    conexao_banco.Fechar_conexao
    
End Sub

Public Sub Alterar(Tabela As String, Valores_SET As String, campo_identificador As String, valor_campo_identificador As String, Aplicacao As String, Form As Object, Optional Banco As String, Optional Nome_Campo_Codigo_Empresa As String, Optional Valor_Campo_Codigo_Empresa As String)
    
    'Indicando o banco à conectar-se
    conexao_banco.Initial_Catalog = Banco
    
    'Estabelecendo conexão com o banco
    conexao_banco.Abrir_conexao (Aplicacao)
    
    'Indica o inicio da transação junto o banco
    conexao_banco.CNConexao.BeginTrans
    
    On Error GoTo Erro
    
    strSQL = "UPDATE " & Tabela & " " & _
             " " & Valores_SET & " " & _
             "WHERE " & campo_identificador & " = '" & valor_campo_identificador & "' "
             
    If Nome_Campo_Codigo_Empresa <> Empty Then
       strSQL = strSQL + "AND " & Nome_Campo_Codigo_Empresa & " = " & Valor_Campo_Codigo_Empresa & " "
    End If
                                     
    conexao_banco.CNConexao.Execute strSQL
                                  
    'Indica o sucesso da transação do banco
    conexao_banco.CNConexao.CommitTrans
    
    'Fecha a conexão com o banco
    conexao_banco.Fechar_conexao
    
    DoEvents
       
    Exit Sub
    
Erro:
  
    Call Erro.Erro(Form, Aplicacao, "Alterar")
    
    'Indica o fracasso da transação do banco
    conexao_banco.CNConexao.RoolbackTrans
    
    'Fecha a conexão com o banco
    conexao_banco.Fechar_conexao
       
End Sub

Public Sub Gravar_Portal(Banco_portal As String, Tabela_portal As String, Campos_tabela_portal As String, Valores_campos_portal As String, Form As Object, Aplicacao_retaguarda As String, Banco_retaguarda As String, Tabela_retaguarda As String, Campo_Integracao_Retaguarda As String, Campo_comparacao_retaguarda As String, Valor_comparacao_retaguarda As String)

    'Indicando o banco à conectar-se
    conexao_banco.Initial_Catalog = Banco_portal

    'Estabelecendo conexão com o banco
    conexao_banco.Abrir_conexao ("Portal")

    'Indica o inicio da transação junto o banco
    conexao_banco.CNConexao.BeginTrans

    On Error GoTo Erro

    strSQL = "INSERT INTO " & Tabela_portal & " (" & Campos_tabela_portal & ") " & _
             "SELECT " & Valores_campos_portal & " "

    conexao_banco.CNConexao.Execute strSQL

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    'Retaguarda
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    conexao_banco_retaguarda.Initial_Catalog = Banco_retaguarda

    'Estabelecendo conexão com o banco
    conexao_banco_retaguarda.Abrir_conexao (Aplicacao_retaguarda)

    conexao_banco_retaguarda.CNConexao.BeginTrans
        
    strSQL = "UPDATE " & Tabela_retaguarda & " SET " & Campo_Integracao_Retaguarda & " = 1 WHERE " & Campo_comparacao_retaguarda & " = " & Valor_comparacao_retaguarda & " "
    
    conexao_banco_retaguarda.CNConexao.Execute strSQL

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'COMITANDO
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Indica o sucesso da transação do banco
    conexao_banco_retaguarda.CNConexao.CommitTrans

    'Indica o sucesso da transação do banco
    conexao_banco.CNConexao.CommitTrans

    'Fecha a conexão com o banco
    conexao_banco_retaguarda.Fechar_conexao

    conexao_banco.Fechar_conexao

    DoEvents

    Exit Sub

Erro:

    Call Erro.Erro(Form, "Portal", "Gravar")

    'Indica o fracasso da transação do banco
    conexao_banco.CNConexao.RollbackTrans
    conexao_banco_retaguarda.CNConexao.RollbackTrans

    'Fecha a conexão com o banco
    conexao_banco.Fechar_conexao
    conexao_banco_retaguarda.Fechar_conexao
    
End Sub

Public Sub Alterar_Portal(Banco_portal As String, Tabela_portal As String, Valores_SET_portal As String, campo_identificador_portal As String, valor_campo_identificador_portal As String, Campo_comparacao_retaguarda As String, Valor_comparacao_retaguarda As String, Aplicacao_retaguarda As String, Form As Object, Banco_retaguarda As String, Tabela_retaguarda As String, Campo_Integracao_Retaguarda As String, Optional Nome_Campo_Codigo_Empresa As String, Optional Valor_Campo_Codigo_Empresa As String)

    'Indicando o banco à conectar-se
    conexao_banco.Initial_Catalog = Banco_portal

    'Estabelecendo conexão com o banco
    conexao_banco.Abrir_conexao ("Portal")

    'Indica o inicio da transação junto o banco
    conexao_banco.CNConexao.BeginTrans

    strSQL = "UPDATE " & Tabela_portal & " " & _
             " " & Valores_SET_portal & " " & _
             "WHERE " & campo_identificador_portal & " = '" & valor_campo_identificador_portal & "' "

    If Nome_Campo_Codigo_Empresa <> Empty Then
       strSQL = strSQL + "AND " & Nome_Campo_Codigo_Empresa & " = " & Valor_Campo_Codigo_Empresa & " "
    End If

    conexao_banco.CNConexao.Execute strSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '
    'Retaguarda
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    conexao_banco_retaguarda.Initial_Catalog = Banco_retaguarda

    'Estabelecendo conexão com o banco
    conexao_banco_retaguarda.Abrir_conexao (Aplicacao_retaguarda)

    conexao_banco_retaguarda.CNConexao.BeginTrans

    strSQL = "UPDATE " & Tabela_retaguarda & " SET " & Campo_Integracao_Retaguarda & " = 1 WHERE " & Campo_comparacao_retaguarda & " = " & Valor_comparacao_retaguarda & " "

    conexao_banco_retaguarda.CNConexao.Execute strSQL

    'Indica o sucesso da transação do banco
    conexao_banco.CNConexao.CommitTrans
    conexao_banco_retaguarda.CNConexao.CommitTrans

    'Fecha a conexão com o banco
    conexao_banco.Fechar_conexao
    conexao_banco_retaguarda.Fechar_conexao

    DoEvents

    Exit Sub

Erro:

    Call Erro.Erro(Form, "Portal", "Alterar")

    'Indica o fracasso da transação do banco
    conexao_banco.CNConexao.RoolbackTrans
    conexao_banco_retaguarda.CNConexao.RollbackTrans

    'Fecha a conexão com o banco
    conexao_banco.Fechar_conexao
    conexao_banco_retaguarda.Fechar_conexao

End Sub
