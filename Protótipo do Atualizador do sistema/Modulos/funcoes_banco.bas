Attribute VB_Name = "funcoes_banco"
Option Explicit
Dim conexao_banco As New DLLConexao_Sistema.conexao
Dim strsql As String

Public Sub Gravar(Tabela As String, campos_tabela As String, valores_campos As String, Aplicacao As String, Form As Object, Optional Banco As String)

    'Indicando o banco à conectar-se
    conexao_banco.Initial_Catalog = Banco
    
    'Estabelecendo conexão com o banco
    conexao_banco.Abrir_conexao (Aplicacao)
    
    'Indica o inicio da transação junto o banco
    conexao_banco.CNConexao.BeginTrans
    
    On Error GoTo Erro
    
    strsql = "INSERT INTO " & Tabela & " (" & campos_tabela & ") " & _
             "SELECT " & valores_campos & " "
                                     
    conexao_banco.CNConexao.Execute strsql
                                  
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
    
    strsql = "DELETE FROM " & Tabela & " " & _
             "WHERE " & campo_identificador & " = '" & valor_campo_identificador & "' "
    
    If Nome_Campo_Codigo_Empresa <> Empty Then
       strsql = strsql + "AND " & Nome_Campo_Codigo_Empresa & " = " & Valor_Campo_Codigo_Empresa & " "
    End If
    
    conexao_banco.CNConexao.Execute strsql
                                  
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
    
    strsql = "UPDATE " & Tabela & " " & _
             " " & Valores_SET & " " & _
             "WHERE " & campo_identificador & " = '" & valor_campo_identificador & "' "
             
    If Nome_Campo_Codigo_Empresa <> Empty Then
       strsql = strsql + "AND " & Nome_Campo_Codigo_Empresa & " = " & Valor_Campo_Codigo_Empresa & " "
    End If
                                     
    conexao_banco.CNConexao.Execute strsql
                                  
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
