VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Grupo Mil - D.S.I - Departamento de Sistemas de Informação                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Movimentações da Baixa                                         '
' Módulo.................: Controle de Balanças                                           '
' Objetivo...............: Projeto de classe                                              '
' Data de Criação........: 04/09/2002                                                     '
' Análista Responsável...: Marcos Baião                                                   '
' Programação............: Marcos Baião                                                   '
' Última Manutenção......:                                                                '
' Data última manutenção.:                                                                '
' Observação.............: Public Property Let --> Recebe um valor para a propriedade     '
'                          (em run time)                                                  '
'                          Public Property Get --> Retorna para a aplicação um valor      '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private mvarcodigo_usuario As Integer 'local copy
Private mvarnome_usuario As String 'local copy
Private mvarsenha_usuario As String 'local copy
Private conexao_usuario As clsConexao
Public usuario_logado As String

Public Function alterar_usuario()
    
    If mvarcodigo_usuario = " " Then
       Err.Raise vbObjectError + 1000, "clsUsuario", "Usuário não Cadastrado"
    End If
    If mvarnome_usuario = " " Then
       Err.Raise vbObjectError + 1001, "clsUsuario", "Nome do usuário não Cadastrado"
    End If
    If mvarsenha_usuario = " " Then
       Err.Raise vbObjectError + 1002, "clsUsuario", "Senha do usuário não Cadastrada"
    End If
    
    On Error GoTo Erro
    
    conexao_usuario.Abrir_conexao
    
    conexao_usuario.CNConexao.BeginTrans
    
    conexao_usuario.CNConexao.Execute "UPDATE TBusuario SET codigo_usuario = mvarcodigo_usuario,nome_usuario = mvarnome_usuario,senha_usuario = mvarsenha_usuario " & _
                              "WHERE codigo_usuario = mvarcodigo_usuario"
    'Inserir um LOG
    'Call Funcoes_Gerais.Gravar_Log(Caption_interface, txtcodigo.Text, "Gravação", frmLogin.strUsuario_Sistema, "Usuário Incluiu/Alterou um registro")
    
    conexao_usuario.CNConexao.CommitTrans
    
    conexao_usuario.CNConexao.Fechar_conexao
    
    DoEvents
       
    Exit Function
    
Erro:
    conexao_usuario.CNConexao.RoolbackTrans
    conexao_usuario.Fechar_conexao
    Err.Raise vbObjectError + 1004, "clsUsuario", "Erros ao Alterar um novo usuário"
End Function

Public Function excluir_usuario()
    On Error GoTo Erro
    
    conexao_usuario.Abrir_conexao
    
    conexao_usuario.CNConexao.BeginTrans
    
    conexao_usuario.CNConexao.Execute "DELETE FROM TBusuario(codigo_usuario,nome_usuario,senha_usuario) " & _
                              "SELECT mvarcodigo_usuario, mvarnome_usuario,mvarsenha_usuario"
    'Inserir um LOG
    'Call Funcoes_Gerais.Gravar_Log(Caption_interface, txtcodigo.Text, "Gravação", frmLogin.strUsuario_Sistema, "Usuário Incluiu/Alterou um registro")
    
    conexao_usuario.CNConexao.CommitTrans
    
    conexao_usuario.Fechar_conexao
    
    DoEvents
    
Erro:
    conexao_usuario.CNConexao.RoolbackTrans
    conexao_usuario.Fechar_conexao
    Err.Raise vbObjectError + 1005, "clsUsuario", "Erros ao excluir um novo usuário"
End Function

Public Function gravar_usuario()
        
    If mvarcodigo_usuario = " " Then
       Err.Raise vbObjectError + 1000, "clsUsuario", "Usuário não Cadastrado"
    End If
    If mvarnome_usuario = " " Then
       Err.Raise vbObjectError + 1001, "clsUsuario", "Nome do usuário não Cadastrado"
    End If
    If mvarsenha_usuario = " " Then
       Err.Raise vbObjectError + 1002, "clsUsuario", "Senha do usuário não Cadastrada"
    End If
    
    On Error GoTo Erro
    
    conexao_usuario.Abrir_conexao
    
    conexao_usuario.CNConexao.BeginTrans
    
    conexao_usuario.CNConexao.Execute "INSERT INTO TBusuario(codigo_usuario,nome_usuario,senha_usuario) " & _
                              "SELECT mvarcodigo_usuario, mvarnome_usuario,mvarsenha_usuario"
    'Inserir um LOG
    'Call Funcoes_Gerais.Gravar_Log(Caption_interface, txtcodigo.Text, "Gravação", frmLogin.strUsuario_Sistema, "Usuário Incluiu/Alterou um registro")
    
    conexao_usuario.CNConexao.CommitTrans
    conexao_usuario.Fechar_conexao
    
    DoEvents
       
    Exit Function
    
Erro:
    conexao_usuario.CNConexao.RoolbackTrans
    conexao_usuario.Fechar_conexao
    Err.Raise vbObjectError + 1003, "clsUsuario", "Erros ao Incluir um novo usuário"
    
End Function

Public Function consultar_usuario() As String

End Function

Public Function validar_usuario() As Boolean
    On Error GoTo Erro
    
    If mvarnome_usuario = " " Then
       Err.Raise vbObjectError + 1011, "clsUsuario", "Nome do usuário não Informado"
    End If
    
    conexao_usuario.Abrir_conexao
       
    Dim rstUsuario As New ADODB.Recordset
    
    rstUsuario.CursorLocation = adUseClient
    rstUsuario.Open "SELECT codigo_usuario FROM TBusuario WHERE nome_usuario = '" & mvarnome_usuario & "'", conexao_usuario.CNConexao, adOpenStatic, adLockReadOnly
    
    'Se a recordset retornar vazia ele retorna false para o client indicando que o usuario não
    'esta cadastrado
    If rstUsuario.BOF And rstUsuario.EOF Then
       validar_usuario = False
    Else
       validar_usuario = True
    End If
    
    rstUsuario.Close
    Set rstUsuario = Nothing
    
    conexao_usuario.Fechar_conexao
    
    Exit Function
    
Erro:
    conexao_usuario.Fechar_conexao
    Err.Raise vbObjectError + 1020, "clsUsuario", "Erros ao tentar validar o usuário"
    
End Function
Public Function validar_senha() As Boolean

    Dim rstSenha As New ADODB.Recordset
    
    On Error GoTo Erro
    
    If mvarsenha_usuario = " " Then
       Err.Raise vbObjectError + 1012, "clsUsuario", "Senha não confere"
    End If
    conexao_usuario.Abrir_conexao
   
    rstSenha.CursorLocation = adUseClient
    rstSenha.Open "SELECT codigo_usuario FROM TBusuario WHERE nome_usuario = '" & mvarnome_usuario & "' AND senha_usuario = " & mvarsenha_usuario & " ", conexao_usuario.CNConexao, adOpenStatic, adLockReadOnly
    
    'Se a recordset retornar vazia ele retorna false para o client indicando que a senha não
    'esta cadastrada
    If rstSenha.BOF And rstSenha.EOF Then
       validar_senha = False
    Else
       validar_senha = True
    End If
    
    rstSenha.Close
    Set rstSenha = Nothing
        
    conexao_usuario.Fechar_conexao
    
    DoEvents
    
    Exit Function
    
Erro:
    conexao_usuario.Fechar_conexao
    Err.Raise vbObjectError + 1020, "clsUsuario", "Erros ao tentar validar o senha"
    
End Function

Public Property Let senha_usuario(ByVal vData As String)
    mvarsenha_usuario = vData
End Property

Public Property Get senha_usuario() As String
    senha_usuario = mvarsenha_usuario
End Property

Public Property Let nome_usuario(ByVal vData As String)
'Recebendo um valor para apropriedade
    mvarnome_usuario = vData
End Property

Public Property Get nome_usuario() As String
'Retornando um valor para a aplicação
    nome_usuario = mvarnome_usuario
End Property

Public Property Let codigo_usuario(ByVal vData As Integer)
    mvarcodigo_usuario = vData
End Property

Public Property Get codigo_usuario() As Integer
    codigo_usuario = mvarcodigo_usuario
End Property

Private Sub Class_Initialize()

    On Error GoTo Erro
    
    Set conexao_usuario = CreateObject("DLLBalanca.clsConexao")
    
    conexao_usuario.Provider = "SQLOLEDB.1"
    conexao_usuario.Data_Source = "INFO-028"
    conexao_usuario.Initial_Catalog = "BDBalanca"
    conexao_usuario.User_ID = "sa"
    
    Exit Sub
    
Erro:
    Err.Raise vbObjectError + 1006, "clsUsuario", "Erro ao Inicializar este objeto"
End Sub

Private Sub Class_Terminate()
    Set conexao_usuario = Nothing
End Sub
