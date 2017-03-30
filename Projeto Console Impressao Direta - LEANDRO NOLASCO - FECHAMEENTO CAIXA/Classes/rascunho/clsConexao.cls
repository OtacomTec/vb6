VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "clsConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarProvider As String 'local copy
Private mvarData_Source As String 'local copy
Private mvarUser_ID As String 'local copy
Private mvarInitial_Catalog As String 'local copy
Public CNConexao As New ADODB.Connection

Public Sub Fechar_conexao()
'Método usado para fechar conexão com o banco
    On Error GoTo Erro
    
    CNConexao.Close
    Set CNConexao = Nothing
    
    Exit Sub
Erro:
    Err.Raise vbObjectError + 1002, "clsConexao", "Erros ao conectar com o provedor de dados! "

End Sub
Public Sub Abrir_conexao()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'D.S.I - Departamento de sistemas de informação - Marcos Baião                                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Este método foi criado para estabelecer conexão com o banco,nele existem os seguintes parametros:  '
'                                                                                                   '
'Provider    --> Usado para estabelecer o provedor da conexão. Ex: SQL Server,Oracle,Interbase,etc..'
'Data_Source --> Parâmetro que indica o servidor de dados(Ex: Sql Server ou Oracle) ou o caminho    '
'                do arquivo(Ex: .mdb do Access)                                                     '
'User_ID /   --> Parâmetros que indicam a senha do usuário que tenta o acesso e o banco à acessar   '
'Initial_Catalog * Parâmetros OPCIONAIS                                                             '
'                ** Mais usado no caso dos SGBD´s                                                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      On Error GoTo Erro
      Dim strConexao As String
      
      'Montará a string de conexão
      If mvarData_Source <> Empty Then
         strConexao = "Data Source=" & mvarData_Source & ";"
      End If
      If mvarInitial_Catalog <> Empty Then
         strConexao = strConexao + "Initial Catalog=" & mvarInitial_Catalog & ";"
      End If
      If mvarUser_ID <> Empty Then
         strConexao = strConexao + "User ID=" & mvarUser_ID & ";"
      End If
      strConexao = strConexao + "Persist Security Info=False"
      
      'Abrirá a conexão
      CNConexao.Provider = mvarProvider
      CNConexao.Open strConexao
      
      Exit Sub
      
Erro:
    Err.Raise vbObjectError + 1001, "clsConexao", "Erros ao conectar com o provedor de dados! "

End Sub

Public Property Let Initial_Catalog(ByVal vData As String)
    mvarInitial_Catalog = vData
End Property
Public Property Get Initial_Catalog() As String
    Initial_Catalog = mvarInitial_Catalog
End Property

Public Property Let User_ID(ByVal vData As String)
    mvarUser_ID = vData
End Property
Public Property Get User_ID() As String
    User_ID = mvarUser_ID
End Property

Public Property Let Data_Source(ByVal vData As String)
    mvarData_Source = vData
End Property
Public Property Get Data_Source() As String
    Data_Source = mvarData_Source
End Property

Public Property Let Provider(ByVal vData As String)
    mvarProvider = vData
End Property
Public Property Get Provider() As String
    Provider = mvarProvider
End Property
