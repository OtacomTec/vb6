VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Grupo Mil - D.S.I - Departamento de Sistemas de Informa��o                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Movimenta��es da Baixa                                         '
' M�dulo.................: Controle de Balan�as                                           '
' Objetivo...............: Projeto de classe                                              '
' Data de Cria��o........: 04/09/2002                                                     '
' An�lista Respons�vel...: Marcos Bai�o                                                   '
' Programa��o............: Marcos Bai�o                                                   '
' �ltima Manuten��o......:                                                                '
' Data �ltima manuten��o.:                                                                '
' Observa��o.............: Public Property Let --> Recebe um valor para a propriedade     '
'                          (em run time)                                                  '
'                          Public Property Get --> Retorna para a aplica��o um valor      '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private mvarcodigo_usuario As Integer 'local copy
Private mvarnome_usuario As String 'local copy
Private mvarsenha_usuario As String 'local copy
Private CNconexao_usuario As New ADODB.Connection
Public Function atualizar_usuario()
End Function

Public Function excluir_usuario()
End Function


Public Function gravar_usuario()
        
    If mvarcodigo_usuario = " " Then
       Err.Raise vbObjectError + 1000, "clsUsuario", "Usu�rio n�o Cadastrado"
    End If
    If mvarnome_usuario = " " Then
       Err.Raise vbObjectError + 1001, "clsUsuario", "Nome do usu�rio n�o Cadastrado"
    End If
    If mvarsenha_usuario = " " Then
       Err.Raise vbObjectError + 1002, "clsUsuario", "Senha do usu�rio n�o Cadastrada"
    End If
    On Error GoTo Erro
    
    CNconexao_usuario.Execute "INSERT INTO TBusuario(codigo_usuario,nome_usuario,senha_usuario) " & _
                              "SELECT mvarcodigo_usuario, mvarnome_usuario,mvarsenha_usuario"
    'Inserir um LOG
    'Call Funcoes_Gerais.Gravar_Log(Caption_interface, txtcodigo.Text, "Grava��o", frmLogin.strUsuario_Sistema, "Usu�rio Incluiu/Alterou um registro")
    
    DoEvents
       
    Exit Function
    
Erro:
    Err.Raise vbObjectError + 1003, "clsUsuario", "Erros ao Incluir um novo usu�rio"

End Function

Public Function consultar_usuario() As String
End Function

Public Function validar_usuario() As Boolean
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
'Retornando um valor para a aplica��o
    nome_usuario = mvarnome_usuario
End Property

Public Property Let codigo_usuario(ByVal vData As Integer)
    mvarcodigo_usuario = vData
End Property

Public Property Get codigo_usuario() As Integer
    codigo_usuario = mvarcodigo_usuario
End Property

Private Sub Class_GetDataMember(DataMember As String, Data As Object)

End Sub

Private Sub Class_Initialize()
    CNconexao_usuario.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDBalanca;Data Source=INFO-028"
End Sub

Private Sub Class_Terminate()
    CNconexao_usuario.Close
    Set CNconexao_usuario = Nothing
End Sub
