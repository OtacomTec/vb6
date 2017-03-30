Attribute VB_Name = "ADOSeguranca"
'*******************************************************************************************
'
'Sistema...........................: Director
'Módulo............................: Nenhum
'Conexão...........................: Nenhuma
'Formulário........................: ADOSeguranca
'Objetivo do formulário............: Seguranca (ADO)
'Análise...........................: Eugênio Gomes
'Programação.......................: Eduardo Cruz, Pablo Souza
'Data..............................: 24/04/2000
'Data da última manutenção.........: 24/04/2000
'Manutenção executada por..........: Eduardo Cruz
'
'*******************************************************************************************
Public booAcesso As Boolean

Public Function Seguranca(Recordset_Memoria As ADODB.Recordset, Nivel_Usuario As Integer, ID_Formulario As Integer, Conexao As ADODB.Connection) As String
    Dim strConsultar As String
    Dim strIncluir As String
    Dim strAlterar As String
    Dim strExcluir As String
    
    Dim strSQLPermissao As String
    
    On Error GoTo erro
    strSQLPermissao = ""
    strSQLPermissao = strSQLPermissao & "SELECT DFconsultar, DFincluir, DFalterar, DFexcluir "
    strSQLPermissao = strSQLPermissao & "FROM TBSeguranca "
    strSQLPermissao = strSQLPermissao & "WHERE DFnivel_usuario "
    strSQLPermissao = strSQLPermissao & " = " & Nivel_Usuario & " "
    strSQLPermissao = strSQLPermissao & "AND DFid_formulario "
    strSQLPermissao = strSQLPermissao & " = " & ID_Formulario & ""
    
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open strSQLPermissao, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
    Set Recordset_Memoria = TBrecordset
    
    strConsultar = TBrecordset("DFconsultar")
    strIncluir = TBrecordset("DFincluir")
    strAlterar = TBrecordset("DFalterar")
    strExcluir = TBrecordset("DFexcluir")
    
    If strConsultar = "S" Then
        booConsultar = True
    Else
        booConsultar = False
    End If
    
    If strIncluir = "S" Then
        booIncluir = True
    Else
        booIncluir = False
    End If
    
    If strAlterar = "S" Then
        booAlterar = True
    Else
        booAlterar = False
    End If
    
    If strExcluir = "S" Then
        booExcluir = True
    Else
        booExcluir = False
    End If
    
    Set TBrecordset = Nothing
    
    If booConsultar = False Then
        MsgBox "Formulário não disponível para este usuário", vbCritical, "Director"
        booAcesso = False
    Else
        booAcesso = True
    End If
    Exit Function
    
erro:
    Call erro.erro
    Resume Next

End Function

Public Function Acesso(Formulario As Form) As String
    
    On Error Resume Next
    DoEvents
    If booIncluir = True Then
        Formulario.cmdIncluir.Enabled = True
    Else
        Formulario.cmdIncluir.Enabled = False
    End If
        
    If booExcluir = True Then
        Formulario.cmdExcluir.Enabled = True
    Else
        Formulario.cmdExcluir.Enabled = False
    End If
        
    If booAlterar = False Then
        MsgBox "Este Formulário é somente para leitura.", vbExclamation, "Director"
    End If
     
End Function
Public Function Termina(Conexao As ADODB.Connection)
    On Error GoTo erro
    If booAlterar = True Then
        Conexao.CommitTrans
    Else
        Conexao.RollbackTrans
    End If
    Exit Function
erro:
    Call erro.erro
    Resume Next
End Function
