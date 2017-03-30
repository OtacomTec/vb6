Attribute VB_Name = "Seguranca"
'*******************************************************************************************
'
'Sistema...........................: Director
'Módulo............................: Nenhum
'Conexão...........................: Nenhuma
'Formulário........................: Seguranca
'Objetivo do formulário............: Seguranca
'Análise...........................: Eugênio Gomes
'Programação.......................: Eduardo Cruz, Pablo Souza
'Data..............................: 24/04/2000
'Data da última manutenção.........: 12/12/2000
'Manutenção executada por..........: Vagner Vilela
'                                    12/12/2000
'                                    - Correção da rotina para
'                                    tratar usuários recem incluidos
'                                    que ainda não tinham permissões
'                                    configuradas em Acessibilidade
'
'*******************************************************************************************
Public booAcesso As Boolean

Public Function Seguranca(Recordset_Memoria As ADODB.Recordset, Nivel_Usuario As Integer, ID_Formulario As Integer, Conexao As ADODB.Connection) As String
    Dim strConsultar As String
    Dim strIncluir As String
    Dim strAlterar As String
    Dim strExcluir As String
    
    Dim strSQLPermissao As String
    
    On Error GoTo Erro
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
    
    If TBrecordset.RecordCount = 0 Then
        'Não achou o registro, resultado da SQL
        '"SELECT DFconsultar, DFincluir, DFalterar, DFexcluir
        'From TBSeguranca
        'Where DFnivel_usuario = a And DFid_formulario = b
        'Ou não existe o nível a não existe o formulario b.
        'O Nivel deve existir pois quando se cadastra um
        'usuário pede-se um nível.
        'O formulario pode não existir. Pois pode não ter
        'sido dada permissão alguma para esse usuário.
        MsgBox "Formulário não disponível para este usuário", vbCritical, "Director"
        booAcesso = False
        Exit Function
    End If
    
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
    
Erro:
    Call Erro.Erro("Seguranca")
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
        
    If booAlterar = True Then
        Formulario.cmdAlterar.Enabled = True
    Else
        Formulario.cmdAlterar.Enabled = False
        MsgBox "Este Formulário é somente para leitura.", vbExclamation, "Director"
    End If
    
    Formulario.cmdIncluir.ToolTipText = "Incluir (atalho: Alt + I)"
    Formulario.cmdConfirmar.ToolTipText = "Confirmar (atalho: Alt + C)"
    Formulario.cmdCancelar.ToolTipText = "Cancelar (atalho: Alt + N)"
    Formulario.cmdExcluir.ToolTipText = "Excluir (atalho: Alt + E)"
    Formulario.cmdAlterar.ToolTipText = "Alterar (atalho: Alt + A)"
    Formulario.cmdImprimir.ToolTipText = "Imprimir (atalho: Alt + P)"
    Formulario.cmdAtualizar.ToolTipText = "Atualizar (atalho: F5)"
    Formulario.cmdSair.ToolTipText = "Sair (atalho: ESC)"
    
End Function

