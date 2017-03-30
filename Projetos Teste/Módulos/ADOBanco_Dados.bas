Attribute VB_Name = "ADOBanco_Dados"
 
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
'Data da última manutenção.........: 07/04/2000
'Manutenção executada por..........: Pablo Souza
'
'*******************************************************************************************

Public CNconexao As ADODB.Connection
Public CNconexaoII As ADODB.Connection
Public conexao As ADODB.Connection
Public TBrecordset As ADODB.Recordset
Public CNemissao As ADODB.Connection
Public TBemissao As ADODB.Recordset


Public Function Abre_Conexao(Caminho_Banco As String, conexao As ADODB.Connection) As String

Set conexao = New ADODB.Connection
    On Error GoTo ErroConexao
    conexao.Open "Provider = Microsoft.jet.OLEDB.4.0;Data source = " & Caminho_Banco & ";Jet OLEDB:Database Password = btr96;"
    Exit Function
    
ErroConexao:
    MsgBox "Houve um problema de conexão. Verifique sua rede ou o caminho para conexão", vbCritical, "Director"
    frmProcura_Caminho.Show 1
    End
End Function

Public Function Fecha_RecordSet(Recordset_Memoria As ADODB.Recordset) As String
    On Error Resume Next
        Recordset_Memoria.Close
    Set Recordset_Memoria = Nothing
End Function

Public Function Inicio(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Ordem As String, Recordset_Memoria As ADODB.Recordset, conexao As ADODB.Connection) As String
    Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "SELECT " & Nome_Campos & " "
    strSQL = strSQL & "FROM " & Nome_Tabela & " "
    strSQL = strSQL & "ORDER BY " & Nome_Campo_Ordem & ""
    
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        On Error GoTo ErroBanco
        TBrecordset.Open strSQL, conexao, adOpenKeyset, adLockOptimistic, adCmdText
    Set Recordset_Memoria = TBrecordset
    
    Set TBrecordset = Nothing
    Exit Function
    
ErroBanco:
    If Err.Number = -2147217865 Then
        MsgBox "Houve um problema de conexão. Verifique sua rede ou o caminho para conexão", vbCritical, "BTR - Bahamas Ticket's Refeição"
        frmProcura_Caminho.Show 1
        End
    End If

End Function

Public Function Condicao_Numerico(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Condicao As String, Operador As String, TextBox As TextBox, Recordset_Memoria As ADODB.Recordset, conexao As ADODB.Connection) As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT " & Nome_Campos & " "
    strSQL = strSQL & "FROM " & Nome_Tabela & " "
    strSQL = strSQL & "WHERE " & Nome_Campo_Condicao & " "
    strSQL = strSQL & "" & Operador & " " & TextBox.Text & ""
        
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open strSQL, conexao, adOpenKeyset, adLockOptimistic, adCmdText
    Set Recordset_Memoria = TBrecordset
    
End Function

Public Function Condicao_String(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Condicao As String, Operador As String, strString As String, Recordset_Memoria As ADODB.Recordset, conexao As ADODB.Connection) As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT " & Nome_Campos & " "
    strSQL = strSQL & "FROM " & Nome_Tabela & " "
    strSQL = strSQL & "WHERE " & Nome_Campo_Condicao & " "
    strSQL = strSQL & "" & Operador & "'" & strString & "'"
                
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open strSQL, conexao, adOpenKeyset, adLockOptimistic, adCmdText
    Set Recordset_Memoria = TBrecordset
    
End Function

'Public Function Condicao_String_ADB(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Condicao As String, Operador As String, adbDataCombo As DataCombo, Recordset_Memoria As ADODB.Recordset, conexao As ADODB.Connection) As String
'    Dim strSQL As String
'
'    strSQL = ""
'    strSQL = strSQL & "SELECT " & Nome_Campos & " "
'    strSQL = strSQL & "FROM " & Nome_Tabela & " "
'    strSQL = strSQL & "WHERE " & Nome_Campo_Condicao & " "
'    strSQL = strSQL & "" & Operador & " '" & adbDataCombo.Text & "'"
'
'    Set TBrecordset = New ADODB.Recordset
'        TBrecordset.CursorLocation = adUseClient
'        TBrecordset.Open strSQL, conexao, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Recordset_Memoria = TBrecordset
'
'End Function
'
'Public Function Condicao_String_CMB(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Condicao As String, Operador As String, Combo As ComboBox, Recordset_Memoria As ADODB.Recordset, conexao As ADODB.Connection) As String
'    Dim strSQL As String
'
'    strSQL = ""
'    strSQL = strSQL & "SELECT " & Nome_Campos & " "
'    strSQL = strSQL & "FROM " & Nome_Tabela & " "
'    strSQL = strSQL & "WHERE " & Nome_Campo_Condicao & " "
'    strSQL = strSQL & "" & Operador & " '" & Combo.Text & "'"
'
'    Set TBrecordset = New ADODB.Recordset
'        TBrecordset.CursorLocation = adUseClient
'        TBrecordset.Open strSQL, conexao, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Recordset_Memoria = TBrecordset
'
'End Function
'
'Public Function Condicao_Integer(Nome_Campos As String, Nome_Tabela As String, Nome_Campo_Condicao As String, Operador As String, Variavel As Integer, Recordset_Memoria As ADODB.Recordset, conexao As ADODB.Connection) As String
'    Dim strSQL As String
'
'    strSQL = ""
'    strSQL = strSQL & "SELECT " & Nome_Campos & " "
'    strSQL = strSQL & "FROM " & Nome_Tabela & " "
'    strSQL = strSQL & "WHERE " & Nome_Campo_Condicao & " "
'    strSQL = strSQL & "" & Operador & " " & Variavel & ""
'
'    Set TBrecordset = New ADODB.Recordset
'        TBrecordset.CursorLocation = adUseClient
'        TBrecordset.Open strSQL, conexao, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Recordset_Memoria = TBrecordset
'
'End Function
