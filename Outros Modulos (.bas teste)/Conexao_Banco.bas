Attribute VB_Name = "Conexao_Banco"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Manager Vers�o 1.0                                                                      '
' M�dulo............: M�dulos de Sistema                                                  '
' Finalidade........: Ger�nciar conex�o com banco e conex�es com RECORDSET�s              '
' Data de Cria��o...: 08/03/2001                                                          '
' Autor.............: Marcos Bai�o                                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Conexao As ADODB.Connection
Public TBrecordset As ADODB.Recordset

Public Function Abre_Conexao(caminho As String) As String

    Set Conexao = New ADODB.Connection
        Conexao.Open " " & caminho & " "
        Exit Function
        
End Function

Public Sub Abrir_conexao_registro(aplicacao As String)

      On Error GoTo Erro

      Dim Registro As New DLLFuncoesGerais.Registro
      Dim strConexao As String
      
      strConexao = Registro.WinRegLerSequ�ncia("HKEY_LOCAL_MACHINE\SOFTWARE\" + aplicacao + "\BD", "String de conexao")
      
      strConexao = strConexao + "Persist Security Info=False"
      
      Set Conexao = New ADODB.Connection
      
      'Abrir� a conex�o
      Conexao.Open strConexao
      
      Exit Sub
      
Erro:
    If Err.Number = -2147467259 Then
       MsgBox "Fonte de dados n�o encontrada!", vbCritical, "Logicx"
       Shell "C:\Projetos\Sistemas\Mercedes\Configurador_provedor_dados.exe", vbNormalFocus
       End
       Exit Sub
    Else
       Call Erro.Erro("Load")
       Exit Sub
    End If
    
End Sub
Public Function Abre_Recordset(tabela As String, TBrecordset As ADODB.Recordset) As String

    Dim strSQl As String
    
    SQL_Pesquisa = "Select * From " & tabela & " "
    
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open SQL_Pesquisa, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
                
End Function

Public Function Fecha_RecordSet(TBrecordset As ADODB.Recordset) As String

    TBrecordset.Close
    Set TBrecordset = Nothing
    
End Function

Public Function Fecha_Conexao() As String

    Conexao.Close
    Set Conexao = Nothing
    
End Function
Public Function SQLgeral(strSQL_comando As String, Recordset_Memoria As ADODB.Recordset)
    On Error GoTo Erro
    Set Recordset_Memoria = New ADODB.Recordset
        Recordset_Memoria.CursorLocation = adUseClient
        Recordset_Memoria.Open strSQL_comando, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
    Exit Function
Erro:
    Call Erro.Erro("SQLgeral")
    
    Resume Next
End Function

