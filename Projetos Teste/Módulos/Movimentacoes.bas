Attribute VB_Name = "Movimentacoes"

Public Function Inicio(Nome_Campo_Codigo As String, Nome_Campo_Descricao As String, Nome_Tabela As String, DataCombo As DataCombo) As String
    Dim strSQLinicio As String
     
    strSQLinicio = ""
    strSQLinicio = strSQLinicio & "SELECT " & Nome_Campo_Codigo & ", " & Nome_Campo_Descricao & " "
    strSQLinicio = strSQLinicio & "FROM " & Nome_Tabela & " "
    strSQLinicio = strSQLinicio & "ORDER BY " & Nome_Campo_Descricao & ""
                    
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open strSQLinicio, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
                
    Set DataCombo.RowSource = TBrecordset
        DataCombo.ListField = Nome_Campo_Descricao
        'BoundColumn -> sendo usado para retornar no TextBox o valor pedido, neste caso
        '               Nome_Campo_Codigo
        DataCombo.BoundColumn = Nome_Campo_Codigo
    
    Set TBrecordset = Nothing
    
'*******************************************************************************************
' OBS.: Sempre que utilizar esta funcao, no evento Click do DataCombo voce tera que colocar
'       uma linha de comando.
'       Exemplo da linha:
'
'   If Area = 2 Then
'       TextBox = DataCombo.BoundText
'   End If
'
'*******************************************************************************************
End Function
Public Function Inicio_SQL(Nome_Campo_Codigo As String, Nome_Campo_Descricao As String, Nome_Tabela As String, DataCombo As DataCombo, String_Sql As String) As String

    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open String_Sql, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
                
    Set DataCombo.RowSource = TBrecordset
        DataCombo.ListField = Nome_Campo_Descricao
        'BoundColumn -> sendo usado para retornar no TextBox o valor pedido, neste caso
        '               Nome_Campo_Codigo
        DataCombo.BoundColumn = Nome_Campo_Codigo
    
    Set TBrecordset = Nothing
    
'*******************************************************************************************
' OBS.: Sempre que utilizar esta funcao, no evento Click do DataCombo voce tera que colocar
'       uma linha de comando.
'       Exemplo da linha:
'
'   If Area = 2 Then
'       TextBox = DataCombo.BoundText
'   End If
'
'*******************************************************************************************
End Function

Public Function Codigo_Perde_Foco(Nome_Campo_Codigo As String, Nome_Campo_Descricao As String, Nome_Tabela As String, TextBox_Codigo As TextBox, DataCombo As DataCombo) As String

    On Error GoTo Erro
    If TextBox_Codigo <> "" Then
        Dim strSQLPerde_Foco As String

        strSQLPerde_Foco = ""
        strSQLPerde_Foco = strSQLPerde_Foco & "SELECT " & Nome_Campo_Codigo & ", " & Nome_Campo_Descricao & " "
        strSQLPerde_Foco = strSQLPerde_Foco & "FROM " & Nome_Tabela & " "
        strSQLPerde_Foco = strSQLPerde_Foco & "WHERE " & Nome_Campo_Codigo & " "
        strSQLPerde_Foco = strSQLPerde_Foco & "= " & TextBox_Codigo & ""

        Set TBrecordset = New ADODB.Recordset
            TBrecordset.CursorLocation = adUseClient
            TBrecordset.Open strSQLPerde_Foco, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
        
        DataCombo = TBrecordset(Nome_Campo_Descricao)
        
        Set TBrecordset = Nothing
    End If
    Exit Function
    
Erro:
    Call Erro.Erro("Codigo_Perde_Foco")
    Resume Next
End Function

Public Function Inicio_DataGrid(strComando_SQL As String, Nome_DataGrid As DataGrid, strTamanho_Campos As String, strCaption_Campos As String) As String
    '*****************************************************************************************
    'Preenche um DataGrid
    'Não faz chamadas a outras funções
    'OBS.:   Se no DataGrid retornar algum campo que voce nao deseja que seja exibido, no
    '        local onde voce indica o Tamanho dele voce deve colocar 0 (zero)  e  no caso
    '        Caption e só deixar em branco
    '*****************************************************************************************
    
    Dim strMatrix(100, 3) As String
    Dim strSQLretorno(3) As String
    Dim intPosicao_FROM As Integer
    Dim intPosicao As Integer
    Dim intCont As Integer
    
    strSQLretorno(1) = UCase$(strComando_SQL)
    strSQLretorno(2) = strTamanho_Campos
    strSQLretorno(3) = strCaption_Campos
        
    'InStr -> Retorna a posicao de uma expressão dentro da string
    'O numero 7 (sete) esta sendo subtraido por que é o número de posições do comando
    'SELECT e espaço até o nome do primeiro campo
    intPosicao_FROM = (InStr(1, strSQLretorno(1), "FROM")) - 7
    
    strSQLretorno(1) = Mid(strSQLretorno(1), 7, intPosicao_FROM)
    
    'Verificao se o valor de strSQLretorno tem um '*'
    'se tiver ira preencher a partir de segunda coluna por isso  intCont = 2
    If InStr(1, strSQLretorno(1), "*") > 0 Then
        intCont = 2
    Else
        intCont = 1
    End If
        
    Do While intCont < 4
        ' - 1 (menos um), pra não pegar a vírgula
        For I = 1 To 100
            intPosicao = (InStr(1, strSQLretorno(intCont), ",")) - 1
            'Iff -> esta sendo usado por que o ultimo valor é um número negativo
            'entao faço a substituição por um valor positivo
            strMatrix(I, intCont) = Mid(strSQLretorno(intCont), 1, IIf(intPosicao < 0, Len(strSQLretorno(intCont)), intPosicao))
            
            strSQLretorno(intCont) = Mid(strSQLretorno(intCont), intPosicao + 2, (Len(strSQLretorno(intCont)) - intPosicao))
            
            If intPosicao < 0 Then
                Exit For
            End If
        Next I
        intCont = intCont + 1
    Loop
    
    DoEvents
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open strComando_SQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
    Set Nome_DataGrid.DataSource = TBrecordset
    
    Dim intColuna As Integer
    intColuna = 0
    
    For I = 1 To 100
        On Error GoTo ErroGrid
        If strMatrix(I, 1) <> "" Then
            Nome_DataGrid.Columns(intColuna).DataField = strMatrix(I, 1)
        End If
        If strMatrix(I, 3) <> "" Then
            Nome_DataGrid.Columns(intColuna).Caption = strMatrix(I, 3)
        End If
        If strMatrix(I, 2) <> "" Then
            Nome_DataGrid.Columns(intColuna).Width = Val(strMatrix(I, 2))
        End If
        intColuna = intColuna + 1
    Next I
        
    Set TBrecordset = Nothing
    Exit Function
ErroGrid:
    If Err.Number = 9 Then
        Exit Function
    End If
End Function

Public Sub Inicio_DataGrid_II(strComando_SQL As String, DataGrid As DataGrid, vetTamanho_Campos As Variant, vetCaption_Campos As Variant, Optional vetCampos As Variant)
    Dim I As Integer
    
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open strComando_SQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
    Set DataGrid.DataSource = TBrecordset
    
    
    If IsMissing(vetCampos) Then
        For I = 0 To UBound(vetCaption_Campos)
            DataGrid.Columns(I).Caption = vetCaption_Campos(I)
            DataGrid.Columns(I).Width = vetTamanho_Campos(I)
        Next I
    Else
        For I = 0 To UBound(vetCampos)
            DataGrid.Columns(I).DataField = vetCampos(I)
            DataGrid.Columns(I).Caption = vetCaption_Campos(I)
            DataGrid.Columns(I).Width = vetTamanho_Campos(I)
        Next I
    End If
    Set TBrecordset = Nothing
End Sub
