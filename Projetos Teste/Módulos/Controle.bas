Attribute VB_Name = "Controle"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Manager Versão 1.0                                                                      '
' Módulo............: Módulos de Sistema                                                  '
' Finalidade........: Manutenção de Tabelas como Inclusão,Exclusão e Alteração            '
' Data de Criação...: 19/04/2001                                                          '
' Autor.............: Marcos Baião                                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Controle(campos As String, tabela As String, DataFields As String, Texts As String, condição As String, rotina As String) As String
    
    On Error GoTo erro_inclusao
    
    Dim strSQL As String
    Dim valor As String
    Dim valor2 As String
    Dim index As Integer
    Dim posiciona(20) As String
    Dim Ultimo As String
    Dim flag As Integer

    If rotina = "Inclusão" Then
       valor = Len(Texts)
       achastring = 1
       flag = 1
       Do While Texts <> ""
          achastring = InStr(achastring, Texts, ",")
          achastring = achastring - 1
          Ultimo = Texts
          posiciona(flag) = Left(Texts, achastring)
          On Error GoTo finaliza
          achastring = achastring + 2
          Texts = Mid(Right(Texts, valor), achastring)
          achastring = 1
           
          flag = flag + 1
       Loop
       
       Exit Function
       
finaliza:

    Ultimo = Texts
    'flag = flag - 1
    posiciona(flag) = Ultimo
    
    flag = 1
    Texts = Empty
    conteudo = posiciona(flag)
    
    Do While posiciona(flag) <> ""
       conteudo = posiciona(flag)
       Texts = Texts & conteudo & ","
       flag = flag + 1
    Loop
    conexao.Execute ("INSERT INTO TBclientes(" & DataFields & ") SELECT " & Texts & "")
        
                
    ElseIf rotina = "Exclusão" Then
    
    ElseIf rotina = "Alteração" Then
    
    End If
    
    Exit Function
    
erro_inclusao:
    MsgBox "Erro ao Incluir o registro!", vbCritical, "Manager"
        
End Function

