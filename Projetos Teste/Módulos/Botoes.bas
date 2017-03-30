Attribute VB_Name = "Botoes"
'*******************************************************************************************
'
'Programa��o.......................: Marcos Bai�o
'Data..............................: 07/04/2000
'Fun��o............................: M�dulo destinado � automatizar as atividades no banco de
'                                    dados
'*******************************************************************************************
Public intValor_Botao As Integer
'Ir� armazenar o valor dos bot�es da interface conforme ordem em que aparecem, come�ando do
'numero 1

Public Function Confirmar(Formulario As Form, Nome_Tabela As String, Nome_Campos As String, Caption_interface As String, txtcodigo As TextBox, strValores As String, Evento As String, Optional Nome_Campo_Comparacao As String, Optional Valor_Comparado As String) As String
'******************************************************************************
'M�dulo............................: Bot�es
'Procedimento/Fun��o...............: Confirmar
'Objetivo:.........................: Confirmar a inclus�o ou a altera��o de
'                                    registros nos formul�rios de cadastro
'Desenvolvimento...................: (...)
'Data de cria��o...................: (...)
'Observa��oes......................:
'   Nome_Campo_Comparacao = Campo da Tabela que sera comparado com o codigo que for
'                           escolhido pelo usuario
'   Valor_Comparado       = Valor que foi escolhido pelo usuario
'******************************************************************************
    On Error GoTo erro
    Dim strSQL As String
       
    strValores = Funcoes_Gerais.Verifica_Apostrofo(strValores)
    
'    strSQL = ""
'    strSQL = strSQL & "SELECT * "
'    strSQL = strSQL & "FROM " & Nome_Tabela & " "
'    strSQL = strSQL & "WHERE " & Nome_Campo_Comparacao & " = '" & Valor_Comparado & "'"
'    Dim adrVerificacao As New ADODB.Recordset
'
'    adrVerificacao.Open strSQL, Conexao, adOpenDynamic, adLockOptimistic
'    If adrVerificacao.EOF And adrVerificacao.BOF Then
    If Evento = "N" Then
        strSQL = ""
        strSQL = strSQL & "INSERT INTO " & Nome_Tabela & Space(1)
        strSQL = strSQL & "(" & Nome_Campos & ") "
        strSQL = strSQL & "SELECT" & Space(1) & strValores
    Else
        Dim strMatrix(100, 2) As String
        Dim strSQLretorno(2) As String
        Dim intPosicao As Integer
        Dim intPosicao2 As Integer  'variavel que armazena a posicao do 1o. apostrofo
        Dim intPosicao3 As Integer  'variavel que armazena a posicao do 1o. apostrofo
        Dim intCont As Integer
        Dim intColuna As Integer

        intCont = 1

        strSQLretorno(1) = Nome_Campos
        strSQLretorno(2) = strValores

        Do While intCont < 3
            For I = 1 To 100
                intPosicao = (InStr(1, strSQLretorno(intCont), ","))

                If intCont = 2 Then
                    intPosicao2 = (InStr(1, strSQLretorno(intCont), "'"))
                    intPosicao3 = (InStr((intPosicao2 + 1), strSQLretorno(intCont), "'"))

                    'Codificacao para verificar se virgula esta no meio do texto
                    Do While intPosicao > intPosicao2 And intPosicao < intPosicao3
                        intPosicao = (InStr((intPosicao + 1), strSQLretorno(intCont), ","))

                        intPosicao2 = (InStr(1, strSQLretorno(intCont), "'"))
                        intPosicao3 = (InStr((intPosicao2 + 1), strSQLretorno(intCont), "'"))

                        DoEvents
                    Loop
                End If

                ' - 1 (menos um), pra n�o pegar a v�rgula
                    intPosicao = intPosicao - 1

                'Iff -> esta sendo usado por que o ultimo valor � um n�mero negativo
                'entao fa�o a substitui��o por um valor positivo
                strMatrix(I, intCont) = Mid(strSQLretorno(intCont), 1, IIf(intPosicao < 0, Len(strSQLretorno(intCont)), intPosicao))

                strSQLretorno(intCont) = Mid(strSQLretorno(intCont), intPosicao + 2, (Len(strSQLretorno(intCont)) - intPosicao))

                If intPosicao < 0 Then
                    Exit For
                End If
            Next I
            intCont = intCont + 1
        Loop

        intColuna = 0

        'Apartir daqui concanetacao dos valores da Matriz, para montar a SQL de atualizacao
        strSQL = ""
        strSQL = strSQL & "UPDATE " & Nome_Tabela & " "
        strSQL = strSQL & "SET " & strMatrix(1, 1) & " = " & strMatrix(1, 2) & " "

        For I = 2 To 100
            If strMatrix(I, 1) <> "" Then
                strSQL = strSQL & ", " & strMatrix(I, 1) & " = " & Trim(strMatrix(I, 2)) & " "
            Else
                Exit For
            End If
            intColuna = intColuna + 1
        Next I
        strSQL = strSQL & "WHERE " & Nome_Campo_Comparacao & " = '" & Valor_Comparado & "' "
    End If
          
    'Execucao da SQL, de inclusao ou alteracao conforme o clique do usuario
    Conexao.Execute strSQL
    
    'LOG
    'Call Funcoes_Gerais.Gravar_Log(Caption_interface, txtcodigo.Text, "Grava��o", frmLogin.strUsuario_Sistema, "Usu�rio Incluiu/Alterou um registro")
    
    DoEvents
       
    Exit Function
    
erro:
    If Err.Number = -2147168242 Then
        Resume Next
    Else
        Call erro.erro("Confirmar")
        Resume Next
    End If
'----------------------------Instru��es-----------------------------------------------------
'Esta � a fun��o de confirmacao, objetos necess�rios e nomenclatura:
'   Formulario - nome do formulario que tera a funcao de confirmacao
'   Recordset_Memoria - contem copia da tabela que ira receber o novo registro
'   Conexao - Nome da conexao que foi conectada ao banco de dados
'
'Como Cham�-la?
'   Use a instru��o CALL, digite o nome do m�dulo e em seguida o nome da fun��o.
'       EX: CALL Botoes.Confirmar(
'
'Como Preench�-la?
'   Dentro do parenteses digite o nome dos objetos pedidos, depois feche o parenteses.
'
'--------------------------------------------------------------------------------------------
End Function

Public Function Cancelar(Formulario As Form, Optional ByRef intBotao As Integer = -1) As String

'******************************************************************************
'M�dulo............................: Bot�es
'Procedimento/Fun��o...............: Cancelar
'Objetivo:.........................: Cancela a inclus�o ou a altera��o de
'                                    registros nos formul�rios de cadastro
'Desenvolvimento...................: (...)
'Data de cria��o...................: (...)
'Observa��oes......................:
'   Nome_Campo_Comparacao = Campo da Tabela que sera comparado com o codigo que for
'                           escolhido pelo usuario
'   Valor_Comparado       = Valor que foi escolhido pelo usuario
'******************************************************************************
        
    'Enabled = True   -> Habilitado
    'Enabled = False  -> Desabilitado
    Formulario.cmdIncluir.Enabled = True
    'Formulario.cmdExcluir.Enabled = True
    Formulario.cmdConfirmar.Enabled = False
    Formulario.cmdCancelar.Enabled = False
    'Formulario.cmdAlterar.Enabled = True
    
    'Indica o botao que recebera o foco
    Formulario.cmdIncluir.SetFocus
    
    'Armazena o valor do botao
    If intBotao = -1 Then
        intValor_Botao = 3
    Else
        intBotao = 3
    End If
    
    Exit Function
erro:
    Call erro.erro("Cancelar")
    Resume Next
'----------------------------Instru��es-----------------------------------------------------
'Esta � a fun��o de cancelamento, objetos necess�rios e nomenclatura:
'   Formulario - nome do formulario que tera a funcao de cancelamento
'   Recordset_Memoria - contem copia da tabela que receberia o novo registro
'   Conexao - Nome da conexao que foi conectada ao banco de dados
'
'Como Cham�-la?
'   Use a instru��o CALL, digite o nome do m�dulo e em seguida o nome da fun��o.
'       EX: CALL Botoes.Cancelar(
'
'Como Preench�-la?
'   Dentro do parenteses digite o nome dos objetos pedidos, depois feche o parenteses.
'
'--------------------------------------------------------------------------------------------
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'M�dulo............................: Bot�es
'Procedimento/Fun��o...............: Excluir
'Objetivo:.........................: Controla a exclus�o dos registros
'Desenvolvimento...................: (...)
'Data de cria��o...................: (...)
'Data da �ltima manuten��o.........: 23/04/2001
'Observa��oes......................:
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Exclui(Nome_Tabela As String, Nome_Campo_Comparador As String, txtcodigo As TextBox, Caption_interface As String) As Boolean

    On Error GoTo erro
    Dim bytResultado As Byte
    Dim lonValor As Long
    
    bytResultado = MsgBox("Aten��o! Antes de confirmar esta opera��o, certifique-se que seu BACKUP esteja atualizado. Confirma a exclus�o?", vbQuestion + vbYesNo, "Controle de Balan�as")

    If bytResultado = vbYes Then Exclui2 = True

    If bytResultado = vbNo Then Exclui2 = False
      
    Dim strDelete As String
    Conexao.BeginTrans
    
    strDelete = "DELETE " & _
                "FROM " & Nome_Tabela & Space(1) & " " & _
                    "WHERE " & Nome_Campo_Comparador & Space(1) & " " & _
                    "= '" & txtcodigo.Text & "' "

    DoEvents

    Conexao.Execute strDelete
   
    Conexao.CommitTrans
    
    'Call Funcoes_Gerais.Gravar_Log(Caption_interface, txtcodigo.Text, "Exclus�o", frmLogin.strUsuario_Sistema, "Usu�rio Excluiu um registro")
    

Exit Function
    
erro:
    Conexao.RollbackTrans
    If Err.Number = 6160 Then
        MsgBox "N�o h� informa��es para serem exclu�das", vbCritical, "Integrador"
        Exit Function
    ElseIf Err.Number = -2147217900 Then
        Call Banco_Dados.Integridade(False)
    Else
        Call erro.erro("Excluir")
    End If
    
    Exit Function
    'Resume Next
'----------------------------Instru��es-----------------------------------------------------
'Esta � a fun��o de cancelamento, objetos necess�rios e nomenclatura:
'   Recordset_Memoria - contem copia da tabela onde que voce ira apagar um registro
'   Conexao - Nome da conexao que foi conectada ao banco de dados
'
'Como Cham�-la?
'   Use a instru��o CALL, digite o nome do m�dulo e em seguida o nome da fun��o.
'       EX: CALL Botoes.Excluir(
'
'Como Preench�-la?
'   Dentro do parenteses digite o nome dos objetos pedidos, depois feche o parenteses.
'
'--------------------------------------------------------------------------------------------

End Function

Public Function Alterar(Formulario As Form, Optional ByRef intBotao As Integer = -1) As String
'******************************************************************************
'M�dulo............................: Bot�es
'Procedimento/Fun��o...............: Alterar
'Objetivo:.........................: Preparar os formul�rio de cadastro para
'                                    aceitar a altera��o de um registro
'Desenvolvimento...................: (...)
'Data de cria��o...................: (...)
'Data da �ltima manuten��o.........: 23/04/2001
'Observa��oes......................:
'    Acrescentei o par�metro opcional 'intBotao' que receber� o valor da vari�vel
'    intValor_Botao declarada como local nos formul�rios do Ambiente.
'    observe que este par�metro s� ser� utilizado pelos cadastros do Ambiente.
'******************************************************************************
    On Error GoTo erro
    
    'Enabled = True   -> Habilitado
    'Enabled = False  -> Desabilitado
    Formulario.cmdIncluir.Enabled = False
    Formulario.cmdExcluir.Enabled = False
    Formulario.cmdConfirmar.Enabled = True
    Formulario.cmdCancelar.Enabled = True
    Formulario.cmdAlterar.Enabled = False
    
    'Armazena o valor do botao
    If intBotao = -1 Then
        intValor_Botao = 5
    Else
        intBotao = 5
    End If
    
    Exit Function
erro:
    Call erro.erro("Alterar")
'----------------------------Instru��es------------------------------------------------
'Esta � a fun��o de alteracao, objetos necess�rios e nomenclatura:
'   Formulario - nome do formulario que tera a funcao de inclusao
'
'Como Cham�-la?
'   Use a instru��o CALL, digite o nome do m�dulo e em seguida o nome da fun��o.
'       EX: CALL Botoes.Alterar(
'
'Como Preench�-la?
'   Dentro do parenteses digite o nome dos objetos pedidos, depois feche o parenteses.
'--------------------------------------------------------------------------------------

End Function

Sub Atalhos(Form As Form, KeyCode As Integer, Shift As Integer, Optional Status As Boolean)
'******************************************************************************
'M�dulo............................: Bot�es
'Procedimento/Fun��o...............: Atalhos
'Objetivo:.........................: trata as teclas de atalho dos bot�es de
'                                    controle dos formul�rios
'Data de cria��o...................: 21/01/2002
'Data da �ltima manuten��o.........:
'Manuten��o executada por..........:
'Observa��oes......................:
'    .o par�metro form recebe o nome do form propriet�rio dos bot�es
'    .os par�metros keycode e shift recebem os valores do eventro keydown ou keyup
'    .o par�metro status sai desta sub informando se uma tecla de atalho foi usada ou n�o
'    .nesta rotina foi usada a fun��o CallByName que chama um m�todo de um objeto qualquer
'     pelo nome em formato de string ( mais informa��es no help do vb - MSDN )
'
'IMPORTANTE........................:
'    .para que est� rotina funcione � necess�rio que o evento click dos bot�es
'     do formul�rio sejam declarados como public
'
'******************************************************************************
    On Error GoTo erro
    Dim strMetodo As String
    
    If Shift = vbAltMask Then
        If KeyCode = vbKeyI Then
            If Form.cmdIncluir.Enabled Then
                strMetodo = "cmdIncluir_Click"
            End If
        ElseIf KeyCode = vbKeyC Then
            If Form.cmdConfirmar.Enabled Then
                strMetodo = "cmdConfirmar_Click"
            End If
        ElseIf KeyCode = vbKeyN Then
            If Form.cmdCancelar.Enabled Then
                strMetodo = "cmdCancelar_Click"
            End If
        ElseIf KeyCode = vbKeyA Then
            If Form.cmdAlterar.Enabled Then
                strMetodo = "cmdAlterar_Click"
            End If
        ElseIf KeyCode = vbKeyE Then
            If Form.cmdExcluir.Enabled Then
                strMetodo = "cmdExcluir_Click"
            End If
        ElseIf KeyCode = vbKeyP Then
            If Form.cmdImprimir.Enabled = True Then
                strMetodo = "cmdImprimir_Click"
            End If
        End If
    ElseIf KeyCode = vbKeyF5 Then
        If Form.cmdAtualizar.Enabled Then
            strMetodo = "cmdAtualizar_Click"
        End If
    End If
    
    If Len(strMetodo) Then
        CallByName Form, strMetodo, VbMethod
        Status = True
    End If

erro:
    ' realmente n�o � necess�rio tratamento de erro
End Sub
