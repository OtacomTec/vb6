Attribute VB_Name = "Botoes"
'*******************************************************************************************
'
'Programação.......................: Marcos Baião
'Data..............................: 07/04/2000
'Função............................: Módulo destinado à automatizar as atividades no banco de
'                                    dados
'*******************************************************************************************
Public intValor_Botao As Integer
'Irá armazenar o valor dos botões da interface conforme ordem em que aparecem, começando do
'numero 1

Public Function Confirmar(Formulario As Form, Nome_Tabela As String, Nome_Campos As String, Caption_interface As String, txtcodigo As TextBox, strValores As String, Evento As String, Optional Nome_Campo_Comparacao As String, Optional Valor_Comparado As String) As String
'******************************************************************************
'Módulo............................: Botões
'Procedimento/Função...............: Confirmar
'Objetivo:.........................: Confirmar a inclusão ou a alteração de
'                                    registros nos formulários de cadastro
'Desenvolvimento...................: (...)
'Data de criação...................: (...)
'Observaçãoes......................:
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

                ' - 1 (menos um), pra não pegar a vírgula
                    intPosicao = intPosicao - 1

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
    'Call Funcoes_Gerais.Gravar_Log(Caption_interface, txtcodigo.Text, "Gravação", frmLogin.strUsuario_Sistema, "Usuário Incluiu/Alterou um registro")
    
    DoEvents
       
    Exit Function
    
erro:
    If Err.Number = -2147168242 Then
        Resume Next
    Else
        Call erro.erro("Confirmar")
        Resume Next
    End If
'----------------------------Instruções-----------------------------------------------------
'Esta é a função de confirmacao, objetos necessários e nomenclatura:
'   Formulario - nome do formulario que tera a funcao de confirmacao
'   Recordset_Memoria - contem copia da tabela que ira receber o novo registro
'   Conexao - Nome da conexao que foi conectada ao banco de dados
'
'Como Chamá-la?
'   Use a instrução CALL, digite o nome do módulo e em seguida o nome da função.
'       EX: CALL Botoes.Confirmar(
'
'Como Preenchê-la?
'   Dentro do parenteses digite o nome dos objetos pedidos, depois feche o parenteses.
'
'--------------------------------------------------------------------------------------------
End Function

Public Function Cancelar(Formulario As Form, Optional ByRef intBotao As Integer = -1) As String

'******************************************************************************
'Módulo............................: Botões
'Procedimento/Função...............: Cancelar
'Objetivo:.........................: Cancela a inclusão ou a alteração de
'                                    registros nos formulários de cadastro
'Desenvolvimento...................: (...)
'Data de criação...................: (...)
'Observaçãoes......................:
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
'----------------------------Instruções-----------------------------------------------------
'Esta é a função de cancelamento, objetos necessários e nomenclatura:
'   Formulario - nome do formulario que tera a funcao de cancelamento
'   Recordset_Memoria - contem copia da tabela que receberia o novo registro
'   Conexao - Nome da conexao que foi conectada ao banco de dados
'
'Como Chamá-la?
'   Use a instrução CALL, digite o nome do módulo e em seguida o nome da função.
'       EX: CALL Botoes.Cancelar(
'
'Como Preenchê-la?
'   Dentro do parenteses digite o nome dos objetos pedidos, depois feche o parenteses.
'
'--------------------------------------------------------------------------------------------
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Módulo............................: Botões
'Procedimento/Função...............: Excluir
'Objetivo:.........................: Controla a exclusão dos registros
'Desenvolvimento...................: (...)
'Data de criação...................: (...)
'Data da última manutenção.........: 23/04/2001
'Observaçãoes......................:
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Exclui(Nome_Tabela As String, Nome_Campo_Comparador As String, txtcodigo As TextBox, Caption_interface As String) As Boolean

    On Error GoTo erro
    Dim bytResultado As Byte
    Dim lonValor As Long
    
    bytResultado = MsgBox("Atenção! Antes de confirmar esta operação, certifique-se que seu BACKUP esteja atualizado. Confirma a exclusão?", vbQuestion + vbYesNo, "Controle de Balanças")

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
    
    'Call Funcoes_Gerais.Gravar_Log(Caption_interface, txtcodigo.Text, "Exclusão", frmLogin.strUsuario_Sistema, "Usuário Excluiu um registro")
    

Exit Function
    
erro:
    Conexao.RollbackTrans
    If Err.Number = 6160 Then
        MsgBox "Não há informações para serem excluídas", vbCritical, "Integrador"
        Exit Function
    ElseIf Err.Number = -2147217900 Then
        Call Banco_Dados.Integridade(False)
    Else
        Call erro.erro("Excluir")
    End If
    
    Exit Function
    'Resume Next
'----------------------------Instruções-----------------------------------------------------
'Esta é a função de cancelamento, objetos necessários e nomenclatura:
'   Recordset_Memoria - contem copia da tabela onde que voce ira apagar um registro
'   Conexao - Nome da conexao que foi conectada ao banco de dados
'
'Como Chamá-la?
'   Use a instrução CALL, digite o nome do módulo e em seguida o nome da função.
'       EX: CALL Botoes.Excluir(
'
'Como Preenchê-la?
'   Dentro do parenteses digite o nome dos objetos pedidos, depois feche o parenteses.
'
'--------------------------------------------------------------------------------------------

End Function

Public Function Alterar(Formulario As Form, Optional ByRef intBotao As Integer = -1) As String
'******************************************************************************
'Módulo............................: Botões
'Procedimento/Função...............: Alterar
'Objetivo:.........................: Preparar os formulário de cadastro para
'                                    aceitar a alteração de um registro
'Desenvolvimento...................: (...)
'Data de criação...................: (...)
'Data da última manutenção.........: 23/04/2001
'Observaçãoes......................:
'    Acrescentei o parâmetro opcional 'intBotao' que receberá o valor da variável
'    intValor_Botao declarada como local nos formulários do Ambiente.
'    observe que este parâmetro só será utilizado pelos cadastros do Ambiente.
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
'----------------------------Instruções------------------------------------------------
'Esta é a função de alteracao, objetos necessários e nomenclatura:
'   Formulario - nome do formulario que tera a funcao de inclusao
'
'Como Chamá-la?
'   Use a instrução CALL, digite o nome do módulo e em seguida o nome da função.
'       EX: CALL Botoes.Alterar(
'
'Como Preenchê-la?
'   Dentro do parenteses digite o nome dos objetos pedidos, depois feche o parenteses.
'--------------------------------------------------------------------------------------

End Function

Sub Atalhos(Form As Form, KeyCode As Integer, Shift As Integer, Optional Status As Boolean)
'******************************************************************************
'Módulo............................: Botões
'Procedimento/Função...............: Atalhos
'Objetivo:.........................: trata as teclas de atalho dos botões de
'                                    controle dos formulários
'Data de criação...................: 21/01/2002
'Data da última manutenção.........:
'Manutenção executada por..........:
'Observaçãoes......................:
'    .o parâmetro form recebe o nome do form proprietário dos botões
'    .os parâmetros keycode e shift recebem os valores do eventro keydown ou keyup
'    .o parâmetro status sai desta sub informando se uma tecla de atalho foi usada ou não
'    .nesta rotina foi usada a função CallByName que chama um método de um objeto qualquer
'     pelo nome em formato de string ( mais informações no help do vb - MSDN )
'
'IMPORTANTE........................:
'    .para que está rotina funcione é necessário que o evento click dos botões
'     do formulário sejam declarados como public
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
    ' realmente não é necessário tratamento de erro
End Sub
