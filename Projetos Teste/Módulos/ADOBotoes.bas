Attribute VB_Name = "ADOBotoes"
'*******************************************************************************************
'
'Análise...........................: Eugênio Gomes
'Programação.......................: Pablo Souza
'Data..............................: 07/04/2000
'Data da última manutenção.........: 25/04/2000
'Manutenção executada por..........: Pablo Souza
'
'Este módulo foi desenvolvido para facilitar durante o desenvolvimento a programação dos
'botões de Inclusao, Confirmação, Cancelar, Excluir e Atualizar os registros do banco de
'dados. Os botões na interface devem ter o nome igual aos usados neste modulo.
'
'*******************************************************************************************
'
Public intValor_Botao As Integer
'Irá armazenar o valor dos botões da interface conforme ordem em que aparecem, começando do
'numero 1

Public Function Incluir(Formulario As Form, Recordset_Memoria As ADODB.Recordset, TextBox_Focus As TextBox) As String
    Formulario.cmdIncluir.Enabled = False
    'Deixa o botão incluir desabilitado
    Formulario.cmdExcluir.Enabled = False
    'Deixa o botão excluir desabilitado
    Formulario.cmdConfirmar.Enabled = True
    'Deixa o botão confirmar habilitado
    Formulario.cmdCancelar.Enabled = True
    'Deixa o botão cancelar habilitado
    Formulario.CmdAtualizar.Enabled = False
    'Deixa o botão alterar desabilitado
    Recordset_Memoria.Requery
    Recordset_Memoria.AddNew
    'Adiciona um novo registro ao banco de dados
    On Error Resume Next
    TextBox_Focus.SetFocus
    'Manda o foco para o objeto desejado
    Formulario.cmdPrimeiro.Enabled = False
    Formulario.cmdAnterior.Enabled = False
    Formulario.cmdProximo.Enabled = False
    Formulario.cmdUltimo.Enabled = False
    intValor_Botao = 1

'----------------------------Instruções-----------------------------------------------------
'
'Esta é a função de inclusão, objetos necessários e nomenclatura:
'   Formulário com banco de dados aberto, pode ter qualquer nome
'   Data, pode ter qualquer nome
'   TextBox que vai receber o foco assim que o botão perder o foco, pode ter qualquer nome
'   CommanButton de acordo com o padrão processa, deve ter a figura correspondente, deve ter
'                ter o nome de: cmdIncluir
'
'Como Chamá-la?
'   Use a instrução CALL, digite o nome do módulo e em seguida o nome da função.
'       EX: CALL Botoes.Incluir(
'
'Como Preenchê-la?
'   Dentro do parênteses digite o nome do Formulário, Data, TextBox. Em seguida feche o pa-
'       rênteses.
'
'-------------------------------------------------------------------------------------------

End Function

Public Function Confirmar(Formulario As Form, Recordset_Memoria As ADODB.Recordset) As String
    On Error GoTo ErroUpdate
    Recordset_Memoria.Update
    Recordset_Memoria.Requery
    'Grava as informações no banco de dados
    Formulario.cmdIncluir.Enabled = True
    'Deixa o botão incluir Habilitado
    Formulario.cmdIncluir.SetFocus
    'Manda o foco para o botão incluir
    Formulario.cmdExcluir.Enabled = True
    'Deixa o botão incluir Habilitado
    Formulario.cmdConfirmar.Enabled = False
    'Deixa o botão confirmar desabilitado
    Formulario.cmdCancelar.Enabled = False
    'Deixa o botão cancelar desabilitado
    Formulario.CmdAtualizar.Enabled = True
    'Deixa o botão alterar Habilitado
    Formulario.cmdPrimeiro.Enabled = True
    Formulario.cmdAnterior.Enabled = True
    Formulario.cmdProximo.Enabled = True
    Formulario.cmdUltimo.Enabled = True
    intValor_Botao = 2

    Exit Function
ErroUpdate:
    If Err.Number = 3201 Or Err.Number = -2147467259 Then
        MsgBox "Erro do Banco. Alguns dos códigos digitados não estão previamente cadastrados. Se o problema persistir entre em contato com o distribuidor do software", vbCritical, "Director"
        Exit Function
    End If
'----------------------------Instruções---------------------------------------'
'                                                                             '
'Esta é a função de confirmação, objetos necessários e nomenclatura:          '
'   Formulário com banco de dados aberto, pode ter qualquer nome              '
'   Data, pode ter qualquer nome                                              '
'   CommanButton de acordo com o padrão processa, deve ter a figura correspon-'
'                dente, deve ter o nome de: cmdConfirmar                      '
'                                                                             '
'Como Chamá-la?                                                               '
'  Use a instrução CALL, digite o nome do módulo e em seguida o nome da função'
'       EX: CALL Botoes.Confirmar(                                            '
'                                                                             '
'Como Preenchê-la?                                                            '
'   Dentro do parênteses digite o nome do Formulário, Data. Em seguida feche o'
'parênteses.                                                                  '
'                                                                             '
'-----------------------------------------------------------------------------'

End Function

Public Function Cancelar(Formulario As Form, Recordset_Memoria As ADODB.Recordset) As String
    On Error GoTo ErroUpdate

    Recordset_Memoria.CancelUpdate
    Recordset_Memoria.Requery
    'Cancela a gravação no banco de dados
    Formulario.cmdIncluir.Enabled = True
    'Deixa o botão incluir habilitado
    Formulario.cmdIncluir.SetFocus
    'manda o foco para o botão incluir
    Formulario.cmdExcluir.Enabled = True
    'Deixa o botão excluir habilitado
    Formulario.cmdConfirmar.Enabled = False
    'Deixa o botão confirmar desabilitado
    Formulario.cmdCancelar.Enabled = False
    'Deixa o botão cancelar desabilitado
    Formulario.CmdAtualizar.Enabled = True
    'Deixa o botão alterar habilitado
    Formulario.cmdPrimeiro.Enabled = True
    Formulario.cmdAnterior.Enabled = True
    Formulario.cmdProximo.Enabled = True
    Formulario.cmdUltimo.Enabled = True
    intValor_Botao = 3
    Exit Function
ErroUpdate:
If Err.Number = -2147217842 Then

    Unload Formulario
    Formulario.Show
  End If



'----------------------------Instruções-----------------------------------------------------
'
'Esta é a função de cancelamento, objetos necessários e nomenclatura:
'   Formulário com banco de dados aberto, pode ter qualquer nome
'   Data, pode ter qualquer nome
'   CommanButton de acordo com o padrão processa, deve ter a figura correspondente, deve ter
'                ter o nome de: cmdCancelar
'
'Como Chamá-la?
'   Use a instrução CALL, digite o nome do módulo e em seguida o nome da função.
'       EX: CALL Botoes.Cancelar(
'
'Como Preenchê-la?
'   Dentro do parênteses digite o nome do Formulário, Data. Em seguida feche o parênteses.
'
'-------------------------------------------------------------------------------------------

End Function

Public Function Excluir(Recordset_Memoria As ADODB.Recordset) As Boolean
    Dim bytResultado As Byte
    'Variável que vai receber o valor da mensagem
    bytResultado = MsgBox("Atenção! Antes de confirmar esta operação, certifique-se que seu BACKUP esteja atualizado. Confirma a exclusão?", vbQuestion + vbYesNo, "Director")
    'Atribuição da variável à caixa de mensagem, que retorna True ou False de acordo com a
    'interação do usuário ao responder à mensagem que será exibida na tela
    If bytResultado = vbYes Then Excluir = True
    'Atribui o retorno da mensagem à variável bytResultado como True
    If bytResultado = vbNo Then Excluir = False
    'Atribui o retorno da mensagem à variável bytResultado como False
    intValor_Botao = 4

    If Excluir = True Then
    'Verifica se o valor for True, ou seja o usuário tem certeza de que quer excluir
        On Error GoTo ErroExclusao
        'Trata o erro quando ocorre
        Recordset_Memoria.Delete
        Recordset_Memoria.Requery
        'Deleta o registro atual
        Recordset_Memoria.MoveFirst
        'Move o banco de daos para o primeiro registro
        If Recordset_Memoria.BOF Then
        'Caso o ponteiro banco de dados vá para antes do primeiro registro... (isto acontece
        'quando o registro deletado é o primeiro
            Recordset_Memoria.MoveNext
            'Move o ponteiro do banco de dados para o próximo registro. Caso não haja um, vai
            'acontecer um erro que é tratado na instrução On erro GoTo acima...
        ElseIf Recordset_Memoria.EOF Then
            Recordset_Memoria.MovePrevious
        End If
    End If
    Exit Function
    'Força a saída da função para que não leia as linhas abaixo
ErroExclusao:
'Tratamento de erro
If Err.Number = 3021 Then
'O erro número 3021 acontece quando não há registro para ser exibido nos controles
    MsgBox "Não há nenhum registro para ser excluído", vbCritical, "Director"
    'Emite mensagem ao usuário
    Exit Function
    'Força a saída da função
End If

'----------------------------Instruções-----------------------------------------------------
'
'Esta é a função de cancelamento, objetos necessários e nomenclatura:
'   Formulário com banco de dados aberto, pode ter qualquer nome
'   Data, pode ter qualquer nome
'   CommanButton de acordo com o padrão processa, deve ter a figura correspondente, deve ter
'                ter o nome de: cmdExcluir
'
'Como Chamá-la?
'   Use a instrução CALL, digite o nome do módulo e em seguida o nome da função.
'       EX: CALL Botoes.Excluir(
'
'Como Preenchê-la?
'   Dentro do parênteses digite o nome do Data. Em seguida feche o parênteses.
'
'-------------------------------------------------------------------------------------------

End Function

Public Function Atualizar(Recordset_Memoria As ADODB.Recordset) As String
    On Error Resume Next
    Recordset_Memoria.Requery
    intValor_Botao = 5
End Function

Public Function Primeiro(Recordset_Memoria As ADODB.Recordset) As String
    On Error Resume Next
    Recordset_Memoria.MoveFirst
End Function

Public Function Anterior(Recordset_Memoria As ADODB.Recordset) As String
    On Error Resume Next
    Recordset_Memoria.MovePrevious
    If Recordset_Memoria.BOF Then Recordset_Memoria.MoveNext
End Function

Public Function Proximo(Recordset_Memoria As ADODB.Recordset) As String
    On Error Resume Next
    Recordset_Memoria.MoveNext
    If Recordset_Memoria.EOF Then Recordset_Memoria.MovePrevious
End Function

Public Function Ultimo(Recordset_Memoria As ADODB.Recordset) As String
    On Error Resume Next
    Recordset_Memoria.MoveLast
End Function
