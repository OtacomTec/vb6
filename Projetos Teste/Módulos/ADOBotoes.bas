Attribute VB_Name = "ADOBotoes"
'*******************************************************************************************
'
'An�lise...........................: Eug�nio Gomes
'Programa��o.......................: Pablo Souza
'Data..............................: 07/04/2000
'Data da �ltima manuten��o.........: 25/04/2000
'Manuten��o executada por..........: Pablo Souza
'
'Este m�dulo foi desenvolvido para facilitar durante o desenvolvimento a programa��o dos
'bot�es de Inclusao, Confirma��o, Cancelar, Excluir e Atualizar os registros do banco de
'dados. Os bot�es na interface devem ter o nome igual aos usados neste modulo.
'
'*******************************************************************************************
'
Public intValor_Botao As Integer
'Ir� armazenar o valor dos bot�es da interface conforme ordem em que aparecem, come�ando do
'numero 1

Public Function Incluir(Formulario As Form, Recordset_Memoria As ADODB.Recordset, TextBox_Focus As TextBox) As String
    Formulario.cmdIncluir.Enabled = False
    'Deixa o bot�o incluir desabilitado
    Formulario.cmdExcluir.Enabled = False
    'Deixa o bot�o excluir desabilitado
    Formulario.cmdConfirmar.Enabled = True
    'Deixa o bot�o confirmar habilitado
    Formulario.cmdCancelar.Enabled = True
    'Deixa o bot�o cancelar habilitado
    Formulario.CmdAtualizar.Enabled = False
    'Deixa o bot�o alterar desabilitado
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

'----------------------------Instru��es-----------------------------------------------------
'
'Esta � a fun��o de inclus�o, objetos necess�rios e nomenclatura:
'   Formul�rio com banco de dados aberto, pode ter qualquer nome
'   Data, pode ter qualquer nome
'   TextBox que vai receber o foco assim que o bot�o perder o foco, pode ter qualquer nome
'   CommanButton de acordo com o padr�o processa, deve ter a figura correspondente, deve ter
'                ter o nome de: cmdIncluir
'
'Como Cham�-la?
'   Use a instru��o CALL, digite o nome do m�dulo e em seguida o nome da fun��o.
'       EX: CALL Botoes.Incluir(
'
'Como Preench�-la?
'   Dentro do par�nteses digite o nome do Formul�rio, Data, TextBox. Em seguida feche o pa-
'       r�nteses.
'
'-------------------------------------------------------------------------------------------

End Function

Public Function Confirmar(Formulario As Form, Recordset_Memoria As ADODB.Recordset) As String
    On Error GoTo ErroUpdate
    Recordset_Memoria.Update
    Recordset_Memoria.Requery
    'Grava as informa��es no banco de dados
    Formulario.cmdIncluir.Enabled = True
    'Deixa o bot�o incluir Habilitado
    Formulario.cmdIncluir.SetFocus
    'Manda o foco para o bot�o incluir
    Formulario.cmdExcluir.Enabled = True
    'Deixa o bot�o incluir Habilitado
    Formulario.cmdConfirmar.Enabled = False
    'Deixa o bot�o confirmar desabilitado
    Formulario.cmdCancelar.Enabled = False
    'Deixa o bot�o cancelar desabilitado
    Formulario.CmdAtualizar.Enabled = True
    'Deixa o bot�o alterar Habilitado
    Formulario.cmdPrimeiro.Enabled = True
    Formulario.cmdAnterior.Enabled = True
    Formulario.cmdProximo.Enabled = True
    Formulario.cmdUltimo.Enabled = True
    intValor_Botao = 2

    Exit Function
ErroUpdate:
    If Err.Number = 3201 Or Err.Number = -2147467259 Then
        MsgBox "Erro do Banco. Alguns dos c�digos digitados n�o est�o previamente cadastrados. Se o problema persistir entre em contato com o distribuidor do software", vbCritical, "Director"
        Exit Function
    End If
'----------------------------Instru��es---------------------------------------'
'                                                                             '
'Esta � a fun��o de confirma��o, objetos necess�rios e nomenclatura:          '
'   Formul�rio com banco de dados aberto, pode ter qualquer nome              '
'   Data, pode ter qualquer nome                                              '
'   CommanButton de acordo com o padr�o processa, deve ter a figura correspon-'
'                dente, deve ter o nome de: cmdConfirmar                      '
'                                                                             '
'Como Cham�-la?                                                               '
'  Use a instru��o CALL, digite o nome do m�dulo e em seguida o nome da fun��o'
'       EX: CALL Botoes.Confirmar(                                            '
'                                                                             '
'Como Preench�-la?                                                            '
'   Dentro do par�nteses digite o nome do Formul�rio, Data. Em seguida feche o'
'par�nteses.                                                                  '
'                                                                             '
'-----------------------------------------------------------------------------'

End Function

Public Function Cancelar(Formulario As Form, Recordset_Memoria As ADODB.Recordset) As String
    On Error GoTo ErroUpdate

    Recordset_Memoria.CancelUpdate
    Recordset_Memoria.Requery
    'Cancela a grava��o no banco de dados
    Formulario.cmdIncluir.Enabled = True
    'Deixa o bot�o incluir habilitado
    Formulario.cmdIncluir.SetFocus
    'manda o foco para o bot�o incluir
    Formulario.cmdExcluir.Enabled = True
    'Deixa o bot�o excluir habilitado
    Formulario.cmdConfirmar.Enabled = False
    'Deixa o bot�o confirmar desabilitado
    Formulario.cmdCancelar.Enabled = False
    'Deixa o bot�o cancelar desabilitado
    Formulario.CmdAtualizar.Enabled = True
    'Deixa o bot�o alterar habilitado
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



'----------------------------Instru��es-----------------------------------------------------
'
'Esta � a fun��o de cancelamento, objetos necess�rios e nomenclatura:
'   Formul�rio com banco de dados aberto, pode ter qualquer nome
'   Data, pode ter qualquer nome
'   CommanButton de acordo com o padr�o processa, deve ter a figura correspondente, deve ter
'                ter o nome de: cmdCancelar
'
'Como Cham�-la?
'   Use a instru��o CALL, digite o nome do m�dulo e em seguida o nome da fun��o.
'       EX: CALL Botoes.Cancelar(
'
'Como Preench�-la?
'   Dentro do par�nteses digite o nome do Formul�rio, Data. Em seguida feche o par�nteses.
'
'-------------------------------------------------------------------------------------------

End Function

Public Function Excluir(Recordset_Memoria As ADODB.Recordset) As Boolean
    Dim bytResultado As Byte
    'Vari�vel que vai receber o valor da mensagem
    bytResultado = MsgBox("Aten��o! Antes de confirmar esta opera��o, certifique-se que seu BACKUP esteja atualizado. Confirma a exclus�o?", vbQuestion + vbYesNo, "Director")
    'Atribui��o da vari�vel � caixa de mensagem, que retorna True ou False de acordo com a
    'intera��o do usu�rio ao responder � mensagem que ser� exibida na tela
    If bytResultado = vbYes Then Excluir = True
    'Atribui o retorno da mensagem � vari�vel bytResultado como True
    If bytResultado = vbNo Then Excluir = False
    'Atribui o retorno da mensagem � vari�vel bytResultado como False
    intValor_Botao = 4

    If Excluir = True Then
    'Verifica se o valor for True, ou seja o usu�rio tem certeza de que quer excluir
        On Error GoTo ErroExclusao
        'Trata o erro quando ocorre
        Recordset_Memoria.Delete
        Recordset_Memoria.Requery
        'Deleta o registro atual
        Recordset_Memoria.MoveFirst
        'Move o banco de daos para o primeiro registro
        If Recordset_Memoria.BOF Then
        'Caso o ponteiro banco de dados v� para antes do primeiro registro... (isto acontece
        'quando o registro deletado � o primeiro
            Recordset_Memoria.MoveNext
            'Move o ponteiro do banco de dados para o pr�ximo registro. Caso n�o haja um, vai
            'acontecer um erro que � tratado na instru��o On erro GoTo acima...
        ElseIf Recordset_Memoria.EOF Then
            Recordset_Memoria.MovePrevious
        End If
    End If
    Exit Function
    'For�a a sa�da da fun��o para que n�o leia as linhas abaixo
ErroExclusao:
'Tratamento de erro
If Err.Number = 3021 Then
'O erro n�mero 3021 acontece quando n�o h� registro para ser exibido nos controles
    MsgBox "N�o h� nenhum registro para ser exclu�do", vbCritical, "Director"
    'Emite mensagem ao usu�rio
    Exit Function
    'For�a a sa�da da fun��o
End If

'----------------------------Instru��es-----------------------------------------------------
'
'Esta � a fun��o de cancelamento, objetos necess�rios e nomenclatura:
'   Formul�rio com banco de dados aberto, pode ter qualquer nome
'   Data, pode ter qualquer nome
'   CommanButton de acordo com o padr�o processa, deve ter a figura correspondente, deve ter
'                ter o nome de: cmdExcluir
'
'Como Cham�-la?
'   Use a instru��o CALL, digite o nome do m�dulo e em seguida o nome da fun��o.
'       EX: CALL Botoes.Excluir(
'
'Como Preench�-la?
'   Dentro do par�nteses digite o nome do Data. Em seguida feche o par�nteses.
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
