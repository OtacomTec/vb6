Attribute VB_Name = "Funcoes_botoes"
Public Function Habilita_Desabilita_Botoes(Evento As String, Formulario As Form) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Grupo Mil - Setor de Desenvolvimento
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'M�dulo............................: Navega��o
'Procedimento/Fun��o...............: Navega��o
'Objetivo:.........................: Faz o controle Habilita/Desabilita dos Botoes
'                                    do barra de ferramentas
'Desenvolvimento...................: Marcos Bai�o
'Data de cria��o...................: 10/07/2002
'Observa��oes......................:
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Evento = "Load" Then
        Formulario.tlbBotoes.Buttons.Item(1).Enabled = True
        Formulario.tlbBotoes.Buttons.Item(2).Enabled = False
        Formulario.tlbBotoes.Buttons.Item(3).Enabled = False
        Formulario.tlbBotoes.Buttons.Item(4).Enabled = False
    End If
    
    If Evento = "DataGrid" Then
        Formulario.tlbBotoes.Buttons.Item(1).Enabled = False
        Formulario.tlbBotoes.Buttons.Item(2).Enabled = True
        Formulario.tlbBotoes.Buttons.Item(3).Enabled = True
        Formulario.tlbBotoes.Buttons.Item(4).Enabled = True
    End If
    
    Exit Function
Erro:
    Call Erro.Erro(Habilita_Desabilita_Botoes)
    Resume Next

End Function

