Attribute VB_Name = "Temp_Folha"
''''''''''''Option Explicit
''''''''''''
''''''''''''Private Sub Calculo_Evento_Salario_Liquido()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Calculo_Evento_Salario_Liquido
'''''''''''''Objetivo:.........................: Calcula o evento Salario Liquido.
'''''''''''''Desenvolvimento...................: Wascley Costa
'''''''''''''Data de criação...................: 24/07/2001
'''''''''''''Data da última manutenção.........: 21/09/200124/07/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''    Dim intCod_Evento As Integer
''''''''''''    Dim datData As Date
''''''''''''    Dim strSQL As String
''''''''''''
''''''''''''    'Busca na tabela de opções o código do evento salário base.
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 108")
''''''''''''    intCod_Evento = adrOpcoes("DFvalor")
''''''''''''
''''''''''''    booSalario_Calculado = False
''''''''''''    'Verifica se o evento valor atual é o Salário Base.
''''''''''''    If adrFiltro_secundario("DFcod_evento") = intCod_Evento Then
''''''''''''        booSalario_Calculado = True
''''''''''''
''''''''''''        'Tratando agora o conteúdo do campo DFvalor_referencia relativo a este evento.
''''''''''''        strSQL = "SELECT TBregime.DFquantidade "
''''''''''''        strSQL = strSQL & "From TBregime "
''''''''''''        strSQL = strSQL & "INNER JOIN TBcargo "
''''''''''''        strSQL = strSQL & "ON TBregime.DFid_regime = TBcargo.DFid_regime "
''''''''''''        strSQL = strSQL & "INNER JOIN TBfuncionario "
''''''''''''        strSQL = strSQL & "ON TBcargo.DFcod_cargo = TBfuncionario.DFcod_cargo "
''''''''''''        strSQL = strSQL & "WHERE TBfuncionario.DFmatricula = " & adrFiltro_secundario("DFmatricula")
''''''''''''        'Gravando na ADR principal o valor da referência referente a esse evento.
''''''''''''        adrFiltro_secundario("DFvalor_referencia") = CStr(CNConexao.Execute(strSQL).Fields("DFquantidade"))
''''''''''''
''''''''''''        'Busca na tabela de funcionario qual a sua data de admissão.
''''''''''''        datData = CNConexao.Execute("SELECT DFdata_admissao FROM TBfuncionario WHERE DFmatricula = " & adrFiltro_secundario("DFmatricula")).Fields("DFData_admissao")
''''''''''''        'Verifica se o funcionario foi admitido no mês atual.
''''''''''''        If (Year(datData) = Year(dtpMes_ano)) And (Month(datData) = Month(dtpMes_ano)) Then
''''''''''''            'Verifica a quantidade de dias que o funcionario trabalhou e calcula o seu salario.
''''''''''''            adrFiltro_secundario("DFvalor") = IIf((30 - Day(datData) + 1) = 0, 1 * (adrFiltro_secundario("DFvalor") / 30), (30 - Day(datData) + 1)) * (adrFiltro_secundario("DFvalor") / 30)
''''''''''''            adrFiltro_secundario("DFvalor_referencia") = (30 - Day(datData) + 1)
''''''''''''        End If
''''''''''''
''''''''''''        'Busca na tabela de funcionario qual a sua data de demissão se ele foi demitido.
''''''''''''        If IsNull(CNConexao.Execute("SELECT DFData_demissao FROM TBfuncionario WHERE DFmatricula = " & adrFiltro_secundario("DFmatricula")).Fields("DFData_demissao")) Then
''''''''''''            Exit Sub
''''''''''''        Else
''''''''''''            datData = CNConexao.Execute("SELECT DFData_demissao FROM TBfuncionario WHERE DFmatricula = " & adrFiltro_secundario("DFmatricula")).Fields("DFData_demissao")
''''''''''''        End If
''''''''''''        'Verifica se o funcionario foi demitido no mês atual.
''''''''''''        If (Year(datData) = Year(Date)) And (Month(datData) = Month(Date)) Then
''''''''''''            'Verifica a quantidade de dias que o funcionario trabalhou e calcula o seu salario.
''''''''''''            adrFiltro_secundario("DFvalor") = Day(datData) * (adrFiltro_secundario("DFvalor") / 30)
''''''''''''        End If
''''''''''''
''''''''''''    End If
''''''''''''
''''''''''''    Exit Sub
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Calculo_Evento_Salario_Liquido")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''Private Sub Grava_Recibo_Novo()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Grava_Recibo_Novo
'''''''''''''Objetivo:.........................: Grava um novo ID na tabela TBrecibo_pagamento
'''''''''''''Desenvolvimento...................: Wascley Costa
'''''''''''''Data de criação...................: 09/03/2001
'''''''''''''Data da última manutenção.........: 21/09/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''    Dim strSQL As String
''''''''''''    Dim adrRecibo_pagamento As ADODB.Recordset
''''''''''''
''''''''''''    'Esta função deleta um calculo anterior do mesmo funcionario e no mesmo periodo,
''''''''''''    'caso o usuário solicite calcular novamente.
''''''''''''    Call Deleta_Recibo_Antigo
''''''''''''
''''''''''''    'Criação da Sql para inserção de um novo recibo.
''''''''''''    strSQL = "INSERT INTO TBrecibo_pagamento " & _
''''''''''''             "([DFmatricula], " & _
''''''''''''             "[DFmes_ano_calculo], " & _
''''''''''''             "[DFsalario_base], " & _
''''''''''''             "[DFsalario_contribuicao_inss], " & _
''''''''''''             "[DFbase_calculo_fgts], " & _
''''''''''''             "[DFfgts_mes], " & _
''''''''''''             "[DFbase_calculo_irrf], " & _
''''''''''''             "[DFfaixa_irrf], " & _
''''''''''''             "[DFtotal_vencimento], " & _
''''''''''''             "[DFtotal_descontos]) " & _
''''''''''''             "SELECT " & adrFiltro_secundario("DFmatricula") & _
''''''''''''             ", '" & Format(datMes_ano, "yyyyMMdd") & "'" & _
''''''''''''             ", 0, 0, 0, 0, 0, 0 ,0 ,0"
''''''''''''
''''''''''''    'Executa a Sql criada acima.
''''''''''''    CNConexao.Execute strSQL
''''''''''''
''''''''''''    'Criação da ADR para guardar os Id's da tabela TBrecibo_pagamento.
''''''''''''    Call Banco_Dados.SQLgeral("SELECT DFid_recibo_pagamento FROM TBrecibo_pagamento ORDER BY DFid_recibo_pagamento", adrRecibo_pagamento)
''''''''''''    'Posiciona a ADR no seu último registro o qual foi gravado neste momento.
''''''''''''    adrRecibo_pagamento.MoveLast
''''''''''''
''''''''''''    'Atribui o Id do Recibo atual para uma variavel de memória.
''''''''''''    lonId_recibo = adrRecibo_pagamento("DFid_recibo_pagamento")
''''''''''''
''''''''''''    'Limpa e destrói a adrRecibo_pagamento.
''''''''''''    adrRecibo_pagamento.Close
''''''''''''    Set adrRecibo_pagamento = Nothing
''''''''''''
''''''''''''    Exit Sub
''''''''''''
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Grava_Recibo_Novo")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''Private Sub Deleta_Recibo_Antigo()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Deleta_Recibo_Antigo
'''''''''''''Objetivo:.........................: Se já existir o calculo solicitado esta função,
'''''''''''''                                    deleta o calculo antigo.
'''''''''''''Desenvolvimento...................: Wascley Costa
'''''''''''''Data de criação...................: 14/03/2001
'''''''''''''Data da última manutenção.........: 21/09/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''    Dim strSQL As String
''''''''''''
''''''''''''    'Criação da Sql para verificar se existe algum recibo já gravado deste funcionario deste periodo.
''''''''''''    strSQL = "SELECT DFid_recibo_pagamento FROM TBrecibo_pagamento " & _
''''''''''''             "WHERE DFmatricula = " & adrFiltro_secundario("DFmatricula") & " " & _
''''''''''''             "AND DFmes_ano_calculo = '" & Format(datMes_ano, "yyyyMMdd") & "'"
''''''''''''
''''''''''''    If CNConexao.Execute(Replace(strSQL, "DFid_recibo_pagamento", "COUNT(*)"))(0) <> 0 Then
''''''''''''        'Se houver registros então deleta-se o Recibo e os Eventos.
''''''''''''        CNConexao.Execute "DELETE FROM TBevento_recibo WHERE DFid_recibo_pagamento = " & CNConexao.Execute(strSQL)("DFid_recibo_pagamento")
''''''''''''        CNConexao.Execute "DELETE FROM TBrecibo_pagamento WHERE DFid_recibo_pagamento = " & CNConexao.Execute(strSQL)("DFid_recibo_pagamento")
''''''''''''        'Esta variavel é usada para saber se será necessário atualizar o atributo DFparcela dos eventos,
''''''''''''        'sendo que se o recibo já existir essa ação não será executada.
''''''''''''        booRecibo_existente = True
''''''''''''    Else
''''''''''''        booRecibo_existente = False
''''''''''''    End If
''''''''''''
''''''''''''    Exit Sub
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Deleta_Recibo_Antigo")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''Private Sub Calculo_Percentual_Salario_Base()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Calcula_Percentual_Salario_Base
'''''''''''''Objetivo:.........................: Calcular o atributo valor com base no percentual
'''''''''''''                                    sobre o salario base.
'''''''''''''Desenvolvimento...................: Wascley Costa
'''''''''''''Data de criação...................: 12/03/2001
'''''''''''''Data da última manutenção.........: 21/09/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''    Dim strSQL As String
''''''''''''
''''''''''''    'Criação da Sql para busca do salario base do funcionario atual.
''''''''''''    strSQL = "SELECT DFvalor FROM TBfixo " & _
''''''''''''             "WHERE DFmatricula = " & lonFuncionario_atual & " " & _
''''''''''''             "AND DFcod_evento = " & intCodigo_Salario_Base
''''''''''''    'Calculo do valor com base no percentual sobre o salario base.
''''''''''''    adrFiltro_secundario("DFvalor") = (CNConexao.Execute(strSQL)("DFvalor") * adrFiltro_secundario("DFpercentual_sob_salario_base")) / 100
''''''''''''
''''''''''''    Exit Sub
''''''''''''
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Calculo_Percentual_Salario_Base")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''Private Sub Calcular_Afastamentos()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Afastamento
'''''''''''''Objetivo:.........................: Verifica o afastamento do Funcionario
'''''''''''''Desenvolvimento...................: Vitor Constâncio da Silva
'''''''''''''Data de criação...................: 02/10/2001
'''''''''''''Data da última manutenção.........: 00/00/0000
'''''''''''''Manutenção executada por..........:
'''''''''''''******************************************************************************
''''''''''''    Dim strValores As String
''''''''''''    Dim strAfastamentos As String
''''''''''''    Dim intDias_Trabalhados As Integer
''''''''''''    Dim intDias_Afastado As Integer
''''''''''''    Dim intDias_Restantes As Integer
''''''''''''    Dim intInicio_Mes As Integer
''''''''''''    Dim intFim_Mes As Integer
''''''''''''    Dim curSal_afastamento As Integer
''''''''''''    Dim strSQL As String
''''''''''''    Dim booMes_Calculo_Afast As Boolean
''''''''''''
''''''''''''    'BUSCANDO VALORES NECESSÁRIOS PARA O CALCULO
''''''''''''
''''''''''''    strSQL = "SELECT COUNT(*) FROM TBvariavel " & _
''''''''''''             "WHERE DFmes_ano_lancamento = '" & Format(dtpMes_ano, "yyyyMM\0\1") & "' AND DFmatricula = " & adrFiltro_primario("DFmatricula") & " " & _
''''''''''''             "AND DFcod_evento IN (" & intCodigo_Auxilio_Acidente & "," & intCodigo_Auxilio_Doenca & "," & intCodigo_Salario_maternidade & ")"
''''''''''''
''''''''''''    If CNConexao.Execute(strSQL)(0) <> 0 Then
''''''''''''        booExiste_Afastamento = False
''''''''''''        booExiste_Salario_Base = True
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''
''''''''''''    strSQL = "SELECT COUNT(*) FROM TBfixo " & _
''''''''''''             "WHERE DFmatricula = " & adrFiltro_primario("DFmatricula") & " " & _
''''''''''''             "AND DFcod_evento IN (" & intCodigo_Auxilio_Acidente & "," & intCodigo_Auxilio_Doenca & "," & intCodigo_Salario_maternidade & ")"
''''''''''''
''''''''''''    If CNConexao.Execute(strSQL)(0) <> 0 Then
''''''''''''        booExiste_Afastamento = False
''''''''''''        booExiste_Salario_Base = True
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''
''''''''''''
'''''''''''''==//==//==//==//==//==//==//==//==//==//==//==//==//==//==//==//==//==//==//==//==//
''''''''''''    'VERIFICANDO A QUANTIDADE DE DIAS TRABALHADOS NO MES E/OU SE TRABALHOU
''''''''''''    Call Verificar_Dias_Movimentacoes(dtpMes_ano, Str(adrFiltro_primario("DFmatricula")), 30, strValores)
''''''''''''    If strValores <> Empty Then
''''''''''''        Do While strValores <> Empty
''''''''''''            intInicio_Mes = Mid(strValores, 1, 2)
''''''''''''            intFim_Mes = Mid(strValores, 4, 2)
''''''''''''            If intInicio_Mes = intFim_Mes Then
''''''''''''                intDias_Trabalhados = intDias_Trabalhados + (intFim_Mes - intInicio_Mes)
''''''''''''            Else
''''''''''''                intDias_Trabalhados = intDias_Trabalhados + (intFim_Mes - intInicio_Mes) + 1
''''''''''''            End If
''''''''''''            strValores = IIf(Len(strValores) > 5, Mid(strValores, 7), Empty)
''''''''''''        Loop
''''''''''''
''''''''''''        If intDias_Trabalhados <> 0 Then
''''''''''''            booExiste_Afastamento = True
''''''''''''            booExiste_Salario_Base = True
''''''''''''        Else
''''''''''''            booExiste_Afastamento = True
''''''''''''            booExiste_Salario_Base = False
''''''''''''        End If
''''''''''''        booMes_Calculo_Afast = True
''''''''''''    Else
''''''''''''    'VERIFICANDO SE O FUNCIONARIO ESTÁ OU NÃO AFASTADO PARA MESES ANTERIORES
''''''''''''        strSQL = "SELECT DFtipo FROM TBafastamento " & _
''''''''''''                 "WHERE DFcod_afastamento = " & _
''''''''''''                    "(SELECT DFcod_movimentacao FROM TBafastamento_funcionario " & _
''''''''''''                    "WHERE DFdata_inicio = " & _
''''''''''''                        "(SELECT MAX(DFdata_inicio) AS DFdata_inicio " & _
''''''''''''                        "FROM TBafastamento_funcionario " & _
''''''''''''                        "WHERE (CONVERT(CHAR(4),YEAR(DFdata_inicio)))+ " & _
''''''''''''                        "(CONVERT(CHAR(2),MONTH(DFdata_inicio))) <= " & _
''''''''''''                        "'" & Format(dtpMes_ano, "yyyyM") & "' " & _
''''''''''''                            "AND DFmatricula = " & adrFiltro_primario("DFmatricula") & "))"
''''''''''''
''''''''''''        If CNConexao.Execute(strSQL).EOF Then
''''''''''''            booExiste_Afastamento = False
''''''''''''            booExiste_Salario_Base = True
''''''''''''            Exit Sub
''''''''''''        End If
''''''''''''
''''''''''''        If CNConexao.Execute(strSQL).Fields("DFtipo") = "R" Then
''''''''''''            booExiste_Salario_Base = True
''''''''''''            booExiste_Afastamento = False
''''''''''''            Exit Sub
''''''''''''        Else
''''''''''''            Call Verificar_Dias_Movimentacoes(DateAdd("m", -1, dtpMes_ano), Str(adrFiltro_primario("DFmatricula")), 30, strValores)
''''''''''''            If strValores <> Empty Then
''''''''''''                Do While strValores <> Empty
''''''''''''                    intInicio_Mes = Mid(strValores, 1, 2)
''''''''''''                    intFim_Mes = Mid(strValores, 4, 2)
''''''''''''                    If intInicio_Mes = intFim_Mes Then
''''''''''''                        intDias_Trabalhados = intDias_Trabalhados + (intFim_Mes - intInicio_Mes)
''''''''''''                    Else
''''''''''''                        intDias_Trabalhados = intDias_Trabalhados + (intFim_Mes - intInicio_Mes) + 1
''''''''''''                    End If
''''''''''''                    strValores = IIf(Len(strValores) > 5, Mid(strValores, 7), Empty)
''''''''''''                Loop
''''''''''''                If intDias_Trabalhados >= 16 Then
''''''''''''                    booExiste_Salario_Base = False
''''''''''''                    booExiste_Afastamento = True
''''''''''''                Else
''''''''''''                    booExiste_Salario_Base = False
''''''''''''                    booExiste_Afastamento = False
''''''''''''                    Exit Sub
''''''''''''                End If
''''''''''''            Else
''''''''''''                booExiste_Salario_Base = False
''''''''''''                booExiste_Afastamento = False
''''''''''''                Exit Sub
''''''''''''            End If
''''''''''''        End If
''''''''''''        booMes_Calculo_Afast = False
''''''''''''    End If
''''''''''''
''''''''''''    If intDias_Trabalhados > 30 Then
''''''''''''        intDias_Trabalhados = 30
''''''''''''    End If
''''''''''''
''''''''''''    If adrFiltro_secundario("DFcod_evento") = intCodigo_Auxilio_Acidente Then
''''''''''''        If intCont_Acidente <> 0 Then
''''''''''''            If Not adrAcidente.EOF Then
''''''''''''                adrAcidente.MoveNext
''''''''''''            End If
''''''''''''        Else
''''''''''''            strSQL = "SELECT * FROM TBafastamento_funcionario " & _
''''''''''''                     "WHERE DFmatricula = " & adrFiltro_primario("DFmatricula") & " " & _
''''''''''''                     "AND DFcod_movimentacao = '" & strCodigo_Auxilio_Acidente & "' " & _
''''''''''''                     "AND DFdata_inicio > DATEADD( d, 15, DATEADD( m, -1, '" & Format(dtpMes_ano, "yyyyMM\0\1") & "' )) " & _
''''''''''''                     "AND DFdata_inicio < DATEADD( m, +1, '" & Format(dtpMes_ano, "yyyyMM\0\1") & "' ) " & _
''''''''''''                     "AND DFid_retorno  NOT IN (SELECT DFid_afastamento_funcionario " & _
''''''''''''                                               "FROM TBafastamento_funcionario " & _
''''''''''''                                               "WHERE DFmatricula = " & adrFiltro_primario("DFmatricula") & " " & _
''''''''''''                                               "AND DFcod_movimentacao = '" & strCodigo_Retorno_Acidente & "' " & _
''''''''''''                                               "AND MONTH( DFdata_inicio )  =  MONTH( '" & Format(dtpMes_ano, "yyyyMM\0\1") & "' ) -1) "
''''''''''''            Call Banco_Dados.SQLgeral(strSQL, adrAcidente)
''''''''''''        End If
''''''''''''
''''''''''''        If Not IsNull(adrAcidente("DFid_retorno")) Then
''''''''''''            Dim datFinal As Date
''''''''''''            Dim intDias_Pagar As Integer
''''''''''''
''''''''''''            datFinal = CNConexao.Execute("SELECT DFdata_inicio FROM TBafastamento_funcionario WHERE DFid_afastamento_funcionario = " & adrAcidente("DFid_retorno"))(0)
''''''''''''
''''''''''''            If Month(adrAcidente("DFdata_inicio")) = Month(datFinal) And Year(adrAcidente("DFdata_inicio")) = Year(datFinal) Then
''''''''''''                intDias_Afastado = DateDiff("d", adrAcidente("DFdata_inicio"), datFinal)
''''''''''''            ElseIf Month(adrAcidente("DFdata_inicio")) <> Month(dtpMes_ano) And Year(adrAcidente("DFdata_inicio")) <> Year(dtpMes_ano) Then
''''''''''''                intDias_Pagar = 15 - DateDiff("d", adrAcidente("DFdata_inicio"), Format(adrAcidente("DFdata_inicio"), "yyyyMM\3\0"))
''''''''''''                intDias_Afastado = IIf(DateDiff("d", Format(dtpMes_ano, "yyyyMM\0\1"), datFinal) > intDias_Pagar, intDias_Pagar, DateDiff("d", Format(dtpMes_ano, "yyyyMM\0\1"), datFinal))
''''''''''''            Else
''''''''''''                intDias_Afastado = DateDiff("d", adrAcidente("DFdata_inicio"), Format(dtpMes_ano, "\3\0/MM/yyyy"))
''''''''''''            End If
''''''''''''        Else
''''''''''''            If Month(adrAcidente("DFdata_inicio")) = Month(dtpMes_ano) And Year(adrAcidente("DFdata_inicio")) = Year(dtpMes_ano) Then
''''''''''''                intDias_Afastado = DateDiff("d", adrAcidente("DFdata_inicio"), Format(dtpMes_ano, "yyyyMM\3\0"))
''''''''''''            ElseIf Month(adrAcidente("DFdata_inicio")) <> Month(dtpMes_ano) And Year(adrAcidente("DFdata_inicio")) <> Year(dtpMes_ano) Then
''''''''''''                intDias_Pagar = 15 - DateDiff("d", adrAcidente("DFdata_inicio"), Format(adrAcidente("DFdata_inicio"), "yyyyMM\3\0"))
''''''''''''                intDias_Afastado = IIf(DateDiff("d", Format(dtpMes_ano, "yyyyMM\0\1"), datFinal) > intDias_Pagar, intDias_Pagar, DateDiff("d", Format(dtpMes_ano, "yyyyMM\0\1"), datFinal))
''''''''''''            End If
''''''''''''        End If
''''''''''''        intDias_Afastado = IIf(intDias_Afastado >= 15, 15, intDias_Afastado + 1)
''''''''''''        curSal_afastamento = (CNConexao.Execute("SELECT DFsalario FROM TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")).Fields("DFsalario") / 30) * intDias_Afastado
''''''''''''        intCont_Acidente = intCont_Acidente + 1
''''''''''''
''''''''''''    ElseIf adrFiltro_secundario("DFcod_evento") = intCodigo_Auxilio_Doenca Then
''''''''''''        intCont_Doenca = intCont_Doenca + 1
''''''''''''
''''''''''''        If intCont_Doenca <= 1 Then
''''''''''''            strSQL = "SELECT * FROM TBafastamento_funcionario " & _
''''''''''''                     "WHERE DFmatricula = " & adrFiltro_primario("DFmatricula") & " " & _
''''''''''''                     "AND ( DFcod_movimentacao = " & strCodigo_Auxilio_Acidente & " " & _
''''''''''''                        "OR DFid_afastamento = ( SELECT DFid_afastamento_funcionario FROM TBafastamento_funcionario WHERE DFcod_movimentacao = 'O1') ) " & _
''''''''''''                     "AND MONTH( DFdata_inicio ) = MONTH( '" & Format(dtpMes_ano, "yyyyMMdd") & "' ) " & _
''''''''''''                     "AND YEAR( DFdata_inicio ) = YEAR( '" & Format(dtpMes_ano, "yyyyMMdd") & "' ) "
''''''''''''            Call Banco_Dados.SQLgeral(strSQL, adrDoenca)
''''''''''''        End If
''''''''''''
''''''''''''        If adrDoenca.RecordCount = 1 Then
''''''''''''            If IsNull(adrDoenca("DFid_afastamento")) Then
''''''''''''                intDias_Afastado = DateDiff("d", adrDoenca("DFdata_inicio"), Format(dtpMes_ano, "3\0\/MM/yyyy"))
''''''''''''            ElseIf Not IsNull(adrDoenca("DFid_afastamento")) Then
''''''''''''                If adrDoenca.EOF Then
''''''''''''                    Exit Sub
''''''''''''                End If
''''''''''''                If CNConexao.Execute("SELECT DFcod_movimentacao FROM TBafastamento_funcionario WHERE DFid_afastamento_funcionario = " & adrDoenca("DFid_afastamento"))(0) = strCodigo_Auxilio_Doenca Then
''''''''''''                    intDias_Afastado = (CNConexao.Execute("SELECT DFdata_inicio FROM TBafastamento_funcionario WHERE DFid_afastamento_funcionario = " & adrDoenca("DFid_afastamento"))(0) + 15) - 30
''''''''''''                    If intDias_Afastado = 0 Then
''''''''''''                        Exit Sub
''''''''''''                    End If
''''''''''''                Else
''''''''''''                    Exit Sub
''''''''''''                End If
''''''''''''            End If
''''''''''''        ElseIf adrDoenca.RecordCount = 2 Then
''''''''''''            If IsNull(adrDoenca("DFid_afastamento")) Then
''''''''''''                If intCont_Doenca = 1 Then
''''''''''''                    intDias_Afastado = 15
''''''''''''                Else
''''''''''''                    intDias_Afastado = DateDiff("d", adrDoenca("DFdata_inicio"), Format(dtpMes_ano, "3\0\/MM/yyyy"))
''''''''''''                End If
''''''''''''            ElseIf Not IsNull(adrDoenca("DFid_afastamento")) Then
''''''''''''
''''''''''''                If adrDoenca.EOF Then
''''''''''''                    Exit Sub
''''''''''''                End If
''''''''''''
''''''''''''                If CNConexao.Execute("SELECT DFcod_movimentacao FROM TBafastamento_funcionario WHERE DFid_afastamento_funcionario = " & adrDoenca("DFid_afastamento"))(0) = strCodigo_Auxilio_Doenca Then
''''''''''''
''''''''''''                    If intCont_Doenca = 1 Then
''''''''''''                        intDias_Afastado = (CNConexao.Execute("SELECT CONVERT(INT,DAY(DFdata_inicio)) FROM TBafastamento_funcionario WHERE DFid_afastamento_funcionario = " & adrDoenca("DFid_afastamento"))(0) + 15) - 30
''''''''''''                    Else
''''''''''''                        intDias_Afastado = DateDiff("d", adrDoenca("DFdata_inicio"), Format(dtpMes_ano, "3\0\/MM/yyyy"))
''''''''''''                    End If
''''''''''''
''''''''''''                    If intDias_Afastado = 0 Then
''''''''''''                        adrDoenca.MoveNext
''''''''''''                        Exit Sub
''''''''''''                    End If
''''''''''''                Else
''''''''''''                    adrDoenca.MoveNext
''''''''''''                    Exit Sub
''''''''''''                End If
''''''''''''            End If
''''''''''''            adrDoenca.MoveNext
''''''''''''        ElseIf adrDoenca.RecordCount = 3 Then
''''''''''''            If IsNull(adrDoenca("DFid_afastamento")) Then
''''''''''''                If intCont_Doenca = 1 Then
''''''''''''                    intDias_Afastado = 15
''''''''''''                Else
''''''''''''                    intDias_Afastado = DateDiff("d", adrDoenca("DFdata_inicio"), Format(dtpMes_ano, "3\0\/MM/yyyy"))
''''''''''''                End If
''''''''''''            ElseIf Not IsNull(adrDoenca("DFid_afastamento")) Then
''''''''''''                If adrDoenca.EOF Then
''''''''''''                    Exit Sub
''''''''''''                End If
''''''''''''                If CNConexao.Execute("SELECT DFcod_movimentacao FROM TBafastamento_funcionario WHERE DFid_afastamento_funcionario = " & adrDoenca("DFid_afastamento"))(0) = strCodigo_Auxilio_Doenca Then
''''''''''''
''''''''''''                    If intCont_Doenca = 1 Then
''''''''''''                        intDias_Afastado = (CNConexao.Execute("SELECT DFdata_inicio FROM TBafastamento_funcionario WHERE DFid_afastamento_funcionario = " & adrDoenca("DFid_afastamento"))(0) + 15) - 30
''''''''''''                    Else
''''''''''''                        intDias_Afastado = 15
''''''''''''                    End If
''''''''''''
''''''''''''                    If intDias_Afastado = 0 Then
''''''''''''                        adrDoenca.MoveNext
''''''''''''                        Exit Sub
''''''''''''                    End If
''''''''''''                Else
''''''''''''                    adrDoenca.MoveNext
''''''''''''                    Exit Sub
''''''''''''                End If
''''''''''''            End If
''''''''''''            adrDoenca.MoveNext
''''''''''''        ElseIf adrDoenca.RecordCount = 4 Then
''''''''''''            If IsNull(adrDoenca("DFid_afastamento")) Then
''''''''''''                intDias_Afastado = 15
''''''''''''            ElseIf Not IsNull(adrDoenca("DFid_afastamento")) Then
''''''''''''                If adrDoenca.EOF Then
''''''''''''                    Exit Sub
''''''''''''                End If
''''''''''''                If CNConexao.Execute("SELECT DFcod_movimentacao FROM TBafastamento_funcionario WHERE DFid_afastamento_funcionario = " & adrDoenca("DFid_afastamento"))(0) = strCodigo_Auxilio_Doenca Then
''''''''''''
''''''''''''                    If intCont_Doenca = 1 Then
''''''''''''                        intDias_Afastado = (CNConexao.Execute("SELECT DFdata_inicio FROM TBafastamento_funcionario WHERE DFid_afastamento_funcionario = " & adrDoenca("DFid_afastamento"))(0) + 15) - 30
''''''''''''                    ElseIf intCont_Doenca = 2 Then
''''''''''''                        intDias_Afastado = 15
''''''''''''                    ElseIf intCont_Doenca = 3 Then
''''''''''''                        intDias_Afastado = DateDiff("d", adrDoenca("DFdata_inicio"), Format(dtpMes_ano, "3\0\/MM/yyyy"))
''''''''''''                    End If
''''''''''''
''''''''''''                    If intDias_Afastado = 0 Then
''''''''''''                        adrDoenca.MoveNext
''''''''''''                        Exit Sub
''''''''''''                    End If
''''''''''''                Else
''''''''''''                    adrDoenca.MoveNext
''''''''''''                    Exit Sub
''''''''''''                End If
''''''''''''            End If
''''''''''''            adrDoenca.MoveNext
''''''''''''        Else
''''''''''''            Exit Sub
''''''''''''        End If
''''''''''''
''''''''''''        'CALCULANDO O SALARIO DE ACORDO COM OS DIAS TRABALHADOS E AFASTADOS CALCULADOS ACIMA
''''''''''''        intDias_Afastado = IIf(intDias_Afastado >= 15, 15, intDias_Afastado + 1)
''''''''''''        curSal_afastamento = (CNConexao.Execute("SELECT DFsalario FROM TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")).Fields("DFsalario") / 30) * intDias_Afastado
''''''''''''
''''''''''''    End If
''''''''''''
''''''''''''    curSalario = (CNConexao.Execute("SELECT DFsalario FROM TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")).Fields("DFsalario") / 30) * intDias_Trabalhados
''''''''''''    'atualizando o evento Salário Base na TBtemp_calculo_folha
''''''''''''    If adrFiltro_secundario("DFcod_evento") = intCodigo_Salario_Base Then
''''''''''''        adrFiltro_secundario("DFvalor") = Grava_Moeda(curSalario)
''''''''''''        adrFiltro_secundario("DFvalor_referencia") = intDias_Trabalhados
''''''''''''    End If
''''''''''''
''''''''''''    'atualizando o evento Auxilio Doença na TBtemp_calculo_folha
''''''''''''    If adrFiltro_secundario("DFcod_evento") = intCodigo_Auxilio_Doenca Or adrFiltro_secundario("DFcod_evento") = intCodigo_Auxilio_Acidente Then
''''''''''''        adrFiltro_secundario("DFvalor") = Grava_Moeda(curSal_afastamento)
''''''''''''        adrFiltro_secundario("DFvalor_referencia") = intDias_Afastado
''''''''''''    End If
''''''''''''
''''''''''''
''''''''''''    Exit Sub
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Calcular_Afastamentos")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''
''''''''''''Private Sub Grava_Recibo_Novo()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Grava_Recibo_Novo
'''''''''''''Objetivo:.........................: Grava um novo ID na tabela TBrecibo_pagamento
'''''''''''''Desenvolvimento...................: Wascley Costa
'''''''''''''Data de criação...................: 09/03/2001
'''''''''''''Data da última manutenção.........: 21/09/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''Comentado por tempo indeterminado 29/10/2001 José Braga
''''''
''''''''''''    Dim strSQL As String
''''''''''''    Dim adrRecibo_pagamento As ADODB.Recordset
''''''''''''
''''''''''''    'Esta função deleta um calculo anterior do mesmo funcionario e no mesmo periodo,
''''''''''''    'caso o usuário solicite calcular novamente.
''''''''''''    Call Deleta_Recibo_Antigo
''''''''''''
''''''''''''    'Criação da Sql para inserção de um novo recibo.
''''''''''''    strSQL = "INSERT INTO TBrecibo_pagamento " & _
''''''''''''             "([DFmatricula], " & _
''''''''''''             "[DFmes_ano_calculo], " & _
''''''''''''             "[DFsalario_base], " & _
''''''''''''             "[DFsalario_contribuicao_inss], " & _
''''''''''''             "[DFbase_calculo_fgts], " & _
''''''''''''             "[DFfgts_mes], " & _
''''''''''''             "[DFbase_calculo_irrf], " & _
''''''''''''             "[DFfaixa_irrf], " & _
''''''''''''             "[DFtotal_vencimento], " & _
''''''''''''             "[DFtotal_descontos]) " & _
''''''''''''             "SELECT " & adrFiltro_secundario("DFmatricula") & _
''''''''''''             ", '" & Format(datMes_ano, "yyyyMMdd") & "'" & _
''''''''''''             ", 0, 0, 0, 0, 0, 0 ,0 ,0"
''''''''''''
''''''''''''    'Executa a Sql criada acima.
''''''''''''    CNConexao.Execute strSQL
''''''''''''
''''''''''''    'Criação da ADR para guardar os Id's da tabela TBrecibo_pagamento.
''''''''''''    Call Banco_Dados.SQLgeral("SELECT DFid_recibo_pagamento FROM TBrecibo_pagamento ORDER BY DFid_recibo_pagamento", adrRecibo_pagamento)
''''''''''''    'Posiciona a ADR no seu último registro o qual foi gravado neste momento.
''''''''''''    adrRecibo_pagamento.MoveLast
''''''''''''
''''''''''''    'Atribui o Id do Recibo atual para uma variavel de memória.
''''''''''''    lonId_recibo = adrRecibo_pagamento("DFid_recibo_pagamento")
''''''''''''
''''''''''''    'Limpa e destrói a adrRecibo_pagamento.
''''''''''''    adrRecibo_pagamento.Close
''''''''''''    Set adrRecibo_pagamento = Nothing
''''''''''''
''''''''''''    Exit Sub
''''''''''''
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Grava_Recibo_Novo")
''''''''''''    Resume Next
''''''''''''End Sub
''''''
''''''''''''Public Function Verifica_Parcela_Evento() As Boolean
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Verifica_Parcela_Evento
'''''''''''''Objetivo:.........................: Se o evento tiver parcelas subtrai uma unidade.
'''''''''''''                                    Chamada por cmdConfirmar_click
'''''''''''''Desenvolvimento...................: Wascley Costa
'''''''''''''Data de criação...................: 14/03/2001
'''''''''''''Data da última manutenção.........: 21/09/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''    Dim strSQL As String
''''''''''''    Dim adrRecibo_exitente As ADODB.Recordset
''''''''''''
''''''''''''    'Se o atributo DFparcela for nulo então este evento será calculado.
''''''''''''    If IsNull(adrFiltro_secundario("DFparcela")) Then
''''''''''''        Verifica_Parcela_Evento = True
''''''''''''
''''''''''''    'Se o atributo DFparcela for zero então este evento não será calculado.
''''''''''''    ElseIf adrFiltro_secundario("DFparcela") = 0 Then
''''''''''''        Verifica_Parcela_Evento = False
''''''''''''
''''''''''''    'Se o atributo DFparcela tiver valor então este evento será calculado,
''''''''''''    'e será retirado do numero de parcelas uma unidade.
''''''''''''    Else
''''''''''''
''''''''''''        'Se o calculo já existir então não é necessário atualizar a TBfixo, pois
''''''''''''        'ela já deve ter sido atualizada anteriormente.
''''''''''''        If booRecibo_existente = True Then
''''''''''''            Verifica_Parcela_Evento = True
''''''''''''            Exit Function
''''''''''''
''''''''''''        End If
''''''''''''
''''''''''''        strSQL = "UPDATE TBfixo " & _
''''''''''''                 "SET DFparcela = " & adrFiltro_secundario("DFparcela") - 1 & " " & _
''''''''''''                 "WHERE DFid_fixo = " & adrFiltro_secundario("DFid_fixo")
''''''''''''        'Executa a Sql acima, a qual atualiza a tabela TBfixo.
''''''''''''        CNConexao.Execute strSQL
''''''''''''        Verifica_Parcela_Evento = True
''''''''''''
''''''''''''    End If
''''''''''''
''''''''''''    Exit Function
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Verifica_Parcela_Evento")
''''''''''''    Resume Next
''''''''''''End Function
''''''''''''
''''''''''''Private Sub Evento_Automatico_Adicional_Insalubridade()
'''''''''''''*******************************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Evento_Automatico_Adicional_Insalubridade
'''''''''''''Objetivo:.........................: Insere na Tabela temp o evento Adicional Insalubridade,
'''''''''''''                                    se ele não foi cadastrado como fixo ou variavel.
'''''''''''''Desenvolvimento...................: José Braga
'''''''''''''Data de criação...................: 01/10/2001
'''''''''''''Data da última manutenção.........: 01/10/2001
'''''''''''''Manutenção executada por..........: José Braga
'''''''''''''Observaçãoes......................:
'''''''''''''*******************************************************************************************
''''''''''''
''''''''''''    On Error GoTo Erro
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    Dim adrEvento As ADODB.Recordset
''''''''''''    Dim strSQL As String
''''''''''''    Dim strCodigo_Adicional As String
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    If adrFiltro_primario("DFperc_insalubridade") = 0 Then
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    'Busca o código do evento Adicional de Insalubridade na TBopcoes.
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 114")
''''''''''''    strCodigo_Adicional = adrOpcoes("DFvalor")
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    'SQL para verificar se o evento Adicional de Insalubridade já existe para este funcionário.
''''''''''''    strSQL = "SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha " & _
''''''''''''             "WHERE DFcod_evento = " & strCodigo_Adicional
''''''''''''    If CNConexao.Execute(strSQL)(0) = 0 Then
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''        'Busca na TBevento o registro referente ao Adicional de Insalubridade
''''''''''''        strSQL = "SELECT * FROM TBevento WHERE DFcod_evento = " & strCodigo_Adicional
'''''''''''''        Set adrEvento = New ADODB.Recordset
'''''''''''''        adrEvento.Open strSQL, CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''        Call Banco_Dados.SQLgeral(strSQL, adrEvento)
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''        Dim intValor_campo(5) As Integer
''''''''''''        Dim intIndice As Integer, intCont As Integer
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''        'Joga os valores dos campos do tipo boleano para uma variavel bidimencional,
''''''''''''        'transformando de True/False para 1/0.
''''''''''''        intIndice = 4
''''''''''''        For intCont = 0 To 5
''''''''''''            If adrEvento(intIndice) = True Then
''''''''''''                intValor_campo(intCont) = 1
''''''''''''            ElseIf adrEvento(intIndice) = False Then
''''''''''''                intValor_campo(intCont) = 0
''''''''''''            End If
''''''''''''            intIndice = intIndice + 1
''''''''''''        Next
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''        If CNConexao.Execute("SELECT COUNT(DFcod_evento) FROM TBevento WHERE DFcod_evento = " & strCodigo_Adicional)(0) <> 0 Then
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''            'SQL de inserção de registro na TB temp de eventos.
''''''''''''            strSQL = _
''''''''''''                "INSERT INTO TBtemp_calculo_folha (" & _
''''''''''''                    "[DFmatricula]," & _
''''''''''''                    "[DFid_fixo]," & _
''''''''''''                    "[DFcod_evento]," & _
''''''''''''                    "[DFparcela]," & _
''''''''''''                    "[DFpercentual_sob_salario_base]," & _
''''''''''''                    "[DFpercentual_sob_dias_trabalhados]," & _
''''''''''''                    "[DFvalor]," & _
''''''''''''                    "[DFdesconta_faltas]," & _
''''''''''''                    "[DFdescricao]," & _
''''''''''''                    "[DFtipo]," & _
''''''''''''                    "[DFreferencia]," & _
''''''''''''                    "[DFvalor_referencia]," & _
''''''''''''                    "[DFimprime_referencia]," & _
''''''''''''                    "[DFincide_inss]," & _
''''''''''''                    "[DFincide_fgts]," & _
''''''''''''                    "[DFincide_irrf]," & _
''''''''''''                    "[DFincide_rais]," & _
''''''''''''                    "[DFincide_informe_rendimentos]," & _
''''''''''''                    "[DFmultiplicador]) "
''''''''''''
''''''''''''            strSQL = strSQL & _
''''''''''''                "SELECT " & _
''''''''''''                    adrFiltro_primario("DFmatricula") & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    "," & strCodigo_Adicional & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    "," & "0" & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    ",'" & adrEvento("DFdescricao") & "'" & _
''''''''''''                    ",'" & adrEvento("DFtipo") & "'" & _
''''''''''''                    ",'" & adrEvento("DFreferencia") & "'" & _
''''''''''''                    "," & CInt(adrFiltro_primario("DFperc_insalubridade")) & _
''''''''''''                    "," & intValor_campo(0) & _
''''''''''''                    "," & intValor_campo(1) & _
''''''''''''                    "," & intValor_campo(2) & _
''''''''''''                    "," & intValor_campo(3) & _
''''''''''''                    "," & intValor_campo(4) & _
''''''''''''                    "," & intValor_campo(5) & _
''''''''''''                    "," & adrEvento("DFmultiplicador")
''''''''''''
''''''''''''            CNConexao.Execute strSQL
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''        End If
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    Else
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''        'Testa se existe na tabela dois eventos duplicados do Salario Base
''''''''''''        strSQL = "SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha WHERE DFcod_evento = " & strCodigo_Adicional
''''''''''''        If CNConexao.Execute(strSQL)(0) <> 1 Then
''''''''''''            strSQL = "DELETE FROM TBtemp_calculo_folha " & _
''''''''''''                     "WHERE DFcod_evento = " & strCodigo_Adicional & " " & _
''''''''''''                     "AND DFid_fixo <> 0"
''''''''''''            CNConexao.Execute strSQL
''''''''''''        End If
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    End If
''''''''''''
''''''''''''    Exit Sub
''''''''''''
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Evento_Automatico_Adicional_Insalubridade")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''Private Sub Evento_Automatico_Adicional_Periculosidade()
'''''''''''''*******************************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Evento_Automatico_Adicional_Insalubridade
'''''''''''''Objetivo:.........................: Insere na Tabela temp o evento Adicional Insalubridade,
'''''''''''''                                    se ele não foi cadastrado como fixo ou variavel.
'''''''''''''Desenvolvimento...................: José Braga
'''''''''''''Data de criação...................: 01/10/2001
'''''''''''''Data da última manutenção.........: 01/10/2001
'''''''''''''Manutenção executada por..........: José Braga
'''''''''''''Observaçãoes......................:
'''''''''''''*******************************************************************************************
''''''''''''
''''''''''''    On Error GoTo Erro
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    Dim adrEvento As ADODB.Recordset
''''''''''''    Dim strSQL As String
''''''''''''    Dim strCodigo_Adicional As String
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    If adrFiltro_primario("DFperc_periculosidade") = 0 Then
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    'Busca o código do evento Adicional de Insalubridade na TBopcoes.
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 128")
''''''''''''    strCodigo_Adicional = adrOpcoes("DFvalor")
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    'SQL para verificar se o evento Adicional de Insalubridade já existe para este funcionário.
''''''''''''    strSQL = "SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha " & _
''''''''''''             "WHERE DFcod_evento = " & strCodigo_Adicional
''''''''''''    If CNConexao.Execute(strSQL)(0) = 0 Then
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''        'Busca na TBevento o registro referente ao Adicional de Insalubridade
''''''''''''        strSQL = "SELECT * FROM TBevento WHERE DFcod_evento = " & strCodigo_Adicional
'''''''''''''        Set adrEvento = New ADODB.Recordset
'''''''''''''        adrEvento.Open strSQL, CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''        Call Banco_Dados.SQLgeral(strSQL, adrEvento)
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''        Dim intValor_campo(5) As Integer
''''''''''''        Dim intIndice As Integer, intCont As Integer
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''        'Joga os valores dos campos do tipo boleano para uma variavel bidimencional,
''''''''''''        'transformando de True/False para 1/0.
''''''''''''        intIndice = 4
''''''''''''        For intCont = 0 To 5
''''''''''''            If adrEvento(intIndice) = True Then
''''''''''''                intValor_campo(intCont) = 1
''''''''''''            ElseIf adrEvento(intIndice) = False Then
''''''''''''                intValor_campo(intCont) = 0
''''''''''''            End If
''''''''''''            intIndice = intIndice + 1
''''''''''''        Next
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''        If CNConexao.Execute("SELECT COUNT(DFcod_evento) FROM TBevento WHERE DFcod_evento = " & strCodigo_Adicional)(0) <> 0 Then
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''            'SQL de inserção de registro na TB temp de eventos.
''''''''''''            strSQL = _
''''''''''''                "INSERT INTO TBtemp_calculo_folha (" & _
''''''''''''                    "[DFmatricula]," & _
''''''''''''                    "[DFid_fixo]," & _
''''''''''''                    "[DFcod_evento]," & _
''''''''''''                    "[DFparcela]," & _
''''''''''''                    "[DFpercentual_sob_salario_base]," & _
''''''''''''                    "[DFpercentual_sob_dias_trabalhados]," & _
''''''''''''                    "[DFvalor]," & _
''''''''''''                    "[DFdesconta_faltas]," & _
''''''''''''                    "[DFdescricao]," & _
''''''''''''                    "[DFtipo]," & _
''''''''''''                    "[DFreferencia]," & _
''''''''''''                    "[DFvalor_referencia]," & _
''''''''''''                    "[DFimprime_referencia]," & _
''''''''''''                    "[DFincide_inss]," & _
''''''''''''                    "[DFincide_fgts]," & _
''''''''''''                    "[DFincide_irrf]," & _
''''''''''''                    "[DFincide_rais]," & _
''''''''''''                    "[DFincide_informe_rendimentos]," & _
''''''''''''                    "[DFmultiplicador]) "
''''''''''''
''''''''''''            strSQL = strSQL & _
''''''''''''                "SELECT " & _
''''''''''''                    adrFiltro_primario("DFmatricula") & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    "," & strCodigo_Adicional & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    "," & "0" & _
''''''''''''                    "," & "NULL" & _
''''''''''''                    ",'" & adrEvento("DFdescricao") & "'" & _
''''''''''''                    ",'" & adrEvento("DFtipo") & "'" & _
''''''''''''                    ",'" & adrEvento("DFreferencia") & "'" & _
''''''''''''                    "," & CInt(adrFiltro_primario("DFperc_periculosidade")) & _
''''''''''''                    "," & intValor_campo(0) & _
''''''''''''                    "," & intValor_campo(1) & _
''''''''''''                    "," & intValor_campo(2) & _
''''''''''''                    "," & intValor_campo(3) & _
''''''''''''                    "," & intValor_campo(4) & _
''''''''''''                    "," & intValor_campo(5) & _
''''''''''''                    "," & adrEvento("DFmultiplicador")
''''''''''''
''''''''''''            CNConexao.Execute strSQL
''''''''''''        '-----------------------------------------------------------------------------------
''''''''''''        End If
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''    Else
''''''''''''    '---------------------------------------------------------------------------------------
''''''''''''        'Testa se existe na tabela dois eventos duplicados do Salario Base
''''''''''''        strSQL = "SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha WHERE DFcod_evento = " & strCodigo_Adicional
''''''''''''        If CNConexao.Execute(strSQL)(0) <> 1 Then
''''''''''''            strSQL = "DELETE FROM TBtemp_calculo_folha " & _
''''''''''''                     "WHERE DFcod_evento = " & strCodigo_Adicional & " " & _
''''''''''''                     "AND DFid_fixo <> 0"
''''''''''''            CNConexao.Execute strSQL
''''''''''''        End If
''''''''''''        '----------------------------------------------------------------------------------
''''''''''''    End If
''''''''''''
''''''''''''    Exit Sub
''''''''''''
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Evento_Automatico_Adicional_Periculosidade")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''Private Sub Evento_Automatico_Auxilio_Acidente()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Evento_Automatico_Auxilio_Acidente
'''''''''''''Objetivo:.........................: Insere na Tabela temp o evento Auxilio Acidente,
'''''''''''''                                    se ele não foi cadastrado como fixo ou variavel.
'''''''''''''Desenvolvimento...................: Vitor Constâncio da Silva
'''''''''''''Data de criação...................: 21/09/2001
'''''''''''''Data da última manutenção.........: 21/09/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''
''''''''''''    Dim strSQL As String
''''''''''''    Dim intCont As Integer
''''''''''''    Dim intIndice As Integer
''''''''''''    Dim intValor_campo(5) As Integer
''''''''''''    Dim adrValor_salario As ADODB.Recordset
''''''''''''    Dim adrCampos_evento As ADODB.Recordset
''''''''''''    Dim datInicio_licenca As Date
''''''''''''    Dim datFinal_licenca As Date
''''''''''''    Dim intDias_Referentes As Integer
''''''''''''    Dim strCodigo_Retorno_Acidente As String
''''''''''''
''''''''''''
''''''''''''    'Busca o código do evento Salário Base na TBopcoes.
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 183")
''''''''''''    strCodigo_Auxilio_Acidente = adrOpcoes("DFvalor")
''''''''''''
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 190")
''''''''''''    strCodigo_Retorno_Acidente = adrOpcoes("DFvalor")
''''''''''''
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 184")
''''''''''''    intCodigo_Auxilio_Acidente = adrOpcoes("DFvalor")
''''''''''''
''''''''''''    strSQL = "SELECT TBafastamento_funcionario.DFid_afastamento_funcionario, TBafastamento_Funcionario.DFcod_movimentacao, TBafastamento_funcionario.DFdata_inicio, TBafastamento.DFtipo " & _
''''''''''''             "FROM TBfuncionario " & _
''''''''''''                "INNER JOIN TBafastamento_funcionario " & _
''''''''''''                     "ON TBfuncionario.DFmatricula = TBafastamento_funcionario.DFmatricula " & _
''''''''''''                "INNER JOIN TBafastamento " & _
''''''''''''                     "ON TBafastamento_funcionario.DFcod_movimentacao = TBafastamento.DFcod_afastamento " & _
''''''''''''             "WHERE (CONVERT(CHAR(4),YEAR(DFdata_inicio)))+ " & _
''''''''''''                 "(CONVERT(CHAR(2),MONTH(DFdata_inicio))) = " & _
''''''''''''                 "'" & Format(dtpMes_ano, "yyyyM") & "' " & _
''''''''''''                 "AND TBafastamento_funcionario.DFmatricula = " & adrFiltro_primario("DFmatricula") & " " & _
''''''''''''                 "AND (TBafastamento_funcionario.DFcod_movimentacao = '" & strCodigo_Auxilio_Acidente & "' " & _
''''''''''''                 "OR TBafastamento_funcionario.DFcod_movimentacao = '" & strCodigo_Retorno_Acidente & "') "
''''''''''''
''''''''''''    If CNConexao.Execute(strSQL).EOF Then
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''
''''''''''''    datInicio_licenca = CNConexao.Execute("SELECT DFdata_inicio FROM TBafastamento_funcionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")).Fields("DFdata_inicio")
''''''''''''    datFinal_licenca = CNConexao.Execute("SELECT DFdata_final FROM TBafastamento_funcionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")).Fields("DFdata_final")
''''''''''''
''''''''''''    intDias_Referentes = DateDiff("d", datInicio_licenca, datFinal_licenca)
''''''''''''    If intDias_Referentes > 15 Then
''''''''''''        intDias_Referentes = 15
''''''''''''    End If
''''''''''''    If Month(datInicio_licenca) > Month(dtpMes_ano) And Year(datInicio_licenca) = Year(dtpMes_ano) Then
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''
''''''''''''    'SQL para verificar se o evento Auxilio Acidente já existe para este funcionário.
''''''''''''    strSQL = "SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha " & _
''''''''''''             "WHERE DFcod_evento = " & adrOpcoes("DFvalor")
''''''''''''    If CNConexao.Execute(strSQL)(0) = 0 Then
''''''''''''        Dim curAuxilio_Acidente As Currency
''''''''''''        'Busca dados de outras tabelas que serão necessários.
'''''''''''''        Set adrValor_salario = New ADODB.Recordset
'''''''''''''        adrValor_salario.Open "SELECT DFsalario From TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula"), CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''        Call Banco_Dados.SQLgeral("SELECT DFsalario From TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula"), adrValor_salario)
''''''''''''
''''''''''''        curAuxilio_Acidente = (adrValor_salario("DFsalario") / 30) * intDias_Referentes
''''''''''''
'''''''''''''        Set adrCampos_evento = New ADODB.Recordset
'''''''''''''        adrCampos_evento.Open "SELECT * FROM TBevento WHERE DFcod_evento = " & adrOpcoes("DFvalor"), CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''        Call Banco_Dados.SQLgeral("SELECT * FROM TBevento WHERE DFcod_evento = " & adrOpcoes("DFvalor"), adrCampos_evento)
''''''''''''
''''''''''''        'Joga os valores dos campos do tipo boleano para uma variavel bidimencional,
''''''''''''        'transformando de True/False para 1/0.
''''''''''''        intIndice = 4
''''''''''''        For intCont = 0 To 5
''''''''''''            If adrCampos_evento(intIndice) = True Then
''''''''''''                intValor_campo(intCont) = 1
''''''''''''            ElseIf adrCampos_evento(intIndice) = False Then
''''''''''''                intValor_campo(intCont) = 0
''''''''''''            End If
''''''''''''            intIndice = intIndice + 1
''''''''''''        Next
''''''''''''
''''''''''''        'SQL de inserção de registro na TB temp de eventos.
''''''''''''        strSQL = ""
''''''''''''        strSQL = strSQL & "INSERT INTO TBtemp_calculo_folha "
''''''''''''        strSQL = strSQL & "([DFmatricula],[DFid_fixo],"
''''''''''''        strSQL = strSQL & "[DFcod_evento],[DFparcela],"
''''''''''''        strSQL = strSQL & "[DFpercentual_sob_salario_base],"
''''''''''''        strSQL = strSQL & "[DFpercentual_sob_dias_trabalhados],"
''''''''''''        strSQL = strSQL & "[DFvalor],"
''''''''''''        strSQL = strSQL & "[DFdesconta_faltas],[DFdescricao],"
''''''''''''        strSQL = strSQL & "[DFtipo],[DFreferencia],"
''''''''''''        strSQL = strSQL & "[DFimprime_referencia],[DFincide_inss],"
''''''''''''        strSQL = strSQL & "[DFincide_fgts],[DFincide_irrf],"
''''''''''''        strSQL = strSQL & "[DFincide_rais],[DFincide_informe_rendimentos],"
''''''''''''        strSQL = strSQL & "[DFmultiplicador]) "
''''''''''''
''''''''''''        strSQL = strSQL & "SELECT " & adrFiltro_primario("DFmatricula") & ", "
''''''''''''        strSQL = strSQL & "NULL," & adrOpcoes("DFvalor") & ", NULL, NULL, NULL, "
''''''''''''        strSQL = strSQL & Grava_Moeda(curAuxilio_Acidente) & ", "
''''''''''''        strSQL = strSQL & "NULL, '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFdescricao") & "', '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFtipo") & "', '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFreferencia") & "', "
''''''''''''        strSQL = strSQL & intValor_campo(0) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(1) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(2) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(3) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(4) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(5) & ", "
''''''''''''        strSQL = strSQL & adrCampos_evento("DFmultiplicador")
''''''''''''
''''''''''''        'Executa a SQL acima construida.
''''''''''''        CNConexao.Execute strSQL
''''''''''''    Else
''''''''''''        'Testa se existe na tabela dois eventos duplicados do Salario Base
''''''''''''        If CNConexao.Execute("SELECT COUNT(*) FROM TBtemp_calculo_folha WHERE DFcod_evento = " & strCodigo_Auxilio_Acidente)(0) <> 1 Then
''''''''''''            CNConexao.Execute ("DELETE FROM TBtemp_calculo_folha WHERE DFcod_evento = " & strCodigo_Auxilio_Acidente & " AND DFid_fixo <> 0")
''''''''''''        End If
''''''''''''    End If
''''''''''''
''''''''''''    Dim datInicio_Afastamento As Date
''''''''''''    Dim datRetorno As Date
''''''''''''    If Not CNConexao.Execute("SELECT DFcod_afastamento FROM TBfuncionario WHERE DFcod_afastamento IS NULL AND DFmatricula = " & adrFiltro_primario("DFmatricula")).EOF Then
''''''''''''        datInicio_Afastamento = CNConexao.Execute("SELECT DFdata_inicio FROM TBafastamento_funcionario WHERE DFid_retorno IS NOT NULL").Fields("DFdata_inicio")
''''''''''''        datRetorno = CNConexao.Execute("SELECT DFdata_inicio FROM TBafastamento_funcionario WHERE DFid_afastamento IS NOT NULL").Fields("DFdata_inicio")
''''''''''''        If Month(datInicio_Afastamento) <> Month(datRetorno) Then
''''''''''''            If Day(datInicio_Afastamento) + 15 <= 30 Then
''''''''''''                If Month(datInicio_Afastamento) <> Month(dtpMes_ano) Then
''''''''''''                    CNConexao.Execute ("DELETE FROM TBtemp_calculo_folha WHERE DFcod_evento = " & intCodigo_Auxilio_Acidente)
''''''''''''                End If
''''''''''''            End If
''''''''''''        End If
''''''''''''    End If
''''''''''''
''''''''''''    Exit Sub
''''''''''''
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Evento_Automatico_Auxilio_Acidente")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''
''''''''''''Private Sub Evento_Automatico_Vale_Transporte()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Evento_Automatico_Vale_Transporte
'''''''''''''Objetivo:.........................: Insere na Tabela temp o evento Vale-Transporte,
'''''''''''''                                    se ele não foi cadastrado como fixo ou variavel.
'''''''''''''Desenvolvimento...................: Vitor
'''''''''''''Data de criação...................: 18/09/2001
'''''''''''''Data da última manutenção.........: 21/09/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''    Dim strSQL As String
''''''''''''    Dim intCont As Integer
''''''''''''    Dim intIndice As Integer
''''''''''''    Dim intValor_campo(5) As Integer
''''''''''''    Dim sinPercentual As Single
''''''''''''    Dim curValor_Desconto As Currency
''''''''''''    Dim adrValor_salario As ADODB.Recordset
''''''''''''    Dim adrCampos_evento As ADODB.Recordset
''''''''''''    Dim strCodigo_Vale_Transporte As String
''''''''''''    Dim strOpcao_Vale As String
''''''''''''    Dim adrTipo_Vale As ADODB.Recordset
''''''''''''    Dim adrVale_Transporte As ADODB.Recordset
''''''''''''    Dim intReferencia_Vale As Integer
''''''''''''    Dim strDescricao_evento As String
''''''''''''
''''''''''''
''''''''''''    'Verifica se o funcionário é beneficiado com o vale_transporte.
''''''''''''    If CNConexao.Execute("SELECT DFvale_transporte FROM TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")).Fields("DFvale_transporte") = 0 Then
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''
''''''''''''    'Busca o código do evento Vale-Transporte na TBopcoes.
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 166")
''''''''''''    strCodigo_Vale_Transporte = adrOpcoes("DFvalor")
''''''''''''
'''''''''''''    Set adrCampos_evento = New ADODB.Recordset
'''''''''''''    adrCampos_evento.Open "SELECT * FROM TBevento WHERE DFcod_evento = " & strCodigo_Vale_Transporte, CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''    Call Banco_Dados.SQLgeral("SELECT * FROM TBevento WHERE DFcod_evento = " & strCodigo_Vale_Transporte, adrCampos_evento)
''''''''''''
''''''''''''    'transformando de True/False para 1/0.
''''''''''''    intIndice = 5
''''''''''''    For intCont = 0 To 5
''''''''''''        If adrCampos_evento(intIndice) = True Then
''''''''''''            intValor_campo(intCont) = 1
''''''''''''        ElseIf adrCampos_evento(intIndice) = False Then
''''''''''''            intValor_campo(intCont) = 0
''''''''''''        End If
''''''''''''        intIndice = intIndice + 1
''''''''''''    Next
''''''''''''
''''''''''''    'SQL para verificar se o evento Vale-Transporte já existe para este funcionário.
''''''''''''    strSQL = "SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha " & _
''''''''''''             "WHERE DFcod_evento = " & strCodigo_Vale_Transporte
''''''''''''    If CNConexao.Execute(strSQL)(0) = 0 Then
''''''''''''        strOpcao_Vale = CNConexao.Execute("SELECT DFmodo_calc_vt FROM TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")).Fields("DFmodo_calc_vt")
''''''''''''        'Verifica por qual tipo de calculo do desconto de Vale-Transporte o funcionário optou.
''''''''''''        If strOpcao_Vale = "P" Then
''''''''''''            'Busca dados de outras tabelas que serão necessários.
''''''''''''            adrOpcoes.MoveFirst
''''''''''''            adrOpcoes.Find ("DFcodigo = 108")
''''''''''''
'''''''''''''            Set adrValor_salario = New ADODB.Recordset
'''''''''''''            adrValor_salario.Open "SELECT DFvalor From TBtemp_calculo_folha WHERE DFcod_evento = " & adrOpcoes("DFvalor")
''''''''''''            Call Banco_Dados.SQLgeral("SELECT DFvalor From TBtemp_calculo_folha WHERE DFcod_evento = " & adrOpcoes("DFvalor"), adrValor_salario)
''''''''''''
''''''''''''            strSQL = "SELECT TBempresa.DFpercentual_vt FROM TBempresa "
''''''''''''            strSQL = strSQL & "INNER JOIN TBfuncionario "
''''''''''''            strSQL = strSQL & "ON TBempresa.DFcod_empresa = TBfuncionario.DFcod_empresa "
''''''''''''            strSQL = strSQL & "WHERE TBfuncionario.DFmatricula = " & adrFiltro_primario("DFmatricula")
''''''''''''            sinPercentual = CNConexao.Execute(strSQL).Fields("DFpercentual_vt")
''''''''''''
''''''''''''            curValor_Desconto = Format(adrValor_salario("DFvalor") * (sinPercentual / 100), "Currency")
''''''''''''            intReferencia_Vale = sinPercentual
''''''''''''            strDescricao_evento = adrCampos_evento("DFdescricao")
''''''''''''
''''''''''''            strSQL = "INSERT INTO TBtemp_calculo_folha " & _
''''''''''''                            "([DFmatricula],[DFid_fixo]," & _
''''''''''''                            "[DFcod_evento],[DFparcela]," & _
''''''''''''                            "[DFpercentual_sob_salario_base]," & _
''''''''''''                            "[DFpercentual_sob_dias_trabalhados]," & _
''''''''''''                            "[DFvalor],[DFdesconta_faltas],[DFdescricao]," & _
''''''''''''                            "[DFtipo],[DFreferencia],[DFvalor_referencia]," & _
''''''''''''                            "[DFimprime_referencia],[DFincide_inss]," & _
''''''''''''                            "[DFincide_fgts],[DFincide_irrf]," & _
''''''''''''                            "[DFincide_rais],[DFincide_informe_rendimentos]," & _
''''''''''''                            "[DFmultiplicador]) " & _
''''''''''''                     "SELECT " & adrFiltro_primario("DFmatricula") & ", " & _
''''''''''''                            "NULL," & strCodigo_Vale_Transporte & ", NULL, NULL, NULL, " & _
''''''''''''                            Grava_Moeda(curValor_Desconto) & ", NULL, '" & _
''''''''''''                            strDescricao_evento & "', '" & _
''''''''''''                            adrCampos_evento("DFtipo") & "', '" & _
''''''''''''                            adrCampos_evento("DFreferencia") & "', '" & _
''''''''''''                            intReferencia_Vale & "', " & _
''''''''''''                            intValor_campo(0) & ", " & _
''''''''''''                            intValor_campo(1) & ", " & _
''''''''''''                            intValor_campo(2) & ", " & _
''''''''''''                            intValor_campo(3) & ", " & _
''''''''''''                            intValor_campo(4) & ", " & _
''''''''''''                            intValor_campo(5) & ", " & _
''''''''''''                            adrCampos_evento("DFmultiplicador")
''''''''''''            CNConexao.Execute strSQL
''''''''''''        ElseIf strOpcao_Vale = "V" Then
''''''''''''
''''''''''''            'Calculo do Valor do desconto de Vale-Transporte com base no Valor dos Vales.
''''''''''''            'O código abaixo inclusive a estrutura repetitiva são utilizados para calcular
''''''''''''            'a referencia do evento, e caso o funcionario optar pelo calculo de desconto de vale
''''''''''''            'a partir do valor dos vales esta operação também é executada abaixo.
''''''''''''
''''''''''''            strSQL = "SELECT * FROM TBtipo_vale_transporte_funcionario " & _
''''''''''''                     "WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")
'''''''''''''            Set adrTipo_Vale = New ADODB.Recordset
'''''''''''''            adrTipo_Vale.Open strSQL, CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''            Call Banco_Dados.SQLgeral(strSQL, adrTipo_Vale)
''''''''''''
''''''''''''            strSQL = "SELECT * FROM TBtemp_vale_transporte " & _
''''''''''''                     "WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")
'''''''''''''            Set adrVale_Transporte = New ADODB.Recordset
'''''''''''''            adrVale_Transporte.Open strSQL, CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''            Call Banco_Dados.SQLgeral(strSQL, adrVale_Transporte)
''''''''''''
''''''''''''            Do While Not adrTipo_Vale.EOF
''''''''''''
''''''''''''                curValor_Desconto = adrVale_Transporte("DFvalor_mes")
''''''''''''                intReferencia_Vale = adrVale_Transporte("DFqtde_mes")
''''''''''''                strDescricao_evento = adrVale_Transporte("DFtipo_vale")
''''''''''''
'''''''''''''==================//==================//===================//===================//=========
''''''''''''
''''''''''''                'SQL de inserção de registro na TBtemp de eventos.
''''''''''''                strSQL = "INSERT INTO TBtemp_calculo_folha " & _
''''''''''''                           "([DFmatricula],[DFid_fixo]," & _
''''''''''''                           "[DFcod_evento],[DFparcela]," & _
''''''''''''                           "[DFpercentual_sob_salario_base]," & _
''''''''''''                           "[DFpercentual_sob_dias_trabalhados]," & _
''''''''''''                           "[DFvalor],[DFdesconta_faltas],[DFdescricao]," & _
''''''''''''                           "[DFtipo],[DFreferencia],[DFvalor_referencia]," & _
''''''''''''                           "[DFimprime_referencia],[DFincide_inss]," & _
''''''''''''                           "[DFincide_fgts],[DFincide_irrf]," & _
''''''''''''                           "[DFincide_rais],[DFincide_informe_rendimentos]," & _
''''''''''''                           "[DFmultiplicador]) " & _
''''''''''''                "SELECT " & adrFiltro_primario("DFmatricula") & ", " & _
''''''''''''                           "NULL," & strCodigo_Vale_Transporte & ", NULL, NULL, NULL, " & _
''''''''''''                           Grava_Moeda(curValor_Desconto) & ", NULL, '" & _
''''''''''''                           strDescricao_evento & "', '" & _
''''''''''''                           adrCampos_evento("DFtipo") & "', '" & _
''''''''''''                           adrCampos_evento("DFreferencia") & "', '" & _
''''''''''''                           intReferencia_Vale & "', " & _
''''''''''''                           intValor_campo(0) & ", " & _
''''''''''''                           intValor_campo(1) & ", " & _
''''''''''''                           intValor_campo(2) & ", " & _
''''''''''''                           intValor_campo(3) & ", " & _
''''''''''''                           intValor_campo(4) & ", " & _
''''''''''''                           intValor_campo(5) & ", " & _
''''''''''''                           adrCampos_evento("DFmultiplicador")
''''''''''''                CNConexao.Execute strSQL
''''''''''''                adrTipo_Vale.MoveNext
''''''''''''                adrVale_Transporte.MoveNext
''''''''''''            Loop
''''''''''''        End If
''''''''''''    End If
''''''''''''
''''''''''''    Exit Sub
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Evento_Automatico_Vale_Transporte")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''Private Sub Evento_Automatico_Auxilio_Doenca()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Evento_Automatico_Auxilio_Doenca
'''''''''''''Objetivo:.........................: Insere na Tabela temp o evento Auxilio Doenca,
'''''''''''''                                    se ele não foi cadastrado como fixo ou variavel.
'''''''''''''Desenvolvimento...................: Vitor Constâncio da Silva
'''''''''''''Data de criação...................: 21/09/2001
'''''''''''''Data da última manutenção.........: 21/09/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''
''''''''''''    Dim strSQL As String
''''''''''''    Dim intCont As Integer
''''''''''''    Dim intIndice As Integer
''''''''''''    Dim intValor_campo(5) As Integer
''''''''''''    Dim adrValor_salario As ADODB.Recordset
''''''''''''    Dim adrCampos_evento As ADODB.Recordset
''''''''''''    Dim datInicio_licenca As Date
''''''''''''
''''''''''''    'Busca o código do evento Salário Base na TBopcoes.
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 180")
''''''''''''    strCodigo_Auxilio_Doenca = adrOpcoes("DFvalor")
''''''''''''
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 181")
''''''''''''    intCodigo_Auxilio_Doenca = adrOpcoes("DFvalor")
''''''''''''
''''''''''''    strSQL = "SELECT TBafastamento_funcionario.DFid_afastamento_funcionario, TBafastamento_Funcionario.DFcod_movimentacao, TBafastamento_funcionario.DFdata_inicio, TBafastamento.DFtipo " & _
''''''''''''             "FROM TBfuncionario " & _
''''''''''''                "INNER JOIN TBafastamento_funcionario " & _
''''''''''''                     "ON TBfuncionario.DFmatricula = TBafastamento_funcionario.DFmatricula " & _
''''''''''''                "INNER JOIN TBafastamento " & _
''''''''''''                     "ON TBafastamento_funcionario.DFcod_movimentacao = TBafastamento.DFcod_afastamento " & _
''''''''''''             "WHERE (CONVERT(CHAR(4),YEAR(DFdata_inicio)))+ " & _
''''''''''''                 "(CONVERT(CHAR(2),MONTH(DFdata_inicio))) = " & _
''''''''''''                 "'" & Format(dtpMes_ano, "yyyyM") & "' " & _
''''''''''''                 "AND TBafastamento_funcionario.DFmatricula = " & adrFiltro_primario("DFmatricula") & " " & _
''''''''''''                 "AND TBafastamento_funcionario.DFcod_movimentacao = '" & strCodigo_Auxilio_Doenca & "' "
''''''''''''
''''''''''''    If CNConexao.Execute(strSQL).EOF Then
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''
''''''''''''    datInicio_licenca = CNConexao.Execute("SELECT DFdata_inicio FROM TBafastamento_funcionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")).Fields("DFdata_inicio")
''''''''''''    If Month(datInicio_licenca) > Month(dtpMes_ano) And Year(datInicio_licenca) = Year(dtpMes_ano) Then
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''
''''''''''''    'SQL para verificar se o evento Auxilio Doença já existe para este funcionário.
''''''''''''    strSQL = "SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha " & _
''''''''''''             "WHERE DFcod_evento = " & adrOpcoes("DFvalor")
''''''''''''    If CNConexao.Execute(strSQL)(0) = 0 Then
''''''''''''        Dim curAuxilio_Doenca As Currency
''''''''''''        'Busca dados de outras tabelas que serão necessários.
'''''''''''''        Set adrValor_salario = New ADODB.Recordset
'''''''''''''        adrValor_salario.Open "SELECT DFsalario From TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula"), CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''        Call Banco_Dados.SQLgeral("SELECT DFsalario From TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula"), adrValor_salario)
''''''''''''
''''''''''''        curAuxilio_Doenca = (adrValor_salario("DFsalario") / 30) * 15
''''''''''''
'''''''''''''        Set adrCampos_evento = New ADODB.Recordset
'''''''''''''        adrCampos_evento.Open "SELECT * FROM TBevento WHERE DFcod_evento = " & adrOpcoes("DFvalor"), CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''        Call Banco_Dados.SQLgeral("SELECT * FROM TBevento WHERE DFcod_evento = " & adrOpcoes("DFvalor"), adrCampos_evento)
''''''''''''
''''''''''''        'Joga os valores dos campos do tipo boleano para uma variavel bidimencional,
''''''''''''        'transformando de True/False para 1/0.
''''''''''''        intIndice = 4
''''''''''''        For intCont = 0 To 5
''''''''''''            If adrCampos_evento(intIndice) = True Then
''''''''''''                intValor_campo(intCont) = 1
''''''''''''            ElseIf adrCampos_evento(intIndice) = False Then
''''''''''''                intValor_campo(intCont) = 0
''''''''''''            End If
''''''''''''            intIndice = intIndice + 1
''''''''''''        Next
''''''''''''
''''''''''''        'SQL de inserção de registro na TB temp de eventos.
''''''''''''        strSQL = ""
''''''''''''        strSQL = strSQL & "INSERT INTO TBtemp_calculo_folha "
''''''''''''        strSQL = strSQL & "([DFmatricula],[DFid_fixo],"
''''''''''''        strSQL = strSQL & "[DFcod_evento],[DFparcela],"
''''''''''''        strSQL = strSQL & "[DFpercentual_sob_salario_base],"
''''''''''''        strSQL = strSQL & "[DFpercentual_sob_dias_trabalhados],"
''''''''''''        strSQL = strSQL & "[DFvalor],"
''''''''''''        strSQL = strSQL & "[DFdesconta_faltas],[DFdescricao],"
''''''''''''        strSQL = strSQL & "[DFtipo],[DFreferencia],"
''''''''''''        strSQL = strSQL & "[DFimprime_referencia],[DFincide_inss],"
''''''''''''        strSQL = strSQL & "[DFincide_fgts],[DFincide_irrf],"
''''''''''''        strSQL = strSQL & "[DFincide_rais],[DFincide_informe_rendimentos],"
''''''''''''        strSQL = strSQL & "[DFmultiplicador]) "
''''''''''''
''''''''''''        strSQL = strSQL & "SELECT " & adrFiltro_primario("DFmatricula") & ", "
''''''''''''        strSQL = strSQL & "NULL," & adrOpcoes("DFvalor") & ", NULL, NULL, NULL, "
''''''''''''        strSQL = strSQL & Grava_Moeda(curAuxilio_Doenca) & ", "
''''''''''''        strSQL = strSQL & "NULL, '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFdescricao") & "', '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFtipo") & "', '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFreferencia") & "', "
''''''''''''        strSQL = strSQL & intValor_campo(0) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(1) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(2) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(3) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(4) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(5) & ", "
''''''''''''        strSQL = strSQL & adrCampos_evento("DFmultiplicador")
''''''''''''
''''''''''''        'Executa a SQL acima construida.
''''''''''''        CNConexao.Execute strSQL
''''''''''''    Else
''''''''''''        'Testa se existe na tabela dois eventos duplicados do Salario Base
''''''''''''        If CNConexao.Execute("SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha WHERE DFcod_evento = " & intCodigo_Auxilio_Doenca)(0) <> 1 Then
''''''''''''            CNConexao.Execute ("DELETE FROM TBtemp_calculo_folha WHERE DFcod_evento = " & strCodigo_Auxilio_Doenca & " AND DFid_fixo <> 0")
''''''''''''        End If
''''''''''''    End If
''''''''''''
''''''''''''    Exit Sub
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Evento_Automatico_Auxilio_Doenca")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''Private Sub Evento_Automatico_Salario_Base()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Evento_Automatico_Salario_Base
'''''''''''''Objetivo:.........................: Insere na Tabela temp o evento Salario Base,
'''''''''''''                                    se ele não foi cadastrado como fixo ou variavel.
'''''''''''''                                    Chamada por Calculo_Eventos_Automaticos
'''''''''''''Desenvolvimento...................: Wascley Costa
'''''''''''''Data de criação...................: 23/04/2001
'''''''''''''Data da última manutenção.........: 21/09/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''
''''''''''''    Dim strSQL As String
''''''''''''    Dim intCont As Integer
''''''''''''    Dim intIndice As Integer
''''''''''''    Dim intValor_campo(5) As Integer
''''''''''''
''''''''''''    'Criação de ADRs para obtenção de dados que serão necessários para se inserir uma nova
''''''''''''    'linha de registro na lista de eventos.
''''''''''''    Dim adrValor_salario As ADODB.Recordset
''''''''''''    Dim adrCampos_evento As ADODB.Recordset
''''''''''''
''''''''''''    'Busca o código do evento Salário Base na TBopcoes.
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 108")
''''''''''''    intCodigo_Salario_Base = adrOpcoes("DFvalor")
''''''''''''
''''''''''''    'SQL para verificar se o evento Salario Base já existe para este funcionário.
''''''''''''    strSQL = "SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha " & _
''''''''''''             "WHERE DFcod_evento = " & intCodigo_Salario_Base
''''''''''''
''''''''''''    If CNConexao.Execute(strSQL)(0) = 0 Then
''''''''''''        'Busca dados de outras tabelas que serão necessários.
'''''''''''''        Set adrValor_salario = New ADODB.Recordset
'''''''''''''        adrValor_salario.Open "SELECT DFsalario From TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula"), CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''        Call Banco_Dados.SQLgeral("SELECT DFsalario From TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula"), adrValor_salario)
'''''''''''''        Set adrCampos_evento = New ADODB.Recordset
'''''''''''''        adrCampos_evento.Open "SELECT * FROM TBevento WHERE DFcod_evento = " & intCodigo_Salario_Base, CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''        Call Banco_Dados.SQLgeral("SELECT * FROM TBevento WHERE DFcod_evento = " & intCodigo_Salario_Base, adrCampos_evento)
''''''''''''
''''''''''''        'Joga os valores dos campos do tipo boleano para uma variavel bidimensional,
''''''''''''        'transformando de True/False para 1/0.
''''''''''''        intIndice = 4
''''''''''''        For intCont = 0 To 5
''''''''''''            If adrCampos_evento(intIndice) = True Then
''''''''''''                intValor_campo(intCont) = 1
''''''''''''            ElseIf adrCampos_evento(intIndice) = False Then
''''''''''''                intValor_campo(intCont) = 0
''''''''''''            End If
''''''''''''            intIndice = intIndice + 1
''''''''''''        Next
''''''''''''
''''''''''''        'SQL de inserção de registro na TB temp de eventos.
''''''''''''        strSQL = ""
''''''''''''        strSQL = strSQL & "INSERT INTO TBtemp_calculo_folha "
''''''''''''        strSQL = strSQL & "([DFmatricula],[DFid_fixo],"
''''''''''''        strSQL = strSQL & "[DFcod_evento],[DFparcela],"
''''''''''''        strSQL = strSQL & "[DFpercentual_sob_salario_base],"
''''''''''''        strSQL = strSQL & "[DFpercentual_sob_dias_trabalhados],"
''''''''''''        strSQL = strSQL & "[DFvalor],"
''''''''''''        strSQL = strSQL & "[DFdesconta_faltas],[DFdescricao],"
''''''''''''        strSQL = strSQL & "[DFtipo],[DFreferencia],"
''''''''''''        strSQL = strSQL & "[DFimprime_referencia],[DFincide_inss],"
''''''''''''        strSQL = strSQL & "[DFincide_fgts],[DFincide_irrf],"
''''''''''''        strSQL = strSQL & "[DFincide_rais],[DFincide_informe_rendimentos],"
''''''''''''        strSQL = strSQL & "[DFmultiplicador]) "
''''''''''''
''''''''''''        strSQL = strSQL & "SELECT " & adrFiltro_primario("DFmatricula") & ", "
''''''''''''        strSQL = strSQL & "NULL," & intCodigo_Salario_Base & ", NULL, NULL, NULL, "
''''''''''''        strSQL = strSQL & Grava_Moeda(adrValor_salario("DFsalario")) & ", "
''''''''''''        strSQL = strSQL & "NULL, '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFdescricao") & "', '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFtipo") & "', '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFreferencia") & "', "
''''''''''''        strSQL = strSQL & intValor_campo(0) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(1) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(2) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(3) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(4) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(5) & ", "
''''''''''''        strSQL = strSQL & adrCampos_evento("DFmultiplicador")
''''''''''''
''''''''''''        'Executa a SQL acima construida.
''''''''''''        CNConexao.Execute strSQL
''''''''''''    Else
''''''''''''        'Testa se existe na tabela dois eventos duplicados do Salario Base
''''''''''''        If CNConexao.Execute("SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha WHERE DFcod_evento = " & intCodigo_Salario_Base)(0) <> 1 Then
''''''''''''            CNConexao.Execute ("DELETE FROM TBtemp_calculo_folha WHERE DFcod_evento = " & intCodigo_Salario_Base & " AND DFid_fixo <> 0")
''''''''''''        End If
''''''''''''    End If
''''''''''''
''''''''''''    Exit Sub
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Evento_Automatico_Salario_Base")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''Private Sub Evento_Automatico_Salario_Familia()
'''''''''''''******************************************************************************
'''''''''''''Sistema...........................: Director
'''''''''''''Módulo............................: Pessoal
'''''''''''''Procedimento/Função...............: Evento_Automatico_Salario_Familia
'''''''''''''Objetivo:.........................: Calcula o valor do Salario Familia de acordo com o numero de filhos
'''''''''''''Desenvolvimento...................: Vitor Constâncio da Silva
'''''''''''''Data de criação...................: 01/08/2001
'''''''''''''Data da última manutenção.........: 21/09/200101/08/2001
'''''''''''''Manutenção executada por..........: Vitor Constâncio da Silva
'''''''''''''Observaçãoes......................:
'''''''''''''******************************************************************************
''''''''''''    On Error GoTo Erro
''''''''''''    Dim strSQL As String
''''''''''''    Dim sinNumFilhos As Single
''''''''''''    Dim curValor_Salario_Familia As Currency
''''''''''''    Dim intCod_evento As Integer
''''''''''''    Dim curSalario As Currency
''''''''''''    Dim adrCampos_evento As ADODB.Recordset
''''''''''''    'Buscando informações necessárias para a inclusão de um novo evento.
''''''''''''    'Atribui-se os valores encontrados às variaveis de memória.
''''''''''''    adrOpcoes.MoveFirst
''''''''''''    adrOpcoes.Find ("DFcodigo = 109")
''''''''''''    intCod_evento = adrOpcoes("DFvalor")
''''''''''''
''''''''''''    'SQL para verificar se o evento Salario Familia já existe para este funcionário.
''''''''''''    strSQL = "SELECT COUNT(DFcod_evento) FROM TBtemp_calculo_folha " & _
''''''''''''             "WHERE DFcod_evento = " & intCod_evento
''''''''''''    If CNConexao.Execute(strSQL)(0) = 0 Then
''''''''''''        'Faz a Contagem de todos os dependentes Ativos do funcionário
''''''''''''        strSQL = "SELECT COUNT(DFmatricula) FROM TBdependente " & _
''''''''''''                 "WHERE DFmatricula = " & adrFiltro_primario("DFmatricula") & " " & _
''''''''''''                 "AND DFativo = 1 "
''''''''''''        sinNumFilhos = CNConexao.Execute(strSQL)(0)
''''''''''''
''''''''''''        'verifica se existe dependente
''''''''''''        If sinNumFilhos = 0 Then
''''''''''''            Exit Sub
''''''''''''        End If
''''''''''''        '---------------------------------------------------------------------------------
''''''''''''        If CNConexao.Execute("SELECT DFsalario FROM TBfuncionario WHERE DFmatricula = " & adrFiltro_primario("DFmatricula")).Fields("DFsalario") > CNConexao.Execute("SELECT DFlimite FROM TBevento WHERE DFcod_evento = " & intCod_evento).Fields("DFlimite") Then
''''''''''''            Exit Sub
''''''''''''        End If
''''''''''''        'Faz o calculo do valor do Salario Familia multiplicado
''''''''''''        'pela quantidade de dependentes ativos do funcionário
''''''''''''        curValor_Salario_Familia = CNConexao.Execute("SELECT DFvalor FROM TBevento WHERE DFcod_evento = " & intCod_evento).Fields("DFvalor") * sinNumFilhos
'''''''''''''        Set adrCampos_evento = New ADODB.Recordset
'''''''''''''        adrCampos_evento.Open "SELECT * FROM TBevento WHERE DFcod_evento = " & intCod_Evento, CNConexao, adOpenForwardOnly, adLockOptimistic
''''''''''''        Call Banco_Dados.SQLgeral("SELECT * FROM TBevento WHERE DFcod_evento = " & intCod_evento, adrCampos_evento)
''''''''''''
''''''''''''        'transformando de True/False para 1/0.
''''''''''''        Dim intIndice As Integer
''''''''''''        Dim intCont As Integer
''''''''''''        Dim intValor_campo(5) As Integer
''''''''''''        intIndice = 5
''''''''''''        For intCont = 0 To 5
''''''''''''            If adrCampos_evento(intIndice) = True Then
''''''''''''                intValor_campo(intCont) = 1
''''''''''''            ElseIf adrCampos_evento(intIndice) = False Then
''''''''''''                intValor_campo(intCont) = 0
''''''''''''            End If
''''''''''''            intIndice = intIndice + 1
''''''''''''        Next
''''''''''''
''''''''''''        'SQL de inserção de registro na TBtemp de eventos.
''''''''''''        strSQL = "INSERT INTO TBtemp_calculo_folha "
''''''''''''        strSQL = strSQL & "([DFmatricula],[DFid_fixo],"
''''''''''''        strSQL = strSQL & "[DFcod_evento],[DFparcela],"
''''''''''''        strSQL = strSQL & "[DFpercentual_sob_salario_base],"
''''''''''''        strSQL = strSQL & "[DFpercentual_sob_dias_trabalhados],"
''''''''''''        strSQL = strSQL & "[DFvalor],[DFdesconta_faltas],[DFdescricao],"
''''''''''''        strSQL = strSQL & "[DFtipo],[DFreferencia],[DFvalor_referencia],"
''''''''''''        strSQL = strSQL & "[DFimprime_referencia],[DFincide_inss],"
''''''''''''        strSQL = strSQL & "[DFincide_fgts],[DFincide_irrf],"
''''''''''''        strSQL = strSQL & "[DFincide_rais],[DFincide_informe_rendimentos],"
''''''''''''        strSQL = strSQL & "[DFmultiplicador]) "
''''''''''''
''''''''''''        strSQL = strSQL & "SELECT " & adrFiltro_primario("DFmatricula") & ", "
''''''''''''        strSQL = strSQL & "NULL," & intCod_evento & ", NULL, NULL, NULL, "
''''''''''''        strSQL = strSQL & Grava_Moeda(curValor_Salario_Familia) & ", NULL, '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFdescricao") & "', '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFtipo") & "', '"
''''''''''''        strSQL = strSQL & adrCampos_evento("DFreferencia") & "', '"
''''''''''''        strSQL = strSQL & sinNumFilhos & "', "
''''''''''''        strSQL = strSQL & intValor_campo(0) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(1) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(2) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(3) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(4) & ", "
''''''''''''        strSQL = strSQL & intValor_campo(5) & ", "
''''''''''''        strSQL = strSQL & adrCampos_evento("DFmultiplicador")
''''''''''''
''''''''''''        'Executa a SQL acima construida.
''''''''''''        CNConexao.Execute strSQL
''''''''''''
''''''''''''
''''''''''''    End If
''''''''''''    Exit Sub
''''''''''''Erro:
''''''''''''    Call Erro.Erro("Evento_Automatico_Salario_Familia")
''''''''''''    Resume Next
''''''''''''End Sub
''''''''''''
''''''''''''
''''''''''''Public Sub Verificar_Dias_Movimentacoes(Data_Emissao As Date, Matricula As String, Ultimo_Dia_Mes As Integer, Valores As String)
''''''''''''    Dim adrMovimentacoes As ADODB.Recordset
''''''''''''    Dim strSQL As String
''''''''''''
''''''''''''    strSQL = _
''''''''''''        "SELECT TBafastamento_funcionario.DFid_afastamento_funcionario, TBafastamento_Funcionario.DFcod_movimentacao, TBafastamento_funcionario.DFdata_inicio, TBafastamento.DFtipo " & _
''''''''''''        "FROM TBfuncionario " & _
''''''''''''             "INNER JOIN TBafastamento_funcionario " & _
''''''''''''                     "ON TBfuncionario.DFmatricula = TBafastamento_funcionario.DFmatricula " & _
''''''''''''             "INNER JOIN TBafastamento " & _
''''''''''''                     "ON TBafastamento_funcionario.DFcod_movimentacao = TBafastamento.DFcod_afastamento " & _
''''''''''''        "WHERE (CONVERT(CHAR(4),YEAR(DFdata_inicio)))+ " & _
''''''''''''              "(CONVERT(CHAR(2),MONTH(DFdata_inicio))) = " & _
''''''''''''              "'" & Format(Data_Emissao, "yyyyM") & "' " & _
''''''''''''          "AND TBafastamento_funcionario.DFmatricula = " & Matricula
''''''''''''    Call Banco_Dados.SQLgeral(strSQL, adrMovimentacoes)
''''''''''''
''''''''''''    If CNConexao.Execute(Replace(strSQL, "TBafastamento_funcionario.DFid_afastamento_funcionario, TBafastamento_Funcionario.DFcod_movimentacao, TBafastamento_funcionario.DFdata_inicio, TBafastamento.DFtipo", "Count(*)"))(0) <> 0 Then
''''''''''''        adrMovimentacoes.MoveFirst
''''''''''''        Do While Not adrMovimentacoes.EOF
''''''''''''            If Day(adrMovimentacoes("DFdata_inicio")) > 1 Then
''''''''''''                Valores = Valores & _
''''''''''''                          IIf(Valores <> Empty, ",", Empty) & _
''''''''''''                          Format(Day(IIf(adrMovimentacoes("DFtipo") = "A", _
''''''''''''                                        (adrMovimentacoes("DFdata_inicio") - 1), _
''''''''''''                                         adrMovimentacoes("DFdata_inicio"))), "00")
''''''''''''            Else
''''''''''''                Valores = Valores & _
''''''''''''                          IIf(Valores <> Empty, ",", Empty) & _
''''''''''''                          Format(Day(IIf(adrMovimentacoes("DFtipo") = "A", _
''''''''''''                                        (adrMovimentacoes("DFdata_inicio")), _
''''''''''''                                         adrMovimentacoes("DFdata_inicio"))), "00")
''''''''''''            End If
''''''''''''            adrMovimentacoes.MoveNext
''''''''''''        Loop
''''''''''''        adrMovimentacoes.MoveFirst
''''''''''''
''''''''''''        If adrMovimentacoes("DFtipo") = "A" And (adrMovimentacoes.RecordCount Mod 2) = 0 Then
''''''''''''            Valores = "01," & Valores & "," & Ultimo_Dia_Mes
''''''''''''        ElseIf adrMovimentacoes("DFtipo") = "A" And (adrMovimentacoes.RecordCount Mod 2) <> 0 Then
''''''''''''            Valores = "01," & Valores
''''''''''''        ElseIf adrMovimentacoes("DFtipo") = "R" And (adrMovimentacoes.RecordCount Mod 2) <> 0 Then
''''''''''''            Valores = Valores & "," & Ultimo_Dia_Mes
''''''''''''        End If
''''''''''''
''''''''''''    Else
''''''''''''        Valores = Empty
''''''''''''    End If
''''''''''''
''''''''''''End Sub




