Attribute VB_Name = "Erro"
'*******************************************************************************************
'Programa��o.......................: Marcos Bai�o
'Data..............................: 00/00/2000
'
'Este m�dulo foi desenvolvido para o tratamento dos erros que podem
'ocorrer no programa sendo assim mais facilmente efetuada uma manuten��o
'para a corre��o deste erro, j� traduzido
'
'Par�metros:Modulo       (Armazena o nome do Form ou do M�dulo onde o erro esta acontecendo)
'           Procedimento (Armazena a fun��o ou evento onde o erro est� acontecendo)
'           DataError    (Armazena o numero do erro da fun��o Error do DataGrid)
'
'*******************************************************************************************
Dim log As New DLLSystemManager.log

'Public Function Erro(Form As Object, Aplicacao As String, Optional Evento As String, Optional DataError As Integer) As String 'Optional Interface As String,
'        'Fun��o aqui s� inserida para o tratamento dos erros do DataGrid
'        If DataError <> Empty Then
'            Select Case DataError
'                Case 7007
'                    MsgBox "Tipo de dado inv�lido.", vbCritical, "Only Tech"
'                Case 13
'                    MsgBox "Tipo de dado incompat�vel.", vbCritical, "Only Tech"
'                Case 6153
'                    MsgBox "Informa��o de coluna insuficiente para atualizar.", vbCritical, "Only Tech"
'                Case Else
'                    MsgBox "Erro do Data Grid n� " & DataError, vbCritical, "Only Tech"
'            End Select
'            Exit Function
'        End If
'
'        Select Case Err.Number
'
'            Case 20507
'                MsgBox "Nome de Arquivo Inv�lido", vbCritical, "Only Tech"
'            Case -2147217904
'                MsgBox "Texto em campo num�rico", vbCritical, "Only Tech"
'            Case -2147217900
'                MsgBox "Erro de Sintaxe ou Chave duplicada!Ou Viola��o de Integridade Referencial", vbCritical, "Only Tech"
'                Exit Function
'            Case -2147217871
'                MsgBox "Tempo limite de opera��o excedido", vbCritical, "Only Tech"
'            Case -2147217865
'                MsgBox "Houve um problema de conex�o. Verifique sua rede ou o caminho para conex�o", vbCritical, "Only Tech"
'            Case -2147217833
'                MsgBox "Tipo de dado inv�lido", vbCritical, "Only Tech"
'            Case -2147467259
'                MsgBox "Falha na conex�o com o servidor. Pode ser necess�rio reiniciar o Only Tech.", vbInformation, "Only Tech"
'            Case -2147217843
'                MsgBox "Falha no login do usu�rio.", vbInformation, "Only Tech"
'
'            Case 53
'                MsgBox "Arquivo n�o encontrado ou caminho incorreto, altere e tente novamente", vbCritical, "Only Tech"
'            Case 91
'                MsgBox "Todas as altera��es foram canceladas", vbInformation, "Only Tech"
'            Case 3021
'                MsgBox "Um erro foi encontrado na pesquisa, tente novamente informando os dados corretamente", vbCritical, "Only Tech"
'
'           'Erros de ADO Decimal Positivo
'
'            Case 3001
'                MsgBox "O Aplicativo esta usando argumentos de algum tipo incorreto, estao fora do limite permitido, ou em conflito com um outro.", vbCritical, "Only Tech"
'            Case 3002
'                MsgBox "Erro ocorrido durante tentativa de abertura de arquivo.", vbCritical, "Only Tech"
'            Case 3003
'                MsgBox "Erro ocorrido durante tentativa de leitura de arquivo.", vbCritical, "Only Tech"
'            Case 3004
'                MsgBox "Erro ocorrido durante tentativa de gravacao de arquivo.", vbCritical, "Only Tech"
'            Case 3219
'                MsgBox "A operacao requerida pela aplicacao nao e permitida.", vbCritical, "Only Tech"
'            Case 3246
'                MsgBox "O Aplicativo nao pode fechar um objeto de Conexao no meio de uma transacao", vbCritical, "Only Tech"
'            Case 3251
'                MsgBox "A operacao requerida pela aplicacao nao e suportada pelo provedor.", vbCritical, "Only Tech"
'            Case 3265
'                MsgBox "ADO nao encontrou o objeto na colecao correspondente ao nome ou referencia solicitada pelo aplicativo.", vbCritical, "Only Tech"
'            Case 3367
'                MsgBox "Objeto nao pode ser adicionado. O objeto ja esta na colecao.", vbCritical, "Only Tech"
'            Case 3420
'                MsgBox "O objeto referenciado pelo aplicativo nao mais aponta para um objeto valido.", vbCritical, "Only Tech"
'            Case 3421
'                MsgBox "O aplicativo esta usando o valor de um tipo incorreto para a operacao atual.", vbCritical, "Only Tech"
'            Case 3704
'                MsgBox "A operacao solicitada pelo aplicativo nao e permitida se o objeto esta fechado.", vbCritical, "Only Tech"
'            Case 3705
'                MsgBox "A operacao requerida pela aplicacao nao e permitida se o objeto estiver aberto.", vbCritical, "Only Tech"
'            Case 3706
'                MsgBox "ADO nao pode encontrar o provedor especificado.", vbCritical, "Only Tech"
'            Case 3707
'                MsgBox "O Aplicativo nao pode alterar a propriedade ActiveConnection do objeto Recordset com o objeto Command como sua fonte de dados.", vbCritical, "Only Tech"
'            Case 3708
'                MsgBox "O Aplicativo definiu impropramente um objeto parametro.", vbCritical, "Only Tech"
'            Case 3709
'                MsgBox "O Aplicativo solicitou uma operacao em um objeto com referencia a um objeto Connection que foi fechado ou e invalido.", vbCritical, "Only Tech"
'            Case 3710
'                MsgBox "Operacao invalida no objeto durente processamento do evento.", vbCritical, "Only Tech"
'            Case 3711
'                MsgBox "Operacao invalida no objeto enquanto processa um outro comando.", vbCritical, "Only Tech"
'            Case 3712
'                MsgBox "Operacao cancelada pelo usuario.", vbCritical, "Only Tech"
'            Case 3713
'                MsgBox "Operacao invalida no objeto enquanto ainda estiver conectado.", vbCritical, "Only Tech"
'            Case 3715
'                MsgBox "Operacao invalida no objeto enquanto nao e executado.", vbCritical, "Only Tech"
'            Case 3716
'                MsgBox "A operacao solicitada pela aplicacao nao e segura para as configuracoes da maquina", vbCritical, "Only Tech"
'
'           'Erros de ADO Decimal Negativo
'
'            Case -2146824581
'                MsgBox "O Aplicativo nao pode alterar a propriedade ActiveConnection do objeto Recordset com o objeto Command como sua fonte de dados.", vbCritical, "Only Tech"
'            Case -2146824867
'                MsgBox "O aplicativo esta usando o valor de um tipo incorreto para a operacao atual.", vbCritical, "Only Tech"
'            Case -2146825037
'                MsgBox "A operacao requerida pela aplicacao nao e suportada pelo provedor.", vbCritical, "Only Tech"
'            Case -2146825037
'                MsgBox "A operacao requerida pela aplicacao nao e permitida.", vbCritical, "Only Tech"
'            Case -2146825042
'                MsgBox "O Aplicativo nao pode fechar um objeto de Conexao no meio de uma transacao", vbCritical, "Only Tech"
'            Case -2146825287
'                MsgBox "O Aplicativo esta usando argumentos de algum tipo incorreto, estao fora do limite permitido, ou em conflito com um outro.", vbCritical, "Only Tech"
'            Case -2146824579
'                MsgBox "O Aplicativo solicitou uma operacao em um objeto com referencia a um objeto Connection que foi fechado ou e invalido.", vbCritical, "Only Tech"
'            Case -2146824580
'                MsgBox "O Aplicativo definiu impropramente um objeto parametro.", vbCritical, "Only Tech"
'            Case -2146825023
'                MsgBox "ADO nao encontrou o objeto na colecao correspondente ao nome ou referencia solicitada pelo aplicativo.", vbCritical, "Only Tech"
'           'Referente ao erro 3021
'            Case -2146825267
'                MsgBox "O registro corrente foi excluido,a operacao solicitada pelo aplicativo requer uma registro corrente.", vbCritical, "Only Tech"
'            Case -2146824573
'                MsgBox "Operacao invalida no objeto enquanto nao e executado.", vbCritical, "Only Tech"
'            Case -2146824578
'                MsgBox "Operacao invalida no objeto durente processamento do evento.", vbCritical, "Only Tech"
'            Case -2146824584
'                MsgBox "A operacao solicitada pelo aplicativo nao e permitida se o objeto esta fechado.", vbCritical, "Only Tech"
'            Case -2146824921
'                MsgBox "Objeto nao pode ser adicionado. O objeto ja esta na colecao.", vbCritical, "Only Tech"
'            Case -2146824868
'                MsgBox "O objeto referenciado pelo aplicativo nao mais aponta para um objeto valido.", vbCritical, "Only Tech"
'            Case -2146824583
'                MsgBox "A operacao requerida pela aplicacao nao e permitida se o objeto estiver aberto.", vbCritical, "Only Tech"
'            Case -2146825286
'                MsgBox "Erro ocorrido durante tentativa de abertura de arquivo.", vbCritical, "Only Tech"
'            Case -2146824576
'                MsgBox "Operacao cancelada pelo usuario.", vbCritical, "Only Tech"
'            Case -2146824582
'                MsgBox "ADO nao pode encontrar o provedor especificado.", vbCritical, "Only Tech"
'            Case -2146824285
'                MsgBox "Erro ocorrido durante tentativa de leitura de arquivo.", vbCritical, "Only Tech"
'            Case -2146824575
'                MsgBox "Operacao invalida no objeto enquanto ainda estiver conectado.", vbCritical, "Only Tech"
'            Case -2146824577
'                MsgBox "Operacao invalida no objeto enquanto processa um outro comando.", vbCritical, "Only Tech"
'            Case -2146824572
'                MsgBox "A operacao solicitada pela aplicacao nao e segura para as configuracoes da maquina", vbCritical, "Only Tech"
'            Case -2146825284
'                MsgBox "Erro ocorrido durante tentativa de gravacao de arquivo.", vbCritical, "Only Tech"
'            Case -2147217873
'                MsgBox "Erro de integridade refer�ncial.Este registro n�o pode ser INCLUIDO/DELETADO.", vbCritical, "Only Tech"
'           'Erros Interceptaveis
'            Case 3
'                MsgBox "Return sem GoSub", vbCritical, "Only Tech"
'            Case 5
'                MsgBox "Chamada de procedimento inv�lida", vbCritical, "Only Tech"
'            Case 6
'                MsgBox "Sobrecarga", vbCritical, "Only Tech"
'            Case 7
'                MsgBox "Mem�ria insuficiente", vbCritical, "Only Tech"
'            Case 9
'                MsgBox "Subscrito fora do intervalo", vbCritical, "Only Tech"
'            Case 10
'                MsgBox "Esta matriz � fixa ou est� temporariamente bloqueada", vbCritical, "Only Tech"
'            Case 11
'                MsgBox "Divis�o por zero", vbCritical, "Only Tech"
'            Case 13
'                MsgBox "Tipo incompat�vel", vbCritical, "Only Tech"
'            Case 14
'                MsgBox "Espa�o insuficiente para seq��ncia de caracteres", vbCritical, "Only Tech"
'            Case 16
'                MsgBox "Express�o muito complexa", vbCritical, "Only Tech"
'            Case 17
'                MsgBox "N�o � poss�vel executar a opera��o solicitada", vbCritical, "Only Tech"
'            Case 18
'                MsgBox "Ocorreu uma interrup��o do usu�rio", vbCritical, "Only Tech"
'            Case 20
'                MsgBox "Recome�ar sem erro", vbCritical, "Only Tech"
'            Case 28
'                MsgBox "Espa�o insuficiente para pilha", vbCritical, "Only Tech"
'            Case 35
'                MsgBox "Sub, Function ou Property n�o definida", vbCritical, "Only Tech"
'            Case 47
'                MsgBox "N�mero excessivo de clientes do aplicativo DLL", vbCritical, "Only Tech"
'            Case 48
'                MsgBox "Erro ao carregar DLL", vbCritical, "Only Tech"
'            Case 49
'                MsgBox "Conven��o de chamada DLL inv�lida", vbCritical, "Only Tech"
'            Case 51
'                MsgBox "erro interno", vbCritical, "Only Tech"
'            Case 52
'                MsgBox "Nome ou n�mero de arquivo inv�lido", vbCritical, "Only Tech"
'            Case 54
'                MsgBox "Modo de arquivo inv�lido", vbCritical, "Only Tech"
'            Case 55
'                MsgBox "O arquivo j� est� aberto", vbCritical, "Only Tech"
'            Case 57
'                MsgBox "Erro de dispositivo de E/S", vbCritical, "Only Tech"
'            Case 58
'                MsgBox "O arquivo j� existe", vbCritical, "Only Tech"
'            Case 59
'                MsgBox "Comprimento de registro inv�lido", vbCritical, "Only Tech"
'            Case 61
'                MsgBox "disco cheio", vbCritical, "Only Tech"
'            Case 62
'                MsgBox "Entrada depois do fim do arquivo", vbCritical, "Only Tech"
'            Case 63
'                MsgBox "N�mero de registro inv�lido", vbCritical, "Only Tech"
'            Case 67
'                MsgBox "N�mero excessivo de arquivos", vbCritical, "Only Tech"
'            Case 68
'                MsgBox "Dispositivo n�o dispon�vel", vbCritical, "Only Tech"
'            Case 70
'                MsgBox "Permiss�o negada", vbCritical, "Only Tech"
'            Case 71
'                MsgBox "O disco n�o est� pronto", vbCritical, "Only Tech"
'            Case 74
'                MsgBox "N�o � poss�vel renomear com unidade de disco diferente", vbCritical, "Only Tech"
'            Case 75
'                MsgBox "Erro de acesso a caminho/arquivo", vbCritical, "Only Tech"
'            Case 76
'                MsgBox "Caminho n�o encontrado", vbCritical, "Only Tech"
'            Case 92
'                MsgBox "Loop �For� n�o inicializado", vbCritical, "Only Tech"
'            Case 93
'                MsgBox "Seq��ncia de caracteres padr�o inv�lida", vbCritical, "Only Tech"
'            Case 94
'                MsgBox "Uso de Null inv�lido", vbCritical, "Only Tech"
'            Case 97
'                MsgBox "N�o � poss�vel chamar procedimento Friend para um objeto que n�o � uma inst�ncia da classe de defini��o", vbCritical, "Only Tech"
'            Case 298
'                MsgBox "DLL do sistema n�o pode ser carregada", vbCritical, "Only Tech"
'            Case 320
'                MsgBox "N�o � poss�vel utilizar nomes de dispositivos em nomes de arquivos espec�ficos", vbCritical, "Only Tech"
'            Case 321
'                MsgBox "Formato de arquivo inv�lido", vbCritical, "Only Tech"
'            Case 322
'                MsgBox "N�o � poss�vel criar arquivo tempor�rio necess�rio", vbCritical, "Only Tech"
'            Case 325
'                MsgBox "Formato inv�lido no arquivo de recursos", vbCritical, "Only Tech"
'            Case 327
'                MsgBox "Nome do valor de dados n�o encontrado", vbCritical, "Only Tech"
'            Case 328
'                MsgBox "Par�metro ilegal; n�o � poss�vel gravar matrizes", vbCritical, "Only Tech"
'            Case 355
'                MsgBox "N�o foi poss�vel acessar registro do sistema", vbCritical, "Only Tech"
'            Case 336
'                MsgBox "Componente ActiveX n�o foi registrado corretamente", vbCritical, "Only Tech"
'            Case 337
'                MsgBox "Componente ActiveX n�o foi encontrado", vbCritical, "Only Tech"
'            Case 338
'                MsgBox "Componente ActiveX n�o executou corretamente", vbCritical, "Only Tech"
'            Case 360
'                MsgBox "Objeto j� carregado", vbCritical, "Only Tech"
'            Case 361
'                MsgBox "N�o � poss�vel carregar ou descarregar este objeto", vbCritical, "Only Tech"
'            Case 363
'                MsgBox "Controle ActiveX especificado n�o foi encontrado", vbCritical, "Only Tech"
'            Case 364
'                MsgBox "Objeto foi descarregado", vbCritical, "Only Tech"
'            Case 365
'                MsgBox "N�o � poss�vel carregar dentro desse contexto", vbCritical, "Only Tech"
'            Case 368
'                MsgBox "O arquivo especificado est� desatualizado. Este programa exige uma vers�o posterior", vbCritical, "Only Tech"
'            Case 371
'                MsgBox "O objeto especificado n�o pode ser utilizado como um formul�rio propriet�rio de Show", vbCritical, "Only Tech"
'            Case 380
'                MsgBox "Valor de propriedade inv�lido", vbCritical, "Only Tech"
'            Case 381
'                MsgBox "�ndice de matriz de propriedades inv�lido", vbCritical, "Only Tech"
'            Case 382
'                MsgBox "Propriedade Set n�o pode ser executada em tempo de execu��o", vbCritical, "Only Tech"
'            Case 383
'                MsgBox "Propriedade Set n�o pode ser utilizada com uma propriedade somente leitura", vbCritical, "Only Tech"
'            Case 385
'                MsgBox "� necess�rio o �ndice de matriz de propriedade", vbCritical, "Only Tech"
'            Case 387
'                MsgBox "Propriedade Set n�o permitida", vbCritical, "Only Tech"
'            Case 393
'                MsgBox "Propriedade Get n�o pode ser executada em tempo de execu��o", vbCritical, "Only Tech"
'            Case 394
'                MsgBox "Propriedade Get n�o pode ser executada em propriedade somente grava��o", vbCritical, "Only Tech"
'            Case 400
'                MsgBox "Formul�rio j� exibido; imposs�vel exibir de forma modal", vbCritical, "Only Tech"
'            Case 402
'                MsgBox "C�digo deve fechar o formul�rio modal superior", vbCritical, "Only Tech"
'            Case 419
'                MsgBox "Permiss�o para utilizar objeto negada", vbCritical, "Only Tech"
'            Case 422
'                MsgBox "Propriedade n�o encontrada", vbCritical, "Only Tech"
'            Case 423
'                MsgBox "Propriedade ou m�todo n�o foi encontrado", vbCritical, "Only Tech"
'            Case 424
'                MsgBox "Objeto � obrigat�rio", vbCritical, "Only Tech"
'            Case 425
'                MsgBox "Uso inv�lido de objeto", vbCritical, "Only Tech"
'            Case 429
'                'O programa precisa de um objeto que n�o est� registrado ou n�o
'                'existe no dico r�gido. O m�dulo n�o inciar� ou n�o funcionar�
'                'corretamente.
'                MsgBox "O componente ActiveX n�o pode criar um objeto ou retornar refer�ncia a esse objeto", vbCritical, "Only Tech"
'                End
'            Case 430
'                MsgBox "A classe n�o aceita Automa��o", vbCritical, "Only Tech"
'            Case 432
'                MsgBox "O nome do arquivo ou o nome da classe n�o foi encontrado durante a opera��o de Automa��o", vbCritical, "Only Tech"
'            Case 438
'                MsgBox "O objeto n�o aceita esta propriedade ou m�todo", vbCritical, "Only Tech"
'            Case 440
'                MsgBox "Erro de automa��o", vbCritical, "Only Tech"
'            Case 442
'                MsgBox "A conex�o � biblioteca de objetos ou de tipos para processo remoto foi perdida", vbCritical, "Only Tech"
'            Case 443
'                MsgBox "O objeto de Automa��o n�o possui um valor padr�o", vbCritical, "Only Tech"
'            Case 445
'                MsgBox "O objeto n�o suporta esta a��o", vbCritical, "Only Tech"
'            Case 446
'                MsgBox "O objeto n�o suporta argumentos nomeados", vbCritical, "Only Tech"
'            Case 447
'                MsgBox "O objeto n�o aceita a defini��o atual de localidade", vbCritical, "Only Tech"
'            Case 448
'                MsgBox "O argumento nomeado n�o foi encontrado", vbCritical, "Only Tech"
'            Case 449
'                MsgBox "Argumento n�o opcional ou atribui��o de propriedade inv�lida", vbCritical, "Only Tech"
'            Case 450
'                MsgBox "N�mero de argumentos incorreto ou atribui��o de propriedade inv�lida", vbCritical, "Only Tech"
'            Case 451
'                MsgBox "Object n�o � uma cole��o", vbCritical, "Only Tech"
'            Case 452
'                MsgBox "Ordinal inv�lido", vbCritical, "Only Tech"
'            Case 453
'                MsgBox "A fun��o DLL especificada n�o foi encontrada", vbCritical, "Only Tech"
'            Case 454
'                MsgBox "O recurso de c�digo n�o foi encontrado", vbCritical, "Only Tech"
'            Case 455
'                MsgBox "Erro de prote��o de recurso de c�digo", vbCritical, "Only Tech"
'            Case 457
'                MsgBox "Esta tecla j� est� associada a um elemento desta cole��o", vbCritical, "Only Tech"
'            Case 458
'                MsgBox "A vari�vel utiliza um tipo de automa��o n�o suportada no Visual Basic", vbCritical, "Only Tech"
'            Case 459
'                MsgBox "Este componente n�o suporta eventos", vbCritical, "Only Tech"
'            Case 460
'                MsgBox "Formato da �rea de transfer�ncia inv�lido", vbCritical, "Only Tech"
'            Case 461
'                MsgBox "Formato especificado n�o corresponde ao formato dos dados", vbCritical, "Only Tech"
'            Case 480
'                MsgBox "N�o � poss�vel criar imagem AutoRedraw", vbCritical, "Only Tech"
'            Case 481
'                MsgBox "figura inv�lida", vbCritical, "Only Tech"
'            Case 482
'                MsgBox "Erro na impressora", vbCritical, "Only Tech"
'            Case 483
'                MsgBox "Driver da impressora n�o aceita a propriedade especificada", vbCritical, "Only Tech"
'            Case 484
'                MsgBox "Problemas ao obter informa��es da impressora a partir do sistema. Certifique-se de que a impressora esteja instalada corretamente", vbCritical, "Only Tech"
'            Case 485
'                MsgBox "Tipo de figura inv�lido", vbCritical, "Only Tech"
'            Case 486
'                MsgBox "N�o � poss�vel imprimir imagem de formul�rio neste tipo de impressora", vbCritical, "Only Tech"
'            Case 520
'                MsgBox "N�o � poss�vel esvaziar a �rea de transfer�ncia", vbCritical, "Only Tech"
'            Case 521
'                MsgBox "N�o � poss�vel abrir a �rea de transfer�ncia", vbCritical, "Only Tech"
'            Case 735
'                MsgBox "N�o � poss�vel salvar arquivo no diret�rio TEMP", vbCritical, "Only Tech"
'            Case 744
'                MsgBox "Texto procurado n�o encontrado", vbCritical, "Only Tech"
'            Case 746
'                MsgBox "Substitui��es muito longas", vbCritical, "Only Tech"
'            Case 31001
'                MsgBox "Mem�ria insuficiente", vbCritical, "Only Tech"
'            Case 31004
'                MsgBox "Nenhum objeto", vbCritical, "Only Tech"
'            Case 31018
'                MsgBox "Classe n�o est� definida", vbCritical, "Only Tech"
'            Case 31027
'                MsgBox "N�o � poss�vel ativar objeto", vbCritical, "Only Tech"
'            Case 31032
'                MsgBox "N�o foi poss�vel criar objeto incorporado", vbCritical, "Only Tech"
'            Case 31036
'                MsgBox "Erro ao salvar o arquivo", vbCritical, "Only Tech"
'            Case 31037
'                MsgBox "Erro ao carregar do arquivo", vbCritical, "Only Tech"
'
'           'Erros intercept�veis do Microsoft Jet e do objeto de acesso a dados (DAO, Data Access Object)
'
'            Case 2420
'                MsgBox "Erro de sintaxe em n�mero.", vbCritical, "Only Tech"
'            Case 2421
'                MsgBox "Erro de sintaxe em data.", vbCritical, "Only Tech"
'            Case 2422
'                MsgBox "Erro de sintaxe em seq��ncia.", vbCritical, "Only Tech"
'            Case 2423
'                MsgBox "Utiliza��o inv�lida de '.', '!' ou '()'.", vbCritical, "Only Tech"
'            Case 2424
'                MsgBox "Nome desconhecido.", vbCritical, "Only Tech"
'            Case 2425
'                MsgBox "Nome de fun��o desconhecido.", vbCritical, "Only Tech"
'            Case 2426
'                MsgBox "Fun��o n�o-dispon�vel em express�es.", vbCritical, "Only Tech"
'            Case 2427
'                MsgBox "Objeto sem valor.", vbCritical, "Only Tech"
'            Case 2428
'                MsgBox "Argumentos inv�lidos utilizados com a fun��o de dom�nio.", vbCritical, "Only Tech"
'            Case 2429
'                MsgBox "Operador In sem ().", vbCritical, "Only Tech"
'            Case 2430
'                MsgBox "Operador Between sem And.", vbCritical, "Only Tech"
'            Case 2431
'                MsgBox "Erro de sintaxe (operador ausente).", vbCritical, "Only Tech"
'            Case 2432
'                MsgBox "Erro de sintaxe (v�rgula).", vbCritical, "Only Tech"
'            Case 2433
'                MsgBox "Erro de sintaxe.", vbCritical, "Only Tech"
'            Case 2434
'                MsgBox "Erro de sintaxe (operador ausente).", vbCritical, "Only Tech"
'            Case 2435
'                MsgBox ") extra.", vbCritical, "Only Tech"
'            Case 2436
'                MsgBox "), ] ou item ausentes.", vbCritical, "Only Tech"
'            Case 2437
'                MsgBox "Utiliza��o inv�lida de barras verticais.", vbCritical, "Only Tech"
'            Case 2438
'                MsgBox "Erro de sintaxe.", vbCritical, "Only Tech"
'            Case 2439
'                MsgBox "N�mero incorreto de argumentos utilizados com a fun��o.", vbCritical, "Only Tech"
'            Case 2440
'                MsgBox "Fun��o IIF sem ().", vbCritical, "Only Tech"
'            Case 2442
'                MsgBox "Utiliza��o inv�lida de par�nteses.", vbCritical, "Only Tech"
'            Case 2443
'                MsgBox "Utiliza��o inv�lida do operador Is.", vbCritical, "Only Tech"
'            Case 2445
'                MsgBox "Express�o muito complexa.", vbCritical, "Only Tech"
'            Case 2446
'                MsgBox "Mem�ria insuficiente durante o c�lculo.", vbCritical, "Only Tech"
'            Case 2447
'                MsgBox "Utiliza��o inv�lida de '.', '!' ou '()'.", vbCritical, "Only Tech"
'            Case 2448
'                MsgBox "N�o � poss�vel definir o valor.", vbCritical, "Only Tech"
'            Case 3000
'                MsgBox "Erro <Item> reservado; n�o existe mensagem para este erro.", vbCritical, "Only Tech"
'            Case 3001
'                MsgBox "Argumento inv�lido.", vbCritical, "Only Tech"
'            Case 3002
'                MsgBox "N�o foi poss�vel iniciar a sess�o.", vbCritical, "Only Tech"
'            Case 3003
'                MsgBox "N�o foi poss�vel iniciar a transa��o; j� existem muitas transa��es aninhadas.", vbCritical, "Only Tech"
'            Case 3005
'                MsgBox "<Nome do banco de dados> n�o � um nome de banco de dados v�lido.", vbCritical, "Only Tech"
'            Case 3006
'                MsgBox "O banco de dados <nome> est� bloqueado exclusivamente.", vbCritical, "Only Tech"
'            Case 3007
'                MsgBox "N�o � poss�vel abrir o banco de dados da biblioteca <nome>.", vbCritical, "Only Tech"
'            Case 3008
'                MsgBox "A tabela <nome> j� est� aberta exclusivamente por outro usu�rio ou j� est� aberta atrav�s da interface do usu�rio e n�o pode ser manipulada programaticamente.", vbCritical, "Only Tech"
'            Case 3009
'                MsgBox "Voc� tentou bloquear a tabela <tabela> enquanto a abria, mas ela n�o pode ser bloqueada porque est� em uso no momento. Aguarde um instante e, em seguida, tente a opera��o novamente.", vbCritical, "Only Tech"
'            Case 3010
'                MsgBox "A tabela <nome> j� existe.", vbCritical, "Only Tech"
'            Case 3011
'                MsgBox "O mecanismo de banco de dados Microsoft Jet n�o p�de encontrar o objeto <nome>. Certifique-se de que o objeto existe e que voc� digitou o seu nome e o nome do caminho corretamente.", vbCritical, "Only Tech"
'            Case 3012
'                MsgBox "O objeto <nome> j� existe.", vbCritical, "Only Tech"
'            Case 3013
'                MsgBox "N�o foi poss�vel renomear o arquivo ISAM instal�vel.", vbCritical, "Only Tech"
'            Case 3014
'                MsgBox "N�o � poss�vel abrir mais tabelas.", vbCritical, "Only Tech"
'            Case 3015
'                MsgBox "<Nome do �ndice> n�o � um �ndice nesta tabela. Consulte a cole��o Indexes do objeto TableDef para determinar os nomes de �ndice v�lidos.", vbCritical, "Only Tech"
'            Case 3016
'                MsgBox "O campo n�o caber� no registro.", vbCritical, "Only Tech"
'            Case 3017
'                MsgBox "O tamanho do campo � grande demais.", vbCritical, "Only Tech"
'            Case 3018
'                MsgBox "N�o foi poss�vel encontrar o campo <nome>.", vbCritical, "Only Tech"
'            Case 3020
'                MsgBox "Voc� tentou chamar Update ou CancelUpdate ou tentou atualizar um Field em um conjunto de registros sem chamar primeiro AddNew ou Edit.", vbCritical, "Only Tech"
'            Case 3022
'                MsgBox "As altera��es que voc� solicitou � tabela n�o foram bem-sucedidas porque criariam valores duplicados no �ndice, na chave prim�ria ou na rela��o. Altere os dados no campo ou campos que cont�m dados duplicados, remova o �ndice ou redefina-o para permitir entradas duplicadas e tente novamente.", vbCritical, "Only Tech"
'            Case 3023
'                MsgBox "AddNew ou Edit j� utilizado.", vbCritical, "Only Tech"
'            Case 3024
'                MsgBox "N�o foi poss�vel encontrar <nome>.", vbCritical, "Only Tech"
'            Case 3025
'                MsgBox "N�o � poss�vel abrir mais arquivos.", vbCritical, "Only Tech"
'            Case 3026
'                MsgBox "Espa�o insuficiente em disco.", vbCritical, "Only Tech"
'            Case 3027
'                MsgBox "N�o foi poss�vel atualizar. O banco de dados ou objeto � somente leitura.", vbCritical, "Only Tech"
'            Case 3028
'                MsgBox "N�o � poss�vel iniciar seu aplicativo. O arquivo de informa��es do grupo de trabalho est� ausente ou aberto exclusivamente por outro usu�rio.", vbCritical, "Only Tech"
'            Case 3029
'                MsgBox "Nome de conta ou senha inv�lidos.", vbCritical, "Only Tech"
'            Case 3030
'                MsgBox "<Nome da conta> n�o � um nome de conta v�lido.", vbCritical, "Only Tech"
'            Case 3031
'                MsgBox "Senha inv�lida.", vbCritical, "Only Tech"
'            Case 3032
'                MsgBox "N�o � poss�vel executar esta opera��o.", vbCritical, "Only Tech"
'            Case 3033
'                MsgBox "Voc� n�o tem as permiss�es necess�rias para utilizar o objeto <nome>. Fa�a o seu administrador do sistema ou a pessoa que criou este objeto estabelecer as permiss�es apropriadas para voc�.", vbCritical, "Only Tech"
'            Case 3034
'                MsgBox "Voc� tentou aceitar ou cancelar uma transa��o sem utilizar primeiro BeginTrans.", vbCritical, "Only Tech"
'            Case 3036
'                MsgBox "O banco de dados alcan�ou o tamanho m�ximo.", vbCritical, "Only Tech"
'            Case 3037
'                MsgBox "N�o � poss�vel abrir mais tabelas ou consultas.", vbCritical, "Only Tech"
'            Case 3039
'                MsgBox "N�o foi poss�vel criar o �ndice; muitos �ndices j� definidos.", vbCritical, "Only Tech"
'            Case 3040
'                MsgBox "Erro de E/S em disco durante a leitura.", vbCritical, "Only Tech"
'            Case 3041
'                MsgBox "N�o � poss�vel abrir um banco de dados criado com uma vers�o anterior do seu aplicativo.", vbCritical, "Only Tech"
'            Case 3042
'                MsgBox "Sem identificadores de arquivo do MS-DOS.", vbCritical, "Only Tech"
'            Case 3043
'                MsgBox "Erro de disco ou rede.", vbCritical, "Only Tech"
'            Case 3044
'                MsgBox "<Caminho> n�o � um caminho v�lido. Certifique-se de que o nome do caminho est� digitado corretamente e que voc� est� conectado ao servidor no qual se encontra o arquivo.", vbCritical, "Only Tech"
'            Case 3045
'                MsgBox "N�o foi poss�vel utilizar <nome>; o arquivo j� est� em utiliza��o.", vbCritical, "Only Tech"
'            Case 3046
'                MsgBox "N�o foi poss�vel salvar; atualmente bloqueado por outro usu�rio.", vbCritical, "Only Tech"
'            Case 3047
'                MsgBox "O registro � grande demais.", vbCritical, "Only Tech"
'            Case 3048
'                MsgBox "N�o � poss�vel abrir mais bancos de dados.", vbCritical, "Only Tech"
'            Case 3049
'                MsgBox "N�o � poss�vel abrir o banco de dados <nome>. Ele pode n�o ser um banco de dados que o seu aplicativo reconhe�a ou o arquivo pode estar corrompido.", vbCritical, "Only Tech"
'            Case 3051
'                MsgBox "O mecanismo de banco de dados Microsoft Jet n�o pode abrir o arquivo <nome>. Ele j� est� aberto exclusivamente por outro usu�rio ou voc� precisa de permiss�o para visualizar seus dados.", vbCritical, "Only Tech"
'            Case 3052
'                MsgBox "O n�mero de bloqueios de compartilhamento de arquivos do MS-DOS foi excedido. Voc� precisa aumentar o n�mero de bloqueios instalados com Share.exe.", vbCritical, "Only Tech"
'            Case 3053
'                MsgBox "Tarefas cliente em excesso.", vbCritical, "Only Tech"
'            Case 3054
'                MsgBox "Campos Memorando ou �Objeto OLE� em excesso.", vbCritical, "Only Tech"
'            Case 3055
'                MsgBox "Nome de campo inv�lido.", vbCritical, "Only Tech"
'            Case 3056
'                MsgBox "N�o foi poss�vel reparar este banco de dados.", vbCritical, "Only Tech"
'            Case 3057
'                MsgBox "Opera��o n�o suportada em tabelas vinculadas.", vbCritical, "Only Tech"
'            Case 3058
'                MsgBox "O �ndice ou chave prim�ria n�o pode conter um valor Null.", vbCritical, "Only Tech"
'            Case 3059
'                MsgBox "Opera��o cancelada pelo usu�rio.", vbCritical, "Only Tech"
'            Case 3060
'                MsgBox "Tipo de dados incorreto para o par�metro <par�metro>.", vbCritical, "Only Tech"
'            Case 3061
'                MsgBox "Muito poucos par�metros. Eram esperados <n�mero>.", vbCritical, "Only Tech"
'            Case 3062
'                MsgBox "Alias de sa�da <nome> duplicado.", vbCritical, "Only Tech"
'            Case 3063
'                MsgBox "Destino de sa�da <nome> duplicado.", vbCritical, "Only Tech"
'            Case 3064
'                MsgBox "N�o � poss�vel abrir a consulta a��o <nome>.", vbCritical, "Only Tech"
'            Case 3065
'                MsgBox "N�o � poss�vel executar uma consulta sele��o.", vbCritical, "Only Tech"
'            Case 3066
'                MsgBox "A consulta deve ter pelo menos um campo de destino.", vbCritical, "Only Tech"
'            Case 3067
'                MsgBox "A entrada da consulta deve conter pelo menos uma tabela ou consulta.", vbCritical, "Only Tech"
'            Case 3068
'                MsgBox "Nome de alias inv�lido.", vbCritical, "Only Tech"
'            Case 3069
'                MsgBox "A consulta a��o <nome> n�o pode ser utilizada como origem da linha.", vbCritical, "Only Tech"
'            Case 3070
'                MsgBox "O mecanismo de banco de dados Microsoft Jet n�o reconhece <nome> como um nome de campo ou express�o v�lida.", vbCritical, "Only Tech"
'            Case 3071
'                MsgBox "Esta express�o est� digitada incorretamente ou � complexa demais para ser avaliada. Por exemplo, uma express�o num�rica pode conter muitos elementos complicados. Tente simplificar a express�o atribuindo partes da express�o a vari�veis.", vbCritical, "Only Tech"
'            Case 3073
'                MsgBox "A opera��o deve utilizar uma consulta atualiz�vel.", vbCritical, "Only Tech"
'            Case 3074
'                MsgBox "N�o � poss�vel repetir o nome da tabela <nome> na cl�usula FROM.", vbCritical, "Only Tech"
'            Case 3075
'                MsgBox "<Mensagem> na express�o de consulta <express�o>.", vbCritical, "Only Tech"
'            Case 3076
'                MsgBox "<Nome> na express�o de crit�rio.", vbCritical, "Only Tech"
'            Case 3077
'                MsgBox "<Mensagem> na express�o.", vbCritical, "Only Tech"
'            Case 3078
'                MsgBox "O mecanismo de banco de dados Microsoft Jet n�o consegue encontrar a tabela de entrada ou a consulta <nome>. Certifique-se de que ela existe e que o seu nome est� digitado corretamente.", vbCritical, "Only Tech"
'            Case 3079
'                MsgBox "O campo especificado <campo> poderia se referir a mais de uma tabela listada na cl�usula FROM da sua instru��o SQL.", vbCritical, "Only Tech"
'            Case 3080
'                MsgBox "A tabela associada <nome> n�o est� listada na cl�usula FROM.", vbCritical, "Only Tech"
'            Case 3081
'                MsgBox "N�o � poss�vel associar mais de uma tabela com o mesmo nome <nome>.", vbCritical, "Only Tech"
'            Case 3082
'                MsgBox "A opera��o JOIN <opera��o> refere-se a um campo que n�o est� em uma das tabelas associadas.", vbCritical, "Only Tech"
'            Case 3083
'                MsgBox "N�o � poss�vel utilizar consulta de relat�rio interno.", vbCritical, "Only Tech"
'            Case 3084
'                MsgBox "N�o � poss�vel inserir dados com a consulta a��o.", vbCritical, "Only Tech"
'            Case 3085
'                MsgBox "Fun��o <nome> indefinida na express�o.", vbCritical, "Only Tech"
'            Case 3086
'                MsgBox "N�o foi poss�vel excluir das tabelas especificadas.", vbCritical, "Only Tech"
'            Case 3087
'                MsgBox "Express�es em excesso na cl�usula GROUP BY.", vbCritical, "Only Tech"
'            Case 3088
'                MsgBox "Express�es em excesso na cl�usula ORDER BY.", vbCritical, "Only Tech"
'            Case 3089
'                MsgBox "Express�es em excesso na sa�da DISTINCT.", vbCritical, "Only Tech"
'            Case 3090
'                MsgBox "A tabela resultante n�o pode ter mais de um campo AutoNumera��o.", vbCritical, "Only Tech"
'            Case 3092
'                MsgBox "N�o � poss�vel utilizar a cl�usula HAVING na instru��o TRANSFORM.", vbCritical, "Only Tech"
'            Case 3093
'                MsgBox "A cl�usula ORDER BY <cl�usula> entra em conflito com DISTINCT.", vbCritical, "Only Tech"
'            Case 3094
'                MsgBox "A cl�usula ORDER BY <cl�usula> entra em conflito com a cl�usula GROUP BY.", vbCritical, "Only Tech"
'            Case 3095
'                MsgBox "N�o � poss�vel ter uma fun��o agregada na express�o <express�o>.", vbCritical, "Only Tech"
'            Case 3096
'                MsgBox "N�o � poss�vel ter uma fun��o agregada na cl�usula WHERE <cl�usula>.", vbCritical, "Only Tech"
'            Case 3097
'                MsgBox "N�o � poss�vel ter uma fun��o agregada na cl�usula ORDER BY <cl�usula>.", vbCritical, "Only Tech"
'            Case 3098
'                MsgBox "N�o � poss�vel ter uma fun��o agregada na cl�usula GROUP BY <cl�usula>.", vbCritical, "Only Tech"
'            Case 3099
'                MsgBox "N�o � poss�vel ter uma fun��o agregada na opera��o JOIN <opera��o>.", vbCritical, "Only Tech"
'            Case 3100
'                MsgBox "N�o � poss�vel definir o campo <nome> na chave de associa��o como Null.", vbCritical, "Only Tech"
'            Case 3101
'                MsgBox "O mecanismo de banco de dados Microsoft Jet n�o consegue encontrar um registro na tabela <nome> com campo(s) <nome> de correspond�ncia de chave.", vbCritical, "Only Tech"
'            Case 3102
'                MsgBox "Refer�ncia circular causada pela <refer�ncia da consulta>.", vbCritical, "Only Tech"
'            Case 3103
'                MsgBox "Refer�ncia circular causada pelo alias <nome> na lista SELECT da defini��o da consulta.", vbCritical, "Only Tech"
'            Case 3104
'                MsgBox "N�o � poss�vel especificar mais de uma vez o <valor> do t�tulo de colunas fixas em uma consulta de tabela de refer�ncia cruzada.", vbCritical, "Only Tech"
'            Case 3105
'                MsgBox "Nome do campo de destino ausente na instru��o SELECT INTO <instru��o>.", vbCritical, "Only Tech"
'            Case 3106
'                MsgBox "Nome do campo de destino ausente na instru��o UPDATE <instru��o>.", vbCritical, "Only Tech"
'            Case 3107
'                MsgBox "Registro(s) n�o pode(m) ser adicionado(s); sem permiss�o de inser��o no <nome>.", vbCritical, "Only Tech"
'            Case 3108
'                MsgBox "Registro(s) n�o pode(m) ser editado(s); sem permiss�o de atualiza��o em <nome>.", vbCritical, "Only Tech"
'            Case 3109
'                MsgBox "Registro(s) n�o pode(m) ser exclu�dos, sem permiss�o de exclus�o em <nome>.", vbCritical, "Only Tech"
'            Case 3110
'                MsgBox "N�o foi poss�vel ler defini��es; sem permiss�o de leitura de defini��es da tabela ou consulta <nome>.", vbCritical, "Only Tech"
'            Case 3111
'                MsgBox "N�o foi poss�vel criar; sem permiss�o de modifica��o da estrutura da tabela ou consulta <nome>.", vbCritical, "Only Tech"
'            Case 3112
'                MsgBox "Registro(s) n�o pode(m) ser lido(s); sem permiss�o de leitura em <nome>.", vbCritical, "Only Tech"
'            Case 3113
'                MsgBox "N�o � poss�vel atualizar <nome do campo>; campo n�o atualiz�vel.", vbCritical, "Only Tech"
'            Case 3114
'                MsgBox "N�o � poss�vel incluir Memorando ou Objeto OLE quando forem selecionados valores exclusivos <instru��o>.", vbCritical, "Only Tech"
'            Case 3115
'                MsgBox "N�o � poss�vel ter campos Memorando ou Objeto OLE no argumento agregado <instru��o>.", vbCritical, "Only Tech"
'            Case 3116
'                MsgBox "N�o � poss�vel ter campos Memorando ou Objeto OLE no crit�rio <crit�rio> para a fun��o agregada.", vbCritical, "Only Tech"
'            Case 3117
'                MsgBox "N�o � poss�vel classificar em Memorando ou Objeto OLE <cl�usula>.", vbCritical, "Only Tech"
'            Case 3118
'                MsgBox "N�o � poss�vel associar em Memorando ou Objeto OLE <nome>.", vbCritical, "Only Tech"
'            Case 3119
'                MsgBox "N�o � poss�vel agrupar em Memorando ou Objeto OLE <cl�usula>.", vbCritical, "Only Tech"
'            Case 3120
'                MsgBox "N�o � poss�vel agrupar em campos selecionados com '*' <nome da tabela>.", vbCritical, "Only Tech"
'            Case 3121
'                MsgBox "N�o � poss�vel agrupar em campos selecionados com '*'.", vbCritical, "Only Tech"
'            Case 3122
'                MsgBox "Voc� tentou executar uma consulta que n�o inclui a express�o <nome> especificada como parte de uma fun��o agregada.", vbCritical, "Only Tech"
'            Case 3123
'                MsgBox "N�o � poss�vel utilizar '*' em consulta de tabela de refer�ncia cruzada.", vbCritical, "Only Tech"
'            Case 3124
'                MsgBox "N�o � poss�vel obter a entrada pela consulta de relat�rio interno <nome>.", vbCritical, "Only Tech"
'            Case 3125
'                MsgBox "O mecanismo de banco de dados n�o consegue encontrar <nome>. Certifique-se de que � um nome de par�metro ou alias v�lido, que n�o inclui caracteres nem pontua��o inv�lida e que o nome n�o � grande demais.", vbCritical, "Only Tech"
'            Case 3126
'                MsgBox "Colchetes inv�lidos no nome <nome>.", vbCritical, "Only Tech"
'            Case 3127
'                MsgBox "A instru��o INSERT INTO cont�m o seguinte nome de campo desconhecido: <nome do campo>. Certifique-se de que voc� digitou o nome corretamente e tente a opera��o novamente.", vbCritical, "Only Tech"
'            Case 3128
'                MsgBox "Especifique a tabela que cont�m os registros que deseja excluir.", vbCritical, "Only Tech"
'            Case 3129
'                MsgBox "Instru��o SQL inv�lida; era esperado 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT' ou 'UPDATE'.", vbCritical, "Only Tech"
'            Case 3130
'                MsgBox "Erro de sintaxe na instru��o DELETE.", vbCritical, "Only Tech"
'            Case 3131
'                MsgBox "Erro de sintaxe na cl�usula FROM.", vbCritical, "Only Tech"
'            Case 3132
'                MsgBox "Erro de sintaxe na cl�usula GROUP BY.", vbCritical, "Only Tech"
'            Case 3133
'                MsgBox "Erro de sintaxe na cl�usula HAVING.", vbCritical, "Only Tech"
'            Case 3134
'                MsgBox "Erro de sintaxe na instru��o INSERT INTO.", vbCritical, "Only Tech"
'            Case 3135
'                MsgBox "Erro de sintaxe na opera��o JOIN.", vbCritical, "Only Tech"
'            Case 3136
'                MsgBox "A cl�usula LEVEL inclui uma palavra ou argumento reservado que est� digitado incorretamente ou est� ausente, ou a pontua��o est� incorreta.", vbCritical, "Only Tech"
'            Case 3138
'                MsgBox "Erro de sintaxe na cl�usula ORDER BY.", vbCritical, "Only Tech"
'            Case 3139
'                MsgBox "Erro de sintaxe na cl�usula PARAMETER.", vbCritical, "Only Tech"
'            Case 3140
'                MsgBox "Erro de sintaxe na cl�usula PROCEDURE.", vbCritical, "Only Tech"
'            Case 3141
'                MsgBox "A instru��o SELECT inclui uma palavra ou argumento reservado ou um nome de argumento que est� digitado incorretamente ou est� ausente, ou a pontua��o est� incorreta.", vbCritical, "Only Tech"
'            Case 3143
'                MsgBox "Erro de sintaxe na instru��o TRANSFORM.", vbCritical, "Only Tech"
'            Case 3144
'                MsgBox "Erro de sintaxe na instru��o UPDATE.", vbCritical, "Only Tech"
'            Case 3145
'                MsgBox "Erro de sintaxe na cl�usula WHERE.", vbCritical, "Only Tech"
'            Case 3146
'                MsgBox "ODBC � a chamada falhou.", vbCritical, "Only Tech"
'            Case 3151
'                MsgBox "ODBC�� a conex�o a <nome> falhou.", vbCritical, "Only Tech"
'            Case 3154
'                MsgBox "ODBC�� n�o foi poss�vel encontrar DLL <nome>.", vbCritical, "Only Tech"
'            Case 3155
'                MsgBox "ODBC�� a inser��o em uma tabela vinculada <tabela> falhou.", vbCritical, "Only Tech"
'            Case 3156
'                MsgBox "ODBC�� a exclus�o em uma tabela vinculada <tabela> falhou.", vbCritical, "Only Tech"
'            Case 3157
'                MsgBox "ODBC�� a atualiza��o em uma tabela vinculada <tabela> falhou.", vbCritical, "Only Tech"
'            Case 3158
'                MsgBox "N�o foi poss�vel salvar o registro; bloqueado no momento por outro usu�rio.", vbCritical, "Only Tech"
'            Case 3159
'                MsgBox "Indicador inv�lido.", vbCritical, "Only Tech"
'            Case 3160
'                MsgBox "A tabela n�o est� aberta.", vbCritical, "Only Tech"
'            Case 3161
'                MsgBox "N�o foi poss�vel descriptografar o arquivo.", vbCritical, "Only Tech"
'            Case 3162
'                MsgBox "Voc� tentou atribuir o valor Null a uma vari�vel que n�o � um tipo de dados Variant.", vbCritical, "Only Tech"
'            Case 3163
'                MsgBox "O campo � pequeno demais para aceitar a quantidade de dados que voc� tentou adicionar. Tente inserir ou colar menos dados.", vbCritical, "Only Tech"
'            Case 3164
'                MsgBox "O campo n�o pode ser atualizado porque outro usu�rio ou processo bloqueou o registro ou tabela correspondente.", vbCritical, "Only Tech"
'            Case 3165
'                MsgBox "N�o foi poss�vel abrir o arquivo .inf.", vbCritical, "Only Tech"
'            Case 3166
'                MsgBox "N�o � poss�vel localizar o arquivo de memorando Xbase solicitado.", vbCritical, "Only Tech"
'            Case 3167
'                MsgBox "Registro exclu�do.", vbCritical, "Only Tech"
'            Case 3168
'                MsgBox "Arquivo .inf inv�lido.", vbCritical, "Only Tech"
'            Case 3169
'                MsgBox "O mecanismo de banco de dados Microsoft Jet n�o p�de executar a instru��o SQL porque ela cont�m um campo que possui um tipo de dados inv�lido.", vbCritical, "Only Tech"
'            Case 3170
'                MsgBox "N�o foi poss�vel encontrar o ISAM instal�vel.", vbCritical, "Only Tech"
'            Case 3171
'                MsgBox "N�o foi poss�vel encontrar o caminho da rede ou o nome de usu�rio.", vbCritical, "Only Tech"
'            Case 3172
'                MsgBox "N�o foi poss�vel abrir o Paradox.net.", vbCritical, "Only Tech"
'            Case 3173
'                MsgBox "N�o foi poss�vel abrir a tabela 'MSysAccounts' no arquivo de informa��es do grupo de trabalho.", vbCritical, "Only Tech"
'            Case 3174
'                MsgBox "N�o foi poss�vel abrir a tabela 'MSysGroups' no arquivo de informa��es do grupo de trabalho.", vbCritical, "Only Tech"
'            Case 3175
'                MsgBox "A data est� fora do intervalo ou est� em um formato inv�lido.", vbCritical, "Only Tech"
'            Case 3176
'                MsgBox "N�o foi poss�vel abrir o arquivo <nome>.", vbCritical, "Only Tech"
'            Case 3177
'                MsgBox "Nome de tabela inv�lido.", vbCritical, "Only Tech"
'            Case 3179
'                MsgBox "Encontrado fim de arquivo inesperado.", vbCritical, "Only Tech"
'            Case 3180
'                MsgBox "N�o foi poss�vel gravar no arquivo <nome>.", vbCritical, "Only Tech"
'            Case 3181
'                MsgBox "Intervalo inv�lido.", vbCritical, "Only Tech"
'            Case 3182
'                MsgBox "Formato de arquivo inv�lido.", vbCritical, "Only Tech"
'            Case 3183
'                MsgBox "Espa�o insuficiente no disco tempor�rio.", vbCritical, "Only Tech"
'            Case 3184
'                MsgBox "N�o foi poss�vel executar a consulta; n�o foi poss�vel encontrar a tabela vinculada.", vbCritical, "Only Tech"
'            Case 3185
'                MsgBox "SELECT INTO em um banco de dados remoto tentou produzir campos demais.", vbCritical, "Only Tech"
'            Case 3186
'                MsgBox "SELECT INTO em um banco de dados remoto tentou produzir campos demais.", vbCritical, "Only Tech"
'            Case 3187
'                MsgBox "N�o foi poss�vel ler; atualmente bloqueado pelo usu�rio <nome> na m�quina <nome>.", vbCritical, "Only Tech"
'            Case 3188
'                MsgBox "N�o foi poss�vel atualizar; atualmente bloqueado por outra sess�o nesta m�quina.", vbCritical, "Only Tech"
'            Case 3189
'                MsgBox "Tabela <nome> � bloqueada exclusivamente pelo usu�rio <nome> na m�quina <nome>.", vbCritical, "Only Tech"
'            Case 3190
'                MsgBox "Definidos campos em excesso.", vbCritical, "Only Tech"
'            Case 3191
'                MsgBox "N�o � poss�vel definir o campo mais de uma vez.", vbCritical, "Only Tech"
'            Case 3192
'                MsgBox "N�o foi poss�vel encontrar a tabela de sa�da <nome>.", vbCritical, "Only Tech"
'            Case 3196
'                MsgBox "O banco de dados <nome do banco de dados> j� est� em uso por outra pessoa ou processo. Quando o banco de dados estiver dispon�vel, tente a opera��o novamente.", vbCritical, "Only Tech"
'            Case 3197
'                MsgBox "O mecanismo de banco de dados Microsoft Jet parou o processo porque voc� e outro usu�rio est�o tentando alterar os mesmos dados ao mesmo tempo.", vbCritical, "Only Tech"
'            Case 3198
'                MsgBox "N�o foi poss�vel iniciar a sess�o. J� existem sess�es em excesso ativas.", vbCritical, "Only Tech"
'            Case 3199
'                MsgBox "N�o foi poss�vel encontrar refer�ncia.", vbCritical, "Only Tech"
'            Case 3200
'                MsgBox "O registro n�o pode ser exclu�do nem alterado porque a tabela <nome> inclui registros relacionados.", vbCritical, "Only Tech"
'            Case 3201
'                MsgBox "Voc� n�o pode adicionar nem alterar um registro porque um registro relacionado � requerido na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3202
'                MsgBox "N�o foi poss�vel salvar; atualmente bloqueado por outro usu�rio.", vbCritical, "Only Tech"
'            Case 3203
'                MsgBox "Subconsultas n�o podem ser utilizadas na express�o <express�o>.", vbCritical, "Only Tech"
'            Case 3204
'                MsgBox "O banco de dados j� existe.", vbCritical, "Only Tech"
'            Case 3205
'                MsgBox "T�tulos de coluna da tabela de refer�ncia cruzada <valor> em excesso.", vbCritical, "Only Tech"
'            Case 3206
'                MsgBox "N�o � poss�vel criar uma rela��o entre um campo e ele mesmo.", vbCritical, "Only Tech"
'            Case 3207
'                MsgBox "Opera��o n�o suportada em uma tabela do Paradox sem chave prim�ria.", vbCritical, "Only Tech"
'            Case 3208
'                MsgBox "Configura��o Deleted inv�lida na chave Xbase do Registro do Windows.", vbCritical, "Only Tech"
'            Case 3210
'                MsgBox "A seq��ncia de conex�o � longa demais.", vbCritical, "Only Tech"
'            Case 3211
'                MsgBox "O mecanismo de banco de dados n�o p�de bloquear a tabela <nome> porque ela j� est� em uso por outra pessoa ou processo.", vbCritical, "Only Tech"
'            Case 3212
'                MsgBox "N�o foi poss�vel bloquear a tabela <nome>; atualmente em uso pelo usu�rio <nome> na m�quina <nome>.", vbCritical, "Only Tech"
'            Case 3213
'                MsgBox "Configura��o Date inv�lida na chave Xbase do Registro do Windows.", vbCritical, "Only Tech"
'            Case 3214
'                MsgBox "Configura��o Mark inv�lida na chave Xbase do Registro do Windows.", vbCritical, "Only Tech"
'            Case 3215
'                MsgBox "Tarefas Btrieve em excesso.", vbCritical, "Only Tech"
'            Case 3216
'                MsgBox "Par�metro <nome> especificado onde � requerido um nome de tabela.", vbCritical, "Only Tech"
'            Case 3217
'                MsgBox "Par�metro <nome> especificado onde � requerido um nome de banco de dados.", vbCritical, "Only Tech"
'            Case 3218
'                MsgBox "N�o foi poss�vel atualizar; atualmente bloqueado.", vbCritical, "Only Tech"
'             Case 3219
'                MsgBox "Opera��o inv�lida.", vbCritical, "Only Tech"
'             Case 3220
'                MsgBox "Seq��ncia de agrupamento incorreta.", vbCritical, "Only Tech"
'             Case 3221
'                MsgBox "Configura��es inv�lidas na chave Btrieve do Registro do Windows.", vbCritical, "Only Tech"
'             Case 3222
'                MsgBox "A consulta n�o pode conter um par�metro Database.", vbCritical, "Only Tech"
'             Case 3223
'                MsgBox "<Nome do par�metro> � inv�lido porque � longo demais ou cont�m caracteres inv�lidos.", vbCritical, "Only Tech"
'             Case 3224
'                MsgBox "N�o � poss�vel ler o dicion�rio de dados do Btrieve.", vbCritical, "Only Tech"
'             Case 3225
'                MsgBox "Encontrado um conflito de prote��o de registro durante a execu��o de uma opera��o Btrieve.", vbCritical, "Only Tech"
'             Case 3226
'                MsgBox "Erros encontrados durante a utiliza��o da DLL do Btrieve.", vbCritical, "Only Tech"
'             Case 3227
'                MsgBox "Configura��o Century inv�lida na chave Xbase do Registro do Windows.", vbCritical, "Only Tech"
'             Case 3228
'                MsgBox "Configura��o CollatingSequence inv�lida na chave Paradox do Registro do Windows.", vbCritical, "Only Tech"
'             Case 3229
'                MsgBox "Btrieve�� n�o foi poss�vel alterar o campo.", vbCritical, "Only Tech"
'             Case 3230
'                MsgBox "Arquivo de prote��o do Paradox desatualizado.", vbCritical, "Only Tech"
'             Case 3231
'                MsgBox "ODBC�� o campo ficaria longo demais; dados truncados.", vbCritical, "Only Tech"
'             Case 3232
'                MsgBox "ODBC�� n�o p�de criar tabela.", vbCritical, "Only Tech"
'             Case 3234
'                MsgBox "ODBC�� o tempo limite de consulta remota expirou.", vbCritical, "Only Tech"
'             Case 3235
'                MsgBox "ODBC�� tipo de dados n�o suportado no servidor.", vbCritical, "Only Tech"
'             Case 3238
'                MsgBox "ODBC�� dados fora do intervalo.", vbCritical, "Only Tech"
'             Case 3239
'                MsgBox "Usu�rios ativos em excesso.", vbCritical, "Only Tech"
'             Case 3240
'                MsgBox "Btrieve�� mecanismo Btrieve ausente.", vbCritical, "Only Tech"
'             Case 3241
'                MsgBox "Btrieve�� sem recursos.", vbCritical, "Only Tech"
'             Case 3242
'                MsgBox "Refer�ncia inv�lida na instru��o SELECT.", vbCritical, "Only Tech"
'             Case 3243
'                MsgBox "Nenhum dos nomes de campo de importa��o corresponde aos campos na tabela acrescentada.", vbCritical, "Only Tech"
'             Case 3244
'                MsgBox "N�o � poss�vel importar planilha protegida por senha.", vbCritical, "Only Tech"
'             Case 3245
'                MsgBox "N�o foi poss�vel analisar os nomes de campo da primeira linha da tabela de importa��o.", vbCritical, "Only Tech"
'             Case 3246
'                MsgBox "Opera��o n�o suportada em transa��es.", vbCritical, "Only Tech"
'             Case 3247
'                MsgBox "ODBC�� a defini��o da tabela vinculada mudou.", vbCritical, "Only Tech"
'             Case 3248
'                MsgBox "Configura��o NetworkAccess inv�lida no Registro do Windows.", vbCritical, "Only Tech"
'             Case 3249
'                MsgBox "Configura��o PageTimeout inv�lida no Registro do Windows.", vbCritical, "Only Tech"
'             Case 3250
'                MsgBox "N�o foi poss�vel construir chave.", vbCritical, "Only Tech"
'             Case 3251
'                MsgBox "A opera��o n�o � suportada para esse tipo de objeto.", vbCritical, "Only Tech"
'             Case 3252
'                MsgBox "N�o � poss�vel abrir um formul�rio cuja consulta base cont�m uma fun��o definida pelo usu�rio que tenta definir ou obter a propriedade Recordsetclose do formul�rio.", vbCritical, "Only Tech"
'             Case 3254
'                MsgBox "ODBC�� N�o � poss�vel bloquear todos os registros.", vbCritical, "Only Tech"
'             Case 3256
'                MsgBox "Arquivo de �ndice n�o encontrado.", vbCritical, "Only Tech"
'             Case 3257
'                MsgBox "Erro de sintaxe na declara��o WITH OWNERACCESS OPTION.", vbCritical, "Only Tech"
'             Case 3258
'                MsgBox "A instru��o SQL n�o poderia ser executada porque cont�m associa��es externas amb�guas. Para for�ar uma das associa��es a ser executada primeiro, crie uma consulta separada que execute a primeira associa��o e, em seguida, inclua essa consulta na sua instru��o SQL.", vbCritical, "Only Tech"
'             Case 3259
'                MsgBox "Tipo de dados de campo inv�lido.", vbCritical, "Only Tech"
'             Case 3260
'                MsgBox "N�o foi poss�vel atualizar; atualmente bloqueado pelo usu�rio <nome> na m�quina <nome>.", vbCritical, "Only Tech"
'             Case 3261
'                MsgBox "A tabela <nome> � bloqueada exclusivamente pelo usu�rio <nome> na m�quina <nome>.", vbCritical, "Only Tech"
'             Case 3262
'                MsgBox "N�o foi poss�vel bloquear a tabela <nome>; atualmente em uso pelo usu�rio <nome> na m�quina <nome>.", vbCritical, "Only Tech"
'             Case 3264
'                MsgBox "Sem campo definido � n�o � poss�vel acrescentar TableDef nem Index.", vbCritical, "Only Tech"
'             Case 3265
'                MsgBox "Item n�o encontrado nesta cole��o.", vbCritical, "Only Tech"
'             Case 3266
'                MsgBox "N�o � poss�vel acrescentar um Field que j� fa�a parte de uma cole��o Fields.", vbCritical, "Only Tech"
'             Case 3267
'                MsgBox "A propriedade somente pode ser definida quando o Field faz parte da cole��o Fields de um objeto Recordset.", vbCritical, "Only Tech"
'             Case 3268
'                MsgBox "N�o � poss�vel definir esta propriedade uma vez que o objeto faz parte de uma cole��o.", vbCritical, "Only Tech"
'             Case 3269
'                MsgBox "N�o � poss�vel acrescentar um Index que j� fa�a parte de uma cole��o Indexes.", vbCritical, "Only Tech"
'             Case 3270
'                MsgBox "Propriedade n�o encontrada.", vbCritical, "Only Tech"
'             Case 3271
'                MsgBox "Valor de propriedade inv�lido.", vbCritical, "Only Tech"
'             Case 3272
'                MsgBox "O objeto n�o � uma cole��o.", vbCritical, "Only Tech"
'             Case 3273
'                MsgBox "M�todo n�o aplic�vel a este objeto.", vbCritical, "Only Tech"
'             Case 3274
'                MsgBox "A tabela externa n�o est� no formato esperado.", vbCritical, "Only Tech"
'             Case 3275
'                MsgBox "Erro inesperado do driver de banco de dados externo <n�mero do erro>.", vbCritical, "Only Tech"
'             Case 3276
'                MsgBox "Refer�ncia inv�lida a objeto de banco de dados.", vbCritical, "Only Tech"
'             Case 3277
'                MsgBox "N�o � poss�vel ter mais de 10 campos em um �ndice.", vbCritical, "Only Tech"
'             Case 3278
'                MsgBox "O mecanismo de banco de dados Microsoft Jet n�o foi inicializado.", vbCritical, "Only Tech"
'             Case 3279
'                MsgBox "O mecanismo de banco de dados Microsoft Jet j� foi inicializado.", vbCritical, "Only Tech"
'             Case 3280
'                MsgBox "N�o � poss�vel excluir um campo que fa�a parte de um �ndice ou que seja necess�rio ao sistema.", vbCritical, "Only Tech"
'             Case 3281
'                MsgBox "N�o � poss�vel excluir este �ndice ou tabela. � o �ndice atual ou � utilizado em uma rela��o.", vbCritical, "Only Tech"
'            Case 3282
'                MsgBox "Opera��o n�o suportada em uma tabela que cont�m dados.", vbCritical, "Only Tech"
'            Case 3283
'                MsgBox "J� existe chave prim�ria.", vbCritical, "Only Tech"
'            Case 3284
'                MsgBox "J� existe �ndice.", vbCritical, "Only Tech"
'            Case 3285
'                MsgBox "Defini��o de �ndice inv�lida.", vbCritical, "Only Tech"
'            Case 3286
'                MsgBox "O formato do arquivo de memorando n�o corresponde ao formato do banco de dados externo especificado.", vbCritical, "Only Tech"
'            Case 3287
'                MsgBox "N�o � poss�vel criar o �ndice no campo fornecido.", vbCritical, "Only Tech"
'            Case 3288
'                MsgBox "O �ndice do Paradox n�o � prim�rio.", vbCritical, "Only Tech"
'            Case 3289
'                MsgBox "Erro de sintaxe na cl�usula CONSTRAINT.", vbCritical, "Only Tech"
'            Case 3290
'                MsgBox "Erro de sintaxe na instru��o CREATE TABLE.", vbCritical, "Only Tech"
'            Case 3291
'                MsgBox "Erro de sintaxe na instru��o CREATE INDEX.", vbCritical, "Only Tech"
'            Case 3292
'                MsgBox "Erro de sintaxe na defini��o do campo.", vbCritical, "Only Tech"
'            Case 3293
'                MsgBox "Erro de sintaxe na instru��o ALTER TABLE.", vbCritical, "Only Tech"
'            Case 3294
'                MsgBox "Erro de sintaxe na instru��o DROP INDEX.", vbCritical, "Only Tech"
'            Case 3295
'                MsgBox "Erro de sintaxe em DROP TABLE ou DROP INDEX.", vbCritical, "Only Tech"
'            Case 3296
'                MsgBox "Express�o de associa��o n�o-suportada.", vbCritical, "Only Tech"
'            Case 3297
'                MsgBox "N�o � poss�vel importar tabela nem consulta. Nenhum registro encontrado ou todos os registros cont�m erros.", vbCritical, "Only Tech"
'            Case 3298
'                MsgBox "H� diversas tabelas com este nome. Especifique o propriet�rio no formato �propriet�rio.tabela�.", vbCritical, "Only Tech"
'            Case 3299
'                MsgBox "Erro de conformidade com a especifica��o ODBC <mensagem>. Relate este erro ao profissional da �rea de desenvolvimento do seu aplicativo.", vbCritical, "Only Tech"
'            Case 3300
'                MsgBox "N�o � poss�vel criar uma rela��o.", vbCritical, "Only Tech"
'            Case 3301
'                MsgBox "N�o � poss�vel executar esta opera��o; os recursos desta vers�o n�o est�o dispon�veis em bancos de dados com formatos mais antigos.", vbCritical, "Only Tech"
'            Case 3302
'                MsgBox "N�o � poss�vel alterar um regra enquanto as regras desta tabela estiverem em uso.", vbCritical, "Only Tech"
'            Case 3303
'                MsgBox "N�o � poss�vel excluir este campo. Ele faz parte de uma ou mais rela��es.", vbCritical, "Only Tech"
'            Case 3304
'                MsgBox "Voc� deve inserir um identificador pessoal (PID) que consista em no m�nimo 4 e no m�ximo 20 caracteres e d�gitos.", vbCritical, "Only Tech"
'            Case 3305
'                MsgBox "Seq��ncia de conex�o inv�lida na consulta passagem.", vbCritical, "Only Tech"
'            Case 3306
'                MsgBox "Voc� gravou uma subconsulta que pode retornar mais de um campo sem utilizar a palavra reservada EXISTS na cl�usula FROM da consulta principal. Altere a instru��o SELECT da subconsulta para solicitar somente um campo.", vbCritical, "Only Tech"
'            Case 3307
'                MsgBox "O n�mero de colunas nas duas tabelas ou consultas selecionadas de uma consulta uni�o n�o coincide.", vbCritical, "Only Tech"
'            Case 3308
'                MsgBox "Argumento TOP inv�lido na consulta sele��o.", vbCritical, "Only Tech"
'            Case 3309
'                MsgBox "A configura��o da propriedade n�o pode ter mais de 2K.", vbCritical, "Only Tech"
'            Case 3310
'                MsgBox "Esta propriedade n�o � suportada em fontes de dados externas ou em bancos de dados criados com uma vers�o anterior do Microsoft Jet.", vbCritical, "Only Tech"
'            Case 3311
'                MsgBox "A propriedade especificada j� existe.", vbCritical, "Only Tech"
'            Case 3312
'                MsgBox "As regras de valida��o e os valores padr�o n�o podem ser inseridos em tabelas do sistema ou vinculadas.", vbCritical, "Only Tech"
'            Case 3313
'                MsgBox "N�o � poss�vel inserir esta express�o de valida��o neste campo.", vbCritical, "Only Tech"
'            Case 3314
'                MsgBox "O campo <nome> n�o pode conter um valor Null porque a propriedade Required deste campo est� definida como True. Insira um valor neste campo.", vbCritical, "Only Tech"
'            Case 3315
'                MsgBox "O campo <nome> n�o pode ser uma seq��ncia de comprimento zero.", vbCritical, "Only Tech"
'            Case 3316
'                MsgBox "<Texto de valida��o em n�vel de tabela>.", vbCritical, "Only Tech"
'            Case 3317
'                MsgBox "Um ou mais valores s�o proibidos pela regra de valida��o <regra> definida para <nome>. Insira um valor que a express�o deste campo possa aceitar.", vbCritical, "Only Tech"
'            Case 3318
'                MsgBox "Os valores especificados em uma cl�usula TOP n�o s�o permitidos em consultas exclus�o e nem em relat�rios.", vbCritical, "Only Tech"
'            Case 3319
'                MsgBox "Erro de sintaxe na consulta uni�o.", vbCritical, "Only Tech"
'            Case 3320
'                MsgBox "<Erro> em express�o de valida��o em n�vel de tabela.", vbCritical, "Only Tech"
'            Case 3321
'                MsgBox "Sem banco de dados especificado na seq��ncia de conex�o ou cl�usula IN.", vbCritical, "Only Tech"
'            Case 3322
'                MsgBox "A consulta de tabela de refer�ncia cruzada cont�m um ou mais t�tulos fixos e inv�lidos de colunas.", vbCritical, "Only Tech"
'            Case 3323
'                MsgBox "A consulta n�o pode ser utilizada como origem da linha.", vbCritical, "Only Tech"
'            Case 3324
'                MsgBox "A consulta � uma consulta DDL e n�o pode ser utilizada como origem da linha.", vbCritical, "Only Tech"
'            Case 3325
'                MsgBox "A consulta passagem com a propriedade ReturnsRecords definida como True n�o retornou registros.", vbCritical, "Only Tech"
'            Case 3326
'                MsgBox "Este Recordset n�o � atualiz�vel.", vbCritical, "Only Tech"
'            Case 3334
'                MsgBox "Somente pode estar presente no formato da vers�o 1.0.", vbCritical, "Only Tech"
'            Case 3336
'                MsgBox "Btrieve: op��o IndexDDF inv�lida na configura��o da inicializa��o.", vbCritical, "Only Tech"
'            Case 3337
'                MsgBox "Op��o DataCodePage inv�lida na configura��o da inicializa��o.", vbCritical, "Only Tech"
'            Case 3338
'                MsgBox "Btrieve: as op��es Xtrieve n�o est�o corretas na configura��o da inicializa��o.", vbCritical, "Only Tech"
'            Case 3339
'                MsgBox "Btrieve: op��o IndexDeleteRenumber inv�lida na configura��o da inicializa��o.", vbCritical, "Only Tech"
'            Case 3340
'                MsgBox "A consulta <nome> est� corrompida.", vbCritical, "Only Tech"
'            Case 3341
'                MsgBox "O campo atual deve corresponder � chave de associa��o <nome> na tabela que serve como lado �um� da rela��o um-para-muitos. Insira um registro no lado �um� da tabela com o valor de chave desejado e, em seguida, fa�a a entrada com a chave de associa��o desejada na tabela �somente-muitos�.", vbCritical, "Only Tech"
'            Case 3342
'                MsgBox "Memorando ou Objeto OLE inv�lido na subconsulta <nome>.", vbCritical, "Only Tech"
'            Case 3343
'                MsgBox "Formato de banco de dados <nome do arquivo> n�o-reconhecido.", vbCritical, "Only Tech"
'            Case 3344
'                MsgBox "O mecanismo de banco de dados n�o reconhece o campo <nome> em uma express�o de valida��o ou o valor padr�o na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3345
'                MsgBox "Refer�ncia de campo <nome> desconhecida ou inv�lida.", vbCritical, "Only Tech"
'            Case 3346
'                MsgBox "O n�mero de valores de consulta e de campos de destino n�o � o mesmo.", vbCritical, "Only Tech"
'            Case 3349
'                MsgBox "Sobrecarga de campo num�rico.", vbCritical, "Only Tech"
'            Case 3350
'                MsgBox "O objeto � inv�lido para a opera��o.", vbCritical, "Only Tech"
'            Case 3351
'                MsgBox "A express�o ORDER BY <express�o> inclui campos que n�o s�o selecionados pela consulta. Somente os campos solicitados na primeira consulta podem ser inclu�dos em uma express�o ORDER BY.", vbCritical, "Only Tech"
'            Case 3352
'                MsgBox "Sem nome de campo de destino na instru��o INSERT INTO <instru��o>.", vbCritical, "Only Tech"
'            Case 3353
'                MsgBox "Btrieve: n�o � poss�vel encontrar o arquivo Field.ddf.", vbCritical, "Only Tech"
'            Case 3354
'                MsgBox "No m�ximo um registro pode ser retornado por esta subconsulta.", vbCritical, "Only Tech"
'            Case 3355
'                MsgBox "Erro de sintaxe no valor padr�o.", vbCritical, "Only Tech"
'            Case 3356
'                MsgBox "Voc� tentou abrir um banco de dados que j� est� aberto exclusivamente pelo usu�rio <nome> na m�quina <nome>. Tente novamente quando o banco de dados estiver dispon�vel.", vbCritical, "Only Tech"
'            Case 3357
'                MsgBox "Esta consulta n�o � uma consulta defini��o de dados devidamente formada.", vbCritical, "Only Tech"
'            Case 3358
'                MsgBox "N�o � poss�vel abrir o arquivo de informa��es do grupo de trabalho do mecanismo Microsoft Jet.", vbCritical, "Only Tech"
'            Case 3359
'                MsgBox "A consulta passagem deve conter pelo menos um caractere.", vbCritical, "Only Tech"
'            Case 3360
'                MsgBox "A consulta � complexa demais.", vbCritical, "Only Tech"
'            Case 3361
'                MsgBox "Uni�es n�o-permitidas em uma subconsulta.", vbCritical, "Only Tech"
'            Case 3362
'                MsgBox "A atualiza��o/exclus�o de linha �nica afetou mais de uma linha de uma tabela vinculada. O �ndice exclusivo cont�m valores duplicados.", vbCritical, "Only Tech"
'            Case 3364
'                MsgBox "N�o � poss�vel utilizar o campo Memorando ou Objeto OLE <nome> na cl�usula SELECT de uma consulta uni�o.", vbCritical, "Only Tech"
'            Case 3365
'                MsgBox "N�o � poss�vel definir esta propriedade para objetos remotos.", vbCritical, "Only Tech"
'            Case 3366
'                MsgBox "N�o � poss�vel acrescentar uma rela��o sem campos definidos.", vbCritical, "Only Tech"
'            Case 3367
'                MsgBox "N�o � poss�vel acrescentar. J� existe na cole��o um objeto com este nome.", vbCritical, "Only Tech"
'            Case 3368
'                MsgBox "A rela��o deve ser no mesmo n�mero de campos com os mesmos tipos de dados.", vbCritical, "Only Tech"
'            Case 3370
'                MsgBox "N�o � poss�vel modificar a estrutura da tabela <nome>. Ela est� em um banco de dados somente leitura.", vbCritical, "Only Tech"
'            Case 3371
'                MsgBox "N�o � poss�vel encontrar tabela ou restri��o.", vbCritical, "Only Tech"
'            Case 3372
'                MsgBox "N�o h� �ndice <nome> na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3373
'                MsgBox "N�o � poss�vel criar uma rela��o. A tabela referenciada <nome> n�o tem uma chave prim�ria.", vbCritical, "Only Tech"
'            Case 3374
'                MsgBox "Os campos especificados n�o s�o indexados exclusivamente na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3375
'                MsgBox "A tabela <nome> j� tem um �ndice chamado <nome>.", vbCritical, "Only Tech"
'            Case 3376
'                MsgBox "A tabela <nome> n�o existe.", vbCritical, "Only Tech"
'            Case 3377
'                MsgBox "N�o h� rela��o <nome> na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3378
'                MsgBox "J� existe uma rela��o chamada <nome> no banco de dados atual.", vbCritical, "Only Tech"
'            Case 3379
'                MsgBox "N�o � poss�vel criar rela��es para impor integridade referencial. Os dados existentes na tabela <nome> violam as regras de integridade referencial na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3380
'                MsgBox "O campo <nome> j� existe na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3381
'                MsgBox "N�o h� campo chamado <nome> na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3382
'                MsgBox "O tamanho do campo <nome> � longo demais.", vbCritical, "Only Tech"
'            Case 3383
'                MsgBox "N�o � poss�vel excluir o campo <nome>. Ele faz parte de uma ou mais rela��es.", vbCritical, "Only Tech"
'            Case 3384
'                MsgBox "N�o � poss�vel excluir uma propriedade interna.", vbCritical, "Only Tech"
'            Case 3385
'                MsgBox "As propriedades n�o definidas pelo usu�rio n�o suportam um valor Null.", vbCritical, "Only Tech"
'            Case 3386
'                MsgBox "A propriedade <nome> deve ser definida antes de utilizar este m�todo.", vbCritical, "Only Tech"
'            Case 3388
'                MsgBox "Fun��o <nome> desconhecida na express�o de valida��o ou no valor padr�o em <nome>.", vbCritical, "Only Tech"
'            Case 3389
'                MsgBox "Suporte de consulta n�o-dispon�vel.", vbCritical, "Only Tech"
'            Case 3390
'                MsgBox "O nome da conta j� existe.", vbCritical, "Only Tech"
'            Case 3393
'                MsgBox "N�o � poss�vel executar associa��o, grupo, classifica��o ou restri��o indexada. Um valor que est� sendo procurado ou classificado � longo demais.", vbCritical, "Only Tech"
'            Case 3394
'                MsgBox "N�o � poss�vel salvar a propriedade; ela � uma propriedade de esquema.", vbCritical, "Only Tech"
'            Case 3396
'                MsgBox "N�o � poss�vel executar a opera��o em cascata. Como existem registros relacionados na tabela <nome>, as regras de integridade referencial seriam violadas.", vbCritical, "Only Tech"
'            Case 3397
'                MsgBox "N�o � poss�vel executar a opera��o em cascata. Deve haver um registro relacionado na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3398
'                MsgBox "N�o � poss�vel executar a opera��o em cascata. Isto resultaria em uma chave nula na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3399
'                MsgBox "N�o � poss�vel executar a opera��o em cascata. Isto resultaria em uma chave duplicada na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3400
'                MsgBox "N�o � poss�vel executar a opera��o em cascata. Isto resultaria em duas atualiza��es do campo <nome> na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3401
'                MsgBox "N�o � poss�vel executar a opera��o em cascata. Isto transformaria o campo <nome> em Null, o que n�o � permitido.", vbCritical, "Only Tech"
'            Case 3402
'                MsgBox "N�o � poss�vel executar a opera��o em cascata. Isto transformaria o campo <nome> em uma seq��ncia de comprimento zero, o que n�o � permitido.", vbCritical, "Only Tech"
'            Case 3403
'                MsgBox "N�o � poss�vel executar a opera��o em cascata: <texto de valida��o>.", vbCritical, "Only Tech"
'            Case 3404
'                MsgBox "N�o � poss�vel executar a opera��o em cascata. O valor inserido � proibido pela regra de valida��o <regra> definida para <nome>.", vbCritical, "Only Tech"
'            Case 3405
'                MsgBox "Erro <texto de erro> na regra de valida��o.", vbCritical, "Only Tech"
'            Case 3406
'                MsgBox "A express�o que voc� est� tentando utilizar na propriedade DefaultValue � inv�lida porque <texto de erro>. Utilize uma express�o v�lida para definir esta propriedade.", vbCritical, "Only Tech"
'            Case 3407
'                MsgBox "A tabela MSysConf do servidor existe, mas est� em um formato incorreto. Entre em contato com o seu administrador do sistema.", vbCritical, "Only Tech"
'            Case 3409
'                MsgBox "Nome de campo <nome> inv�lido na defini��o de �ndice ou rela��o.", vbCritical, "Only Tech"
'            Case 3411
'                MsgBox "Entrada inv�lida. N�o � poss�vel executar a opera��o em cascata na tabela <nome> porque o valor inserido � grande demais para o campo <nome>.", vbCritical, "Only Tech"
'            Case 3412
'                MsgBox "N�o � poss�vel executar a atualiza��o em cascata na tabela porque ela est� atualmente em uso por um outro usu�rio.", vbCritical, "Only Tech"
'            Case 3414
'                MsgBox "N�o � poss�vel executar a opera��o em cascata na tabela <nome> porque ela est� atualmente em uso.", vbCritical, "Only Tech"
'            Case 3415
'                MsgBox "A seq��ncia de comprimento zero � v�lida somente em um campo Texto ou Memorando.", vbCritical, "Only Tech"
'            Case 3416
'                MsgBox "<alerta de erro reservado>", vbCritical, "Only Tech"
'            Case 3417
'                MsgBox "Uma consulta a��o n�o pode ser utilizada como origem de linha.", vbCritical, "Only Tech"
'            Case 3418
'                MsgBox "N�o � poss�vel abrir <nome da tabela>. Outro usu�rio est� com a tabela aberta utilizando um arquivo de controle de rede ou um estilo de bloqueio diferente.", vbCritical, "Only Tech"
'            Case 3419
'                MsgBox "N�o � poss�vel abrir esta tabela do Paradox 4.x ou 5.x porque o ParadoxNetStyle est� definido como 3.x no Registro do Windows.", vbCritical, "Only Tech"
'            Case 3420
'                MsgBox "O objeto � inv�lido ou n�o est� mais definido.", vbCritical, "Only Tech"
'            Case 3421
'                MsgBox "Erro de convers�o do tipo de dados.", vbCritical, "Only Tech"
'            Case 3422
'                MsgBox "N�o � poss�vel modificar a estrutura da tabela. Outro usu�rio est� com a tabela aberta.", vbCritical, "Only Tech"
'            Case 3423
'                MsgBox "Voc� n�o pode utilizar o ODBC para importar de, exportar para ou vincular uma tabela de banco de dados externa do Microsoft Jet ou ISAM para o seu banco de dados.", vbCritical, "Only Tech"
'            Case 3424
'                MsgBox "N�o � poss�vel criar o banco de dados porque a localidade � inv�lida.", vbCritical, "Only Tech"
'            Case 3428
'                MsgBox "Ocorreu um problema no seu banco de dados. Corrija-o reparando e compactando o banco de dados.", vbCritical, "Only Tech"
'            Case 3429
'                MsgBox "Vers�o incompat�vel de um ISAM instal�vel.", vbCritical, "Only Tech"
'            Case 3430
'                MsgBox "Enquanto carregava o ISAM instal�vel do Microsoft Excel, a OLE n�o conseguia inicializar.", vbCritical, "Only Tech"
'            Case 3431
'                MsgBox "Este n�o � um arquivo do Microsoft Excel 5.0.", vbCritical, "Only Tech"
'            Case 3432
'                MsgBox "Erro na abertura de um arquivo do Microsoft Excel 5.0.", vbCritical, "Only Tech"
'            Case 3433
'                MsgBox "Configura��o inv�lida na chave do Excel da se��o Engines do Registro do Windows.", vbCritical, "Only Tech"
'            Case 3434
'                MsgBox "N�o � poss�vel expandir intervalo nomeado.", vbCritical, "Only Tech"
'            Case 3435
'                MsgBox "N�o � poss�vel excluir c�lulas da planilha.", vbCritical, "Only Tech"
'            Case 3436
'                MsgBox "Falha na cria��o do arquivo.", vbCritical, "Only Tech"
'            Case 3437
'                MsgBox "A planilha est� cheia.", vbCritical, "Only Tech"
'            Case 3438
'                MsgBox "Os dados que est�o sendo exportados n�o correspondem ao formato descrito no arquivo Schema.ini.", vbCritical, "Only Tech"
'            Case 3439
'                MsgBox "Voc� tentou vincular ou importar um arquivo de mala direta do Microsoft Word. Apesar de poder exportar esses arquivos, voc� n�o pode vincul�-los nem import�-los.", vbCritical, "Only Tech"
'            Case 3440
'                MsgBox "Foi feita uma tentativa de importar ou vincular um arquivo de texto vazio. Para importar ou vincular um arquivo de texto, o arquivo deve conter dados.", vbCritical, "Only Tech"
'            Case 3441
'                MsgBox "O separador de campo de especifica��o do arquivo de texto corresponde ao separador decimal ou delimitador de texto.", vbCritical, "Only Tech"
'            Case 3442
'                MsgBox "Na especifica��o <nome> do arquivo de texto, a op��o <nome> � inv�lida.", vbCritical, "Only Tech"
'            Case 3443
'                MsgBox "A especifica��o <nome> de largura fixa n�o cont�m larguras de coluna.", vbCritical, "Only Tech"
'            Case 3444
'                MsgBox "Na especifica��o <nome> de largura fixa, a coluna <coluna> n�o especifica uma largura.", vbCritical, "Only Tech"
'            Case 3445
'                MsgBox "Foi encontrada a vers�o incorreta do arquivo DLL <nome>.", vbCritical, "Only Tech"
'            Case 3446
'                MsgBox "O arquivo VBA do Jet (VBAJET.dll para vers�es de 16 bits ou VBAJET32.dll para vers�es de 32 bits) est� ausente. Tente reinstalar o aplicativo que retornou o erro.", vbCritical, "Only Tech"
'            Case 3447
'                MsgBox "O arquivo VBA do Jet (VBAJET.dll para vers�es de 16 bits ou VBAJET32.dll para vers�es de 32 bits) n�o conseguiu inicializar quando chamado. Tente reinstalar o aplicativo que retornou o erro.", vbCritical, "Only Tech"
'            Case 3448
'                MsgBox "Uma chamada a uma fun��o do sistema OLE n�o foi bem-sucedida. Tente reinstalar o aplicativo que retornou o erro.", vbCritical, "Only Tech"
'            Case 3449
'                MsgBox "Nenhum c�digo de pa�s encontrado na seq��ncia de conex�o.", vbCritical, "Only Tech"
'            Case 3452
'                MsgBox "Voc� n�o pode fazer altera��es na estrutura do banco de dados nesta r�plica.", vbCritical, "Only Tech"
'            Case 3453
'                MsgBox "Voc� n�o pode estabelecer ou manter uma rela��o imposta entre uma tabela replicada e uma tabela local.", vbCritical, "Only Tech"
'            Case 3455
'                MsgBox "N�o � poss�vel tornar o banco de dados replic�vel.", vbCritical, "Only Tech"
'            Case 3456
'                MsgBox "O objeto chamado <nome> na cole��o <nome> n�o pode se tornar replic�vel.", vbCritical, "Only Tech"
'            Case 3457
'                MsgBox "Voc� n�o pode definir a propriedade KeepLocal para um objeto que j� est� replicado.", vbCritical, "Only Tech"
'            Case 3458
'                MsgBox "A propriedade KeepLocal n�o pode ser definida em um banco de dados; ela pode ser definida somente nos objetos em um banco de dados.", vbCritical, "Only Tech"
'            Case 3459
'                MsgBox "Depois que um banco de dados � replicado, voc� n�o pode remover os seus recursos de replica��o.", vbCritical, "Only Tech"
'            Case 3460
'                MsgBox "A opera��o que voc� tentou entra em conflito com uma opera��o existente que envolve este membro do conjunto de r�plicas.", vbCritical, "Only Tech"
'            Case 3461
'                MsgBox "A propriedade de replica��o que voc� est� tentando definir ou excluir � somente leitura e n�o pode ser alterada.", vbCritical, "Only Tech"
'            Case 3462
'                MsgBox "N�o foi poss�vel carregar a DLL.", vbCritical, "Only Tech"
'            Case 3463
'                MsgBox "N�o � poss�vel encontrar o .dll <nome>.", vbCritical, "Only Tech"
'            Case 3464
'                MsgBox "Os tipos de dados n�o correspondem na express�o de crit�rio.", vbCritical, "Only Tech"
'            Case 3465
'                MsgBox "A unidade de disco que voc� est� tentando acessar � ileg�vel.", vbCritical, "Only Tech"
'            Case 3468
'                MsgBox "O acesso foi negado enquanto acessava a pasta dropbox <nome>.", vbCritical, "Only Tech"
'            Case 3469
'                MsgBox "O disco da pasta dropbox <nome> est� cheio.", vbCritical, "Only Tech"
'            Case 3470
'                MsgBox "Falha no disco durante o acesso � pasta dropbox <nome>.", vbCritical, "Only Tech"
'            Case 3471
'                MsgBox "N�o foi poss�vel gravar no arquivo de registro Sincronizador.", vbCritical, "Only Tech"
'            Case 3472
'                MsgBox "Disco cheio para caminho <nome>.", vbCritical, "Only Tech"
'            Case 3473
'                MsgBox "Falha no disco durante o acesso ao arquivo de registro <nome>.", vbCritical, "Only Tech"
'            Case 3474
'                MsgBox "N�o � poss�vel abrir o arquivo de registro <nome> para grava��o.", vbCritical, "Only Tech"
'            Case 3475
'                MsgBox "Viola��o de compartilhamento durante a tentativa de abrir o arquivo de registro <nome> no modo Deny Write.", vbCritical, "Only Tech"
'            Case 3476
'                MsgBox "Caminho da dropbox <nome> inv�lido.", vbCritical, "Only Tech"
'            Case 3477
'                MsgBox "Endere�o da dropbox <nome> � sintaticamente inv�lido.", vbCritical, "Only Tech"
'            Case 3478
'                MsgBox "A r�plica n�o � parcial.", vbCritical, "Only Tech"
'            Case 3479
'                MsgBox "N�o � poss�vel designar uma r�plica parcial como Estrutura-Mestre para o conjunto de r�plicas.", vbCritical, "Only Tech"
'            Case 3480
'                MsgBox "A rela��o <nome> na express�o de filtro parcial � inv�lida.", vbCritical, "Only Tech"
'            Case 3481
'                MsgBox "O nome de tabela <nome> na express�o parcial de filtro � inv�lido.", vbCritical, "Only Tech"
'            Case 3482
'                MsgBox "A express�o de filtro para a r�plica parcial � inv�lida.", vbCritical, "Only Tech"
'            Case 3483
'                MsgBox "A senha fornecida para a pasta dropbox <nome> � inv�lida.", vbCritical, "Only Tech"
'            Case 3484
'                MsgBox "A senha utilizada pelo Sincronizador para gravar em uma pasta dropbox de destino � inv�lida.", vbCritical, "Only Tech"
'            Case 3485
'                MsgBox "O objeto n�o pode ser replicado porque o banco de dados n�o � replicado.", vbCritical, "Only Tech"
'            Case 3486
'                MsgBox "Voc� n�o pode adicionar um segundo campo AutoNumera��o do C�digo da Replica��o a uma tabela.", vbCritical, "Only Tech"
'            Case 3487
'                MsgBox "O banco de dados que voc� est� tentando replicar n�o pode ser convertido.", vbCritical, "Only Tech"
'            Case 3488
'                MsgBox "O valor especificado n�o � um C�digoDaReplica��o para qualquer membro do conjunto de r�plicas.", vbCritical, "Only Tech"
'            Case 3489
'                MsgBox "O objeto especificado n�o pode ser replicado porque falta nele um recurso necess�rio.", vbCritical, "Only Tech"
'            Case 3490
'                MsgBox "N�o � poss�vel criar uma nova r�plica porque o objeto <nome> no recipiente <nome> n�o p�de ser replicado.", vbCritical, "Only Tech"
'            Case 3491
'                MsgBox "O banco de dados deve ser aberto no modo exclusivo antes que ele possa ser replicado.", vbCritical, "Only Tech"
'            Case 3492
'                MsgBox "A sincroniza��o falhou porque uma altera��o de estrutura n�o p�de ser aplicada a uma das r�plicas.", vbCritical, "Only Tech"
'            Case 3493
'                MsgBox "N�o � poss�vel definir o par�metro Registro especificado para o Sincronizador.", vbCritical, "Only Tech"
'            Case 3494
'                MsgBox "N�o foi poss�vel recuperar o par�metro Registro especificado para o Sincronizador.", vbCritical, "Only Tech"
'            Case 3495
'                MsgBox "N�o h� sincroniza��es agendadas entre os dois Sincronizadores.", vbCritical, "Only Tech"
'            Case 3496
'                MsgBox "O Gerenciador de Replica��o n�o consegue encontrar o C�digoDaTroca na tabela MSysExchangeLog.", vbCritical, "Only Tech"
'            Case 3497
'                MsgBox "N�o foi poss�vel definir uma agenda para o Sincronizador.", vbCritical, "Only Tech"
'            Case 3499
'                MsgBox "N�o � poss�vel recuperar as informa��es completas de caminho para um membro do conjunto de r�plicas.", vbCritical, "Only Tech"
'            Case 3500
'                MsgBox "N�o � permitido definir uma troca com o mesmo Sincronizador.", vbCritical, "Only Tech"
'            Case 3502
'                MsgBox "A Estrutura-Mestre ou r�plica n�o est� sendo gerenciada por um Sincronizador.", vbCritical, "Only Tech"
'            Case 3503
'                MsgBox "O Registro do Sincronizador n�o tem valor definido para a chave que voc� consultou.", vbCritical, "Only Tech"
'            Case 3504
'                MsgBox "O c�digo do Sincronizador n�o corresponde a um c�digo existente na tabela MSysTranspAddress.", vbCritical, "Only Tech"
'            Case 3506
'                MsgBox "O Sincronizador � incapaz de abrir o registro do Sincronizador.", vbCritical, "Only Tech"
'            Case 3507
'                MsgBox "N�o foi poss�vel gravar no registro do Sincronizador.", vbCritical, "Only Tech"
'            Case 3508
'                MsgBox "N�o h� transporte ativo para o Sincronizador.", vbCritical, "Only Tech"
'            Case 3509
'                MsgBox "N�o foi poss�vel encontrar um transporte v�lido para este Sincronizador.", vbCritical, "Only Tech"
'            Case 3510
'                MsgBox "O membro do conjunto de r�plicas que voc� est� tentando sincronizar est� atualmente sendo utilizado em outra sincroniza��o.", vbCritical, "Only Tech"
'            Case 3512
'                MsgBox "N�o foi poss�vel ler a pasta dropbox.", vbCritical, "Only Tech"
'            Case 3513
'                MsgBox "N�o foi poss�vel gravar na pasta dropbox.", vbCritical, "Only Tech"
'            Case 3514
'                MsgBox "O Sincronizador n�o conseguiu encontrar sincroniza��es agendadas nem a serem solicitadas para processar.", vbCritical, "Only Tech"
'            Case 3515
'                MsgBox "O mecanismo de banco de dados Microsoft Jet n�o conseguiu ler o rel�gio do sistema no seu computador.", vbCritical, "Only Tech"
'            Case 3516
'                MsgBox "N�o foi poss�vel encontrar o endere�o de transporte.", vbCritical, "Only Tech"
'            Case 3517
'                MsgBox "O Sincronizador n�o conseguiu encontrar mensagens para serem processadas.", vbCritical, "Only Tech"
'            Case 3518
'                MsgBox "N�o foi poss�vel encontrar o Sincronizador na tabela MSysTranspAddress.", vbCritical, "Only Tech"
'            Case 3519
'                MsgBox "N�o foi poss�vel enviar a mensagem.", vbCritical, "Only Tech"
'            Case 3520
'                MsgBox "O nome ou c�digo da r�plica n�o corresponde a um membro atualmente gerenciado do conjunto de r�plicas.", vbCritical, "Only Tech"
'            Case 3521
'                MsgBox "Dois membros do conjunto de r�plicas n�o podem ser sincronizados porque n�o h� um ponto comum para iniciar a sincroniza��o.", vbCritical, "Only Tech"
'            Case 3522
'                MsgBox "O Sincronizador n�o consegue encontrar o registro de uma sincroniza��o espec�fica na tabela MSysExchangeLog.", vbCritical, "Only Tech"
'            Case 3523
'                MsgBox "O Sincronizador n�o consegue encontrar um n�mero de vers�o espec�fico na tabela MSysSchChange.", vbCritical, "Only Tech"
'            Case 3524
'                MsgBox "O hist�rico de altera��es de estrutura na r�plica n�o corresponde ao hist�rico na Estrutura-Mestre.", vbCritical, "Only Tech"
'            Case 3525
'                MsgBox "O Sincronizador n�o conseguiu acessar o banco de dados de mensagens.", vbCritical, "Only Tech"
'            Case 3526
'                MsgBox "O nome selecionado para o objeto do sistema j� est� em uso.", vbCritical, "Only Tech"
'            Case 3527
'                MsgBox "O Sincronizador ou Gerenciador de Replica��o n�o conseguiu encontrar o objeto do sistema.", vbCritical, "Only Tech"
'            Case 3528
'                MsgBox "N�o h� dados novos na mem�ria compartilhada para que o Sincronizador ou Gerenciador de Replica��o os leiam.", vbCritical, "Only Tech"
'            Case 3529
'                MsgBox "O Sincronizador ou Gerenciador de Replica��o encontrou dados n�o lidos na mem�ria compartilhada. Os dados existentes ser�o sobrescritos.", vbCritical, "Only Tech"
'            Case 3530
'                MsgBox "O Sincronizador j� est� servindo um cliente.", vbCritical, "Only Tech"
'            Case 3531
'                MsgBox "O per�odo de espera de um evento se esgotou.", vbCritical, "Only Tech"
'            Case 3532
'                MsgBox "O Sincronizador n�o conseguiu ser inicializado.", vbCritical, "Only Tech"
'            Case 3533
'                MsgBox "O objeto do sistema utilizado por um processo continua existindo depois que o processo parou.", vbCritical, "Only Tech"
'            Case 3534
'                MsgBox "O Sincronizador procurou por um evento do sistema, mas n�o encontrou nenhum para relatar ao cliente.", vbCritical, "Only Tech"
'            Case 3535
'                MsgBox "O cliente pediu ao Sincronizador que terminasse a opera��o.", vbCritical, "Only Tech"
'            Case 3536
'                MsgBox "O Sincronizador recebeu uma mensagem inv�lida para um membro do conjunto de r�plicas que ele gerencia.", vbCritical, "Only Tech"
'            Case 3538
'                MsgBox "N�o � poss�vel inicializar o Sincronizador porque h� aplicativos demais em execu��o.", vbCritical, "Only Tech"
'            Case 3539
'                MsgBox "Ocorreu um erro de sistema ou o seu arquivo de troca alcan�ou seu limite.", vbCritical, "Only Tech"
'            Case 3540
'                MsgBox "Seu arquivo de troca alcan�ou seu limite ou est� corrompido.", vbCritical, "Only Tech"
'            Case 3541
'                MsgBox "O Sincronizador n�o p�de ser fechado apropriadamente e continua ativo.", vbCritical, "Only Tech"
'            Case 3542
'                MsgBox "O processo parou quando se tentava terminar o cliente do Sincronizador.", vbCritical, "Only Tech"
'            Case 3543
'                MsgBox "O Sincronizador n�o foi configurado.", vbCritical, "Only Tech"
'            Case 3544
'                MsgBox "O Sincronizador j� est� sendo executado.", vbCritical, "Only Tech"
'            Case 3545
'                MsgBox "As duas r�plicas que voc� est� tentando sincronizar s�o de diferentes conjuntos de r�plicas.", vbCritical, "Only Tech"
'            Case 3546
'                MsgBox "O tipo de sincroniza��o que voc� est� tentando n�o � v�lido.", vbCritical, "Only Tech"
'            Case 3547
'                MsgBox "O Sincronizador n�o conseguiu encontrar uma r�plica do conjunto correto para concluir a sincroniza��o.", vbCritical, "Only Tech"
'            Case 3549
'                MsgBox "O nome de arquivo que voc� forneceu � longo demais.", vbCritical, "Only Tech"
'            Case 3550
'                MsgBox "N�o h� �ndice na coluna GUID.", vbCritical, "Only Tech"
'            Case 3551
'                MsgBox "N�o foi poss�vel excluir o par�metro Registro do Sincronizador.", vbCritical, "Only Tech"
'            Case 3552
'                MsgBox "O tamanho do par�metro Registro excede o m�ximo permitido.", vbCritical, "Only Tech"
'            Case 3553
'                MsgBox "O GUID n�o p�de ser criado.", vbCritical, "Only Tech"
'            Case 3555
'                MsgBox "Todos os apelidos das r�plicas j� est�o em uso.", vbCritical, "Only Tech"
'            Case 3556
'                MsgBox "Caminho inv�lido para a pasta dropbox de destino.", vbCritical, "Only Tech"
'            Case 3557
'                MsgBox "Endere�o inv�lido para a pasta dropbox de destino.", vbCritical, "Only Tech"
'            Case 3558
'                MsgBox "Erro de E/S em disco na pasta dropbox de destino.", vbCritical, "Only Tech"
'            Case 3559
'                MsgBox "N�o foi poss�vel gravar porque o disco de destino est� cheio.", vbCritical, "Only Tech"
'            Case 3560
'                MsgBox "Os dois membros do conjunto de r�plicas que voc� est� tentando sincronizar t�m o mesmo C�digoDaReplica��o.", vbCritical, "Only Tech"
'            Case 3561
'                MsgBox "Os dois membros do conjunto de r�plicas que voc� est� tentando sincronizar s�o ambos Estruturas-Mestre.", vbCritical, "Only Tech"
'            Case 3562
'                MsgBox "Acesso negado na pasta dropbox de destino.", vbCritical, "Only Tech"
'            Case 3563
'                MsgBox "Erro fatal ao acessar uma pasta dropbox local.", vbCritical, "Only Tech"
'            Case 3564
'                MsgBox "O sincronizador n�o consegue encontrar o arquivo de origem das mensagens.", vbCritical, "Only Tech"
'            Case 3565
'                MsgBox "H� uma viola��o de compartilhamento na pasta dropbox de origem porque o banco de dados de mensagens est� aberto em outro aplicativo.", vbCritical, "Only Tech"
'            Case 3566
'                MsgBox "Erro de E/S na rede.", vbCritical, "Only Tech"
'            Case 3567
'                MsgBox "A mensagem na pasta dropbox pertence ao Sincronizador errado.", vbCritical, "Only Tech"
'            Case 3568
'                MsgBox "O Sincronizador n�o conseguiu excluir um arquivo.", vbCritical, "Only Tech"
'            Case 3569
'                MsgBox "Este membro do conjunto de r�plicas foi logicamente removido do conjunto e n�o est� mais dispon�vel.", vbCritical, "Only Tech"
'            Case 3571
'                MsgBox "A tentativa de definir uma coluna em uma r�plica parcial violou uma regra que governa r�plicas parciais.", vbCritical, "Only Tech"
'            Case 3572
'                MsgBox "Ocorreu um erro de E/S em disco durante a leitura ou grava��o no diret�rio TEMP.", vbCritical, "Only Tech"
'            Case 3574
'                MsgBox "O C�digoDaReplica��o deste membro do conjunto de r�plicas foi reatribu�do durante um procedimento de movimenta��o ou c�pia.", vbCritical, "Only Tech"
'            Case 3575
'                MsgBox "A unidade de disco na qual voc� est� tentando gravar est� cheia.", vbCritical, "Only Tech"
'            Case 3576
'                MsgBox "O banco de dados que voc� est� tentando abrir j� est� em uso por outro aplicativo.", vbCritical, "Only Tech"
'            Case 3577
'                MsgBox "N�o � poss�vel atualizar a coluna do sistema de replica��o.", vbCritical, "Only Tech"
'            Case 3578
'                MsgBox "N�o foi poss�vel replicar o banco de dados; n�o � poss�vel determinar se o banco de dados est� aberto no modo exclusivo.", vbCritical, "Only Tech"
'            Case 3581
'                MsgBox "N�o � poss�vel abrir a tabela <nome> do sistema de replica��o porque ela j� est� em uso.", vbCritical, "Only Tech"
'            Case 3583
'                MsgBox "N�o � poss�vel tornar o objeto <nome> no recipiente <nome> replic�vel.", vbCritical, "Only Tech"
'            Case 3584
'                MsgBox "Mem�ria insuficiente para concluir a opera��o.", vbCritical, "Only Tech"
'            Case 3586
'                MsgBox "Erro de sintaxe na express�o de filtro parcial na tabela <nome>.", vbCritical, "Only Tech"
'            Case 3587
'                MsgBox "Express�o inv�lida na propriedade ReplicaFilter.", vbCritical, "Only Tech"
'            Case 3588
'                MsgBox "Erro ao avaliar a express�o de filtro parcial.", vbCritical, "Only Tech"
'            Case 3589
'                MsgBox "A express�o de filtro parcial cont�m uma fun��o desconhecida.", vbCritical, "Only Tech"
'            Case 3592
'                MsgBox "Voc� n�o pode replicar um banco de dados protegido por senha nem definir prote��o por senha em um banco de dados replicado.", vbCritical, "Only Tech"
'            Case 3593
'                MsgBox "Voc� n�o pode alterar o atributo-mestre de dados do conjunto de r�plicas.", vbCritical, "Only Tech"
'            Case 3594
'                MsgBox "Voc� n�o pode alterar o atributo-mestre de dados do conjunto de r�plicas. Permite altera��es de dados somente na Estrutura-Mestre.", vbCritical, "Only Tech"
'            Case 3595
'                MsgBox "As tabelas de sistema na sua r�plica n�o s�o mais confi�veis e n�o devem ser utilizadas.", vbCritical, "Only Tech"
'            Case 3605
'                MsgBox "A sincroniza��o com um banco de dados n�o-replicado n�o � permitida. O banco de dados <nome> n�o � uma Estrutura-Mestre nem uma r�plica.", vbCritical, "Only Tech"
'            Case 3607
'                MsgBox "A propriedade de replica��o que voc� est� tentando excluir � somente leitura e n�o pode ser removida.", vbCritical, "Only Tech"
'            Case 3608
'                MsgBox "O comprimento do registro � longo demais para uma tabela indexada do Paradox.", vbCritical, "Only Tech"
'            Case 3609
'                MsgBox "Nenhum �ndice exclusivo encontrado para o campo referenciado da tabela prim�ria.", vbCritical, "Only Tech"
'            Case 3610
'                MsgBox "Mesma tabela <tabela> referenciada tanto como origem quanto destino em uma consulta criar tabela.", vbCritical, "Only Tech"
'            Case 3611
'                MsgBox "N�o � poss�vel executar instru��es de defini��o de dados em fontes de dados vinculadas.", vbCritical, "Only Tech"
'            Case 3612
'                MsgBox "A cl�usula GROUP BY de v�rios n�veis n�o � permitida em uma subconsulta.", vbCritical, "Only Tech"
'            Case 3613
'                MsgBox "N�o � poss�vel criar uma rela��o em tabelas ODBC vinculadas.", vbCritical, "Only Tech"
'            Case 3614
'                MsgBox "GUID n�o permitido na express�o de crit�rio do m�todo Find.", vbCritical, "Only Tech"
'            Case 3615
'                MsgBox "O tipo n�o corresponde na express�o JOIN.", vbCritical, "Only Tech"
'            Case 3616
'                MsgBox "A atualiza��o de dados em uma tabela vinculada n�o � suportada por este ISAM.", vbCritical, "Only Tech"
'            Case 3617
'                MsgBox "A exclus�o de dados em uma tabela vinculada n�o � suportada por este ISAM.", vbCritical, "Only Tech"
'            Case 3618
'                MsgBox "A tabela de exce��es n�o p�de ser criada na importa��o/exporta��o.", vbCritical, "Only Tech"
'            Case 3619
'                MsgBox "Os registros n�o puderam ser adicionados � tabela de exce��es.", vbCritical, "Only Tech"
'            Case 3620
'                MsgBox "A conex�o para a visualiza��o da sua planilha vinculada do Microsoft Excel foi perdida.", vbCritical, "Only Tech"
'            Case 3621
'                MsgBox "N�o � poss�vel alterar a senha em um banco de dados compartilhado aberto.", vbCritical, "Only Tech"
'            Case 3622
'                MsgBox "Voc� deve utilizar a op��o dbSeeChanges com OpenRecordset quando acessar uma tabela do SQL Server que tenha uma coluna IDENTITY.", vbCritical, "Only Tech"
'            Case 3623
'                MsgBox "N�o � poss�vel acessar o arquivo DBF acoplado <nome do arquivo> do FoxPro 3.0.", vbCritical, "Only Tech"
'            Case 3624
'                MsgBox "N�o foi poss�vel ler o registro; atualmente bloqueado por outro usu�rio.", vbCritical, "Only Tech"
'            Case 3625
'                MsgBox "A especifica��o <nome> do arquivo de texto n�o existe. Voc� n�o pode importar, exportar e nem vincular utilizando a especifica��o.", vbCritical, "Only Tech"
'            Case 3626
'                MsgBox "A opera��o falhou. H� �ndices demais na tabela <nome>. Exclua alguns dos �ndices da tabela e tente a opera��o novamente.", vbCritical, "Only Tech"
'            Case 3627
'                MsgBox "N�o � poss�vel encontrar o arquivo execut�vel do Sincronizador (mstran35.exe).", vbCritical, "Only Tech"
'            Case 3628
'                MsgBox "A r�plica do parceiro n�o � gerenciada por um Sincronizador.", vbCritical, "Only Tech"
'            Case 3629
'                MsgBox "Este Sincronizador e o Sincronizador <nome> t�m a mesma dropbox do Sistema de arquivos � <nome>.", vbCritical, "Only Tech"
'            Case 3631
'                MsgBox "Nome de tabela inv�lido no filtro.", vbCritical, "Only Tech"
'            Case 3632
'                MsgBox "O Sincronizador remoto n�o est� configurado para sincroniza��o remota.", vbCritical, "Only Tech"
'            Case 3633
'                MsgBox "N�o � poss�vel carregar a DLL <nome>.", vbCritical, "Only Tech"
'            Case 3634
'                MsgBox "N�o � poss�vel criar uma r�plica utilizando uma r�plica parcial.", vbCritical, "Only Tech"
'            Case 3635
'                MsgBox "N�o � poss�vel criar uma r�plica parcial de um arquivo de informa��es do grupo de trabalho.", vbCritical, "Only Tech"
'            Case 3636
'                MsgBox "N�o � poss�vel preencher a r�plica e nem alterar o filtro da r�plica porque ela tem conflitos ou erros de dados.", vbCritical, "Only Tech"
'            Case 3637
'                MsgBox "N�o � poss�vel utilizar a tabela de refer�ncia cruzada de uma coluna n�o fixa como uma subconsulta.", vbCritical, "Only Tech"
'            Case 3638
'                MsgBox "Voc� n�o pode criar um banco de dados replic�vel que esteja sendo utilizado por um programa que controla a modifica��o.", vbCritical, "Only Tech"
'            Case 3639
'                MsgBox "N�o � poss�vel criar uma r�plica de um arquivo de informa��es do grupo de trabalho.", vbCritical, "Only Tech"
'            Case 3640
'                MsgBox "O buffer de recupera��o era pequeno demais para a quantidade de dados que voc� solicitou.", vbCritical, "Only Tech"
'            Case 3641
'                MsgBox "H� menos registros restantes no Recordset do que voc� solicitou.", vbCritical, "Only Tech"
'            Case 3642
'                MsgBox "Foi efetuado um cancelamento na opera��o.", vbCritical, "Only Tech"
'            Case 3643
'                MsgBox "Um dos registros do Recordset foi exclu�do por outro processo.", vbCritical, "Only Tech"
'            Case 3645
'                MsgBox "Um dos par�metros de liga��o est� incorreto.", vbCritical, "Only Tech"
'            Case 3646
'                MsgBox "O comprimento de linha especificado � menor que a soma dos comprimentos de coluna.", vbCritical, "Only Tech"
'            Case 3647
'                MsgBox "Uma coluna solicitada n�o est� sendo retornada ao Recordset.", vbCritical, "Only Tech"
'            Case 3648
'                MsgBox "N�o � poss�vel sincronizar uma r�plica parcial com uma outra r�plica parcial.", vbCritical, "Only Tech"
'            Case 3649
'                MsgBox "A p�gina de c�digo do idioma n�o foi especificada ou n�o p�de ser encontrada.", vbCritical, "Only Tech"
'            Case 3650
'                MsgBox "A Internet est� lenta demais.", vbCritical, "Only Tech"
'            Case 3651
'                MsgBox "Endere�o de Internet inv�lido.", vbCritical, "Only Tech"
'            Case 3652
'                MsgBox "Falha de login da Internet.", vbCritical, "Only Tech"
'            Case 3653
'                MsgBox "Internet n�o-configurada.", vbCritical, "Only Tech"
'            Case 3656
'                MsgBox "Erro na avalia��o de uma express�o parcial.", vbCritical, "Only Tech"
'            Case 3660
'                MsgBox "A troca solicitada falhou porque <descri��o>.", vbCritical, "Only Tech"
'            Case -2147168237
'                '(...)
'
'            Case Else
'                'MsgBox Error & vbNewLine & vbNewLine & "Evento:  " & Nome_do_evento & "  ,  N�:  " & Err, vbCritical
'                'Linha para implementar numera��o do c�digo:
'                MsgBox Error & vbNewLine & vbNewLine & "Evento:  " & Nome_do_evento & "  ,  Linha:  " & Erl & ", N� Erro: " & Err, vbCritical
'        End Select
'
'            'MsgBox Error & vbNewLine & vbNewLine & "Evento:  " & Nome_do_evento & "  ,  N�:  " & Err, vbCritical
'                'Linha para implementar numera��o do c�digo:
'                MsgBox Error & vbNewLine & vbNewLine & "Evento:  " & Nome_do_evento & "  ,  Linha:  " & Erl & ", N� Erro: " & Err, vbCritical
'        End
'        Form.log.Tipo = 3
'        Form.log.Numero_Erro = Err.Number
'        Form.log.Descricao = Replace(Err.Description, "'", " ")
'        Form.log.Hora = Format(Now, "hh:mm:ss")
'
'        'Grava o LOG
'        Form.log.Gravar_log Aplicacao, Form
'
'        'Limpando erros para n�o congestionar a aplica��o
'        Err.Clear
'
'End Function

Public Function Erro(Form As Object, Aplicacao As String, Optional Evento As String, Optional DataError As Integer) As String 'Optional Interface As String,

    Dim strDescricao_erro As String
    Dim dblNumer_erro As Double
    
    strDescricao_erro = Err.Description
    dblNumer_erro = Err.Number
    
    strDescricao_erro = Replace(strDescricao_erro, "'", " ")
    
    'Numero do erro, argumento opcional
    log.Numero_Erro = dblNumer_erro
    log.Descricao = strDescricao_erro
    
    MsgBox "Ocorreu um erro: " & dblNumer_erro & " - " & strDescricao_erro & " no Evento: " & Evento, vbCritical, "Only Tech"
    
    'Informa��es Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Relat�rio Resumo Acerto"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    
    'Informa��es Variaveis para o log
    log.Evento = Evento
    log.Tipo = 3
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando o log
    log.Gravar_log "Otica", Form

    'Limpando erros para n�o congestionar a aplica��o
    Err.Clear
    
End Function
