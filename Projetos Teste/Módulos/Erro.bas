Attribute VB_Name = "Erro"
'*******************************************************************************************
'Programação.......................: Marcos Baião
'Data..............................: 00/00/2000
'
'Este módulo foi desenvolvido para o tratamento dos erros que podem
'ocorrer no programa sendo assim mais facilmente efetuada uma manutenção
'para a correção deste erro, já traduzido
'
'Parâmetros:Modulo       (Armazena o nome do Form ou do Módulo onde o erro esta acontecendo)
'           Procedimento (Armazena a função ou evento onde o erro está acontecendo)
'           DataError    (Armazena o numero do erro da função Error do DataGrid)
'
'*******************************************************************************************

Public Function Erro(Optional Evento As String, Optional DataError As Integer) As String 'Optional Interface As String,
        
        'Função aqui só inserida para o tratamento dos erros do DataGrid
    
        If DataError <> Empty Then
            Select Case DataError
                Case 7007
                    MsgBox "Tipo de dado inválido.", vbCritical, "Integrador"
                Case 13
                    MsgBox "Tipo de dado incompatível.", vbCritical, "Integrador"
                Case 6153
                    MsgBox "Informação de coluna insuficiente para atualizar.", vbCritical, "Integrador"
                Case Else
                    MsgBox "Erro do Data Grid nº " & DataError, vbCritical, "Integrador"
            End Select
            Exit Function
        End If

        Select Case Err.Number
            
            Case 20507
                MsgBox "Nome de Arquivo Inválido", vbCritical, "Integrador"
            Case -2147217904
                MsgBox "Texto em campo numérico", vbCritical, "Integrador"
            Case -2147217900
                MsgBox "Erro de Sintaxe ou Chave duplicada", vbCritical, "Integrador"
            Case -2147217871
                MsgBox "Tempo limite de operação excedido", vbCritical, "Integrador"
            Case -2147217865
                MsgBox "Houve um problema de conexão. Verifique sua rede ou o caminho para conexão", vbCritical, "Integrador"
            Case -2147217833
                MsgBox "Tipo de dado inválido", vbCritical, "Integrador"
            Case -2147467259
                MsgBox "Falha na conexão com o servidor. Pode ser necessário reiniciar o Integrador.", vbInformation, "Integrador"
            Case -2147217843
                MsgBox "Falha no login do usuário.", vbInformation, "Integrador"
                
            Case 53
                MsgBox "Arquivo não encontrado ou caminho incorreto, altere e tente novamente", vbCritical, "Integrador"
            Case 91
                MsgBox "Todas as alterações foram canceladas", vbInformation, "Integrador"
            Case 3021
                MsgBox "Um erro foi encontrado na pesquisa, tente novamente informando os dados corretamente", vbCritical, "Integrador"

           'Erros de ADO Decimal Positivo

            Case 3001
                MsgBox "O Aplicativo esta usando argumentos de algum tipo incorreto, estao fora do limite permitido, ou em conflito com um outro.", vbCritical, "Integrador"
            Case 3002
                MsgBox "Erro ocorrido durante tentativa de abertura de arquivo.", vbCritical, "Integrador"
            Case 3003
                MsgBox "Erro ocorrido durante tentativa de leitura de arquivo.", vbCritical, "Integrador"
            Case 3004
                MsgBox "Erro ocorrido durante tentativa de gravacao de arquivo.", vbCritical, "Integrador"
            Case 3219
                MsgBox "A operacao requerida pela aplicacao nao e permitida.", vbCritical, "Integrador"
            Case 3246
                MsgBox "O Aplicativo nao pode fechar um objeto de Conexao no meio de uma transacao", vbCritical, "Integrador"
            Case 3251
                MsgBox "A operacao requerida pela aplicacao nao e suportada pelo provedor.", vbCritical, "Integrador"
            Case 3265
                MsgBox "ADO nao encontrou o objeto na colecao correspondente ao nome ou referencia solicitada pelo aplicativo.", vbCritical, "Integrador"
            Case 3367
                MsgBox "Objeto nao pode ser adicionado. O objeto ja esta na colecao.", vbCritical, "Integrador"
            Case 3420
                MsgBox "O objeto referenciado pelo aplicativo nao mais aponta para um objeto valido.", vbCritical, "Integrador"
            Case 3421
                MsgBox "O aplicativo esta usando o valor de um tipo incorreto para a operacao atual.", vbCritical, "Integrador"
            Case 3704
                MsgBox "A operacao solicitada pelo aplicativo nao e permitida se o objeto esta fechado.", vbCritical, "Integrador"
            Case 3705
                MsgBox "A operacao requerida pela aplicacao nao e permitida se o objeto estiver aberto.", vbCritical, "Integrador"
            Case 3706
                MsgBox "ADO nao pode encontrar o provedor especificado.", vbCritical, "Integrador"
            Case 3707
                MsgBox "O Aplicativo nao pode alterar a propriedade ActiveConnection do objeto Recordset com o objeto Command como sua fonte de dados.", vbCritical, "Integrador"
            Case 3708
                MsgBox "O Aplicativo definiu impropramente um objeto parametro.", vbCritical, "Integrador"
            Case 3709
                MsgBox "O Aplicativo solicitou uma operacao em um objeto com referencia a um objeto Connection que foi fechado ou e invalido.", vbCritical, "Integrador"
            Case 3710
                MsgBox "Operacao invalida no objeto durente processamento do evento.", vbCritical, "Integrador"
            Case 3711
                MsgBox "Operacao invalida no objeto enquanto processa um outro comando.", vbCritical, "Integrador"
            Case 3712
                MsgBox "Operacao cancelada pelo usuario.", vbCritical, "Integrador"
            Case 3713
                MsgBox "Operacao invalida no objeto enquanto ainda estiver conectado.", vbCritical, "Integrador"
            Case 3715
                MsgBox "Operacao invalida no objeto enquanto nao e executado.", vbCritical, "Integrador"
            Case 3716
                MsgBox "A operacao solicitada pela aplicacao nao e segura para as configuracoes da maquina", vbCritical, "Integrador"
            
           'Erros de ADO Decimal Negativo

            Case -2146824581
                MsgBox "O Aplicativo nao pode alterar a propriedade ActiveConnection do objeto Recordset com o objeto Command como sua fonte de dados.", vbCritical, "Integrador"
            Case -2146824867
                MsgBox "O aplicativo esta usando o valor de um tipo incorreto para a operacao atual.", vbCritical, "Integrador"
            Case -2146825037
                MsgBox "A operacao requerida pela aplicacao nao e suportada pelo provedor.", vbCritical, "Integrador"
            Case -2146825037
                MsgBox "A operacao requerida pela aplicacao nao e permitida.", vbCritical, "Integrador"
            Case -2146825042
                MsgBox "O Aplicativo nao pode fechar um objeto de Conexao no meio de uma transacao", vbCritical, "Integrador"
            Case -2146825287
                MsgBox "O Aplicativo esta usando argumentos de algum tipo incorreto, estao fora do limite permitido, ou em conflito com um outro.", vbCritical, "Integrador"
            Case -2146824579
                MsgBox "O Aplicativo solicitou uma operacao em um objeto com referencia a um objeto Connection que foi fechado ou e invalido.", vbCritical, "Integrador"
            Case -2146824580
                MsgBox "O Aplicativo definiu impropramente um objeto parametro.", vbCritical, "Integrador"
            Case -2146825023
                MsgBox "ADO nao encontrou o objeto na colecao correspondente ao nome ou referencia solicitada pelo aplicativo.", vbCritical, "Integrador"
           'Referente ao erro 3021
            Case -2146825267
                MsgBox "O registro corrente foi excluido,a operacao solicitada pelo aplicativo requer uma registro corrente.", vbCritical, "Integrador"
            Case -2146824573
                MsgBox "Operacao invalida no objeto enquanto nao e executado.", vbCritical, "Integrador"
            Case -2146824578
                MsgBox "Operacao invalida no objeto durente processamento do evento.", vbCritical, "Integrador"
            Case -2146824584
                MsgBox "A operacao solicitada pelo aplicativo nao e permitida se o objeto esta fechado.", vbCritical, "Integrador"
            Case -2146824921
                MsgBox "Objeto nao pode ser adicionado. O objeto ja esta na colecao.", vbCritical, "Integrador"
            Case -2146824868
                MsgBox "O objeto referenciado pelo aplicativo nao mais aponta para um objeto valido.", vbCritical, "Integrador"
            Case -2146824583
                MsgBox "A operacao requerida pela aplicacao nao e permitida se o objeto estiver aberto.", vbCritical, "Integrador"
            Case -2146825286
                MsgBox "Erro ocorrido durante tentativa de abertura de arquivo.", vbCritical, "Integrador"
            Case -2146824576
                MsgBox "Operacao cancelada pelo usuario.", vbCritical, "Integrador"
            Case -2146824582
                MsgBox "ADO nao pode encontrar o provedor especificado.", vbCritical, "Integrador"
            Case -2146824285
                MsgBox "Erro ocorrido durante tentativa de leitura de arquivo.", vbCritical, "Integrador"
            Case -2146824575
                MsgBox "Operacao invalida no objeto enquanto ainda estiver conectado.", vbCritical, "Integrador"
            Case -2146824577
                MsgBox "Operacao invalida no objeto enquanto processa um outro comando.", vbCritical, "Integrador"
            Case -2146824572
                MsgBox "A operacao solicitada pela aplicacao nao e segura para as configuracoes da maquina", vbCritical, "Integrador"
            Case -2146825284
                MsgBox "Erro ocorrido durante tentativa de gravacao de arquivo.", vbCritical, "Integrador"
            Case -2147217873
                MsgBox "Erro de integridade referêncial.Este registro não pode ser INCLUIDO/DELETADO.", vbCritical, "Integrador"
           'Erros Interceptaveis
            Case 3
                MsgBox "Return sem GoSub", vbCritical, "Integrador"
            Case 5
                MsgBox "Chamada de procedimento inválida", vbCritical, "Integrador"
            Case 6
                MsgBox "Sobrecarga", vbCritical, "Integrador"
            Case 7
                MsgBox "Memória insuficiente", vbCritical, "Integrador"
            Case 9
                MsgBox "Subscrito fora do intervalo", vbCritical, "Integrador"
            Case 10
                MsgBox "Esta matriz é fixa ou está temporariamente bloqueada", vbCritical, "Integrador"
            Case 11
                MsgBox "Divisão por zero", vbCritical, "Integrador"
            Case 13
                MsgBox "Tipo incompatível", vbCritical, "Integrador"
            Case 14
                MsgBox "Espaço insuficiente para seqüência de caracteres", vbCritical, "Integrador"
            Case 16
                MsgBox "Expressão muito complexa", vbCritical, "Integrador"
            Case 17
                MsgBox "Não é possível executar a operação solicitada", vbCritical, "Integrador"
            Case 18
                MsgBox "Ocorreu uma interrupção do usuário", vbCritical, "Integrador"
            Case 20
                MsgBox "Recomeçar sem erro", vbCritical, "Integrador"
            Case 28
                MsgBox "Espaço insuficiente para pilha", vbCritical, "Integrador"
            Case 35
                MsgBox "Sub, Function ou Property não definida", vbCritical, "Integrador"
            Case 47
                MsgBox "Número excessivo de clientes do aplicativo DLL", vbCritical, "Integrador"
            Case 48
                MsgBox "Erro ao carregar DLL", vbCritical, "Integrador"
            Case 49
                MsgBox "Convenção de chamada DLL inválida", vbCritical, "Integrador"
            Case 51
                MsgBox "erro interno", vbCritical, "Integrador"
            Case 52
                MsgBox "Nome ou número de arquivo inválido", vbCritical, "Integrador"
            Case 54
                MsgBox "Modo de arquivo inválido", vbCritical, "Integrador"
            Case 55
                MsgBox "O arquivo já está aberto", vbCritical, "Integrador"
            Case 57
                MsgBox "Erro de dispositivo de E/S", vbCritical, "Integrador"
            Case 58
                MsgBox "O arquivo já existe", vbCritical, "Integrador"
            Case 59
                MsgBox "Comprimento de registro inválido", vbCritical, "Integrador"
            Case 61
                MsgBox "disco cheio", vbCritical, "Integrador"
            Case 62
                MsgBox "Entrada depois do fim do arquivo", vbCritical, "Integrador"
            Case 63
                MsgBox "Número de registro inválido", vbCritical, "Integrador"
            Case 67
                MsgBox "Número excessivo de arquivos", vbCritical, "Integrador"
            Case 68
                MsgBox "Dispositivo não disponível", vbCritical, "Integrador"
            Case 70
                MsgBox "Permissão negada", vbCritical, "Integrador"
            Case 71
                MsgBox "O disco não está pronto", vbCritical, "Integrador"
            Case 74
                MsgBox "Não é possível renomear com unidade de disco diferente", vbCritical, "Integrador"
            Case 75
                MsgBox "Erro de acesso a caminho/arquivo", vbCritical, "Integrador"
            Case 76
                MsgBox "Caminho não encontrado", vbCritical, "Integrador"
            Case 92
                MsgBox "Loop ‘For’ não inicializado", vbCritical, "Integrador"
            Case 93
                MsgBox "Seqüência de caracteres padrão inválida", vbCritical, "Integrador"
            Case 94
                MsgBox "Uso de Null inválido", vbCritical, "Integrador"
            Case 97
                MsgBox "Não é possível chamar procedimento Friend para um objeto que não é uma instância da classe de definição", vbCritical, "Integrador"
            Case 298
                MsgBox "DLL do sistema não pode ser carregada", vbCritical, "Integrador"
            Case 320
                MsgBox "Não é possível utilizar nomes de dispositivos em nomes de arquivos específicos", vbCritical, "Integrador"
            Case 321
                MsgBox "Formato de arquivo inválido", vbCritical, "Integrador"
            Case 322
                MsgBox "Não é possível criar arquivo temporário necessário", vbCritical, "Integrador"
            Case 325
                MsgBox "Formato inválido no arquivo de recursos", vbCritical, "Integrador"
            Case 327
                MsgBox "Nome do valor de dados não encontrado", vbCritical, "Integrador"
            Case 328
                MsgBox "Parâmetro ilegal; não é possível gravar matrizes", vbCritical, "Integrador"
            Case 355
                MsgBox "Não foi possível acessar registro do sistema", vbCritical, "Integrador"
            Case 336
                MsgBox "Componente ActiveX não foi registrado corretamente", vbCritical, "Integrador"
            Case 337
                MsgBox "Componente ActiveX não foi encontrado", vbCritical, "Integrador"
            Case 338
                MsgBox "Componente ActiveX não executou corretamente", vbCritical, "Integrador"
            Case 360
                MsgBox "Objeto já carregado", vbCritical, "Integrador"
            Case 361
                MsgBox "Não é possível carregar ou descarregar este objeto", vbCritical, "Integrador"
            Case 363
                MsgBox "Controle ActiveX especificado não foi encontrado", vbCritical, "Integrador"
            Case 364
                MsgBox "Objeto foi descarregado", vbCritical, "Integrador"
            Case 365
                MsgBox "Não é possível carregar dentro desse contexto", vbCritical, "Integrador"
            Case 368
                MsgBox "O arquivo especificado está desatualizado. Este programa exige uma versão posterior", vbCritical, "Integrador"
            Case 371
                MsgBox "O objeto especificado não pode ser utilizado como um formulário proprietário de Show", vbCritical, "Integrador"
            Case 380
                MsgBox "Valor de propriedade inválido", vbCritical, "Integrador"
            Case 381
                MsgBox "Índice de matriz de propriedades inválido", vbCritical, "Integrador"
            Case 382
                MsgBox "Propriedade Set não pode ser executada em tempo de execução", vbCritical, "Integrador"
            Case 383
                MsgBox "Propriedade Set não pode ser utilizada com uma propriedade somente leitura", vbCritical, "Integrador"
            Case 385
                MsgBox "É necessário o índice de matriz de propriedade", vbCritical, "Integrador"
            Case 387
                MsgBox "Propriedade Set não permitida", vbCritical, "Integrador"
            Case 393
                MsgBox "Propriedade Get não pode ser executada em tempo de execução", vbCritical, "Integrador"
            Case 394
                MsgBox "Propriedade Get não pode ser executada em propriedade somente gravação", vbCritical, "Integrador"
            Case 400
                MsgBox "Formulário já exibido; impossível exibir de forma modal", vbCritical, "Integrador"
            Case 402
                MsgBox "Código deve fechar o formulário modal superior", vbCritical, "Integrador"
            Case 419
                MsgBox "Permissão para utilizar objeto negada", vbCritical, "Integrador"
            Case 422
                MsgBox "Propriedade não encontrada", vbCritical, "Integrador"
            Case 423
                MsgBox "Propriedade ou método não foi encontrado", vbCritical, "Integrador"
            Case 424
                MsgBox "Objeto é obrigatório", vbCritical, "Integrador"
            Case 425
                MsgBox "Uso inválido de objeto", vbCritical, "Integrador"
            Case 429
                'O programa precisa de um objeto que não está registrado ou não
                'existe no dico rígido. O módulo não inciará ou não funcionará
                'corretamente.
                MsgBox "O componente ActiveX não pode criar um objeto ou retornar referência a esse objeto", vbCritical, "Integrador"
                End
            Case 430
                MsgBox "A classe não aceita Automação", vbCritical, "Integrador"
            Case 432
                MsgBox "O nome do arquivo ou o nome da classe não foi encontrado durante a operação de Automação", vbCritical, "Integrador"
            Case 438
                MsgBox "O objeto não aceita esta propriedade ou método", vbCritical, "Integrador"
            Case 440
                MsgBox "Erro de automação", vbCritical, "Integrador"
            Case 442
                MsgBox "A conexão à biblioteca de objetos ou de tipos para processo remoto foi perdida", vbCritical, "Integrador"
            Case 443
                MsgBox "O objeto de Automação não possui um valor padrão", vbCritical, "Integrador"
            Case 445
                MsgBox "O objeto não suporta esta ação", vbCritical, "Integrador"
            Case 446
                MsgBox "O objeto não suporta argumentos nomeados", vbCritical, "Integrador"
            Case 447
                MsgBox "O objeto não aceita a definição atual de localidade", vbCritical, "Integrador"
            Case 448
                MsgBox "O argumento nomeado não foi encontrado", vbCritical, "Integrador"
            Case 449
                MsgBox "Argumento não opcional ou atribuição de propriedade inválida", vbCritical, "Integrador"
            Case 450
                MsgBox "Número de argumentos incorreto ou atribuição de propriedade inválida", vbCritical, "Integrador"
            Case 451
                MsgBox "Object não é uma coleção", vbCritical, "Integrador"
            Case 452
                MsgBox "Ordinal inválido", vbCritical, "Integrador"
            Case 453
                MsgBox "A função DLL especificada não foi encontrada", vbCritical, "Integrador"
            Case 454
                MsgBox "O recurso de código não foi encontrado", vbCritical, "Integrador"
            Case 455
                MsgBox "Erro de proteção de recurso de código", vbCritical, "Integrador"
            Case 457
                MsgBox "Esta tecla já está associada a um elemento desta coleção", vbCritical, "Integrador"
            Case 458
                MsgBox "A variável utiliza um tipo de automação não suportada no Visual Basic", vbCritical, "Integrador"
            Case 459
                MsgBox "Este componente não suporta eventos", vbCritical, "Integrador"
            Case 460
                MsgBox "Formato da área de transferência inválido", vbCritical, "Integrador"
            Case 461
                MsgBox "Formato especificado não corresponde ao formato dos dados", vbCritical, "Integrador"
            Case 480
                MsgBox "Não é possível criar imagem AutoRedraw", vbCritical, "Integrador"
            Case 481
                MsgBox "figura inválida", vbCritical, "Integrador"
            Case 482
                MsgBox "Erro na impressora", vbCritical, "Integrador"
            Case 483
                MsgBox "Driver da impressora não aceita a propriedade especificada", vbCritical, "Integrador"
            Case 484
                MsgBox "Problemas ao obter informações da impressora a partir do sistema. Certifique-se de que a impressora esteja instalada corretamente", vbCritical, "Integrador"
            Case 485
                MsgBox "Tipo de figura inválido", vbCritical, "Integrador"
            Case 486
                MsgBox "Não é possível imprimir imagem de formulário neste tipo de impressora", vbCritical, "Integrador"
            Case 520
                MsgBox "Não é possível esvaziar a Área de transferência", vbCritical, "Integrador"
            Case 521
                MsgBox "Não é possível abrir a Área de transferência", vbCritical, "Integrador"
            Case 735
                MsgBox "Não é possível salvar arquivo no diretório TEMP", vbCritical, "Integrador"
            Case 744
                MsgBox "Texto procurado não encontrado", vbCritical, "Integrador"
            Case 746
                MsgBox "Substituições muito longas", vbCritical, "Integrador"
            Case 31001
                MsgBox "Memória insuficiente", vbCritical, "Integrador"
            Case 31004
                MsgBox "Nenhum objeto", vbCritical, "Integrador"
            Case 31018
                MsgBox "Classe não está definida", vbCritical, "Integrador"
            Case 31027
                MsgBox "Não é possível ativar objeto", vbCritical, "Integrador"
            Case 31032
                MsgBox "Não foi possível criar objeto incorporado", vbCritical, "Integrador"
            Case 31036
                MsgBox "Erro ao salvar o arquivo", vbCritical, "Integrador"
            Case 31037
                MsgBox "Erro ao carregar do arquivo", vbCritical, "Integrador"

           'Erros interceptáveis do Microsoft Jet e do objeto de acesso a dados (DAO, Data Access Object)

            Case 2420
                MsgBox "Erro de sintaxe em número.", vbCritical, "Integrador"
            Case 2421
                MsgBox "Erro de sintaxe em data.", vbCritical, "Integrador"
            Case 2422
                MsgBox "Erro de sintaxe em seqüência.", vbCritical, "Integrador"
            Case 2423
                MsgBox "Utilização inválida de '.', '!' ou '()'.", vbCritical, "Integrador"
            Case 2424
                MsgBox "Nome desconhecido.", vbCritical, "Integrador"
            Case 2425
                MsgBox "Nome de função desconhecido.", vbCritical, "Integrador"
            Case 2426
                MsgBox "Função não-disponível em expressões.", vbCritical, "Integrador"
            Case 2427
                MsgBox "Objeto sem valor.", vbCritical, "Integrador"
            Case 2428
                MsgBox "Argumentos inválidos utilizados com a função de domínio.", vbCritical, "Integrador"
            Case 2429
                MsgBox "Operador In sem ().", vbCritical, "Integrador"
            Case 2430
                MsgBox "Operador Between sem And.", vbCritical, "Integrador"
            Case 2431
                MsgBox "Erro de sintaxe (operador ausente).", vbCritical, "Integrador"
            Case 2432
                MsgBox "Erro de sintaxe (vírgula).", vbCritical, "Integrador"
            Case 2433
                MsgBox "Erro de sintaxe.", vbCritical, "Integrador"
            Case 2434
                MsgBox "Erro de sintaxe (operador ausente).", vbCritical, "Integrador"
            Case 2435
                MsgBox ") extra.", vbCritical, "Integrador"
            Case 2436
                MsgBox "), ] ou item ausentes.", vbCritical, "Integrador"
            Case 2437
                MsgBox "Utilização inválida de barras verticais.", vbCritical, "Integrador"
            Case 2438
                MsgBox "Erro de sintaxe.", vbCritical, "Integrador"
            Case 2439
                MsgBox "Número incorreto de argumentos utilizados com a função.", vbCritical, "Integrador"
            Case 2440
                MsgBox "Função IIF sem ().", vbCritical, "Integrador"
            Case 2442
                MsgBox "Utilização inválida de parênteses.", vbCritical, "Integrador"
            Case 2443
                MsgBox "Utilização inválida do operador Is.", vbCritical, "Integrador"
            Case 2445
                MsgBox "Expressão muito complexa.", vbCritical, "Integrador"
            Case 2446
                MsgBox "Memória insuficiente durante o cálculo.", vbCritical, "Integrador"
            Case 2447
                MsgBox "Utilização inválida de '.', '!' ou '()'.", vbCritical, "Integrador"
            Case 2448
                MsgBox "Não é possível definir o valor.", vbCritical, "Integrador"
            Case 3000
                MsgBox "Erro <Item> reservado; não existe mensagem para este erro.", vbCritical, "Integrador"
            Case 3001
                MsgBox "Argumento inválido.", vbCritical, "Integrador"
            Case 3002
                MsgBox "Não foi possível iniciar a sessão.", vbCritical, "Integrador"
            Case 3003
                MsgBox "Não foi possível iniciar a transação; já existem muitas transações aninhadas.", vbCritical, "Integrador"
            Case 3005
                MsgBox "<Nome do banco de dados> não é um nome de banco de dados válido.", vbCritical, "Integrador"
            Case 3006
                MsgBox "O banco de dados <nome> está bloqueado exclusivamente.", vbCritical, "Integrador"
            Case 3007
                MsgBox "Não é possível abrir o banco de dados da biblioteca <nome>.", vbCritical, "Integrador"
            Case 3008
                MsgBox "A tabela <nome> já está aberta exclusivamente por outro usuário ou já está aberta através da interface do usuário e não pode ser manipulada programaticamente.", vbCritical, "Integrador"
            Case 3009
                MsgBox "Você tentou bloquear a tabela <tabela> enquanto a abria, mas ela não pode ser bloqueada porque está em uso no momento. Aguarde um instante e, em seguida, tente a operação novamente.", vbCritical, "Integrador"
            Case 3010
                MsgBox "A tabela <nome> já existe.", vbCritical, "Integrador"
            Case 3011
                MsgBox "O mecanismo de banco de dados Microsoft Jet não pôde encontrar o objeto <nome>. Certifique-se de que o objeto existe e que você digitou o seu nome e o nome do caminho corretamente.", vbCritical, "Integrador"
            Case 3012
                MsgBox "O objeto <nome> já existe.", vbCritical, "Integrador"
            Case 3013
                MsgBox "Não foi possível renomear o arquivo ISAM instalável.", vbCritical, "Integrador"
            Case 3014
                MsgBox "Não é possível abrir mais tabelas.", vbCritical, "Integrador"
            Case 3015
                MsgBox "<Nome do índice> não é um índice nesta tabela. Consulte a coleção Indexes do objeto TableDef para determinar os nomes de índice válidos.", vbCritical, "Integrador"
            Case 3016
                MsgBox "O campo não caberá no registro.", vbCritical, "Integrador"
            Case 3017
                MsgBox "O tamanho do campo é grande demais.", vbCritical, "Integrador"
            Case 3018
                MsgBox "Não foi possível encontrar o campo <nome>.", vbCritical, "Integrador"
            Case 3020
                MsgBox "Você tentou chamar Update ou CancelUpdate ou tentou atualizar um Field em um conjunto de registros sem chamar primeiro AddNew ou Edit.", vbCritical, "Integrador"
            Case 3022
                MsgBox "As alterações que você solicitou à tabela não foram bem-sucedidas porque criariam valores duplicados no índice, na chave primária ou na relação. Altere os dados no campo ou campos que contêm dados duplicados, remova o índice ou redefina-o para permitir entradas duplicadas e tente novamente.", vbCritical, "Integrador"
            Case 3023
                MsgBox "AddNew ou Edit já utilizado.", vbCritical, "Integrador"
            Case 3024
                MsgBox "Não foi possível encontrar <nome>.", vbCritical, "Integrador"
            Case 3025
                MsgBox "Não é possível abrir mais arquivos.", vbCritical, "Integrador"
            Case 3026
                MsgBox "Espaço insuficiente em disco.", vbCritical, "Integrador"
            Case 3027
                MsgBox "Não foi possível atualizar. O banco de dados ou objeto é somente leitura.", vbCritical, "Integrador"
            Case 3028
                MsgBox "Não é possível iniciar seu aplicativo. O arquivo de informações do grupo de trabalho está ausente ou aberto exclusivamente por outro usuário.", vbCritical, "Integrador"
            Case 3029
                MsgBox "Nome de conta ou senha inválidos.", vbCritical, "Integrador"
            Case 3030
                MsgBox "<Nome da conta> não é um nome de conta válido.", vbCritical, "Integrador"
            Case 3031
                MsgBox "Senha inválida.", vbCritical, "Integrador"
            Case 3032
                MsgBox "Não é possível executar esta operação.", vbCritical, "Integrador"
            Case 3033
                MsgBox "Você não tem as permissões necessárias para utilizar o objeto <nome>. Faça o seu administrador do sistema ou a pessoa que criou este objeto estabelecer as permissões apropriadas para você.", vbCritical, "Integrador"
            Case 3034
                MsgBox "Você tentou aceitar ou cancelar uma transação sem utilizar primeiro BeginTrans.", vbCritical, "Integrador"
            Case 3036
                MsgBox "O banco de dados alcançou o tamanho máximo.", vbCritical, "Integrador"
            Case 3037
                MsgBox "Não é possível abrir mais tabelas ou consultas.", vbCritical, "Integrador"
            Case 3039
                MsgBox "Não foi possível criar o índice; muitos índices já definidos.", vbCritical, "Integrador"
            Case 3040
                MsgBox "Erro de E/S em disco durante a leitura.", vbCritical, "Integrador"
            Case 3041
                MsgBox "Não é possível abrir um banco de dados criado com uma versão anterior do seu aplicativo.", vbCritical, "Integrador"
            Case 3042
                MsgBox "Sem identificadores de arquivo do MS-DOS.", vbCritical, "Integrador"
            Case 3043
                MsgBox "Erro de disco ou rede.", vbCritical, "Integrador"
            Case 3044
                MsgBox "<Caminho> não é um caminho válido. Certifique-se de que o nome do caminho está digitado corretamente e que você está conectado ao servidor no qual se encontra o arquivo.", vbCritical, "Integrador"
            Case 3045
                MsgBox "Não foi possível utilizar <nome>; o arquivo já está em utilização.", vbCritical, "Integrador"
            Case 3046
                MsgBox "Não foi possível salvar; atualmente bloqueado por outro usuário.", vbCritical, "Integrador"
            Case 3047
                MsgBox "O registro é grande demais.", vbCritical, "Integrador"
            Case 3048
                MsgBox "Não é possível abrir mais bancos de dados.", vbCritical, "Integrador"
            Case 3049
                MsgBox "Não é possível abrir o banco de dados <nome>. Ele pode não ser um banco de dados que o seu aplicativo reconheça ou o arquivo pode estar corrompido.", vbCritical, "Integrador"
            Case 3051
                MsgBox "O mecanismo de banco de dados Microsoft Jet não pode abrir o arquivo <nome>. Ele já está aberto exclusivamente por outro usuário ou você precisa de permissão para visualizar seus dados.", vbCritical, "Integrador"
            Case 3052
                MsgBox "O número de bloqueios de compartilhamento de arquivos do MS-DOS foi excedido. Você precisa aumentar o número de bloqueios instalados com Share.exe.", vbCritical, "Integrador"
            Case 3053
                MsgBox "Tarefas cliente em excesso.", vbCritical, "Integrador"
            Case 3054
                MsgBox "Campos Memorando ou ‘Objeto OLE’ em excesso.", vbCritical, "Integrador"
            Case 3055
                MsgBox "Nome de campo inválido.", vbCritical, "Integrador"
            Case 3056
                MsgBox "Não foi possível reparar este banco de dados.", vbCritical, "Integrador"
            Case 3057
                MsgBox "Operação não suportada em tabelas vinculadas.", vbCritical, "Integrador"
            Case 3058
                MsgBox "O índice ou chave primária não pode conter um valor Null.", vbCritical, "Integrador"
            Case 3059
                MsgBox "Operação cancelada pelo usuário.", vbCritical, "Integrador"
            Case 3060
                MsgBox "Tipo de dados incorreto para o parâmetro <parâmetro>.", vbCritical, "Integrador"
            Case 3061
                MsgBox "Muito poucos parâmetros. Eram esperados <número>.", vbCritical, "Integrador"
            Case 3062
                MsgBox "Alias de saída <nome> duplicado.", vbCritical, "Integrador"
            Case 3063
                MsgBox "Destino de saída <nome> duplicado.", vbCritical, "Integrador"
            Case 3064
                MsgBox "Não é possível abrir a consulta ação <nome>.", vbCritical, "Integrador"
            Case 3065
                MsgBox "Não é possível executar uma consulta seleção.", vbCritical, "Integrador"
            Case 3066
                MsgBox "A consulta deve ter pelo menos um campo de destino.", vbCritical, "Integrador"
            Case 3067
                MsgBox "A entrada da consulta deve conter pelo menos uma tabela ou consulta.", vbCritical, "Integrador"
            Case 3068
                MsgBox "Nome de alias inválido.", vbCritical, "Integrador"
            Case 3069
                MsgBox "A consulta ação <nome> não pode ser utilizada como origem da linha.", vbCritical, "Integrador"
            Case 3070
                MsgBox "O mecanismo de banco de dados Microsoft Jet não reconhece <nome> como um nome de campo ou expressão válida.", vbCritical, "Integrador"
            Case 3071
                MsgBox "Esta expressão está digitada incorretamente ou é complexa demais para ser avaliada. Por exemplo, uma expressão numérica pode conter muitos elementos complicados. Tente simplificar a expressão atribuindo partes da expressão a variáveis.", vbCritical, "Integrador"
            Case 3073
                MsgBox "A operação deve utilizar uma consulta atualizável.", vbCritical, "Integrador"
            Case 3074
                MsgBox "Não é possível repetir o nome da tabela <nome> na cláusula FROM.", vbCritical, "Integrador"
            Case 3075
                MsgBox "<Mensagem> na expressão de consulta <expressão>.", vbCritical, "Integrador"
            Case 3076
                MsgBox "<Nome> na expressão de critério.", vbCritical, "Integrador"
            Case 3077
                MsgBox "<Mensagem> na expressão.", vbCritical, "Integrador"
            Case 3078
                MsgBox "O mecanismo de banco de dados Microsoft Jet não consegue encontrar a tabela de entrada ou a consulta <nome>. Certifique-se de que ela existe e que o seu nome está digitado corretamente.", vbCritical, "Integrador"
            Case 3079
                MsgBox "O campo especificado <campo> poderia se referir a mais de uma tabela listada na cláusula FROM da sua instrução SQL.", vbCritical, "Integrador"
            Case 3080
                MsgBox "A tabela associada <nome> não está listada na cláusula FROM.", vbCritical, "Integrador"
            Case 3081
                MsgBox "Não é possível associar mais de uma tabela com o mesmo nome <nome>.", vbCritical, "Integrador"
            Case 3082
                MsgBox "A operação JOIN <operação> refere-se a um campo que não está em uma das tabelas associadas.", vbCritical, "Integrador"
            Case 3083
                MsgBox "Não é possível utilizar consulta de relatório interno.", vbCritical, "Integrador"
            Case 3084
                MsgBox "Não é possível inserir dados com a consulta ação.", vbCritical, "Integrador"
            Case 3085
                MsgBox "Função <nome> indefinida na expressão.", vbCritical, "Integrador"
            Case 3086
                MsgBox "Não foi possível excluir das tabelas especificadas.", vbCritical, "Integrador"
            Case 3087
                MsgBox "Expressões em excesso na cláusula GROUP BY.", vbCritical, "Integrador"
            Case 3088
                MsgBox "Expressões em excesso na cláusula ORDER BY.", vbCritical, "Integrador"
            Case 3089
                MsgBox "Expressões em excesso na saída DISTINCT.", vbCritical, "Integrador"
            Case 3090
                MsgBox "A tabela resultante não pode ter mais de um campo AutoNumeração.", vbCritical, "Integrador"
            Case 3092
                MsgBox "Não é possível utilizar a cláusula HAVING na instrução TRANSFORM.", vbCritical, "Integrador"
            Case 3093
                MsgBox "A cláusula ORDER BY <cláusula> entra em conflito com DISTINCT.", vbCritical, "Integrador"
            Case 3094
                MsgBox "A cláusula ORDER BY <cláusula> entra em conflito com a cláusula GROUP BY.", vbCritical, "Integrador"
            Case 3095
                MsgBox "Não é possível ter uma função agregada na expressão <expressão>.", vbCritical, "Integrador"
            Case 3096
                MsgBox "Não é possível ter uma função agregada na cláusula WHERE <cláusula>.", vbCritical, "Integrador"
            Case 3097
                MsgBox "Não é possível ter uma função agregada na cláusula ORDER BY <cláusula>.", vbCritical, "Integrador"
            Case 3098
                MsgBox "Não é possível ter uma função agregada na cláusula GROUP BY <cláusula>.", vbCritical, "Integrador"
            Case 3099
                MsgBox "Não é possível ter uma função agregada na operação JOIN <operação>.", vbCritical, "Integrador"
            Case 3100
                MsgBox "Não é possível definir o campo <nome> na chave de associação como Null.", vbCritical, "Integrador"
            Case 3101
                MsgBox "O mecanismo de banco de dados Microsoft Jet não consegue encontrar um registro na tabela <nome> com campo(s) <nome> de correspondência de chave.", vbCritical, "Integrador"
            Case 3102
                MsgBox "Referência circular causada pela <referência da consulta>.", vbCritical, "Integrador"
            Case 3103
                MsgBox "Referência circular causada pelo alias <nome> na lista SELECT da definição da consulta.", vbCritical, "Integrador"
            Case 3104
                MsgBox "Não é possível especificar mais de uma vez o <valor> do título de colunas fixas em uma consulta de tabela de referência cruzada.", vbCritical, "Integrador"
            Case 3105
                MsgBox "Nome do campo de destino ausente na instrução SELECT INTO <instrução>.", vbCritical, "Integrador"
            Case 3106
                MsgBox "Nome do campo de destino ausente na instrução UPDATE <instrução>.", vbCritical, "Integrador"
            Case 3107
                MsgBox "Registro(s) não pode(m) ser adicionado(s); sem permissão de inserção no <nome>.", vbCritical, "Integrador"
            Case 3108
                MsgBox "Registro(s) não pode(m) ser editado(s); sem permissão de atualização em <nome>.", vbCritical, "Integrador"
            Case 3109
                MsgBox "Registro(s) não pode(m) ser excluídos, sem permissão de exclusão em <nome>.", vbCritical, "Integrador"
            Case 3110
                MsgBox "Não foi possível ler definições; sem permissão de leitura de definições da tabela ou consulta <nome>.", vbCritical, "Integrador"
            Case 3111
                MsgBox "Não foi possível criar; sem permissão de modificação da estrutura da tabela ou consulta <nome>.", vbCritical, "Integrador"
            Case 3112
                MsgBox "Registro(s) não pode(m) ser lido(s); sem permissão de leitura em <nome>.", vbCritical, "Integrador"
            Case 3113
                MsgBox "Não é possível atualizar <nome do campo>; campo não atualizável.", vbCritical, "Integrador"
            Case 3114
                MsgBox "Não é possível incluir Memorando ou Objeto OLE quando forem selecionados valores exclusivos <instrução>.", vbCritical, "Integrador"
            Case 3115
                MsgBox "Não é possível ter campos Memorando ou Objeto OLE no argumento agregado <instrução>.", vbCritical, "Integrador"
            Case 3116
                MsgBox "Não é possível ter campos Memorando ou Objeto OLE no critério <critério> para a função agregada.", vbCritical, "Integrador"
            Case 3117
                MsgBox "Não é possível classificar em Memorando ou Objeto OLE <cláusula>.", vbCritical, "Integrador"
            Case 3118
                MsgBox "Não é possível associar em Memorando ou Objeto OLE <nome>.", vbCritical, "Integrador"
            Case 3119
                MsgBox "Não é possível agrupar em Memorando ou Objeto OLE <cláusula>.", vbCritical, "Integrador"
            Case 3120
                MsgBox "Não é possível agrupar em campos selecionados com '*' <nome da tabela>.", vbCritical, "Integrador"
            Case 3121
                MsgBox "Não é possível agrupar em campos selecionados com '*'.", vbCritical, "Integrador"
            Case 3122
                MsgBox "Você tentou executar uma consulta que não inclui a expressão <nome> especificada como parte de uma função agregada.", vbCritical, "Integrador"
            Case 3123
                MsgBox "Não é possível utilizar '*' em consulta de tabela de referência cruzada.", vbCritical, "Integrador"
            Case 3124
                MsgBox "Não é possível obter a entrada pela consulta de relatório interno <nome>.", vbCritical, "Integrador"
            Case 3125
                MsgBox "O mecanismo de banco de dados não consegue encontrar <nome>. Certifique-se de que é um nome de parâmetro ou alias válido, que não inclui caracteres nem pontuação inválida e que o nome não é grande demais.", vbCritical, "Integrador"
            Case 3126
                MsgBox "Colchetes inválidos no nome <nome>.", vbCritical, "Integrador"
            Case 3127
                MsgBox "A instrução INSERT INTO contém o seguinte nome de campo desconhecido: <nome do campo>. Certifique-se de que você digitou o nome corretamente e tente a operação novamente.", vbCritical, "Integrador"
            Case 3128
                MsgBox "Especifique a tabela que contém os registros que deseja excluir.", vbCritical, "Integrador"
            Case 3129
                MsgBox "Instrução SQL inválida; era esperado 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT' ou 'UPDATE'.", vbCritical, "Integrador"
            Case 3130
                MsgBox "Erro de sintaxe na instrução DELETE.", vbCritical, "Integrador"
            Case 3131
                MsgBox "Erro de sintaxe na cláusula FROM.", vbCritical, "Integrador"
            Case 3132
                MsgBox "Erro de sintaxe na cláusula GROUP BY.", vbCritical, "Integrador"
            Case 3133
                MsgBox "Erro de sintaxe na cláusula HAVING.", vbCritical, "Integrador"
            Case 3134
                MsgBox "Erro de sintaxe na instrução INSERT INTO.", vbCritical, "Integrador"
            Case 3135
                MsgBox "Erro de sintaxe na operação JOIN.", vbCritical, "Integrador"
            Case 3136
                MsgBox "A cláusula LEVEL inclui uma palavra ou argumento reservado que está digitado incorretamente ou está ausente, ou a pontuação está incorreta.", vbCritical, "Integrador"
            Case 3138
                MsgBox "Erro de sintaxe na cláusula ORDER BY.", vbCritical, "Integrador"
            Case 3139
                MsgBox "Erro de sintaxe na cláusula PARAMETER.", vbCritical, "Integrador"
            Case 3140
                MsgBox "Erro de sintaxe na cláusula PROCEDURE.", vbCritical, "Integrador"
            Case 3141
                MsgBox "A instrução SELECT inclui uma palavra ou argumento reservado ou um nome de argumento que está digitado incorretamente ou está ausente, ou a pontuação está incorreta.", vbCritical, "Integrador"
            Case 3143
                MsgBox "Erro de sintaxe na instrução TRANSFORM.", vbCritical, "Integrador"
            Case 3144
                MsgBox "Erro de sintaxe na instrução UPDATE.", vbCritical, "Integrador"
            Case 3145
                MsgBox "Erro de sintaxe na cláusula WHERE.", vbCritical, "Integrador"
            Case 3146
                MsgBox "ODBC – a chamada falhou.", vbCritical, "Integrador"
            Case 3151
                MsgBox "ODBC – a conexão a <nome> falhou.", vbCritical, "Integrador"
            Case 3154
                MsgBox "ODBC – não foi possível encontrar DLL <nome>.", vbCritical, "Integrador"
            Case 3155
                MsgBox "ODBC – a inserção em uma tabela vinculada <tabela> falhou.", vbCritical, "Integrador"
            Case 3156
                MsgBox "ODBC – a exclusão em uma tabela vinculada <tabela> falhou.", vbCritical, "Integrador"
            Case 3157
                MsgBox "ODBC – a atualização em uma tabela vinculada <tabela> falhou.", vbCritical, "Integrador"
            Case 3158
                MsgBox "Não foi possível salvar o registro; bloqueado no momento por outro usuário.", vbCritical, "Integrador"
            Case 3159
                MsgBox "Indicador inválido.", vbCritical, "Integrador"
            Case 3160
                MsgBox "A tabela não está aberta.", vbCritical, "Integrador"
            Case 3161
                MsgBox "Não foi possível descriptografar o arquivo.", vbCritical, "Integrador"
            Case 3162
                MsgBox "Você tentou atribuir o valor Null a uma variável que não é um tipo de dados Variant.", vbCritical, "Integrador"
            Case 3163
                MsgBox "O campo é pequeno demais para aceitar a quantidade de dados que você tentou adicionar. Tente inserir ou colar menos dados.", vbCritical, "Integrador"
            Case 3164
                MsgBox "O campo não pode ser atualizado porque outro usuário ou processo bloqueou o registro ou tabela correspondente.", vbCritical, "Integrador"
            Case 3165
                MsgBox "Não foi possível abrir o arquivo .inf.", vbCritical, "Integrador"
            Case 3166
                MsgBox "Não é possível localizar o arquivo de memorando Xbase solicitado.", vbCritical, "Integrador"
            Case 3167
                MsgBox "Registro excluído.", vbCritical, "Integrador"
            Case 3168
                MsgBox "Arquivo .inf inválido.", vbCritical, "Integrador"
            Case 3169
                MsgBox "O mecanismo de banco de dados Microsoft Jet não pôde executar a instrução SQL porque ela contém um campo que possui um tipo de dados inválido.", vbCritical, "Integrador"
            Case 3170
                MsgBox "Não foi possível encontrar o ISAM instalável.", vbCritical, "Integrador"
            Case 3171
                MsgBox "Não foi possível encontrar o caminho da rede ou o nome de usuário.", vbCritical, "Integrador"
            Case 3172
                MsgBox "Não foi possível abrir o Paradox.net.", vbCritical, "Integrador"
            Case 3173
                MsgBox "Não foi possível abrir a tabela 'MSysAccounts' no arquivo de informações do grupo de trabalho.", vbCritical, "Integrador"
            Case 3174
                MsgBox "Não foi possível abrir a tabela 'MSysGroups' no arquivo de informações do grupo de trabalho.", vbCritical, "Integrador"
            Case 3175
                MsgBox "A data está fora do intervalo ou está em um formato inválido.", vbCritical, "Integrador"
            Case 3176
                MsgBox "Não foi possível abrir o arquivo <nome>.", vbCritical, "Integrador"
            Case 3177
                MsgBox "Nome de tabela inválido.", vbCritical, "Integrador"
            Case 3179
                MsgBox "Encontrado fim de arquivo inesperado.", vbCritical, "Integrador"
            Case 3180
                MsgBox "Não foi possível gravar no arquivo <nome>.", vbCritical, "Integrador"
            Case 3181
                MsgBox "Intervalo inválido.", vbCritical, "Integrador"
            Case 3182
                MsgBox "Formato de arquivo inválido.", vbCritical, "Integrador"
            Case 3183
                MsgBox "Espaço insuficiente no disco temporário.", vbCritical, "Integrador"
            Case 3184
                MsgBox "Não foi possível executar a consulta; não foi possível encontrar a tabela vinculada.", vbCritical, "Integrador"
            Case 3185
                MsgBox "SELECT INTO em um banco de dados remoto tentou produzir campos demais.", vbCritical, "Integrador"
            Case 3186
                MsgBox "SELECT INTO em um banco de dados remoto tentou produzir campos demais.", vbCritical, "Integrador"
            Case 3187
                MsgBox "Não foi possível ler; atualmente bloqueado pelo usuário <nome> na máquina <nome>.", vbCritical, "Integrador"
            Case 3188
                MsgBox "Não foi possível atualizar; atualmente bloqueado por outra sessão nesta máquina.", vbCritical, "Integrador"
            Case 3189
                MsgBox "Tabela <nome> é bloqueada exclusivamente pelo usuário <nome> na máquina <nome>.", vbCritical, "Integrador"
            Case 3190
                MsgBox "Definidos campos em excesso.", vbCritical, "Integrador"
            Case 3191
                MsgBox "Não é possível definir o campo mais de uma vez.", vbCritical, "Integrador"
            Case 3192
                MsgBox "Não foi possível encontrar a tabela de saída <nome>.", vbCritical, "Integrador"
            Case 3196
                MsgBox "O banco de dados <nome do banco de dados> já está em uso por outra pessoa ou processo. Quando o banco de dados estiver disponível, tente a operação novamente.", vbCritical, "Integrador"
            Case 3197
                MsgBox "O mecanismo de banco de dados Microsoft Jet parou o processo porque você e outro usuário estão tentando alterar os mesmos dados ao mesmo tempo.", vbCritical, "Integrador"
            Case 3198
                MsgBox "Não foi possível iniciar a sessão. Já existem sessões em excesso ativas.", vbCritical, "Integrador"
            Case 3199
                MsgBox "Não foi possível encontrar referência.", vbCritical, "Integrador"
            Case 3200
                MsgBox "O registro não pode ser excluído nem alterado porque a tabela <nome> inclui registros relacionados.", vbCritical, "Integrador"
            Case 3201
                MsgBox "Você não pode adicionar nem alterar um registro porque um registro relacionado é requerido na tabela <nome>.", vbCritical, "Integrador"
            Case 3202
                MsgBox "Não foi possível salvar; atualmente bloqueado por outro usuário.", vbCritical, "Integrador"
            Case 3203
                MsgBox "Subconsultas não podem ser utilizadas na expressão <expressão>.", vbCritical, "Integrador"
            Case 3204
                MsgBox "O banco de dados já existe.", vbCritical, "Integrador"
            Case 3205
                MsgBox "Títulos de coluna da tabela de referência cruzada <valor> em excesso.", vbCritical, "Integrador"
            Case 3206
                MsgBox "Não é possível criar uma relação entre um campo e ele mesmo.", vbCritical, "Integrador"
            Case 3207
                MsgBox "Operação não suportada em uma tabela do Paradox sem chave primária.", vbCritical, "Integrador"
            Case 3208
                MsgBox "Configuração Deleted inválida na chave Xbase do Registro do Windows.", vbCritical, "Integrador"
            Case 3210
                MsgBox "A seqüência de conexão é longa demais.", vbCritical, "Integrador"
            Case 3211
                MsgBox "O mecanismo de banco de dados não pôde bloquear a tabela <nome> porque ela já está em uso por outra pessoa ou processo.", vbCritical, "Integrador"
            Case 3212
                MsgBox "Não foi possível bloquear a tabela <nome>; atualmente em uso pelo usuário <nome> na máquina <nome>.", vbCritical, "Integrador"
            Case 3213
                MsgBox "Configuração Date inválida na chave Xbase do Registro do Windows.", vbCritical, "Integrador"
            Case 3214
                MsgBox "Configuração Mark inválida na chave Xbase do Registro do Windows.", vbCritical, "Integrador"
            Case 3215
                MsgBox "Tarefas Btrieve em excesso.", vbCritical, "Integrador"
            Case 3216
                MsgBox "Parâmetro <nome> especificado onde é requerido um nome de tabela.", vbCritical, "Integrador"
            Case 3217
                MsgBox "Parâmetro <nome> especificado onde é requerido um nome de banco de dados.", vbCritical, "Integrador"
            Case 3218
                MsgBox "Não foi possível atualizar; atualmente bloqueado.", vbCritical, "Integrador"
             Case 3219
                MsgBox "Operação inválida.", vbCritical, "Integrador"
             Case 3220
                MsgBox "Seqüência de agrupamento incorreta.", vbCritical, "Integrador"
             Case 3221
                MsgBox "Configurações inválidas na chave Btrieve do Registro do Windows.", vbCritical, "Integrador"
             Case 3222
                MsgBox "A consulta não pode conter um parâmetro Database.", vbCritical, "Integrador"
             Case 3223
                MsgBox "<Nome do parâmetro> é inválido porque é longo demais ou contém caracteres inválidos.", vbCritical, "Integrador"
             Case 3224
                MsgBox "Não é possível ler o dicionário de dados do Btrieve.", vbCritical, "Integrador"
             Case 3225
                MsgBox "Encontrado um conflito de proteção de registro durante a execução de uma operação Btrieve.", vbCritical, "Integrador"
             Case 3226
                MsgBox "Erros encontrados durante a utilização da DLL do Btrieve.", vbCritical, "Integrador"
             Case 3227
                MsgBox "Configuração Century inválida na chave Xbase do Registro do Windows.", vbCritical, "Integrador"
             Case 3228
                MsgBox "Configuração CollatingSequence inválida na chave Paradox do Registro do Windows.", vbCritical, "Integrador"
             Case 3229
                MsgBox "Btrieve – não foi possível alterar o campo.", vbCritical, "Integrador"
             Case 3230
                MsgBox "Arquivo de proteção do Paradox desatualizado.", vbCritical, "Integrador"
             Case 3231
                MsgBox "ODBC – o campo ficaria longo demais; dados truncados.", vbCritical, "Integrador"
             Case 3232
                MsgBox "ODBC – não pôde criar tabela.", vbCritical, "Integrador"
             Case 3234
                MsgBox "ODBC – o tempo limite de consulta remota expirou.", vbCritical, "Integrador"
             Case 3235
                MsgBox "ODBC – tipo de dados não suportado no servidor.", vbCritical, "Integrador"
             Case 3238
                MsgBox "ODBC – dados fora do intervalo.", vbCritical, "Integrador"
             Case 3239
                MsgBox "Usuários ativos em excesso.", vbCritical, "Integrador"
             Case 3240
                MsgBox "Btrieve – mecanismo Btrieve ausente.", vbCritical, "Integrador"
             Case 3241
                MsgBox "Btrieve – sem recursos.", vbCritical, "Integrador"
             Case 3242
                MsgBox "Referência inválida na instrução SELECT.", vbCritical, "Integrador"
             Case 3243
                MsgBox "Nenhum dos nomes de campo de importação corresponde aos campos na tabela acrescentada.", vbCritical, "Integrador"
             Case 3244
                MsgBox "Não é possível importar planilha protegida por senha.", vbCritical, "Integrador"
             Case 3245
                MsgBox "Não foi possível analisar os nomes de campo da primeira linha da tabela de importação.", vbCritical, "Integrador"
             Case 3246
                MsgBox "Operação não suportada em transações.", vbCritical, "Integrador"
             Case 3247
                MsgBox "ODBC – a definição da tabela vinculada mudou.", vbCritical, "Integrador"
             Case 3248
                MsgBox "Configuração NetworkAccess inválida no Registro do Windows.", vbCritical, "Integrador"
             Case 3249
                MsgBox "Configuração PageTimeout inválida no Registro do Windows.", vbCritical, "Integrador"
             Case 3250
                MsgBox "Não foi possível construir chave.", vbCritical, "Integrador"
             Case 3251
                MsgBox "A operação não é suportada para esse tipo de objeto.", vbCritical, "Integrador"
             Case 3252
                MsgBox "Não é possível abrir um formulário cuja consulta base contém uma função definida pelo usuário que tenta definir ou obter a propriedade Recordsetclose do formulário.", vbCritical, "Integrador"
             Case 3254
                MsgBox "ODBC – Não é possível bloquear todos os registros.", vbCritical, "Integrador"
             Case 3256
                MsgBox "Arquivo de índice não encontrado.", vbCritical, "Integrador"
             Case 3257
                MsgBox "Erro de sintaxe na declaração WITH OWNERACCESS OPTION.", vbCritical, "Integrador"
             Case 3258
                MsgBox "A instrução SQL não poderia ser executada porque contém associações externas ambíguas. Para forçar uma das associações a ser executada primeiro, crie uma consulta separada que execute a primeira associação e, em seguida, inclua essa consulta na sua instrução SQL.", vbCritical, "Integrador"
             Case 3259
                MsgBox "Tipo de dados de campo inválido.", vbCritical, "Integrador"
             Case 3260
                MsgBox "Não foi possível atualizar; atualmente bloqueado pelo usuário <nome> na máquina <nome>.", vbCritical, "Integrador"
             Case 3261
                MsgBox "A tabela <nome> é bloqueada exclusivamente pelo usuário <nome> na máquina <nome>.", vbCritical, "Integrador"
             Case 3262
                MsgBox "Não foi possível bloquear a tabela <nome>; atualmente em uso pelo usuário <nome> na máquina <nome>.", vbCritical, "Integrador"
             Case 3264
                MsgBox "Sem campo definido – não é possível acrescentar TableDef nem Index.", vbCritical, "Integrador"
             Case 3265
                MsgBox "Item não encontrado nesta coleção.", vbCritical, "Integrador"
             Case 3266
                MsgBox "Não é possível acrescentar um Field que já faça parte de uma coleção Fields.", vbCritical, "Integrador"
             Case 3267
                MsgBox "A propriedade somente pode ser definida quando o Field faz parte da coleção Fields de um objeto Recordset.", vbCritical, "Integrador"
             Case 3268
                MsgBox "Não é possível definir esta propriedade uma vez que o objeto faz parte de uma coleção.", vbCritical, "Integrador"
             Case 3269
                MsgBox "Não é possível acrescentar um Index que já faça parte de uma coleção Indexes.", vbCritical, "Integrador"
             Case 3270
                MsgBox "Propriedade não encontrada.", vbCritical, "Integrador"
             Case 3271
                MsgBox "Valor de propriedade inválido.", vbCritical, "Integrador"
             Case 3272
                MsgBox "O objeto não é uma coleção.", vbCritical, "Integrador"
             Case 3273
                MsgBox "Método não aplicável a este objeto.", vbCritical, "Integrador"
             Case 3274
                MsgBox "A tabela externa não está no formato esperado.", vbCritical, "Integrador"
             Case 3275
                MsgBox "Erro inesperado do driver de banco de dados externo <número do erro>.", vbCritical, "Integrador"
             Case 3276
                MsgBox "Referência inválida a objeto de banco de dados.", vbCritical, "Integrador"
             Case 3277
                MsgBox "Não é possível ter mais de 10 campos em um índice.", vbCritical, "Integrador"
             Case 3278
                MsgBox "O mecanismo de banco de dados Microsoft Jet não foi inicializado.", vbCritical, "Integrador"
             Case 3279
                MsgBox "O mecanismo de banco de dados Microsoft Jet já foi inicializado.", vbCritical, "Integrador"
             Case 3280
                MsgBox "Não é possível excluir um campo que faça parte de um índice ou que seja necessário ao sistema.", vbCritical, "Integrador"
             Case 3281
                MsgBox "Não é possível excluir este índice ou tabela. É o índice atual ou é utilizado em uma relação.", vbCritical, "Integrador"
            Case 3282
                MsgBox "Operação não suportada em uma tabela que contém dados.", vbCritical, "Integrador"
            Case 3283
                MsgBox "Já existe chave primária.", vbCritical, "Integrador"
            Case 3284
                MsgBox "Já existe índice.", vbCritical, "Integrador"
            Case 3285
                MsgBox "Definição de índice inválida.", vbCritical, "Integrador"
            Case 3286
                MsgBox "O formato do arquivo de memorando não corresponde ao formato do banco de dados externo especificado.", vbCritical, "Integrador"
            Case 3287
                MsgBox "Não é possível criar o índice no campo fornecido.", vbCritical, "Integrador"
            Case 3288
                MsgBox "O índice do Paradox não é primário.", vbCritical, "Integrador"
            Case 3289
                MsgBox "Erro de sintaxe na cláusula CONSTRAINT.", vbCritical, "Integrador"
            Case 3290
                MsgBox "Erro de sintaxe na instrução CREATE TABLE.", vbCritical, "Integrador"
            Case 3291
                MsgBox "Erro de sintaxe na instrução CREATE INDEX.", vbCritical, "Integrador"
            Case 3292
                MsgBox "Erro de sintaxe na definição do campo.", vbCritical, "Integrador"
            Case 3293
                MsgBox "Erro de sintaxe na instrução ALTER TABLE.", vbCritical, "Integrador"
            Case 3294
                MsgBox "Erro de sintaxe na instrução DROP INDEX.", vbCritical, "Integrador"
            Case 3295
                MsgBox "Erro de sintaxe em DROP TABLE ou DROP INDEX.", vbCritical, "Integrador"
            Case 3296
                MsgBox "Expressão de associação não-suportada.", vbCritical, "Integrador"
            Case 3297
                MsgBox "Não é possível importar tabela nem consulta. Nenhum registro encontrado ou todos os registros contêm erros.", vbCritical, "Integrador"
            Case 3298
                MsgBox "Há diversas tabelas com este nome. Especifique o proprietário no formato ‘proprietário.tabela’.", vbCritical, "Integrador"
            Case 3299
                MsgBox "Erro de conformidade com a especificação ODBC <mensagem>. Relate este erro ao profissional da área de desenvolvimento do seu aplicativo.", vbCritical, "Integrador"
            Case 3300
                MsgBox "Não é possível criar uma relação.", vbCritical, "Integrador"
            Case 3301
                MsgBox "Não é possível executar esta operação; os recursos desta versão não estão disponíveis em bancos de dados com formatos mais antigos.", vbCritical, "Integrador"
            Case 3302
                MsgBox "Não é possível alterar um regra enquanto as regras desta tabela estiverem em uso.", vbCritical, "Integrador"
            Case 3303
                MsgBox "Não é possível excluir este campo. Ele faz parte de uma ou mais relações.", vbCritical, "Integrador"
            Case 3304
                MsgBox "Você deve inserir um identificador pessoal (PID) que consista em no mínimo 4 e no máximo 20 caracteres e dígitos.", vbCritical, "Integrador"
            Case 3305
                MsgBox "Seqüência de conexão inválida na consulta passagem.", vbCritical, "Integrador"
            Case 3306
                MsgBox "Você gravou uma subconsulta que pode retornar mais de um campo sem utilizar a palavra reservada EXISTS na cláusula FROM da consulta principal. Altere a instrução SELECT da subconsulta para solicitar somente um campo.", vbCritical, "Integrador"
            Case 3307
                MsgBox "O número de colunas nas duas tabelas ou consultas selecionadas de uma consulta união não coincide.", vbCritical, "Integrador"
            Case 3308
                MsgBox "Argumento TOP inválido na consulta seleção.", vbCritical, "Integrador"
            Case 3309
                MsgBox "A configuração da propriedade não pode ter mais de 2K.", vbCritical, "Integrador"
            Case 3310
                MsgBox "Esta propriedade não é suportada em fontes de dados externas ou em bancos de dados criados com uma versão anterior do Microsoft Jet.", vbCritical, "Integrador"
            Case 3311
                MsgBox "A propriedade especificada já existe.", vbCritical, "Integrador"
            Case 3312
                MsgBox "As regras de validação e os valores padrão não podem ser inseridos em tabelas do sistema ou vinculadas.", vbCritical, "Integrador"
            Case 3313
                MsgBox "Não é possível inserir esta expressão de validação neste campo.", vbCritical, "Integrador"
            Case 3314
                MsgBox "O campo <nome> não pode conter um valor Null porque a propriedade Required deste campo está definida como True. Insira um valor neste campo.", vbCritical, "Integrador"
            Case 3315
                MsgBox "O campo <nome> não pode ser uma seqüência de comprimento zero.", vbCritical, "Integrador"
            Case 3316
                MsgBox "<Texto de validação em nível de tabela>.", vbCritical, "Integrador"
            Case 3317
                MsgBox "Um ou mais valores são proibidos pela regra de validação <regra> definida para <nome>. Insira um valor que a expressão deste campo possa aceitar.", vbCritical, "Integrador"
            Case 3318
                MsgBox "Os valores especificados em uma cláusula TOP não são permitidos em consultas exclusão e nem em relatórios.", vbCritical, "Integrador"
            Case 3319
                MsgBox "Erro de sintaxe na consulta união.", vbCritical, "Integrador"
            Case 3320
                MsgBox "<Erro> em expressão de validação em nível de tabela.", vbCritical, "Integrador"
            Case 3321
                MsgBox "Sem banco de dados especificado na seqüência de conexão ou cláusula IN.", vbCritical, "Integrador"
            Case 3322
                MsgBox "A consulta de tabela de referência cruzada contém um ou mais títulos fixos e inválidos de colunas.", vbCritical, "Integrador"
            Case 3323
                MsgBox "A consulta não pode ser utilizada como origem da linha.", vbCritical, "Integrador"
            Case 3324
                MsgBox "A consulta é uma consulta DDL e não pode ser utilizada como origem da linha.", vbCritical, "Integrador"
            Case 3325
                MsgBox "A consulta passagem com a propriedade ReturnsRecords definida como True não retornou registros.", vbCritical, "Integrador"
            Case 3326
                MsgBox "Este Recordset não é atualizável.", vbCritical, "Integrador"
            Case 3334
                MsgBox "Somente pode estar presente no formato da versão 1.0.", vbCritical, "Integrador"
            Case 3336
                MsgBox "Btrieve: opção IndexDDF inválida na configuração da inicialização.", vbCritical, "Integrador"
            Case 3337
                MsgBox "Opção DataCodePage inválida na configuração da inicialização.", vbCritical, "Integrador"
            Case 3338
                MsgBox "Btrieve: as opções Xtrieve não estão corretas na configuração da inicialização.", vbCritical, "Integrador"
            Case 3339
                MsgBox "Btrieve: opção IndexDeleteRenumber inválida na configuração da inicialização.", vbCritical, "Integrador"
            Case 3340
                MsgBox "A consulta <nome> está corrompida.", vbCritical, "Integrador"
            Case 3341
                MsgBox "O campo atual deve corresponder à chave de associação <nome> na tabela que serve como lado ‘um’ da relação um-para-muitos. Insira um registro no lado ‘um’ da tabela com o valor de chave desejado e, em seguida, faça a entrada com a chave de associação desejada na tabela ‘somente-muitos’.", vbCritical, "Integrador"
            Case 3342
                MsgBox "Memorando ou Objeto OLE inválido na subconsulta <nome>.", vbCritical, "Integrador"
            Case 3343
                MsgBox "Formato de banco de dados <nome do arquivo> não-reconhecido.", vbCritical, "Integrador"
            Case 3344
                MsgBox "O mecanismo de banco de dados não reconhece o campo <nome> em uma expressão de validação ou o valor padrão na tabela <nome>.", vbCritical, "Integrador"
            Case 3345
                MsgBox "Referência de campo <nome> desconhecida ou inválida.", vbCritical, "Integrador"
            Case 3346
                MsgBox "O número de valores de consulta e de campos de destino não é o mesmo.", vbCritical, "Integrador"
            Case 3349
                MsgBox "Sobrecarga de campo numérico.", vbCritical, "Integrador"
            Case 3350
                MsgBox "O objeto é inválido para a operação.", vbCritical, "Integrador"
            Case 3351
                MsgBox "A expressão ORDER BY <expressão> inclui campos que não são selecionados pela consulta. Somente os campos solicitados na primeira consulta podem ser incluídos em uma expressão ORDER BY.", vbCritical, "Integrador"
            Case 3352
                MsgBox "Sem nome de campo de destino na instrução INSERT INTO <instrução>.", vbCritical, "Integrador"
            Case 3353
                MsgBox "Btrieve: não é possível encontrar o arquivo Field.ddf.", vbCritical, "Integrador"
            Case 3354
                MsgBox "No máximo um registro pode ser retornado por esta subconsulta.", vbCritical, "Integrador"
            Case 3355
                MsgBox "Erro de sintaxe no valor padrão.", vbCritical, "Integrador"
            Case 3356
                MsgBox "Você tentou abrir um banco de dados que já está aberto exclusivamente pelo usuário <nome> na máquina <nome>. Tente novamente quando o banco de dados estiver disponível.", vbCritical, "Integrador"
            Case 3357
                MsgBox "Esta consulta não é uma consulta definição de dados devidamente formada.", vbCritical, "Integrador"
            Case 3358
                MsgBox "Não é possível abrir o arquivo de informações do grupo de trabalho do mecanismo Microsoft Jet.", vbCritical, "Integrador"
            Case 3359
                MsgBox "A consulta passagem deve conter pelo menos um caractere.", vbCritical, "Integrador"
            Case 3360
                MsgBox "A consulta é complexa demais.", vbCritical, "Integrador"
            Case 3361
                MsgBox "Uniões não-permitidas em uma subconsulta.", vbCritical, "Integrador"
            Case 3362
                MsgBox "A atualização/exclusão de linha única afetou mais de uma linha de uma tabela vinculada. O índice exclusivo contém valores duplicados.", vbCritical, "Integrador"
            Case 3364
                MsgBox "Não é possível utilizar o campo Memorando ou Objeto OLE <nome> na cláusula SELECT de uma consulta união.", vbCritical, "Integrador"
            Case 3365
                MsgBox "Não é possível definir esta propriedade para objetos remotos.", vbCritical, "Integrador"
            Case 3366
                MsgBox "Não é possível acrescentar uma relação sem campos definidos.", vbCritical, "Integrador"
            Case 3367
                MsgBox "Não é possível acrescentar. Já existe na coleção um objeto com este nome.", vbCritical, "Integrador"
            Case 3368
                MsgBox "A relação deve ser no mesmo número de campos com os mesmos tipos de dados.", vbCritical, "Integrador"
            Case 3370
                MsgBox "Não é possível modificar a estrutura da tabela <nome>. Ela está em um banco de dados somente leitura.", vbCritical, "Integrador"
            Case 3371
                MsgBox "Não é possível encontrar tabela ou restrição.", vbCritical, "Integrador"
            Case 3372
                MsgBox "Não há índice <nome> na tabela <nome>.", vbCritical, "Integrador"
            Case 3373
                MsgBox "Não é possível criar uma relação. A tabela referenciada <nome> não tem uma chave primária.", vbCritical, "Integrador"
            Case 3374
                MsgBox "Os campos especificados não são indexados exclusivamente na tabela <nome>.", vbCritical, "Integrador"
            Case 3375
                MsgBox "A tabela <nome> já tem um índice chamado <nome>.", vbCritical, "Integrador"
            Case 3376
                MsgBox "A tabela <nome> não existe.", vbCritical, "Integrador"
            Case 3377
                MsgBox "Não há relação <nome> na tabela <nome>.", vbCritical, "Integrador"
            Case 3378
                MsgBox "Já existe uma relação chamada <nome> no banco de dados atual.", vbCritical, "Integrador"
            Case 3379
                MsgBox "Não é possível criar relações para impor integridade referencial. Os dados existentes na tabela <nome> violam as regras de integridade referencial na tabela <nome>.", vbCritical, "Integrador"
            Case 3380
                MsgBox "O campo <nome> já existe na tabela <nome>.", vbCritical, "Integrador"
            Case 3381
                MsgBox "Não há campo chamado <nome> na tabela <nome>.", vbCritical, "Integrador"
            Case 3382
                MsgBox "O tamanho do campo <nome> é longo demais.", vbCritical, "Integrador"
            Case 3383
                MsgBox "Não é possível excluir o campo <nome>. Ele faz parte de uma ou mais relações.", vbCritical, "Integrador"
            Case 3384
                MsgBox "Não é possível excluir uma propriedade interna.", vbCritical, "Integrador"
            Case 3385
                MsgBox "As propriedades não definidas pelo usuário não suportam um valor Null.", vbCritical, "Integrador"
            Case 3386
                MsgBox "A propriedade <nome> deve ser definida antes de utilizar este método.", vbCritical, "Integrador"
            Case 3388
                MsgBox "Função <nome> desconhecida na expressão de validação ou no valor padrão em <nome>.", vbCritical, "Integrador"
            Case 3389
                MsgBox "Suporte de consulta não-disponível.", vbCritical, "Integrador"
            Case 3390
                MsgBox "O nome da conta já existe.", vbCritical, "Integrador"
            Case 3393
                MsgBox "Não é possível executar associação, grupo, classificação ou restrição indexada. Um valor que está sendo procurado ou classificado é longo demais.", vbCritical, "Integrador"
            Case 3394
                MsgBox "Não é possível salvar a propriedade; ela é uma propriedade de esquema.", vbCritical, "Integrador"
            Case 3396
                MsgBox "Não é possível executar a operação em cascata. Como existem registros relacionados na tabela <nome>, as regras de integridade referencial seriam violadas.", vbCritical, "Integrador"
            Case 3397
                MsgBox "Não é possível executar a operação em cascata. Deve haver um registro relacionado na tabela <nome>.", vbCritical, "Integrador"
            Case 3398
                MsgBox "Não é possível executar a operação em cascata. Isto resultaria em uma chave nula na tabela <nome>.", vbCritical, "Integrador"
            Case 3399
                MsgBox "Não é possível executar a operação em cascata. Isto resultaria em uma chave duplicada na tabela <nome>.", vbCritical, "Integrador"
            Case 3400
                MsgBox "Não é possível executar a operação em cascata. Isto resultaria em duas atualizações do campo <nome> na tabela <nome>.", vbCritical, "Integrador"
            Case 3401
                MsgBox "Não é possível executar a operação em cascata. Isto transformaria o campo <nome> em Null, o que não é permitido.", vbCritical, "Integrador"
            Case 3402
                MsgBox "Não é possível executar a operação em cascata. Isto transformaria o campo <nome> em uma seqüência de comprimento zero, o que não é permitido.", vbCritical, "Integrador"
            Case 3403
                MsgBox "Não é possível executar a operação em cascata: <texto de validação>.", vbCritical, "Integrador"
            Case 3404
                MsgBox "Não é possível executar a operação em cascata. O valor inserido é proibido pela regra de validação <regra> definida para <nome>.", vbCritical, "Integrador"
            Case 3405
                MsgBox "Erro <texto de erro> na regra de validação.", vbCritical, "Integrador"
            Case 3406
                MsgBox "A expressão que você está tentando utilizar na propriedade DefaultValue é inválida porque <texto de erro>. Utilize uma expressão válida para definir esta propriedade.", vbCritical, "Integrador"
            Case 3407
                MsgBox "A tabela MSysConf do servidor existe, mas está em um formato incorreto. Entre em contato com o seu administrador do sistema.", vbCritical, "Integrador"
            Case 3409
                MsgBox "Nome de campo <nome> inválido na definição de índice ou relação.", vbCritical, "Integrador"
            Case 3411
                MsgBox "Entrada inválida. Não é possível executar a operação em cascata na tabela <nome> porque o valor inserido é grande demais para o campo <nome>.", vbCritical, "Integrador"
            Case 3412
                MsgBox "Não é possível executar a atualização em cascata na tabela porque ela está atualmente em uso por um outro usuário.", vbCritical, "Integrador"
            Case 3414
                MsgBox "Não é possível executar a operação em cascata na tabela <nome> porque ela está atualmente em uso.", vbCritical, "Integrador"
            Case 3415
                MsgBox "A seqüência de comprimento zero é válida somente em um campo Texto ou Memorando.", vbCritical, "Integrador"
            Case 3416
                MsgBox "<alerta de erro reservado>", vbCritical, "Integrador"
            Case 3417
                MsgBox "Uma consulta ação não pode ser utilizada como origem de linha.", vbCritical, "Integrador"
            Case 3418
                MsgBox "Não é possível abrir <nome da tabela>. Outro usuário está com a tabela aberta utilizando um arquivo de controle de rede ou um estilo de bloqueio diferente.", vbCritical, "Integrador"
            Case 3419
                MsgBox "Não é possível abrir esta tabela do Paradox 4.x ou 5.x porque o ParadoxNetStyle está definido como 3.x no Registro do Windows.", vbCritical, "Integrador"
            Case 3420
                MsgBox "O objeto é inválido ou não está mais definido.", vbCritical, "Integrador"
            Case 3421
                MsgBox "Erro de conversão do tipo de dados.", vbCritical, "Integrador"
            Case 3422
                MsgBox "Não é possível modificar a estrutura da tabela. Outro usuário está com a tabela aberta.", vbCritical, "Integrador"
            Case 3423
                MsgBox "Você não pode utilizar o ODBC para importar de, exportar para ou vincular uma tabela de banco de dados externa do Microsoft Jet ou ISAM para o seu banco de dados.", vbCritical, "Integrador"
            Case 3424
                MsgBox "Não é possível criar o banco de dados porque a localidade é inválida.", vbCritical, "Integrador"
            Case 3428
                MsgBox "Ocorreu um problema no seu banco de dados. Corrija-o reparando e compactando o banco de dados.", vbCritical, "Integrador"
            Case 3429
                MsgBox "Versão incompatível de um ISAM instalável.", vbCritical, "Integrador"
            Case 3430
                MsgBox "Enquanto carregava o ISAM instalável do Microsoft Excel, a OLE não conseguia inicializar.", vbCritical, "Integrador"
            Case 3431
                MsgBox "Este não é um arquivo do Microsoft Excel 5.0.", vbCritical, "Integrador"
            Case 3432
                MsgBox "Erro na abertura de um arquivo do Microsoft Excel 5.0.", vbCritical, "Integrador"
            Case 3433
                MsgBox "Configuração inválida na chave do Excel da seção Engines do Registro do Windows.", vbCritical, "Integrador"
            Case 3434
                MsgBox "Não é possível expandir intervalo nomeado.", vbCritical, "Integrador"
            Case 3435
                MsgBox "Não é possível excluir células da planilha.", vbCritical, "Integrador"
            Case 3436
                MsgBox "Falha na criação do arquivo.", vbCritical, "Integrador"
            Case 3437
                MsgBox "A planilha está cheia.", vbCritical, "Integrador"
            Case 3438
                MsgBox "Os dados que estão sendo exportados não correspondem ao formato descrito no arquivo Schema.ini.", vbCritical, "Integrador"
            Case 3439
                MsgBox "Você tentou vincular ou importar um arquivo de mala direta do Microsoft Word. Apesar de poder exportar esses arquivos, você não pode vinculá-los nem importá-los.", vbCritical, "Integrador"
            Case 3440
                MsgBox "Foi feita uma tentativa de importar ou vincular um arquivo de texto vazio. Para importar ou vincular um arquivo de texto, o arquivo deve conter dados.", vbCritical, "Integrador"
            Case 3441
                MsgBox "O separador de campo de especificação do arquivo de texto corresponde ao separador decimal ou delimitador de texto.", vbCritical, "Integrador"
            Case 3442
                MsgBox "Na especificação <nome> do arquivo de texto, a opção <nome> é inválida.", vbCritical, "Integrador"
            Case 3443
                MsgBox "A especificação <nome> de largura fixa não contém larguras de coluna.", vbCritical, "Integrador"
            Case 3444
                MsgBox "Na especificação <nome> de largura fixa, a coluna <coluna> não especifica uma largura.", vbCritical, "Integrador"
            Case 3445
                MsgBox "Foi encontrada a versão incorreta do arquivo DLL <nome>.", vbCritical, "Integrador"
            Case 3446
                MsgBox "O arquivo VBA do Jet (VBAJET.dll para versões de 16 bits ou VBAJET32.dll para versões de 32 bits) está ausente. Tente reinstalar o aplicativo que retornou o erro.", vbCritical, "Integrador"
            Case 3447
                MsgBox "O arquivo VBA do Jet (VBAJET.dll para versões de 16 bits ou VBAJET32.dll para versões de 32 bits) não conseguiu inicializar quando chamado. Tente reinstalar o aplicativo que retornou o erro.", vbCritical, "Integrador"
            Case 3448
                MsgBox "Uma chamada a uma função do sistema OLE não foi bem-sucedida. Tente reinstalar o aplicativo que retornou o erro.", vbCritical, "Integrador"
            Case 3449
                MsgBox "Nenhum código de país encontrado na seqüência de conexão.", vbCritical, "Integrador"
            Case 3452
                MsgBox "Você não pode fazer alterações na estrutura do banco de dados nesta réplica.", vbCritical, "Integrador"
            Case 3453
                MsgBox "Você não pode estabelecer ou manter uma relação imposta entre uma tabela replicada e uma tabela local.", vbCritical, "Integrador"
            Case 3455
                MsgBox "Não é possível tornar o banco de dados replicável.", vbCritical, "Integrador"
            Case 3456
                MsgBox "O objeto chamado <nome> na coleção <nome> não pode se tornar replicável.", vbCritical, "Integrador"
            Case 3457
                MsgBox "Você não pode definir a propriedade KeepLocal para um objeto que já está replicado.", vbCritical, "Integrador"
            Case 3458
                MsgBox "A propriedade KeepLocal não pode ser definida em um banco de dados; ela pode ser definida somente nos objetos em um banco de dados.", vbCritical, "Integrador"
            Case 3459
                MsgBox "Depois que um banco de dados é replicado, você não pode remover os seus recursos de replicação.", vbCritical, "Integrador"
            Case 3460
                MsgBox "A operação que você tentou entra em conflito com uma operação existente que envolve este membro do conjunto de réplicas.", vbCritical, "Integrador"
            Case 3461
                MsgBox "A propriedade de replicação que você está tentando definir ou excluir é somente leitura e não pode ser alterada.", vbCritical, "Integrador"
            Case 3462
                MsgBox "Não foi possível carregar a DLL.", vbCritical, "Integrador"
            Case 3463
                MsgBox "Não é possível encontrar o .dll <nome>.", vbCritical, "Integrador"
            Case 3464
                MsgBox "Os tipos de dados não correspondem na expressão de critério.", vbCritical, "Integrador"
            Case 3465
                MsgBox "A unidade de disco que você está tentando acessar é ilegível.", vbCritical, "Integrador"
            Case 3468
                MsgBox "O acesso foi negado enquanto acessava a pasta dropbox <nome>.", vbCritical, "Integrador"
            Case 3469
                MsgBox "O disco da pasta dropbox <nome> está cheio.", vbCritical, "Integrador"
            Case 3470
                MsgBox "Falha no disco durante o acesso à pasta dropbox <nome>.", vbCritical, "Integrador"
            Case 3471
                MsgBox "Não foi possível gravar no arquivo de registro Sincronizador.", vbCritical, "Integrador"
            Case 3472
                MsgBox "Disco cheio para caminho <nome>.", vbCritical, "Integrador"
            Case 3473
                MsgBox "Falha no disco durante o acesso ao arquivo de registro <nome>.", vbCritical, "Integrador"
            Case 3474
                MsgBox "Não é possível abrir o arquivo de registro <nome> para gravação.", vbCritical, "Integrador"
            Case 3475
                MsgBox "Violação de compartilhamento durante a tentativa de abrir o arquivo de registro <nome> no modo Deny Write.", vbCritical, "Integrador"
            Case 3476
                MsgBox "Caminho da dropbox <nome> inválido.", vbCritical, "Integrador"
            Case 3477
                MsgBox "Endereço da dropbox <nome> é sintaticamente inválido.", vbCritical, "Integrador"
            Case 3478
                MsgBox "A réplica não é parcial.", vbCritical, "Integrador"
            Case 3479
                MsgBox "Não é possível designar uma réplica parcial como Estrutura-Mestre para o conjunto de réplicas.", vbCritical, "Integrador"
            Case 3480
                MsgBox "A relação <nome> na expressão de filtro parcial é inválida.", vbCritical, "Integrador"
            Case 3481
                MsgBox "O nome de tabela <nome> na expressão parcial de filtro é inválido.", vbCritical, "Integrador"
            Case 3482
                MsgBox "A expressão de filtro para a réplica parcial é inválida.", vbCritical, "Integrador"
            Case 3483
                MsgBox "A senha fornecida para a pasta dropbox <nome> é inválida.", vbCritical, "Integrador"
            Case 3484
                MsgBox "A senha utilizada pelo Sincronizador para gravar em uma pasta dropbox de destino é inválida.", vbCritical, "Integrador"
            Case 3485
                MsgBox "O objeto não pode ser replicado porque o banco de dados não é replicado.", vbCritical, "Integrador"
            Case 3486
                MsgBox "Você não pode adicionar um segundo campo AutoNumeração do Código da Replicação a uma tabela.", vbCritical, "Integrador"
            Case 3487
                MsgBox "O banco de dados que você está tentando replicar não pode ser convertido.", vbCritical, "Integrador"
            Case 3488
                MsgBox "O valor especificado não é um CódigoDaReplicação para qualquer membro do conjunto de réplicas.", vbCritical, "Integrador"
            Case 3489
                MsgBox "O objeto especificado não pode ser replicado porque falta nele um recurso necessário.", vbCritical, "Integrador"
            Case 3490
                MsgBox "Não é possível criar uma nova réplica porque o objeto <nome> no recipiente <nome> não pôde ser replicado.", vbCritical, "Integrador"
            Case 3491
                MsgBox "O banco de dados deve ser aberto no modo exclusivo antes que ele possa ser replicado.", vbCritical, "Integrador"
            Case 3492
                MsgBox "A sincronização falhou porque uma alteração de estrutura não pôde ser aplicada a uma das réplicas.", vbCritical, "Integrador"
            Case 3493
                MsgBox "Não é possível definir o parâmetro Registro especificado para o Sincronizador.", vbCritical, "Integrador"
            Case 3494
                MsgBox "Não foi possível recuperar o parâmetro Registro especificado para o Sincronizador.", vbCritical, "Integrador"
            Case 3495
                MsgBox "Não há sincronizações agendadas entre os dois Sincronizadores.", vbCritical, "Integrador"
            Case 3496
                MsgBox "O Gerenciador de Replicação não consegue encontrar o CódigoDaTroca na tabela MSysExchangeLog.", vbCritical, "Integrador"
            Case 3497
                MsgBox "Não foi possível definir uma agenda para o Sincronizador.", vbCritical, "Integrador"
            Case 3499
                MsgBox "Não é possível recuperar as informações completas de caminho para um membro do conjunto de réplicas.", vbCritical, "Integrador"
            Case 3500
                MsgBox "Não é permitido definir uma troca com o mesmo Sincronizador.", vbCritical, "Integrador"
            Case 3502
                MsgBox "A Estrutura-Mestre ou réplica não está sendo gerenciada por um Sincronizador.", vbCritical, "Integrador"
            Case 3503
                MsgBox "O Registro do Sincronizador não tem valor definido para a chave que você consultou.", vbCritical, "Integrador"
            Case 3504
                MsgBox "O código do Sincronizador não corresponde a um código existente na tabela MSysTranspAddress.", vbCritical, "Integrador"
            Case 3506
                MsgBox "O Sincronizador é incapaz de abrir o registro do Sincronizador.", vbCritical, "Integrador"
            Case 3507
                MsgBox "Não foi possível gravar no registro do Sincronizador.", vbCritical, "Integrador"
            Case 3508
                MsgBox "Não há transporte ativo para o Sincronizador.", vbCritical, "Integrador"
            Case 3509
                MsgBox "Não foi possível encontrar um transporte válido para este Sincronizador.", vbCritical, "Integrador"
            Case 3510
                MsgBox "O membro do conjunto de réplicas que você está tentando sincronizar está atualmente sendo utilizado em outra sincronização.", vbCritical, "Integrador"
            Case 3512
                MsgBox "Não foi possível ler a pasta dropbox.", vbCritical, "Integrador"
            Case 3513
                MsgBox "Não foi possível gravar na pasta dropbox.", vbCritical, "Integrador"
            Case 3514
                MsgBox "O Sincronizador não conseguiu encontrar sincronizações agendadas nem a serem solicitadas para processar.", vbCritical, "Integrador"
            Case 3515
                MsgBox "O mecanismo de banco de dados Microsoft Jet não conseguiu ler o relógio do sistema no seu computador.", vbCritical, "Integrador"
            Case 3516
                MsgBox "Não foi possível encontrar o endereço de transporte.", vbCritical, "Integrador"
            Case 3517
                MsgBox "O Sincronizador não conseguiu encontrar mensagens para serem processadas.", vbCritical, "Integrador"
            Case 3518
                MsgBox "Não foi possível encontrar o Sincronizador na tabela MSysTranspAddress.", vbCritical, "Integrador"
            Case 3519
                MsgBox "Não foi possível enviar a mensagem.", vbCritical, "Integrador"
            Case 3520
                MsgBox "O nome ou código da réplica não corresponde a um membro atualmente gerenciado do conjunto de réplicas.", vbCritical, "Integrador"
            Case 3521
                MsgBox "Dois membros do conjunto de réplicas não podem ser sincronizados porque não há um ponto comum para iniciar a sincronização.", vbCritical, "Integrador"
            Case 3522
                MsgBox "O Sincronizador não consegue encontrar o registro de uma sincronização específica na tabela MSysExchangeLog.", vbCritical, "Integrador"
            Case 3523
                MsgBox "O Sincronizador não consegue encontrar um número de versão específico na tabela MSysSchChange.", vbCritical, "Integrador"
            Case 3524
                MsgBox "O histórico de alterações de estrutura na réplica não corresponde ao histórico na Estrutura-Mestre.", vbCritical, "Integrador"
            Case 3525
                MsgBox "O Sincronizador não conseguiu acessar o banco de dados de mensagens.", vbCritical, "Integrador"
            Case 3526
                MsgBox "O nome selecionado para o objeto do sistema já está em uso.", vbCritical, "Integrador"
            Case 3527
                MsgBox "O Sincronizador ou Gerenciador de Replicação não conseguiu encontrar o objeto do sistema.", vbCritical, "Integrador"
            Case 3528
                MsgBox "Não há dados novos na memória compartilhada para que o Sincronizador ou Gerenciador de Replicação os leiam.", vbCritical, "Integrador"
            Case 3529
                MsgBox "O Sincronizador ou Gerenciador de Replicação encontrou dados não lidos na memória compartilhada. Os dados existentes serão sobrescritos.", vbCritical, "Integrador"
            Case 3530
                MsgBox "O Sincronizador já está servindo um cliente.", vbCritical, "Integrador"
            Case 3531
                MsgBox "O período de espera de um evento se esgotou.", vbCritical, "Integrador"
            Case 3532
                MsgBox "O Sincronizador não conseguiu ser inicializado.", vbCritical, "Integrador"
            Case 3533
                MsgBox "O objeto do sistema utilizado por um processo continua existindo depois que o processo parou.", vbCritical, "Integrador"
            Case 3534
                MsgBox "O Sincronizador procurou por um evento do sistema, mas não encontrou nenhum para relatar ao cliente.", vbCritical, "Integrador"
            Case 3535
                MsgBox "O cliente pediu ao Sincronizador que terminasse a operação.", vbCritical, "Integrador"
            Case 3536
                MsgBox "O Sincronizador recebeu uma mensagem inválida para um membro do conjunto de réplicas que ele gerencia.", vbCritical, "Integrador"
            Case 3538
                MsgBox "Não é possível inicializar o Sincronizador porque há aplicativos demais em execução.", vbCritical, "Integrador"
            Case 3539
                MsgBox "Ocorreu um erro de sistema ou o seu arquivo de troca alcançou seu limite.", vbCritical, "Integrador"
            Case 3540
                MsgBox "Seu arquivo de troca alcançou seu limite ou está corrompido.", vbCritical, "Integrador"
            Case 3541
                MsgBox "O Sincronizador não pôde ser fechado apropriadamente e continua ativo.", vbCritical, "Integrador"
            Case 3542
                MsgBox "O processo parou quando se tentava terminar o cliente do Sincronizador.", vbCritical, "Integrador"
            Case 3543
                MsgBox "O Sincronizador não foi configurado.", vbCritical, "Integrador"
            Case 3544
                MsgBox "O Sincronizador já está sendo executado.", vbCritical, "Integrador"
            Case 3545
                MsgBox "As duas réplicas que você está tentando sincronizar são de diferentes conjuntos de réplicas.", vbCritical, "Integrador"
            Case 3546
                MsgBox "O tipo de sincronização que você está tentando não é válido.", vbCritical, "Integrador"
            Case 3547
                MsgBox "O Sincronizador não conseguiu encontrar uma réplica do conjunto correto para concluir a sincronização.", vbCritical, "Integrador"
            Case 3549
                MsgBox "O nome de arquivo que você forneceu é longo demais.", vbCritical, "Integrador"
            Case 3550
                MsgBox "Não há índice na coluna GUID.", vbCritical, "Integrador"
            Case 3551
                MsgBox "Não foi possível excluir o parâmetro Registro do Sincronizador.", vbCritical, "Integrador"
            Case 3552
                MsgBox "O tamanho do parâmetro Registro excede o máximo permitido.", vbCritical, "Integrador"
            Case 3553
                MsgBox "O GUID não pôde ser criado.", vbCritical, "Integrador"
            Case 3555
                MsgBox "Todos os apelidos das réplicas já estão em uso.", vbCritical, "Integrador"
            Case 3556
                MsgBox "Caminho inválido para a pasta dropbox de destino.", vbCritical, "Integrador"
            Case 3557
                MsgBox "Endereço inválido para a pasta dropbox de destino.", vbCritical, "Integrador"
            Case 3558
                MsgBox "Erro de E/S em disco na pasta dropbox de destino.", vbCritical, "Integrador"
            Case 3559
                MsgBox "Não foi possível gravar porque o disco de destino está cheio.", vbCritical, "Integrador"
            Case 3560
                MsgBox "Os dois membros do conjunto de réplicas que você está tentando sincronizar têm o mesmo CódigoDaReplicação.", vbCritical, "Integrador"
            Case 3561
                MsgBox "Os dois membros do conjunto de réplicas que você está tentando sincronizar são ambos Estruturas-Mestre.", vbCritical, "Integrador"
            Case 3562
                MsgBox "Acesso negado na pasta dropbox de destino.", vbCritical, "Integrador"
            Case 3563
                MsgBox "Erro fatal ao acessar uma pasta dropbox local.", vbCritical, "Integrador"
            Case 3564
                MsgBox "O sincronizador não consegue encontrar o arquivo de origem das mensagens.", vbCritical, "Integrador"
            Case 3565
                MsgBox "Há uma violação de compartilhamento na pasta dropbox de origem porque o banco de dados de mensagens está aberto em outro aplicativo.", vbCritical, "Integrador"
            Case 3566
                MsgBox "Erro de E/S na rede.", vbCritical, "Integrador"
            Case 3567
                MsgBox "A mensagem na pasta dropbox pertence ao Sincronizador errado.", vbCritical, "Integrador"
            Case 3568
                MsgBox "O Sincronizador não conseguiu excluir um arquivo.", vbCritical, "Integrador"
            Case 3569
                MsgBox "Este membro do conjunto de réplicas foi logicamente removido do conjunto e não está mais disponível.", vbCritical, "Integrador"
            Case 3571
                MsgBox "A tentativa de definir uma coluna em uma réplica parcial violou uma regra que governa réplicas parciais.", vbCritical, "Integrador"
            Case 3572
                MsgBox "Ocorreu um erro de E/S em disco durante a leitura ou gravação no diretório TEMP.", vbCritical, "Integrador"
            Case 3574
                MsgBox "O CódigoDaReplicação deste membro do conjunto de réplicas foi reatribuído durante um procedimento de movimentação ou cópia.", vbCritical, "Integrador"
            Case 3575
                MsgBox "A unidade de disco na qual você está tentando gravar está cheia.", vbCritical, "Integrador"
            Case 3576
                MsgBox "O banco de dados que você está tentando abrir já está em uso por outro aplicativo.", vbCritical, "Integrador"
            Case 3577
                MsgBox "Não é possível atualizar a coluna do sistema de replicação.", vbCritical, "Integrador"
            Case 3578
                MsgBox "Não foi possível replicar o banco de dados; não é possível determinar se o banco de dados está aberto no modo exclusivo.", vbCritical, "Integrador"
            Case 3581
                MsgBox "Não é possível abrir a tabela <nome> do sistema de replicação porque ela já está em uso.", vbCritical, "Integrador"
            Case 3583
                MsgBox "Não é possível tornar o objeto <nome> no recipiente <nome> replicável.", vbCritical, "Integrador"
            Case 3584
                MsgBox "Memória insuficiente para concluir a operação.", vbCritical, "Integrador"
            Case 3586
                MsgBox "Erro de sintaxe na expressão de filtro parcial na tabela <nome>.", vbCritical, "Integrador"
            Case 3587
                MsgBox "Expressão inválida na propriedade ReplicaFilter.", vbCritical, "Integrador"
            Case 3588
                MsgBox "Erro ao avaliar a expressão de filtro parcial.", vbCritical, "Integrador"
            Case 3589
                MsgBox "A expressão de filtro parcial contém uma função desconhecida.", vbCritical, "Integrador"
            Case 3592
                MsgBox "Você não pode replicar um banco de dados protegido por senha nem definir proteção por senha em um banco de dados replicado.", vbCritical, "Integrador"
            Case 3593
                MsgBox "Você não pode alterar o atributo-mestre de dados do conjunto de réplicas.", vbCritical, "Integrador"
            Case 3594
                MsgBox "Você não pode alterar o atributo-mestre de dados do conjunto de réplicas. Permite alterações de dados somente na Estrutura-Mestre.", vbCritical, "Integrador"
            Case 3595
                MsgBox "As tabelas de sistema na sua réplica não são mais confiáveis e não devem ser utilizadas.", vbCritical, "Integrador"
            Case 3605
                MsgBox "A sincronização com um banco de dados não-replicado não é permitida. O banco de dados <nome> não é uma Estrutura-Mestre nem uma réplica.", vbCritical, "Integrador"
            Case 3607
                MsgBox "A propriedade de replicação que você está tentando excluir é somente leitura e não pode ser removida.", vbCritical, "Integrador"
            Case 3608
                MsgBox "O comprimento do registro é longo demais para uma tabela indexada do Paradox.", vbCritical, "Integrador"
            Case 3609
                MsgBox "Nenhum índice exclusivo encontrado para o campo referenciado da tabela primária.", vbCritical, "Integrador"
            Case 3610
                MsgBox "Mesma tabela <tabela> referenciada tanto como origem quanto destino em uma consulta criar tabela.", vbCritical, "Integrador"
            Case 3611
                MsgBox "Não é possível executar instruções de definição de dados em fontes de dados vinculadas.", vbCritical, "Integrador"
            Case 3612
                MsgBox "A cláusula GROUP BY de vários níveis não é permitida em uma subconsulta.", vbCritical, "Integrador"
            Case 3613
                MsgBox "Não é possível criar uma relação em tabelas ODBC vinculadas.", vbCritical, "Integrador"
            Case 3614
                MsgBox "GUID não permitido na expressão de critério do método Find.", vbCritical, "Integrador"
            Case 3615
                MsgBox "O tipo não corresponde na expressão JOIN.", vbCritical, "Integrador"
            Case 3616
                MsgBox "A atualização de dados em uma tabela vinculada não é suportada por este ISAM.", vbCritical, "Integrador"
            Case 3617
                MsgBox "A exclusão de dados em uma tabela vinculada não é suportada por este ISAM.", vbCritical, "Integrador"
            Case 3618
                MsgBox "A tabela de exceções não pôde ser criada na importação/exportação.", vbCritical, "Integrador"
            Case 3619
                MsgBox "Os registros não puderam ser adicionados à tabela de exceções.", vbCritical, "Integrador"
            Case 3620
                MsgBox "A conexão para a visualização da sua planilha vinculada do Microsoft Excel foi perdida.", vbCritical, "Integrador"
            Case 3621
                MsgBox "Não é possível alterar a senha em um banco de dados compartilhado aberto.", vbCritical, "Integrador"
            Case 3622
                MsgBox "Você deve utilizar a opção dbSeeChanges com OpenRecordset quando acessar uma tabela do SQL Server que tenha uma coluna IDENTITY.", vbCritical, "Integrador"
            Case 3623
                MsgBox "Não é possível acessar o arquivo DBF acoplado <nome do arquivo> do FoxPro 3.0.", vbCritical, "Integrador"
            Case 3624
                MsgBox "Não foi possível ler o registro; atualmente bloqueado por outro usuário.", vbCritical, "Integrador"
            Case 3625
                MsgBox "A especificação <nome> do arquivo de texto não existe. Você não pode importar, exportar e nem vincular utilizando a especificação.", vbCritical, "Integrador"
            Case 3626
                MsgBox "A operação falhou. Há índices demais na tabela <nome>. Exclua alguns dos índices da tabela e tente a operação novamente.", vbCritical, "Integrador"
            Case 3627
                MsgBox "Não é possível encontrar o arquivo executável do Sincronizador (mstran35.exe).", vbCritical, "Integrador"
            Case 3628
                MsgBox "A réplica do parceiro não é gerenciada por um Sincronizador.", vbCritical, "Integrador"
            Case 3629
                MsgBox "Este Sincronizador e o Sincronizador <nome> têm a mesma dropbox do Sistema de arquivos – <nome>.", vbCritical, "Integrador"
            Case 3631
                MsgBox "Nome de tabela inválido no filtro.", vbCritical, "Integrador"
            Case 3632
                MsgBox "O Sincronizador remoto não está configurado para sincronização remota.", vbCritical, "Integrador"
            Case 3633
                MsgBox "Não é possível carregar a DLL <nome>.", vbCritical, "Integrador"
            Case 3634
                MsgBox "Não é possível criar uma réplica utilizando uma réplica parcial.", vbCritical, "Integrador"
            Case 3635
                MsgBox "Não é possível criar uma réplica parcial de um arquivo de informações do grupo de trabalho.", vbCritical, "Integrador"
            Case 3636
                MsgBox "Não é possível preencher a réplica e nem alterar o filtro da réplica porque ela tem conflitos ou erros de dados.", vbCritical, "Integrador"
            Case 3637
                MsgBox "Não é possível utilizar a tabela de referência cruzada de uma coluna não fixa como uma subconsulta.", vbCritical, "Integrador"
            Case 3638
                MsgBox "Você não pode criar um banco de dados replicável que esteja sendo utilizado por um programa que controla a modificação.", vbCritical, "Integrador"
            Case 3639
                MsgBox "Não é possível criar uma réplica de um arquivo de informações do grupo de trabalho.", vbCritical, "Integrador"
            Case 3640
                MsgBox "O buffer de recuperação era pequeno demais para a quantidade de dados que você solicitou.", vbCritical, "Integrador"
            Case 3641
                MsgBox "Há menos registros restantes no Recordset do que você solicitou.", vbCritical, "Integrador"
            Case 3642
                MsgBox "Foi efetuado um cancelamento na operação.", vbCritical, "Integrador"
            Case 3643
                MsgBox "Um dos registros do Recordset foi excluído por outro processo.", vbCritical, "Integrador"
            Case 3645
                MsgBox "Um dos parâmetros de ligação está incorreto.", vbCritical, "Integrador"
            Case 3646
                MsgBox "O comprimento de linha especificado é menor que a soma dos comprimentos de coluna.", vbCritical, "Integrador"
            Case 3647
                MsgBox "Uma coluna solicitada não está sendo retornada ao Recordset.", vbCritical, "Integrador"
            Case 3648
                MsgBox "Não é possível sincronizar uma réplica parcial com uma outra réplica parcial.", vbCritical, "Integrador"
            Case 3649
                MsgBox "A página de código do idioma não foi especificada ou não pôde ser encontrada.", vbCritical, "Integrador"
            Case 3650
                MsgBox "A Internet está lenta demais.", vbCritical, "Integrador"
            Case 3651
                MsgBox "Endereço de Internet inválido.", vbCritical, "Integrador"
            Case 3652
                MsgBox "Falha de login da Internet.", vbCritical, "Integrador"
            Case 3653
                MsgBox "Internet não-configurada.", vbCritical, "Integrador"
            Case 3656
                MsgBox "Erro na avaliação de uma expressão parcial.", vbCritical, "Integrador"
            Case 3660
                MsgBox "A troca solicitada falhou porque <descrição>.", vbCritical, "Integrador"
            Case -2147168237
                '(...)
            Case Else
                'MsgBox Error & vbNewLine & vbNewLine & "Evento:  " & Nome_do_evento & "  ,  Nº:  " & Err, vbCritical
                'Linha para implementar numeração do código:
                MsgBox Error & vbNewLine & vbNewLine & "Evento:  " & Nome_do_evento & "  ,  Linha:  " & Erl & ", Nº Erro: " & Err, vbCritical
        End Select
    'Call LogError(Interface, Evento, Error, Err)
End Function
