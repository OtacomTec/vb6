Attribute VB_Name = "Comandos_impressoras_fiscais"
Public Const Fabricante_Bematech = "Bematech"
Public Const Fabricante_Sweda = "Sweda"
Public Const Fabricante_Corisco = "Corisco"
Private Declare Function Abre Lib "SWECF.DLL" Alias "ECFOpen" (ByVal Numero As Long, ByVal Tempo As Long, ByVal log As Long, ByVal Mostra As Long) As Long
Private Declare Function Fecha Lib "SWECF.DLL" Alias "ECFClose" () As Long
Private Declare Function Grava Lib "SWECF.DLL" Alias "ECFWrite" (ByVal Comando As String) As Long
Private Declare Function Retorna Lib "SWECF.DLL" Alias "ECFRead" (ByVal Status As String, ByVal Extensao As Long) As Long
Public Function Abre_impressora_fiscal(Fabricante As String)

    If Fabricante = "Bematech" Then
    
         'Verificando a impressora Fiscal ------------------------------------------------------------------
         Dim ACK As Integer
         Dim ST1 As Integer
         Dim ST2 As Integer
         
         LocalRetorno = LeParametrosIni("Sistema", "Retorno")
         
         If LocalRetorno = "-2" Then
             LocalRetorno = "0" 'devolve o retorno na variavel
         Else
             LocalRetorno = Left(LocalRetorno, 1)
         End If
         
         Retorno = Bematech_FI_AbrePortaSerial
        
         If Retorno = -4 Or Retorno = -5 Then
             MsgBox "Erro ao acessar a porta de comunicação com a impressora.Verifique! A aplicação está imposibilitada de ser iniciada", vbCritical, "Only Tech"
             End
         End If
         
         '--- Verificações de periféricos e componentes ---------------------------------------------------------
         
         'Verificar se impressora está ligada.
         Retorno = Bematech_FI_VerificaImpressoraLigada()
         If Retorno = -6 Then
            MsgBox "A Impressora se encontra DESLIGADA.Verifique! A aplicação está imposibilitada de ser iniciada", vbInformation + vbOKOnly, "Atenção"
            End
         End If
         
         'Verifica se a impressora está online ou em intervenção
         Dim strModo As String
         
         strModo = Space(1)
         
         Retorno = Bematech_FI_VerificaModoOperacao(strModo)
         
         If Not strModo = "1" Then
            MsgBox "A Impressora se encontra em Intervenção Técnica.Verifique! A aplicação está imposibilitada de ser iniciada", vbInformation + vbOKOnly, "Atenção"
            Call VerificaRetornoImpressora("", "", "Modo Operação")
         End If
         
         Dim strRetorno_status As String
         Dim strValor_retorno As String
         
         'Verificando a bobina de papel
        ' strRetorno_status = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
        ' strValor_retorno = Str(ACK) & "," & Str(ST1) & "," & Str(ST2)
         
         'Verificando se a bobina está acabando
        ' If (ST1 >= 64) Then
        '     imgInd_pouco_papel.Visible = True
        ' End If
         
         'If (ST1 >= 128) Then
         '    MsgBox "Impressora sem bobina.Troque antes de iniciar a venda.", vbInformation, "Only Tech"
         'End If
         
    End If
    
End Function

Public Function Vende_Item(Fabricante As String, Codigo_produto As String, Descricao_Produto As String, Quantidade_produto As String, Valor_Produto As String, Aliquota_produto As String, Optional Casas_Decimais As Integer, Optional Tipo_Desconto As String, Optional Valor_desconto As Double, Optional Tipo_Quantidade As String, Optional booGaveta_presente As Boolean)
    
    'BEMATECH
    If Fabricante = "Bematech" Then
       Aliquota_produto = Trim(Aliquota_produto)
       
       If Casas_Decimais = 2 Then
          Valor_Produto = Format(Valor_Produto, "#,###0.00")
       End If
       
       If Casas_Decimais = 3 Then
          Valor_Produto = Format(Valor_Produto, "#,###0.000")
       End If
       
       If Tipo_Quantidade = "F" Then
          Quantidade_produto = Format(Quantidade_produto, "#,###0.000")
       End If
       
       Retorno = Bematech_FI_VendeItem(Codigo_produto, Descricao_Produto, Aliquota_produto, Tipo_Quantidade, Quantidade_produto, Casas_Decimais, Valor_Produto, Tipo_Desconto, Valor_desconto)
       
       'Função que analisa o retorno da impressora
       Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
       
       If Retorno <> 1 Then
          frmTela_Venda.booInterrompe_venda = True
          Exit Function
       End If
       
''       'Verificar se gaveta presente
''       If booGaveta_presente = True Then
''          Dim EstadoGaveta As Integer
''          EstadoGaveta = 0
''          Retorno = Bematech_FI_VerificaEstadoGaveta(EstadoGaveta)
''          'Função que analisa o retorno da impressora
''          Call VerificaRetornoImpressora("Estado da Gaveta: ", Str(EstadoGaveta), "Estado da Gaveta")
''          'Verifica se gaveta aberta
''          If EstadoGaveta = 0 Then
''             Retorno = Bematech_FI_AcionaGaveta()
''             'Função que analisa o retorno da impressora
''             Call VerificaRetornoImpressora("", "", "Acionamento da Gaveta")
''          End If
''       End If
       
    End If
    
    'SWEDA
    If Fabricante = "Sweda" Then
       Dim Status As String
       Dim strCodigo As String
       Dim strQuantidade As String
       Dim strPr_unit As String
       Dim strPr_total As String
       Dim strDescricao As String * 24
       Dim strComando As String
       Dim strTrib As String * 3
       strCodigo = Format(Codigo_produto, "0000000000000")
       strQuantidade = Format(Replace(Format(Quantidade_produto, "#,###0.000"), ",", ""), "0000000")
       strPr_unit = Format(Replace(Format(Valor_Produto, "#,###0.00"), ",", ""), "000000000")
       strPr_total = Format(Replace(Format((Quantidade_produto * Valor_Produto), "#,###0.00"), ",", ""), "000000000000")
       strDescricao = Descricao_Produto
       strTrib = Aliquota_produto
       'Iprimi item
Status = Space(512)
       Comando = Chr(27) & ".01" & strCodigo & strQuantidade & strPr_unit & strPr_total & strDescricao & strTrib & "}"
       dll1 = Abre(1, 2, 1, 1)
       dll2 = Grava(Comando)
       dll3 = Retorna(Status, 512)
       If dll1 = 0 Then dll4 = Fecha()
    End If
    
End Function
Public Function Abre_Cupom(Fabricante As String)
    If Fabricante = "Sweda" Then
       Status = Space(512)
       Comando = Chr(27) & ".17}"
       dll1 = Abre(1, 2, 1, 1)
       dll2 = Grava(Comando)
       dll3 = Retorna(Status, 512)
       If dll1 = 0 Then dll4 = Fecha()
    End If
End Function
Public Function Fecha_Cupom(Fabricante As String, Finalizadora As String, Mensagem As String, Optional Total_Pago As Double, Optional ID_Finalizadora As String, Optional TipoAcrescimoDesconto As String, Optional Valor_Total_Bruto As String)
    'Fechando o cupom
    'BEMATECH
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_IniciaFechamentoCupom("A", "%", 0)
       'Função que analisa o retorno da impressora
       Retorno = Bematech_FI_EfetuaFormaPagamento(Finalizadora, Valor_Total_Bruto)
       Retorno = Bematech_FI_TerminaFechamentoCupom(Mensagem)
       Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
       Call Abrir_gaveta(Fabricante)
    End If
    
    'SWEDA
    If Fabricante = "Sweda" Then
         Dim strTotal_pago As String
         Dim strComando As String
         Dim strValor_pago As String
         
         strValor_pago = Format(Replace(Format(Total_Pago, "#,###0.00"), ",", ""), "000000000000")
         
         strFinalizadora = ID_Finalizadora
         
         strComando = ".10" & Trim(strFinalizadora) & strValor_pago
         
         Status = Space(512)
         Comando = Chr(27) & strComando & "}"
         dll1 = Abre(1, 2, 1, 1)
         dll2 = Grava(Comando)
         dll3 = Retorna(Status, 512)
         If dll1 = 0 Then dll4 = Fecha()
         Call Abrir_gaveta(Fabricante)
         Call Rodape_cupom(Fabricante)
    End If
    
End Function

Public Function Cancela_cupom(Fabricante As String)
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_CancelaCupom()
       'Função que analisa o retorno da impressora
       Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")

    End If
    
    If Fabricante = "Sweda" Then
        Status = Space(512)
        Comando = Chr(27) & ".05}"
        dll1 = Abre(1, 2, 1, 1)
        dll2 = Grava(Comando)
        dll3 = Retorna(Status, 512)
        If dll1 = 0 Then dll4 = Fecha()
        Status = Space(512)
    End If
    
End Function

Public Function Leitura_x(Fabricante As String)
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_LeituraX()
       Call VerificaRetornoImpressora("", "", "Leitura X")
    End If
    If Fabricante = "Sweda" Then
        Status = Space(512)
        Comando = Chr(27) & ".13N}"
        dll1 = Abre(1, 2, 1, 1)
        dll2 = Grava(Comando)
        dll3 = Retorna(Status, 512)
        If dll1 = 0 Then dll4 = Fecha()
        Status = Space(512)
    End If
End Function

Public Function Leitura_z(Fabricante As String)
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_FechamentoDoDia()
       Call VerificaRetornoImpressora("", "", "Fechamento do Dia")
    End If
    If Fabricante = "Sweda" Then
        Status = Space(512)
        Comando = Chr(27) & ".14N}"
        dll1 = Abre(1, 2, 1, 1)
        dll2 = Grava(Comando)
        dll3 = Retorna(Status, 512)
        If dll1 = 0 Then dll4 = Fecha()
        Status = Space(512)
    End If
End Function
Public Function Sangria(Fabricante As String, Valor_Sangria)
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_Sangria(Valor_Sangria)
       Call VerificaRetornoImpressora("", "", "Sangria")
    End If
End Function
Public Function Cancela_item(Fabricante As String)
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_CancelaItemAnterior()
       'Função que analisa o retorno da impressora
       Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
    End If
    If Fabricante = "Sweda" Then
        Status = Space(512)
        Comando = Chr(27) & ".04}"
        dll1 = Abre(1, 2, 1, 1)
        dll2 = Grava(Comando)
        dll3 = Retorna(Status, 512)
        If dll1 = 0 Then dll4 = Fecha()
        Status = Space(512)
    End If
End Function
Public Function Abertura_Dia(Finalizadora As String, Fundo_caixa As Double)
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_AberturaDoDia(Fundo_caixa, Finalizadora)
       Call VerificaRetornoImpressora("", "", "Abertura do Dia")
    End If
End Function
Public Function Abrir_gaveta(Fabricante As String)

    If Fabricante = "Sweda" Then
       Status = Space(512)
       Comando = Chr(27) & ".21}"
       dll1 = Abre(1, 2, 1, 1)
       dll2 = Grava(Comando)
       dll3 = Retorna(Status, 512)
       If dll1 = 0 Then dll4 = Fecha()
       Status = Space(512)
    End If
    
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_AcionaGaveta()
       'Função que analisa o retorno da impressora
       Call VerificaRetornoImpressora("", "", "Acionamento da Gaveta")
    End If
    
End Function
Public Function Rodape_cupom(Fabricante As String)

    If Fabricante = "Sweda" Then
         Comando = Empty
         Status = Space(512)
         Comando = Chr(27) & ".12NN0         " & Mensagem & "}"
         dll1 = Abre(1, 2, 1, 1)
         dll2 = Grava(Comando)
         dll3 = Retorna(Status, 512)
         If dll1 = 0 Then dll4 = Fecha()
     End If
     
End Function
Public Function Configuracoes_impressora_fiscal(Fabricante As String)

    If Fabricante = "Bematech" Then
        Retorno = Bematech_FI_ImprimeConfiguracoesImpressora()
        Call VerificaRetornoImpressora("", "", "Configurações da Impressora")
    End If

End Function
Public Function Status_gaveta_aberta(Fabricante As String) As Boolean

    If Fabricante = "Sweda" Then
         Comando = Empty
         Status = Space(512)
         Comando = Chr(27) & ".43}"
         dll1 = Abre(1, 2, 1, 1)
         dll2 = Grava(Comando)
         dll3 = Retorna(Status, 512)
         If dll1 = 0 Then dll4 = Fecha()
         If dll3 = "1" Then
            Status_gaveta_aberta = False
         End If
         If dll3 = "0" Then
            Status_gaveta_aberta = True
         End If
    End If
    
    If Fabricante = "Bematech" Then
       Dim EstadoGaveta As Integer
       Status_gaveta_aberta = 0
       Retorno = Bematech_FI_VerificaEstadoGaveta(EstadoGaveta)
       If EstadoGaveta = 0 Then
          Status_gaveta_aberta = False
       End If
       If EstadoGaveta = 1 Then
          Status_gaveta_aberta = True
       End If
    End If
     
End Function

Public Function Imprime_afinidade_vinculada(Finalizadora As String, Valor_Compra As Double, Valor_limite As Double, Nome_Cliente As String, Fabricante As String, CNPJ_Cliente As String, Observacao As String)

    If Fabricante = "Bematech" Then
        'Abre cupom não vinculado
        Retorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado(Finalizadora, "", "")
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
        'Mensagem 0
        strMensagem = " "
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
        'Mensagem 1
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado("RECONHEÇO E PAGAREI A DÍVIDA REPRESENTADA")
        Call VerificaRetornoImpressora("", "", "BemaFI32")

        
        'Mensagem 2
        strMensagem = Empty
        strMensagem = "VALOR DA COMPRA: R$ " & Valor_Compra
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
        'Mensagem 2.1
        strMensagem = " "
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
        'Mensagem 2.2
        strMensagem = " "
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
        'Varejo
        If frmTela_Venda.intPerfil_ECF = 1 Then
           'Mensagem 3
           strMensagem = Empty
           strMensagem = " ***** LIMITE RESTANTE ***** ----> R$ " & Valor_limite
           Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
           Call VerificaRetornoImpressora("", "", "BemaFI32")
        End If
        
        'Mensagem 4
        strMensagem = " "
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
        'Mensagem 5
        strMensagem = "      __________________________________________"
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
        'Mensagem 6
        strMensagem = "   " & Nome_Cliente
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
        Call VerificaRetornoImpressora("", "", "BemaFI32")

        'Mensagem 7
        strMensagem = Empty
        strMensagem = "CGC/CNPJ: " & CNPJ_Cliente
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
        'Mensagem 8
        strMensagem = Empty
        strMensagem = "OBSERVAÇÃO: " & Observacao
        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
        'Fecha cupom não vinculado
        Retorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado()
        Call VerificaRetornoImpressora("", "", "BemaFI32")
        
    End If

End Function

'''''Public Function Imprime_pagamento_afinidade_vinculada(Num_Titulos As Integer, Numeros_titulos As String, Valor_titulos As String, Finalizadora As String, Valor_Pagamento As Double, Nome_Cliente As String, Fabricante As String)
'''''
'''''    If Fabricante = "Bematech" Then
'''''        'Abre cupom não vinculado
'''''        Finalizadora = "CARTAO"
'''''        Retorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado(Finalizadora, "", "")
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 0
'''''        strMensagem = " "
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        strMensagem = Empty
'''''        strMensagem = "---------------------------------------------"
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 1
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado("PAGAMENTO DE TÍTULOS")
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 1.1
'''''        strMensagem = " "
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 2
'''''        strMensagem = Empty
'''''        strMensagem = "VALOR TOTAL PAGO: R$ " & Valor_Pagamento & "POR " & Num_Titulos & " TITULO(S)"
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 2.1
'''''        strMensagem = " "
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 2.2
'''''        strMensagem = " "
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 4
'''''        strMensagem = " "
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 2
'''''        strMensagem = Empty
'''''        strMensagem = "REF AO(S) TÍTULO(S): " & Numeros_titulos
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 3
'''''        strMensagem = Empty
'''''        strMensagem = "NO VALOR DE R$ " & Valor_titulos
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 4
'''''        strMensagem = " "
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 5
'''''        strMensagem = " "
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 6
'''''        strMensagem = "      ___________________________________"
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''        'Mensagem 7
'''''        strMensagem = "             " & Nome_Cliente
'''''        Retorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(strMensagem)
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''
'''''        'Fecha cupom não vinculado
'''''        Retorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado()
'''''        Call VerificaRetornoImpressora("", "", "BemaFI32")
'''''
'''''    End If
'''''
'''''End Function

