Attribute VB_Name = "Comandos_impressoras_fiscais"
Public Const Fabricante_Bematech = "Bematech"
Public Const Fabricante_Sweda = "Sweda"
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
         strRetorno_status = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
         strValor_retorno = Str(ACK) & "," & Str(ST1) & "," & Str(ST2)
         
         'Verificando se a bobina está acabando
         If (ST1 >= 64) Then
             imgInd_pouco_papel.Visible = True
         End If
         
         If (ST1 >= 128) Then
             MsgBox "Impressora sem bobina.Troque antes de iniciar a venda.", vbInformation, "Only Tech"
         End If
         
    End If
    
End Function
Public Function Vende_Item(Fabricante As String, Codigo_Produto As Long, Descricao_Produto As String, Quantidade_produto As Double, Valor_Produto As Double, Aliquota_produto As Long, Optional Casas_Decimais As Integer, Optional Tipo_Desconto As String, Optional Valor_desconto As Double, Optional Tipo_Quantidade As String)
    'BEMATECH
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_VendeItem(Codigo_Produto, Descricao_Produto, Aliquota_produto, Tipo_Quantidade, Quantidade_produto, Casas_Decimais, Valor_Produto, Tipo_Desconto, Valor_desconto)
       'Função que analisa o retorno da impressora
       Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
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
       'Ver isto
       strCodigo = Format(Codigo_Produto, "0000000000000")
       strQuantidade = Format(Quantidade_produto, "0000000")
       strPr_unit = Format(Valor_Produto, "000000000")
       strPr_total = Format((Quantidade_produto * Valor_Produto), "000000000000")
       strDescricao = Descricao_Produto
       strTrib = "T01"
       'Comando de impressao de item
       'QUANTIDADE FORMATAR PARA 3 CASAS
       'PRECO 2 CASAS DECIMAIS
       'TOTAL 2 CASAS DECIMAIS
       'strComando = ".01" & strCodigo & "0001000" & "000000100" & "000000000100" & strDescricao & strTrib
       'Iprimi item
Status = Space(512)
       'exemplo
       'strComando = Chr(27) & ".01 7891025123454 0001000 000000100 000000000100 Teste Visual Basic      T01}"
       Comando = Chr(27) & ".01" & strCodigo & "0001000" & "000000100" & "000000000100" & strDescricao & strTrib & "}"
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
Public Function Fecha_Cupom(Fabricante As String, Finalizadora As String, Mensagem As String, Optional Total_Pago As Double, Optional ID_Finalizadora As Integer)
    'Fechando o cupom
    'BEMATECH
    If Fabricante = "Bematech" Then
       Retorno = Bematech_FI_FechaCupomResumido(Finalizadora, Mensagem)
       'Função que analisa o retorno da impressora
       Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
    End If
    
    'SWEDA
    If Fabricante = "Sweda" Then
         Dim strTotal_pago As String
         Dim strComando As String
         Dim strValor_pago As String
         
         strValor_pago = Format(Total_Pago, "000000000000")
         strFinalizadora = "01"
         
         strComando = ".10" & strFinalizadora & strValor_pago
         
         Status = Space(512)
         Comando = Chr(27) & ".1001000000000500}"
         'Comando = Chr(27) & ".10" & strFinalizadora & strValor_pago & "}"
         dll1 = Abre(1, 2, 1, 1)
         dll2 = Grava(Comando)
         dll3 = Retorna(Status, 512)
         If dll1 = 0 Then dll4 = Fecha()
         'Fechando o cupom
         Status = Space(512)
        ' Comando = Chr(27) & ".12NN0        " & Mensagem & "      }"
         Comando = Chr(27) & ".12NN0         OBRIGADO PELA PREFERENCIA      }"
         dll1 = Abre(1, 2, 1, 1)
         dll2 = Grava(Comando)
         dll3 = Retorna(Status, 512)
         If dll1 = 0 Then dll4 = Fecha()
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
        Label1.Caption = "Status = " & Status
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
    If Fabricante = Sweda Then
        Status = Space(512)
        Comando = Chr(27) & ".04}"
        dll1 = Abre(1, 2, 1, 1)
        dll2 = Grava(Comando)
        dll3 = Retorna(Status, 512)
        Label1.Caption = "Status = " & Status
        If dll1 = 0 Then dll4 = Fecha()
        Status = Space(512)
    End If
End Function
Public Function Abertura_Dia(Finalizadora As String, Fundo_caixa As Double)
    'If Fabricante = "Bematech" Then
    Retorno = Bematech_FI_AberturaDoDia(Fundo_caixa, Finalizadora)
    Call VerificaRetornoImpressora("", "", "Abertura do Dia")
    'End If
End Function
