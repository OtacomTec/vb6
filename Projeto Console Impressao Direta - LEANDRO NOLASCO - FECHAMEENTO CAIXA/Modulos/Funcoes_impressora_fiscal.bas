Attribute VB_Name = "Funcoes_impressora_fiscal"


'Ler os Valores dos parâmetros nas seções do arquivo ini
Function LeParametrosIni(Secao As String, Label As String) As String
  
   Const TamanhoParametro = 80
   Dim ParametroIni As String * TamanhoParametro
   Dim RetornoFuncao
   Dim ArquivoIni As String
   Dim Contador As Integer
   ParametroIni = ""
     
   RetornoFuncao = GetSystemDirectory(ParametroIni, TamanhoParametro)
   ArquivoIni = Left(ParametroIni, RetornoFuncao) + "\BemaFI32.ini"
   ParametroIni = ""
   RetornoFuncao = GetPrivateProfileString(Secao, Label, "-2", ParametroIni, TamanhoParametro, ArquivoIni)
   RetornoFuncao = Mid(ParametroIni, 1, 2)
   If Val(RetornoFuncao) <> -2 Then
       Contador = 1
       Do
           Tst = Mid(ParametroIni, Contador, 1)
           If Asc(Tst) <> 0 Then
               Contador = Contador + 1
           End If
       Loop While ((Asc(Tst) <> 0) And (Contador < Len(ParametroIni)))
       RetornoFuncao = Mid(ParametroIni, 1, Contador)
   End If
   LeParametrosIni = RetornoFuncao
End Function


Public Function VerificaRetornoImpressora(Label As String, RetornoFuncao As String, TituloJanela As String)
    Dim ACK As Integer
    Dim ST1 As Integer
    Dim ST2 As Integer
    Dim RetornaMensagem As Integer
    Dim StringRetorno As String
    Dim ValorRetorno As String
    Dim RetornoStatus As Integer
    Dim Mensagem As String
    
    'Verificação se acontece algum status que interrompa a impressão do item
    frmTela_Venda.booInterrompe_venda = False
    
    If Retorno = 0 Then
        MsgBox "Erro de comunicação com a impressora.", vbOKOnly + vbCritical, TituloJanela
        Exit Function
    
    ElseIf Retorno = 1 Then
        RetornoStatus = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
        ValorRetorno = Str(ACK) & "," & Str(ST1) & "," & Str(ST2)
        
        If Label <> "" And RetornoFuncao <> "" Then
            RetornaMensagem = 1
        End If
        
        If ACK = 21 Then
            MsgBox "Status da Impressora: 21" & vbCr & vbLf & "Comando não executado", vbOKOnly + vbInformation, TituloJanela
            Exit Function
        End If
        
        If (ST1 <> 0 Or ST2 <> 0) Then
                If (ST1 >= 128) Then
                    Exit Function
                    ST1 = ST1 - 128
                End If
                
                If (ST1 >= 64) Then
                    Exit Function
                    ST1 = ST1 - 64
                End If
                
                If (ST1 >= 32) Then
                    StringRetorno = StringRetorno & "Erro no relógio" & vbCr
                    ST1 = ST1 - 32
                End If
                
                If (ST1 >= 16) Then
                    StringRetorno = StringRetorno & "Impressora em erro" & vbCr
                    ST1 = ST1 - 16
                End If
                    
                If (ST1 >= 8) Then
                    StringRetorno = StringRetorno & "Primeiro dado do comando não foi Esc" & vbCr
                    ST1 = ST1 - 8
                End If
                
                If (ST1 >= 4) Then
                    StringRetorno = StringRetorno & "Comando inexistente" & vbCr
                    ST1 = ST1 - 4
                End If
                    
                If (ST1 >= 2) Then
                    StringRetorno = StringRetorno & "Cupom fiscal aberto" & vbCr
                    ST1 = ST1 - 2
                End If
                
                If (ST1 >= 1) Then
                    StringRetorno = StringRetorno & "Número de parâmetros inválido no comando" & vbCr
                    ST1 = ST1 - 1
                End If
                    
                If (ST2 >= 128) Then
                    StringRetorno = "Tipo de Parâmetro de comando inválido" & vbCr
                    ST2 = ST2 - 128
                End If
                
                If (ST2 >= 64) Then
                    StringRetorno = StringRetorno & "Memória fiscal lotada" & vbCr
                    ST2 = ST2 - 64
                End If
                
                If (ST2 >= 32) Then
                    StringRetorno = StringRetorno & "Erro na CMOS" & vbCr
                    ST2 = ST2 - 32
                End If
                
                If (ST2 >= 16) Then
                    StringRetorno = StringRetorno & "Alíquota não programada" & vbCr
                    ST2 = ST2 - 16
                End If
                    
                If (ST2 >= 8) Then
                    StringRetorno = StringRetorno & "Capacidade de alíquota programáveis lotada" & vbCr
                    ST2 = ST2 - 8
                End If
                
                If (ST2 >= 4) Then
                    StringRetorno = StringRetorno & "Cancelamento não permitido" & vbCr
                    ST2 = ST2 - 4
                End If
                    
                If (ST2 >= 2) Then
                    StringRetorno = StringRetorno & "CGC/IE do proprietário não programados" & vbCr
                    ST2 = ST2 - 2
                End If
                
                If (ST2 >= 1) Then
                    StringRetorno = StringRetorno & "Comando não executado" & vbCr
                    ST2 = ST2 - 1
                End If
                
                If RetornaMensagem Then
                    Mensagem = "Status da Impressora: " & ValorRetorno & _
                           vbCr & vbLf & StringRetorno & vbCr & vbLf & _
                           Label & RetornoFuncao
                Else
                    Mensagem = "Status da Impressora: " & ValorRetorno & _
                       vbCr & vbLf & StringRetorno
                End If
        
                MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
                
                'CANCELAMENTO
                frmTela_Venda.booErro_processamento_impressora = True
                
                Exit Function
        End If 'fim do ST1 <> 0 and ST2 <> 0
        
        If RetornaMensagem Then
            Mensagem = Label & RetornoFuncao
        End If
        
        If Mensagem <> "" Then
            MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
        End If
        Exit Function
    ElseIf Retorno = -1 Then
        MsgBox "Erro de execução da função.", vbOKOnly + vbCritical, TituloJanela
        Exit Function
    
    ElseIf Retorno = -2 Then
        MsgBox "Parâmetro inválido na função.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -3 Then
        MsgBox "Alíquota não programada.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -4 Then
        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório default. " + vbCr + "Por favor, copie esse arquivo para o diretório de sistema do Windows." + vbCr + "Se for o Windows 95 ou 98 é o diretório 'System' se for o Windows NT é o diretório 'System32'.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -6 Then
        MsgBox "Impressora desligada ou cabo de comunicação desconectado.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -7 Then
        MsgBox "Banco não encontrado no arquivo BemaFI32.ini.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -8 Then
        MsgBox "Erro ao criar ou gravar no arquivo status.txt ou retorno.txt.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
        
    ElseIf Retorno = -18 Then
        MsgBox "Não foi possível abrir arquivo INTPOS.001 !", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
        
    ElseIf Retorno = -19 Then
        MsgBox "Parâmetro diferentes !", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -20 Then
        MsgBox "Transação cancelada pelo Operador !", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
        
    ElseIf Retorno = -21 Then
        MsgBox "A Transação não foi aprovada !", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
        
    ElseIf Retorno = -22 Then
        MsgBox "Não foi possível terminal a Impressão !", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
        
    ElseIf Retorno = -23 Then
        MsgBox "Não foi possível terminal a Operação !", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -24 Then
        MsgBox "Forma de pagamento não programada.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -25 Then
        MsgBox "Totalizador não fiscal não programado.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -26 Then
        MsgBox "Transação já realizada.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    
    ElseIf Retorno = -27 Then

        RetornoStatus = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
        ValorRetorno = Str(ACK) & "," & Str(ST1) & "," & Str(ST2)
        
        'Verificando se a bobina está acabando
        If Not (ST1 >= 64) Then
           MsgBox "Status diferente de 6,0,0.", vbOKOnly + vbExclamation, TituloJanela
        End If
        
        If (ST1 >= 128) Then
            MsgBox "Fim de Papel.Troque a bobina para seguir com a venda!", vbCritical, "Only Tech"
            frmTela_Venda.booInterrompe_venda = True
            Exit Function
        End If
        Exit Function
    
    ElseIf Retorno = -28 Then
        MsgBox "Não há dados para serem impressos.", vbOKOnly + vbExclamation, TituloJanela
        Exit Function
    End If
   
End Function

Public Sub CentralizaJanela(Form As Form)
    Form.Top = (Screen.Height - Form.Height) / 2
    Form.Left = (Screen.width - Form.width) / 2
End Sub

Public Function AnalisaFlagsFiscais(FlagFiscal As Integer) As String
    Dim StringRetorno As String
    
    If (FlagFiscal >= 128) Then
        StringRetorno = "Memória fiscal lotada" & vbCr
        FlagFiscal = FlagFiscal - 128
    End If
    
    If (FlagFiscal >= 32) Then
        StringRetorno = StringRetorno & "Permite o cancelamento do cupom" & vbCr
        FlagFiscal = FlagFiscal - 32
    End If
    
    If (FlagFiscal >= 8) Then
        StringRetorno = StringRetorno & "Já houve redução 'Z' no dia" & vbCr
        FlagFiscal = FlagFiscal - 8
    End If
    
    If (FlagFiscal >= 4) Then
        StringRetorno = StringRetorno & "Horário de verão selecionado" & vbCr
        FlagFiscal = FlagFiscal - 4
    End If
        
    If (FlagFiscal >= 2) Then
        StringRetorno = StringRetorno & "Fechamento de formas de pagamento iniciado" & vbCr
        FlagFiscal = FlagFiscal - 2
    End If
    
    If (FlagFiscal >= 1) Then
        StringRetorno = StringRetorno & "Cupom fiscal aberto" & vbCr
        FlagFiscal = FlagFiscal - 1
    End If

    AnalisaFlagsFiscais = StringRetorno

End Function


Public Function AnalisaStatusCheque(StatusCheque As Integer) As String
    Dim StringRetorno As String
    
    If (StatusCheque = 1) Then
        StringRetorno = "Impressora ok." & vbCr
    
    ElseIf (StatusCheque = 2) Then
        StringRetorno = "Cheque em impressão." & vbCr
    
    ElseIf (StatusCheque = 3) Then
        StringRetorno = "Cheque posicionado." & vbCr

    ElseIf (StatusCheque = 4) Then
        StringRetorno = "Aguardando o posicionamento do cheque." & vbCr
    
    End If
    
    AnalisaStatusCheque = StringRetorno

End Function

Public Sub DestacaTexto(Objeto As TextBox)
    Objeto.SelStart = 0
    Objeto.SelLength = Len(Objeto.Text)
End Sub
