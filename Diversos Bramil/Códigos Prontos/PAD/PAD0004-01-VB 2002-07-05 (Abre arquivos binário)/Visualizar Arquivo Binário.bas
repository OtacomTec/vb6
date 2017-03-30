Attribute VB_Name = "Module1"
Private Type tpPreNota
    iCodOperacaotMov                As Integer
    strParTipoOperacaotMov          As Variant
    bIndDoctotMov                   As Byte
    strTipoDoctotMov                As Variant
    strSerieDoctotMov               As Variant
    lNumDoctotMov                   As Long
    dtEmissaoDoctotMov              As Date
    bTipoCadastrotMov               As Byte
    dCGCtMov                        As Double
    bIndTipoPessoatMov              As Byte
    strInscricaoEstadualtMov        As Variant
    strRazaoSocialtMov              As Variant
    strEnderecotMov                 As Variant
    strBairrotMov                   As Variant
    lCEPtMov                        As Long
    strCidadetMov                   As Variant
    strUFtMov                       As Variant
    strFonetMov                     As Variant
    cBaseISStMov                    As Single
    cValorISStMov                   As Single
    cBaseICMStMov                   As Single
    cValorICMStMov                  As Single
    cBaseICMSRetidotMov             As Single
    cValorICMSRetidotMov            As Single
    cBaseNaoTributadatMov           As Single
    cValorMercadoriatMov            As Single
    cValorDescontoCorpotMov         As Single
    sPercDescontoCorpotMov          As Single
    cValorFretetMov                 As Single
    cValorDespesasAcessoriastMov    As Single
    cValorSegurotMov                As Single
    cValorIPItMov                   As Single
    cValorTotalNotatMov             As Single
    iCodCondicaoPagtotMov           As Integer
    bTipoFretetMov                  As Byte
    bStatusMovtotMov                As Byte
    dtStatusMovtotMov               As Date
    lNumItemtMovItem                As Long
    lCodInternoProdutotMovItem      As Long
    strDescrUndMedtMovItem          As Variant
    sQtdeEmbUndMedtMovItem          As Single
    iCodOperacaotMovItem            As Integer
    sQtdeItemtMovItem               As Single
    cPrecoUnitarioItemtMovItem      As Single
    sPercDescontoItemtMovItem       As Single
    cValorDescontoItemtMovItem      As Single
    cValorTotalItemtMovItem         As Single
    sPercICMStMovItem               As Single
    sPercMargemLucroRetidotMovItem  As Single
    sPercICMSRetidotMovItem         As Single
    cValorTotalCustotMovItem        As Single
    sPercIPItMovItem                As Single
End Type

Private Type RegistroTXT   ' Define o tipo definido pelo usuário.
    lCodInternoProdutotPrdBar As Long
    cPrecoVendaAtualtPrdBar As Currency
    dtUltSaidatPrdBar As Date
    dtSaidatGirDia As Date
    
    sQtdeSaidaPDV_e_OutrostGirDia As Single
    sQtdeEstCheckCalculadotPrdBar As Single
    sQtdeEstChecktGirDia As Single
    cTotPrecoCustotGirDia As Currency
    cPrecoCustoTotal As Currency
    cValorTotalVendas As Currency
     
    dtPrecoVendaAtualtPrdBar As Date
    bIndStatusProdutotPrdReg As Byte
    dtStatusProdutotPrdReg As Date
    dtRegistrotPrdReg As Date
    bIndDiaPedidotPrdReg As Byte
End Type
Public Type RegistroZB3   ' Define o tipo definido pelo usuário.
    lCodInternoProdutotPrdBar As Long
    cPrecoVendaAtualtPrdBar As Currency
    dtSaidatGirDia As Date
    sQtdeSaidaPDV_e_OutrostGirDia As Single
    sQtdeEstCheckCalculadotPrdBar As Single
    sQtdeEstChecktGirDia As Single
    cTotPrecoCustotGirDia As Currency
    cPrecoCustoTotal As Currency
    cValorTotalVendas As Currency
    dtPrecoVendaAtualtPrdBar As Date
End Type
Private Type RegistroZB2   ' Define o tipo definido pelo usuário.
    lCodInternoProdutotPrdBar     As Long
    bIndDiaPedidotPrdReg          As Byte
    dtCodInternoAlteracaotPrdBar  As Date
    dtPrecoVendaAtualtPrdBar      As Date
    dtStatusProdutotPrdReg        As Date
    cPrecoVendaAtualtPrdBar       As Currency
    sQtdeEstCheckCalculadotPrdBar As Single
    bIndStatusProdutotPrdReg      As Byte
    dtUltSaidatPrdBar             As Date
End Type

Type TipoReg
    RegAntigo_ As RegistroTXT
    RegZB2_ As RegistroZB2
    RegZB3_ As RegistroZB3
End Type

Public RegAntigo As RegistroTXT
Public RegZB2 As RegistroZB2
Public RegZB3 As RegistroZB3
Public RegGMS005 As tpPreNota


Public CaminhoDoArquivo As String

Public Function ExibirTXTBinário(strNomeDoArquivo As String)   ', Reg As TipoReg)  ' RegAntigo As RegistroTXT)

    strNomeDoArquivoPuro = Mid(strNomeDoArquivo, InStrRev(strNomeDoArquivo, "\") + 1, Len(strNomeDoArquivo) - InStrRev(strNomeDoArquivo, "\"))
    Form1.MSFlexGrid1.Rows = 1
    Form1.StatusBar1.Panels(1).Text = ""
    Form1.StatusBar1.Panels(2).Text = ""
    Form1.StatusBar1.Panels(3).Text = FileDateTime(strNomeDoArquivo)
    
    Select Case Mid(strNomeDoArquivoPuro, 1, InStr(strNomeDoArquivoPuro, "_"))
        Case "txtGMS002_"
            lcanal = FreeFile
            Open strNomeDoArquivo For Random As #lcanal Len = Len(RegAntigo)
            NrRegistro = Seek(lcanal)
            liTotalDeRegistros = FileLen(strNomeDoArquivo) / Len(RegAntigo)
            Form1.StatusBar1.Panels(1).Text = Mid(strNomeDoArquivo, _
                                                  InStrRev(strNomeDoArquivo, "\") + 1, _
                                                  Len(strNomeDoArquivo) - InStrRev(strNomeDoArquivo, "\") + 1)
            Form1.StatusBar1.Panels(2).Text = "0 / " & liTotalDeRegistros
            Form1.ProgressBar1.Max = liTotalDeRegistros
            Form1.ProgressBar1.Min = 0.001
            ctReg = 0
            Get #lcanal, NrRegistro, RegAntigo
            'Monta Cabeçalho do grid
            Form1.MSFlexGrid1.Cols = 14
            Form1.MSFlexGrid1.ColWidth(1) = 800
            Form1.MSFlexGrid1.ColWidth(2) = 1000
            Form1.MSFlexGrid1.ColWidth(3) = 800
            Form1.MSFlexGrid1.ColWidth(4) = 800
            Form1.MSFlexGrid1.ColWidth(5) = 1000
            Form1.MSFlexGrid1.ColWidth(6) = 1100
            Form1.MSFlexGrid1.ColWidth(7) = 1000
            Form1.MSFlexGrid1.ColWidth(8) = 1000
            Form1.MSFlexGrid1.ColWidth(9) = 1200
            Form1.MSFlexGrid1.ColWidth(10) = 700
            Form1.MSFlexGrid1.ColWidth(11) = 1000
            Form1.MSFlexGrid1.ColWidth(12) = 700
            Form1.MSFlexGrid1.ColWidth(13) = 700
            
            Form1.MSFlexGrid1.TextMatrix(0, 1) = "Cod Int"
            Form1.MSFlexGrid1.TextMatrix(0, 2) = "Dt Saída"
            Form1.MSFlexGrid1.TextMatrix(0, 3) = "Qtde"
            Form1.MSFlexGrid1.TextMatrix(0, 4) = "Saldo"
            Form1.MSFlexGrid1.TextMatrix(0, 5) = "Custo Total"
            Form1.MSFlexGrid1.TextMatrix(0, 6) = "Total Vendas"
            Form1.MSFlexGrid1.TextMatrix(0, 7) = "Ult Saída"
            Form1.MSFlexGrid1.TextMatrix(0, 8) = "pç Venda"
            Form1.MSFlexGrid1.TextMatrix(0, 9) = "Dt Reaj"
            Form1.MSFlexGrid1.TextMatrix(0, 10) = "Status"
            Form1.MSFlexGrid1.TextMatrix(0, 11) = "dt Status"
            Form1.MSFlexGrid1.TextMatrix(0, 12) = "Dia"
            Form1.MSFlexGrid1.TextMatrix(0, 13) = "Registro"
            
            
            Do While Not EOF(lcanal)   ' Faz o loop até o fim do arquivo.
                DoEvents
            
                Form1.MSFlexGrid1.AddItem ""
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 1) = RegAntigo.lCodInternoProdutotPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 2) = RegAntigo.dtSaidatGirDia
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 3) = RegAntigo.sQtdeSaidaPDV_e_OutrostGirDia
                        
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 4) = RegAntigo.sQtdeEstChecktGirDia
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 5) = RegAntigo.cPrecoCustoTotal
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 6) = RegAntigo.cValorTotalVendas
            
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 7) = RegAntigo.dtUltSaidatPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 8) = RegAntigo.cPrecoVendaAtualtPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = RegAntigo.dtPrecoVendaAtualtPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 10) = RegAntigo.bIndStatusProdutotPrdReg
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 11) = RegAntigo.dtStatusProdutotPrdReg
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 12) = RegAntigo.bIndDiaPedidotPrdReg
                
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 13) = NrRegistro
                Form1.StatusBar1.Panels(2).Text = NrRegistro & " / " & liTotalDeRegistros
                Form1.ProgressBar1.Value = NrRegistro
                NrRegistro = Seek(lcanal)
                DoEvents
                If Not EOF(lcanal) Then Get #lcanal, NrRegistro, RegAntigo
            Loop
        Case "txtZB2_"
            lcanal = FreeFile
            Open strNomeDoArquivo For Random As #lcanal Len = Len(RegZB2)
            NrRegistro = Seek(lcanal)
            liTotalDeRegistros = FileLen(strNomeDoArquivo) / Len(RegZB2)
            Form1.StatusBar1.Panels(1).Text = Mid(strNomeDoArquivo, _
                                                  InStrRev(strNomeDoArquivo, "\") + 1, _
                                                  Len(strNomeDoArquivo) - InStrRev(strNomeDoArquivo, "\") + 1)
            Form1.StatusBar1.Panels(2).Text = "0 / " & liTotalDeRegistros
            Form1.ProgressBar1.Max = liTotalDeRegistros
            Form1.ProgressBar1.Min = 0.001
            
            ctReg = 0
            Get #lcanal, NrRegistro, RegZB2
            
            'Monta Cabeçalho do grid
            Form1.MSFlexGrid1.Cols = 11
            Form1.MSFlexGrid1.ColWidth(1) = 800
            Form1.MSFlexGrid1.ColWidth(2) = 1000
            Form1.MSFlexGrid1.ColWidth(3) = 800
            Form1.MSFlexGrid1.ColWidth(4) = 800
            Form1.MSFlexGrid1.ColWidth(5) = 1000
            Form1.MSFlexGrid1.ColWidth(6) = 700
            Form1.MSFlexGrid1.ColWidth(7) = 1000
            Form1.MSFlexGrid1.ColWidth(8) = 1000
            Form1.MSFlexGrid1.ColWidth(9) = 1200
            Form1.MSFlexGrid1.ColWidth(10) = 700
            
            Form1.MSFlexGrid1.TextMatrix(0, 1) = "Cod Int"
            Form1.MSFlexGrid1.TextMatrix(0, 2) = "Dt Ult Saída"
            Form1.MSFlexGrid1.TextMatrix(0, 3) = "Estoque"
            Form1.MSFlexGrid1.TextMatrix(0, 4) = "Pc Venda"
            Form1.MSFlexGrid1.TextMatrix(0, 5) = "Dt Pc Venda"
            Form1.MSFlexGrid1.TextMatrix(0, 6) = "Status"
            Form1.MSFlexGrid1.TextMatrix(0, 7) = "Dt Status"
            Form1.MSFlexGrid1.TextMatrix(0, 8) = "Dt Cod Alt"
            Form1.MSFlexGrid1.TextMatrix(0, 9) = "Dia Pedido"
            Form1.MSFlexGrid1.TextMatrix(0, 10) = "Registro"

            Do While Not EOF(lcanal)
                DoEvents
                Form1.MSFlexGrid1.AddItem ""
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 1) = RegZB2.lCodInternoProdutotPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 2) = RegZB2.dtUltSaidatPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 3) = RegZB2.sQtdeEstCheckCalculadotPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 4) = RegZB2.cPrecoVendaAtualtPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 5) = RegZB2.dtPrecoVendaAtualtPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 6) = RegZB2.bIndStatusProdutotPrdReg
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 7) = RegZB2.dtStatusProdutotPrdReg
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 8) = RegZB2.dtCodInternoAlteracaotPrdBar
                Select Case RegZB2.bIndDiaPedidotPrdReg
                    Case 0: Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = "Indefinido"
                    Case 1: Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = "Domingo"
                    Case 2: Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = "Segunda"
                    Case 3: Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = "Terça"
                    Case 4: Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = "Quarta"
                    Case 5: Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = "Quinta"
                    Case 6: Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = "Sexta"
                    Case 7: Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = "Sábado"
                End Select
                'Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = RegZB2.bIndDiaPedidotPrdReg
                
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 10) = NrRegistro
                Form1.StatusBar1.Panels(2).Text = NrRegistro & " / " & liTotalDeRegistros
                Form1.ProgressBar1.Value = NrRegistro
                NrRegistro = Seek(lcanal)
                DoEvents
                If Not EOF(lcanal) Then Get #lcanal, NrRegistro, RegZB2
            Loop
        
        Case "txtGMS005_"
            lcanal = FreeFile
            Open strNomeDoArquivo For Random As #lcanal Len = Len(RegGMS005)
            NrRegistro = Seek(lcanal)
            liTotalDeRegistros = FileLen(strNomeDoArquivo) / Len(RegGMS005)
            Form1.StatusBar1.Panels(1).Text = Mid(strNomeDoArquivo, _
                                                  InStrRev(strNomeDoArquivo, "\") + 1, _
                                                  Len(strNomeDoArquivo) - InStrRev(strNomeDoArquivo, "\") + 1)
            Form1.StatusBar1.Panels(2).Text = "0 / " & liTotalDeRegistros
            Form1.ProgressBar1.Max = liTotalDeRegistros
            Form1.ProgressBar1.Min = 0.001
            
            ctReg = 0
            Get #lcanal, NrRegistro, RegGMS005
            
            
            
            'Monta Cabeçalho do grid
            Form1.MSFlexGrid1.Cols = 12
            Form1.MSFlexGrid1.ColWidth(1) = 800
            Form1.MSFlexGrid1.ColWidth(2) = 1000
            Form1.MSFlexGrid1.ColWidth(3) = 400
            Form1.MSFlexGrid1.ColWidth(4) = 700
            Form1.MSFlexGrid1.ColWidth(5) = 350
            Form1.MSFlexGrid1.ColWidth(6) = 600
            Form1.MSFlexGrid1.ColWidth(7) = 700
            Form1.MSFlexGrid1.ColWidth(8) = 800
            Form1.MSFlexGrid1.ColWidth(9) = 1000
            Form1.MSFlexGrid1.ColWidth(10) = 300
            Form1.MSFlexGrid1.ColWidth(11) = 700
         
            'RegGMS005.sQtdeItemtMovItem
            Form1.MSFlexGrid1.TextMatrix(0, 1) = "Número"
            Form1.MSFlexGrid1.TextMatrix(0, 2) = "Dt Emissão"
            Form1.MSFlexGrid1.TextMatrix(0, 3) = "Item"
            Form1.MSFlexGrid1.TextMatrix(0, 4) = "Produto"
            Form1.MSFlexGrid1.TextMatrix(0, 5) = "UN"
            Form1.MSFlexGrid1.TextMatrix(0, 6) = "QtEmb"
            Form1.MSFlexGrid1.TextMatrix(0, 7) = "P Unit"
            Form1.MSFlexGrid1.TextMatrix(0, 8) = "Valor Item"
            Form1.MSFlexGrid1.TextMatrix(0, 9) = "Total Nota"
            Form1.MSFlexGrid1.TextMatrix(0, 10) = "QT"
            Form1.MSFlexGrid1.TextMatrix(0, 11) = "Reg"

            Do While Not EOF(lcanal)
                DoEvents
                Form1.MSFlexGrid1.AddItem ""
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 1) = RegGMS005.lNumDoctotMov
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 2) = RegGMS005.dtEmissaoDoctotMov
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 3) = RegGMS005.lNumItemtMovItem
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 4) = RegGMS005.lCodInternoProdutotMovItem
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 5) = RegGMS005.strDescrUndMedtMovItem
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 6) = RegGMS005.sQtdeEmbUndMedtMovItem
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 7) = Format(RegGMS005.cPrecoUnitarioItemtMovItem, "#,##0.00")
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 8) = Format(RegGMS005.cValorTotalItemtMovItem, "#,##0.00")
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = Format(RegGMS005.cValorTotalNotatMov, "#,##0.00")
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 10) = RegGMS005.sQtdeItemtMovItem
                
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 11) = NrRegistro
                Form1.StatusBar1.Panels(2).Text = NrRegistro & " / " & liTotalDeRegistros
                Form1.ProgressBar1.Value = NrRegistro
                NrRegistro = Seek(lcanal)
                DoEvents
                If Not EOF(lcanal) Then Get #lcanal, NrRegistro, RegGMS005
            Loop
            
        Case "txtZB3_"
            lcanal = FreeFile
            Open strNomeDoArquivo For Random As #lcanal Len = Len(RegZB3)
            NrRegistro = Seek(lcanal)
            liTotalDeRegistros = FileLen(strNomeDoArquivo) / Len(RegZB3)
            Form1.StatusBar1.Panels(1).Text = Mid(strNomeDoArquivo, _
                                                  InStrRev(strNomeDoArquivo, "\") + 1, _
                                                  Len(strNomeDoArquivo) - InStrRev(strNomeDoArquivo, "\") + 1)
            Form1.StatusBar1.Panels(2).Text = "0 / " & liTotalDeRegistros
            Form1.ProgressBar1.Max = liTotalDeRegistros
            Form1.ProgressBar1.Min = 0.001
            
            ctReg = 0
            Get #lcanal, NrRegistro, RegZB3
            
            'Monta Cabeçalho do grid
            Form1.MSFlexGrid1.Cols = 12
            Form1.MSFlexGrid1.ColWidth(1) = 800
            Form1.MSFlexGrid1.ColWidth(2) = 1000
            Form1.MSFlexGrid1.ColWidth(3) = 800
            Form1.MSFlexGrid1.ColWidth(4) = 1000
            Form1.MSFlexGrid1.ColWidth(5) = 500
            Form1.MSFlexGrid1.ColWidth(6) = 1050
            Form1.MSFlexGrid1.ColWidth(7) = 1000
            Form1.MSFlexGrid1.ColWidth(8) = 1000
            Form1.MSFlexGrid1.ColWidth(9) = 1200
            Form1.MSFlexGrid1.ColWidth(10) = 1000
            Form1.MSFlexGrid1.ColWidth(11) = 700
         
            
            Form1.MSFlexGrid1.TextMatrix(0, 1) = "Cod Int"
            Form1.MSFlexGrid1.TextMatrix(0, 2) = "Dt Saída"
            Form1.MSFlexGrid1.TextMatrix(0, 3) = "Est Calc"
            Form1.MSFlexGrid1.TextMatrix(0, 4) = "Est Chk Gir"
            Form1.MSFlexGrid1.TextMatrix(0, 5) = "Giro"
            Form1.MSFlexGrid1.TextMatrix(0, 6) = "Pc de Venda"
            Form1.MSFlexGrid1.TextMatrix(0, 7) = "Valor Venda"
            Form1.MSFlexGrid1.TextMatrix(0, 8) = "Cust Pratic"
            Form1.MSFlexGrid1.TextMatrix(0, 9) = "Cust Tot Mat"
            Form1.MSFlexGrid1.TextMatrix(0, 10) = "Dt Pc Venda"
            Form1.MSFlexGrid1.TextMatrix(0, 11) = "Registro"

            Do While Not EOF(lcanal)
                DoEvents
                Form1.MSFlexGrid1.AddItem ""
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 1) = RegZB3.lCodInternoProdutotPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 2) = RegZB3.dtSaidatGirDia
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 3) = RegZB3.sQtdeEstCheckCalculadotPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 4) = RegZB3.sQtdeEstChecktGirDia
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 5) = RegZB3.sQtdeSaidaPDV_e_OutrostGirDia
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 6) = RegZB3.cPrecoVendaAtualtPrdBar
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 7) = RegZB3.cValorTotalVendas
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 8) = RegZB3.cTotPrecoCustotGirDia
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 9) = RegZB3.cPrecoCustoTotal
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 10) = RegZB3.dtPrecoVendaAtualtPrdBar
                
                Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 11) = NrRegistro
                Form1.StatusBar1.Panels(2).Text = NrRegistro & " / " & liTotalDeRegistros
                Form1.ProgressBar1.Value = NrRegistro
                NrRegistro = Seek(lcanal)
                DoEvents
                If Not EOF(lcanal) Then Get #lcanal, NrRegistro, RegZB3
            Loop
    End Select

    Close #lcanal
End Function

Public Function AlterarRegistroTXTBinário(strNomeDoArquivo As String, RegAntigo As RegistroTXT, NrRegistro As Long)
    
    lcanal = FreeFile
    Open strNomeDoArquivo For Random As #lcanal Len = Len(RegAntigo)
    Get #lcanal, NrRegistro, RegAntigo
    Select Case Form1.MSFlexGrid1.Col
        Case 1
            RegAntigo.lCodInternoProdutotPrdBar = Form2.Text1.Text
        Case 2
            RegAntigo.dtSaidatGirDia = Form2.Text1.Text
        Case 3
            RegAntigo.sQtdeSaidaPDV_e_OutrostGirDia = Form2.Text1.Text
        Case 4
            RegAntigo.sQtdeEstChecktGirDia = Form2.Text1.Text
        Case 5
            RegAntigo.cPrecoCustoTotal = Form2.Text1.Text
        Case 6
            RegAntigo.cValorTotalVendas = Form2.Text1.Text
        Case 7
            RegAntigo.dtUltSaidatPrdBar = Form2.Text1.Text
        Case 8
            RegAntigo.cPrecoVendaAtualtPrdBar = Form2.Text1.Text
        Case 9
            RegAntigo.dtPrecoVendaAtualtPrdBar = Form2.Text1.Text
        Case 10
            RegAntigo.bIndStatusProdutotPrdReg = Form2.Text1.Text
        Case 11
            RegAntigo.dtStatusProdutotPrdReg = Form2.Text1.Text
        Case 12
            RegAntigo.bIndDiaPedidotPrdReg = Form2.Text1.Text
    
    End Select
    'Carrega o primeiro registro
    Put #lcanal, NrRegistro, RegAntigo
    Close #lcanal
    Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Row, Form1.MSFlexGrid1.Col) = Form2.Text1.Text
    Unload Form2
End Function


