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
    Form1.MSFlexGrid1.Visible = False
    Form1.MSFlexGrid2.Visible = False
    
    Form1.MSFlexGrid1.Height = (Form1.StatusBar1.Top - Form1.Toolbar1.Height) - 130
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
        Case Else
            Select Case strNomeDoArquivoPuro
                Case "CADPROD.txt"
            
                    lcanal = FreeFile
                    Open strNomeDoArquivo For Input As #lcanal
                    Form1.StatusBar1.Panels(1).Text = Mid(strNomeDoArquivo, _
                                                          InStrRev(strNomeDoArquivo, "\") + 1, _
                                                          Len(strNomeDoArquivo) - InStrRev(strNomeDoArquivo, "\") + 1)
                    'Form1.StatusBar1.Panels(2).Text = "0 / " & liTotalDeRegistros
                    'Form1.ProgressBar1.Max = liTotalDeRegistros
                    'Form1.ProgressBar1.Min = 0.001
                    ctReg = 0
                    Dim flx As MSFlexGrid
                    Set flx = Form1.MSFlexGrid1
                    With flx
                        .Cols = 38
                        .ColWidth(0) = 250
                        .ColWidth(1) = 250
                        .ColWidth(2) = 250
                        .ColWidth(3) = 250
                        .ColWidth(4) = 1300
                        .ColWidth(5) = 250
                        .ColWidth(6) = 1000
                        .ColWidth(7) = 250
                        .ColWidth(8) = 2000
                        .ColWidth(9) = 1800
                        .ColWidth(10) = 1300
                        .ColWidth(11) = 250
                        .ColWidth(12) = 1000
                        .ColWidth(13) = 1000
                        
                        .ColWidth(14) = 600
                        .ColWidth(15) = 300
                        .ColWidth(16) = 300
                        
                        .ColWidth(17) = 300
                        .ColWidth(18) = 300
                        .ColWidth(19) = 800
                        
                        .ColWidth(20) = 250
                        .ColWidth(21) = 250
                        
                        .ColWidth(22) = 250
                        .ColWidth(23) = 450
                        
                        .ColWidth(24) = 250
                        .ColWidth(25) = 700
                        .ColWidth(26) = 700
                        
                        .ColWidth(27) = 1200
                        .ColWidth(28) = 450
                        
                        .ColWidth(29) = 700
                        .ColWidth(30) = 700
                        .ColWidth(31) = 700
                        .ColWidth(32) = 700
                        .ColWidth(33) = 700
                        
                        .ColWidth(34) = 1000
                        .ColWidth(35) = 600
                        .ColWidth(36) = 600
                        .ColWidth(37) = 350
                        
                        'Form1.MSFlexGrid1.TextMatrix(0, 1) = "Cod Int"
                        .TextMatrix(0, 1) = "O"
                        .TextMatrix(0, 2) = "AÇ"
                        .TextMatrix(0, 3) = "S"
                        .TextMatrix(0, 4) = "Código de Barra"
                        .TextMatrix(0, 5) = "D"
                        .TextMatrix(0, 6) = "Código Interno"
                        .TextMatrix(0, 7) = "D"
                        .TextMatrix(0, 8) = "Descrição Completa"
                        .TextMatrix(0, 9) = "Descrição Reduzida"
                        .TextMatrix(0, 10) = "Código Vazilhame"
                        .TextMatrix(0, 11) = "D"
                        .TextMatrix(0, 12) = "Preço Venda"
                        .TextMatrix(0, 13) = "Preço Custo"
                        
                        .TextMatrix(0, 14) = "Validade"
                        .TextMatrix(0, 15) = "AL Pdv"
                        .TextMatrix(0, 16) = "AL NF"
                        
                        .TextMatrix(0, 17) = "Cod Emb"
                        .TextMatrix(0, 18) = "Desc Emb"
                        .TextMatrix(0, 19) = "Qtde Emb"
                        .TextMatrix(0, 20) = "Peso Variável"
                        .TextMatrix(0, 21) = "Vende Fracionária"
                        
                        .TextMatrix(0, 22) = "Tipo Etiq Gôndola"
                        .TextMatrix(0, 23) = "Qtde Etiq Gôndola"
                        .TextMatrix(0, 24) = "Tipo Etiq Produto"
                        
                        .TextMatrix(0, 25) = "Ini Promoção"
                        .TextMatrix(0, 26) = "Fim Promoção"
                        .TextMatrix(0, 27) = "Pç Promoção"
                        .TextMatrix(0, 28) = "TP Promoção"
                        
                        .TextMatrix(0, 29) = "Merc 1"
                        .TextMatrix(0, 30) = "Merc 2"
                        .TextMatrix(0, 31) = "Merc 3"
                        .TextMatrix(0, 32) = "Merc 4"
                        .TextMatrix(0, 33) = "Merc 5"
                        
                        .TextMatrix(0, 34) = "Fornecedor"
                        .TextMatrix(0, 35) = "Faixa Pç"
                        .TextMatrix(0, 36) = "Margem Teórica"
                        .TextMatrix(0, 37) = "ST"
                    End With
                    
                    'Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 0) = Mid(lstrLinha, 1, 1)
                    'Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 1) = Mid(lstrLinha, 2, 1)
                    'Form1.MSFlexGrid1.TextMatrix(Form1.MSFlexGrid1.Rows - 1, 2) = Mid(lstrLinha, 3, 1)
                    Line Input #lcanal, lstrLinha
                    Dim ctLinha
                   
                    Do While Not EOF(lcanal)   ' Faz o loop até o fim do arquivo.
                        Line Input #lcanal, lstrLinha
                        
                        DoEvents
                        
                        If ctLinha >= 9000 Then
                            ctLinha = 1
                            flx.Height = flx.Height / 2
                            flx.Visible = True
                            Set flx = Form1.MSFlexGrid2
                            
                            flx.Left = Form1.MSFlexGrid1.Left
                            flx.Width = Form1.MSFlexGrid1.Width
                            flx.Height = Form1.MSFlexGrid1.Height
                            flx.Top = Form1.MSFlexGrid1.Top + Form1.MSFlexGrid1.Height + 50
                            With flx
                                .Cols = 38
                                .ColWidth(0) = 250
                                .ColWidth(1) = 250
                                .ColWidth(2) = 250
                                .ColWidth(3) = 250
                                .ColWidth(4) = 1300
                                .ColWidth(5) = 250
                                .ColWidth(6) = 1000
                                .ColWidth(7) = 250
                                .ColWidth(8) = 2000
                                .ColWidth(9) = 1800
                                .ColWidth(10) = 1300
                                .ColWidth(11) = 250
                                .ColWidth(12) = 1000
                                .ColWidth(13) = 1000
                                
                                .ColWidth(14) = 600
                                .ColWidth(15) = 300
                                .ColWidth(16) = 300
                                
                                .ColWidth(17) = 300
                                .ColWidth(18) = 300
                                .ColWidth(19) = 800
                                
                                .ColWidth(20) = 250
                                .ColWidth(21) = 250
                                
                                .ColWidth(22) = 250
                                .ColWidth(23) = 450
                                
                                .ColWidth(24) = 250
                                .ColWidth(25) = 700
                                .ColWidth(26) = 700
                                
                                .ColWidth(27) = 1200
                                .ColWidth(28) = 450
                                
                                .ColWidth(29) = 700
                                .ColWidth(30) = 700
                                .ColWidth(31) = 700
                                .ColWidth(32) = 700
                                .ColWidth(33) = 700
                                
                                .ColWidth(34) = 1000
                                .ColWidth(35) = 600
                                .ColWidth(36) = 600
                                .ColWidth(37) = 350
                                
                                'Form1.MSFlexGrid1.TextMatrix(0, 1) = "Cod Int"
                                .TextMatrix(0, 1) = "O"
                                .TextMatrix(0, 2) = "AÇ"
                                .TextMatrix(0, 3) = "S"
                                .TextMatrix(0, 4) = "Código de Barra"
                                .TextMatrix(0, 5) = "D"
                                .TextMatrix(0, 6) = "Código Interno"
                                .TextMatrix(0, 7) = "D"
                                .TextMatrix(0, 8) = "Descrição Completa"
                                .TextMatrix(0, 9) = "Descrição Reduzida"
                                .TextMatrix(0, 10) = "Código Vazilhame"
                                .TextMatrix(0, 11) = "D"
                                .TextMatrix(0, 12) = "Preço Venda"
                                .TextMatrix(0, 13) = "Preço Custo"
                                
                                .TextMatrix(0, 14) = "Validade"
                                .TextMatrix(0, 15) = "AL Pdv"
                                .TextMatrix(0, 16) = "AL NF"
                                
                                .TextMatrix(0, 17) = "Cod Emb"
                                .TextMatrix(0, 18) = "Desc Emb"
                                .TextMatrix(0, 19) = "Qtde Emb"
                                .TextMatrix(0, 20) = "Peso Variável"
                                .TextMatrix(0, 21) = "Vende Fracionária"
                                
                                .TextMatrix(0, 22) = "Tipo Etiq Gôndola"
                                .TextMatrix(0, 23) = "Qtde Etiq Gôndola"
                                .TextMatrix(0, 24) = "Tipo Etiq Produto"
                                
                                .TextMatrix(0, 25) = "Ini Promoção"
                                .TextMatrix(0, 26) = "Fim Promoção"
                                .TextMatrix(0, 27) = "Pç Promoção"
                                .TextMatrix(0, 28) = "TP Promoção"
                                
                                .TextMatrix(0, 29) = "Merc 1"
                                .TextMatrix(0, 30) = "Merc 2"
                                .TextMatrix(0, 31) = "Merc 3"
                                .TextMatrix(0, 32) = "Merc 4"
                                .TextMatrix(0, 33) = "Merc 5"
                                
                                .TextMatrix(0, 34) = "Fornecedor"
                                .TextMatrix(0, 35) = "Faixa Pç"
                                .TextMatrix(0, 36) = "Margem Teórica"
                                .TextMatrix(0, 37) = "ST"
                                .Rows = 1
                            End With
                        End If
                            
                        
                    
                        flx.AddItem ""
                        
                        flx.TextMatrix(flx.Rows - 1, 0) = Mid(lstrLinha, 1, 1)
                        flx.TextMatrix(flx.Rows - 1, 1) = Mid(lstrLinha, 2, 1)
                        flx.TextMatrix(flx.Rows - 1, 2) = Mid(lstrLinha, 3, 1)
                        flx.TextMatrix(flx.Rows - 1, 3) = Mid(lstrLinha, 4, 1)
                        
                        flx.TextMatrix(flx.Rows - 1, 4) = Mid(lstrLinha, 5, 13)
                        flx.TextMatrix(flx.Rows - 1, 5) = Mid(lstrLinha, 18, 1)
                        flx.TextMatrix(flx.Rows - 1, 6) = Mid(lstrLinha, 19, 10)
                        flx.TextMatrix(flx.Rows - 1, 7) = Mid(lstrLinha, 29, 1)
                        flx.TextMatrix(flx.Rows - 1, 8) = Mid(lstrLinha, 30, 35)
                        flx.TextMatrix(flx.Rows - 1, 9) = Mid(lstrLinha, 65, 15)
                        flx.TextMatrix(flx.Rows - 1, 10) = Mid(lstrLinha, 80, 13)
                        
                        flx.TextMatrix(flx.Rows - 1, 11) = Mid(lstrLinha, 93, 1)
                        flx.TextMatrix(flx.Rows - 1, 12) = Mid(lstrLinha, 94, 11)
                        flx.TextMatrix(flx.Rows - 1, 13) = Mid(lstrLinha, 105, 11)
                        flx.TextMatrix(flx.Rows - 1, 14) = Mid(lstrLinha, 116, 4)
                        
                        flx.TextMatrix(flx.Rows - 1, 15) = Mid(lstrLinha, 120, 2)
                        flx.TextMatrix(flx.Rows - 1, 16) = Mid(lstrLinha, 122, 2)
                        flx.TextMatrix(flx.Rows - 1, 17) = Mid(lstrLinha, 124, 2)
                        
                        flx.TextMatrix(flx.Rows - 1, 18) = Mid(lstrLinha, 126, 3)
                        flx.TextMatrix(flx.Rows - 1, 19) = Mid(lstrLinha, 129, 7)
                        flx.TextMatrix(flx.Rows - 1, 20) = Mid(lstrLinha, 136, 1)
                        flx.TextMatrix(flx.Rows - 1, 21) = Mid(lstrLinha, 137, 1)
                        flx.TextMatrix(flx.Rows - 1, 22) = Mid(lstrLinha, 138, 2)
                        flx.TextMatrix(flx.Rows - 1, 23) = Mid(lstrLinha, 140, 3)
                        flx.TextMatrix(flx.Rows - 1, 24) = Mid(lstrLinha, 143, 2)
                        
                        flx.TextMatrix(flx.Rows - 1, 25) = Mid(lstrLinha, 145, 8)
                        flx.TextMatrix(flx.Rows - 1, 26) = Mid(lstrLinha, 153, 8)
                        flx.TextMatrix(flx.Rows - 1, 27) = Mid(lstrLinha, 161, 11)
                        flx.TextMatrix(flx.Rows - 1, 28) = Mid(lstrLinha, 172, 3)
                        
                        flx.TextMatrix(flx.Rows - 1, 29) = Mid(lstrLinha, 175, 5)
                        flx.TextMatrix(flx.Rows - 1, 30) = Mid(lstrLinha, 180, 5)
                        flx.TextMatrix(flx.Rows - 1, 31) = Mid(lstrLinha, 185, 5)
                        flx.TextMatrix(flx.Rows - 1, 32) = Mid(lstrLinha, 190, 5)
                        flx.TextMatrix(flx.Rows - 1, 33) = Mid(lstrLinha, 195, 5)
                        
                        flx.TextMatrix(flx.Rows - 1, 34) = Mid(lstrLinha, 200, 8)
                        flx.TextMatrix(flx.Rows - 1, 35) = Mid(lstrLinha, 208, 4)
                        flx.TextMatrix(flx.Rows - 1, 36) = Mid(lstrLinha, 208, 5)
                        flx.TextMatrix(flx.Rows - 1, 37) = Mid(lstrLinha, 213, 2)
                        
                        ctLinha = ctLinha + 1
                        
                    Loop
                    Form1.Frame1.Visible = False
                    flx.Visible = True
            End Select
                
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


