VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmZebra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zebra"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   825
      Left            =   7710
      TabIndex        =   7
      Top             =   3750
      Width           =   3015
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   975
      Left            =   5370
      TabIndex        =   6
      Top             =   2070
      Width           =   2355
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   1095
      Left            =   2970
      TabIndex        =   5
      Top             =   3570
      Width           =   2925
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9840
      Top             =   2670
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   1005
      Left            =   360
      TabIndex        =   4
      Top             =   3540
      Width           =   2355
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1035
      Left            =   420
      TabIndex        =   3
      Top             =   600
      Width           =   2355
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1245
      Left            =   2880
      TabIndex        =   2
      Top             =   1950
      Width           =   2085
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1275
      Left            =   420
      TabIndex        =   1
      Top             =   1950
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1035
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "frmZebra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Private Sub Command1_Click()
    Dim rstImpressora As New ADODB.Recordset
    
    strSql = "SELECT DFCaminho_impressora_via_porta_TBParametros_gerais FROM TBParametros_Gerais WHERE PFKCodigo_TBEmpresa = 100"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstImpressora, "Otica", Me
    
    If IsNull(rstImpressora.Fields("DFCaminho_impressora_via_porta_TBParametros_gerais")) = True Then
       strCaminho_impressora = ""
    Else
       strCaminho_impressora = rstImpressora.Fields("DFCaminho_impressora_via_porta_TBParametros_gerais")
    End If
    
    Set rstImpressora = Nothing
    
    strLinha_Impressao = "-----------------------------------------------------------"
    'MANDANDO O COMANDO DIRETO NA PORTA DA IMPRESSORA
    Funcoes_Gerais.Abre_porta_impressora_via_LPT1 (strCaminho_impressora)
    Print #1, Chr(27) & Chr(15) + strLinha_Impressao 'IMPRIME CARACTER COMPRIMIDO PARA MATRICIAIS EPSON
    'Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
    
    Print #1, "Q184,24" 'Q184 > 184 significa 184 dots, 1 mm = 8 Dots é a altura da etiqueta 184 Dots = 23 mm (2,3 cm) , 24 dots espaço entre etiquetas
    Print #1, "q831"
    Print #1, "rN"
    Print #1, "S9"      'Determina a velocidade da impressão
    Print #1, "D10"     'Determina o fator de escuridao da etiqueta
    Print #1, "ZT"      'Determina a sequencia de impressão T = Top B = Button
    Print #1, "JB"      'Disable Top Of Form Backup
    Print #1, "OD"
    Print #1, "R20,0"   'Determina a margem da impressora
    Print #1, "N"       'Limpa a memoria da impressora a cada nova impressao
    
    Funcoes_Gerais.Fecha_porta_impressora_via_LPT1
End Sub

Private Sub Command2_Click()
    
    Open "\\marcos-pc\ZDesigner S4M-203dpi ZPL" For Output As #1
    
    'Print #1, "N"   'Limpa a memoria da impressora a cada nova impressao
    'Print #1, "D10" 'Determina o fator de escuridao da etiqueta
    'Print #1, "P1"

    For i = 0 To 9
            Print #1, "Q184,24" 'Q184 > 184 significa 184 dots, 1 mm = 8 Dots é a altura da etiqueta 184 Dots = 23 mm (2,3 cm) , 24 dots espaço entre etiquetas
            Print #1, "q831"
            Print #1, "rN"
            Print #1, "S9"      'Determina a velocidade da impressão
            Print #1, "D10"     'Determina o fator de escuridao da etiqueta
            Print #1, "ZT"      'Determina a sequencia de impressão T = Top B = Button
            Print #1, "JB"      'Disable Top Of Form Backup
            Print #1, "OD"
            Print #1, "R20,0"   'Determina a margem da impressora
            Print #1, "N"       'Limpa a memoria da impressora a cada nova impressao
            
            'A0 > COLUNA
            '0  > LINHA
            '0  > ROTAÇÃO
            '3  > TIPO DE FONTE
            '1  > MUTIPLICADOR ALTURA CARACTERES
            '1  > MUTIPLICADOR LARGURA CARACTERES
            'N  >
            
            'Primeira etiqueta
            Print #1, "A0,0,0,3,2,2,N," & Chr(34) & "2" & Chr(34)   'Preço
            Print #1, "B25,50,0,1,2,5,72,N," & Chr(34) & "12345669" & Chr(34)
            Print #1, "A25,130,0,3,1,1,N," & Chr(34) & "12345669" & " " & "456789" & Chr(34)   'Preço
            
            'Segunda etiqueta
            Print #1, "A275,0,0,3,2,2,N," & Chr(34) & "2" & Chr(34)   'Preço
            Print #1, "B295,50,0,1,2,5,72,N," & Chr(34) & "12345669" & Chr(34)
            Print #1, "A295,130,0,3,1,1,N," & Chr(34) & "12345669" & " " & "456789" & Chr(34)   'Preço
    
            'Terceira etiqueta
            Print #1, "A550,0,0,3,2,2,N," & Chr(34) & "2" & Chr(34)   'Preço
            Print #1, "B570,50,0,1,2,5,72,N," & Chr(34) & "12345669" & Chr(34)
            Print #1, "A570,130,0,3,1,1,N," & Chr(34) & "12345669" & " " & "456789" & Chr(34)   'Preço
    
            Print #1, "P" & "35"
        
    Next i

    Close #1
End Sub

Private Sub Command3_Click()
    
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 12
    Printer.FontBold = False
    Printer.CurrentX = 800
    Printer.CurrentY = 80
    'Printer.ScaleMode
    Printer.Orientation = 2
    Printer.Print "MARCOS"
'
'    Printer.Font.Name = "owcode128c"
'    Printer.CurrentX = 800
'    Printer.FontItalic = False
'    Printer.Print "7892840231149"
'
'    Printer.Font.Name = "Tahoma"
'    Printer.Font.Size = 8
'    Printer.FontBold = True
'    Printer.FontItalic = True
'    Printer.CurrentX = 800
'    Printer.Print "7892840231149"
'
'    Printer.Font.Name = "Circled"
'    Printer.Font.Size = 12
'    Printer.FontBold = True
'    Printer.FontItalic = True
'    Printer.CurrentX = 800
'    Printer.Print "3"
'    Printer.CurrentX = 800
'    Printer.Print "17"
'    Printer.CurrentX = 800
'    Printer.Print "30"

    
'    Printer.Print "0156"
'    Printer.Print "Q184,24" 'Q184 > 184 significa 184 dots, 1 mm = 8 Dots é a altura da etiqueta 184 Dots = 23 mm (2,3 cm) , 24 dots espaço entre etiquetas
'    Printer.Print "q831"
'    Printer.Print "rN"
'    Printer.Print "S9"      'Determina a velocidade da impressão
'    Printer.Print "D10"     'Determina o fator de escuridao da etiqueta
'    Printer.Print "ZT"      'Determina a sequencia de impressão T = Top B = Button
'    Printer.Print "JB"      'Disable Top Of Form Backup
'    Printer.Print "OD"
'    Printer.Print "R10,0"   'Determina a margem da impressora
'    Printer.Print "N"       'Limpa a memoria da impressora a cada nova impressao
'
    Printer.EndDoc
    
End Sub

Private Sub Command4_Click()
'      Comando = "^N"
'      Printer.Print Comando
'
'      strDescricao = "^FO20,05^AD,50,15^FD" & "Marcos2" & "^FS"
'      StrCod_Barras = "^FO20,80^BEN,60^FD" & "2425" & "^FS"
'      StrPreco = "^FO350,80^AD,80,25^FDR$ " & Format(CDbl("2"), "#########0.00") & "^FS"
'
'      Comando = "^XA" & Chr(13) & _
'               "^LH30,10" & Chr(13) & _
'               strDescricao & Chr(13) & _
'               StrCod_Barras & Chr(13) & _
'               StrPreco & Chr(13) & _
'               "^XZ"
'
'
'     'Envia o Comando para a Impressora
'     Printer.Print Comando
'     VB.Printer.EndDoc
'     Printer.KillDoc
'
     
    Open "\\marcos-pc\ZDesigner S4M-203dpi ZPL" For Output As #1
    
'    Print #1, "^XA"
'    Print #1, "^PR1"
'    Print #1, "^FO100,100"
'    Print #1, "^GB70,70,70,,3^FS"
'    Print #1, "^FO200,100"
'    Print #1, "^GB70,70,70,,3^FS"
'    Print #1, "^FO300,100"
'    Print #1, "^GB70,70,70,,3^FS"
'    Print #1, "^FO400,100"
'    Print #1, "^GB70,70,70,,3^FS"
'    Print #1, "^FO107,110^CF0,70,93"
'    Print #1, "^FR^FDREVERSE^FS"
'    Print #1, "^XZ"
        
    'a = "^XA^LH0,0^LL113^FDTESTE DE IMPRESSÃO^FS"
    'Print #1, "^FO150,90^A0N,25,20^FDZebra Technologies^FS"
    'Print #1, "^FO50,50"
    'Print #1, "^FO150,115^A0N,25,20^FD333 Corporate Woods Parkway^FS"
    'Print #1, "^GB300,200,10^FS"
    'b = "^FDTESTE DE IMPRESSÃO^FS"
   
    'Print #1, "^XA^FDTESTE DE IMPRESSÃO^FS"
    
      Data1 = "^XA" & _
              "^FDTESTE DE IMPRESSÃO^FS" & _
              "^XZ"
    
    
'    Print #1, "^CI0"
'    Print #1, "^FO1,50"
'    Print #1, "^FB546,1,0,C"
'    Print #1, "^A0N,20,15"
'    Print #1, "^FD{{BAIÃO}}"
'    Print #1, "^FS"
'
'    Print #1, "^FO1,160"
'    Print #1, "^FB546,2,0,L"
'    Print #1, "^A0N,20,15"
'    Print #1, "^FDLOTE: {{Lote}} FAB: {{Dt Fabricação}} VAL: {{Dt Validade}} Reg MS: {{Registro Anvisa}}"
'    Print #1, "^FS"
'
'    Print #1, "^FO298,205"
'    Print #1, "^BEN,30,Y,N"
'    Print #1, "^FD{{Cod Barra}}"
'    Print #1, "^FS"
    Print #1, Data1
    Close #1
     
    
End Sub

Private Sub Command5_Click()

   Data1 = "^XA^LH0,0^LL113" & _
           "^FO89,20^A0N,25,18 ^FDWaranty void if removed^1,1^FS" & _
           "^FO93,45^A0N,24,17^FDGarantie Annule si enleve^1,1^FS" & _
           "^FO117,75^AB^FD" & SNum & "^1,1^FS" & _
           "^XZ"
    
    Open "\\OTACOM-10\ZDesigner S4M-203dpi ZPL" For Output As #1
    Print #1, Data1
    Close #1
    
End Sub

Private Sub Command6_Click()

    Data1 = "^XA" & _
            "^FWR" & _
            "^FO150,90^A0N,25,20^FDFDZebra Technologies^FS" & _
            "^FO115,75^A0,25,20^FD0123456789^FS" & _
            "^FO150,115^A0N,25,20^FD333 Corporate Woods Parkway^FS" & _
            "^FO400,75^A0,25,20^FDXXXXXXXXX^FS" & _
            "^XZ"

            
    Open "\\OTACOM-10\ZDesigner S4M-203dpi ZPL" For Output As #1
    Print #1, Data1
    Close #1
End Sub

Private Sub Command7_Click()
    
    strReferencia = "MPGLC00"
    strLote = "701631"
    strProduto = "GLICEROL DA SILVA SAURO DA MANDIBULA"
    strDestino = "RUA DIOGENES PADILHA, 106 "
    strQtde = "100"
    strPesoBruto = "1.2356"
    strTara = "2.80"
    strFabricacao = "01/01/2017"
    strVal_Int = "01/01/2017"
    strValidade = "01/01/2019"
    strDataAtual = "09/05/2017"
'
'    Data1 = "^XA^MD20" & _
'            "^CI0" & _
'            "^FO20,35^A0N,22,20^FDREFERENCIA: " & strReferencia & "^FS" & _
'            "^FO110,35^A0N,22,20^FDLOTE:  " & strLote & "^FS" & _
'            "^FO40,35^A0N,22,22^FDPRODUTO: " & strProduto & "^FS" & _
'            "^FO60,35^A0N,22,22^FDDESTINO: " & strDestino & "^FS" & _
'            "^FO80,35^A0N,22,22^FDQtde.: " & strQtde & "^FS" & _
'            "^FO80,85^A0N,22,22^FDP.Bruto: " & strPesoBruto & "^FS" & _
'            "^FO80,135^A0N,22,22^FDTara: " & strTara & "^FS" & _
'            "^FO100,35^A0N,22,22^FDFab.: " & strFabricacao & "^FS" & _
'            "^FO100,85^A0N,22,22^FDVal. Int.: " & strVal_Int & "^FS" & _
'            "^FO100,135^A0N,22,22^FDVal: " & strValidade & "^FS" & _
'            "^FO120,35^A0N,22,22^FDData.: " & strDataAtual & "^FS" & _
'            "^FO120,85^A0N,22,22^FDvisto/Resp^FS" & _
'            "^XZ"

    Data1 = "^XA^MD20" & _
            "^CI0" & _
            "^FO20,35^A0N,38,30^FDPRODUTO: " & strProduto & "^FS" & _
            "^FO20,85^A0N,30,30^FDREFERENCIA: " & strReferencia & "^FS" & _
            "^FO380,85^A0N,30,30^FDLOTE:  " & strLote & "^FS" & _
            "^FO20,115^A0N,30,30^FDQTDE.: " & strQtde & "^FS" & _
            "^FO380,115^A0N,28,28^FDP.BRUTO: " & strPesoBruto & "^FS" & _
            "^FO660,115^A0N,28,28^FDTARA: " & strTara & "^FS" & _
            "^FO20,145^A0N,30,30^FDFAB.: " & strFabricacao & "^FS" & _
            "^FO380,145^A0N,25,25^FDVAL. Int.: " & strVal_Int & "^FS" & _
            "^FO660,145^A0N,25,25^FDVAL: " & strValidade & "^FS" & _
            "^FO20,230^A0N,35,30^FDDESTINO: " & strDestino & "^FS" & _
            "^FO430,300^A0N,35,30^FD___________________________^FS" & _
            "^FO20,350^A0N,30,30^FDDATA.: " & strDataAtual & "^FS" & _
            "^FO550,350^A0N,30,30^FDVISTO RESP^FS" & _
            "^XZ"
    
    Open "\\OTACOM-10\ZDesigner S4M-203dpi ZPL" For Output As #1
    Print #1, Data1
    Close #1
        
End Sub

Private Sub Command8_Click()
    
    strCodigoCliente = "1174"
    strCliente = "PREFEITURA MUNICIPAL DE PIRAPOZINHO"
    strEndereco = "RUA DIOGENES PADILHA"
    strComplementto = "BLOCO 5 AP 201"
    strNumero = "106"
    strBairro = "CENTRO"
    strCidade = "TRES RIOS"
    strEstado = "RJ"
    strCep = "25807-010"
    strNota = "125897"
    
    strLinha1 = strCodigoCliente & " - " & strCliente
    strLinha2 = strEndereco & " , " & strComplementto & " , " & strNumero
    strLinha3 = strBairro & " - " & strCidade & " - " & strEstado & " - " & strCep
    strLinha4 = "NOTA: " & strNota
    
    Data1 = "^XA^MD20" & _
            "^CI0" & _
            "^FO20,80^A0N,38,30^FD" & strLinha1 & "^FS" & _
            "^FO20,140^A0N,38,30^FD" & strLinha2 & "^FS" & _
            "^FO20,180^A0N,38,30^FD" & strLinha3 & "^FS" & _
            "^FO20,240^A0N,38,30^FD" & strLinha4 & "^FS" & _
            "^XZ"
    
    Open "\\OTACOM-10\ZDesigner S4M-203dpi ZPL" For Output As #1
    Print #1, Data1
    Close #1
End Sub
