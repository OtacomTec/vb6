VERSION 5.00
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
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1035
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   2955
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1275
      Left            =   6390
      TabIndex        =   2
      Top             =   3390
      Width           =   4305
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1275
      Left            =   1170
      TabIndex        =   1
      Top             =   2490
      Width           =   4305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1065
      Left            =   5040
      TabIndex        =   0
      Top             =   600
      Width           =   4065
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
    
    Open "\\MARCOS-PC\ZDesigner S4M-203dpi ZPL" For Output As #1
    
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
    Printer.CurrentY = 500
    Printer.Print "MARCOS"

    Printer.Font.Name = "owcode128c"
    Printer.CurrentX = 800
    Printer.FontItalic = False
    Printer.Print "7892840231149"
    
    Printer.Font.Name = "Tahoma"
    Printer.Font.Size = 8
    Printer.FontBold = True
    Printer.FontItalic = True
    Printer.CurrentX = 800
    Printer.Print "7892840231149"
    
    Printer.Font.Name = "Circled"
    Printer.Font.Size = 12
    Printer.FontBold = True
    Printer.FontItalic = True
    Printer.CurrentX = 800
    Printer.Print "3"
    Printer.CurrentX = 800
    Printer.Print "17"
    Printer.CurrentX = 800
    Printer.Print "30"

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
      Comando = "^N"
      Printer.Print Comando
      
      strDescricao = "^FO20,05^AD,50,15^FD" & "Marcos2" & "^FS"
      StrCod_Barras = "^FO20,80^BEN,60^FD" & "2425" & "^FS"
      StrPreco = "^FO350,80^AD,80,25^FDR$ " & Format(CDbl("2"), "#########0.00") & "^FS"
      
      Comando = "^XA" & Chr(13) & _
               "^LH30,10" & Chr(13) & _
               strDescricao & Chr(13) & _
               StrCod_Barras & Chr(13) & _
               StrPreco & Chr(13) & _
               "^XZ"
      
     
     'Envia o Comando para a Impressora
     Printer.Print Comando
     VB.Printer.EndDoc
     Printer.KillDoc
     
    
End Sub
