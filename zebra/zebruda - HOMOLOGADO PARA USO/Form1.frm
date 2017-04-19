VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   525
      Left            =   2940
      TabIndex        =   1
      Top             =   2340
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   840
      TabIndex        =   0
      Top             =   990
      Width           =   2325
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
      
    Printer.Print "Q184,24" 'Q184 > 184 significa 184 dots, 1 mm = 8 Dots é a altura da etiqueta 184 Dots = 23 mm (2,3 cm) , 24 dots espaço entre etiquetas
    Printer.Print "q831"
    Printer.Print "rN"
    Printer.Print "S9"      'Determina a velocidade da impressão
    Printer.Print "D10"     'Determina o fator de escuridao da etiqueta
    Printer.Print "ZT"      'Determina a sequencia de impressão T = Top B = Button
    Printer.Print "JB"      'Disable Top Of Form Backup
    Printer.Print "OD"
    Printer.Print "R20,0"   'Determina a margem da impressora
    Printer.Print "N"       'Limpa a memoria da impressora a cada nova impressao
    'VB.Printer.EndDoc
      'Comando = "^N"
      'Printer.Print Comando
      
      
      strDescricao = "FO20,05AD,50,15FD" & "OSAMA" & "FS"
      Printer.Print strDescricao
'      StrCod_Barras = "FO20,80^BEN,60^FD" & "2425" & "^FS"
'      StrPreco = "^FO350,80^AD,80,25^FDR$ " & Format(CDbl("2"), "#########0.00") & "^FS"
      
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
     VB.Printer.EndDoc
     Printer.KillDoc
     
End Sub

Private Sub Command2_Click()
    
    Dim I As Integer, WCOD_BARRA As String, WCOD_AUX As String, WPRECO As String, WQTDE As Single
   
    WCOD_BARRA = "123456789"
    WCOD_AUX = "123456879"
    WPRECO = "123,99"
    WQTDE = 100
    
    Open "\\marcos-pc\ZDesigner S4M-203dpi ZPL" For Output As #1
    
    'Print #1, "N"   'Limpa a memoria da impressora a cada nova impressao
    'Print #1, "D10" 'Determina o fator de escuridao da etiqueta
    'Print #1, "P1"

    For I = 0 To 9
        'If Val(Text4(I)) > 0 Then
        
'            WCOD_BARRA = Mid(Text1(I), 1, 4)
'
'            WCOD_AUX = Mid(Text1(I), 5, 4)
'            WPRECO = "R$ " & Format(CDbl(Text5(I)), "###,###0.00")
'
'
'            WQTDE = Int(Text4(I) / 3)
            
            Dim nValor As Double
            Dim StrDecimal As String
            nValor = Val(WPRECO + (I * 8)) / 3

            StrDecimal = (nValor - Int(nValor))

            If StrDecimal > 0 Then
                WQTDE = WQTDE + 1
            End If
            
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
            Print #1, "A0,0,0,3,2,2,N," & Chr(34) & WPRECO & Chr(34)   'Preço
            Print #1, "B25,50,0,1,2,5,72,N," & Chr(34) & WCOD_BARRA & Chr(34)
            Print #1, "A25,130,0,3,1,1,N," & Chr(34) & WCOD_BARRA & " " & WCOD_AUX & Chr(34)   'Preço
            
            'Segunda etiqueta
            Print #1, "A275,0,0,3,2,2,N," & Chr(34) & WPRECO & Chr(34)   'Preço
            Print #1, "B295,50,0,1,2,5,72,N," & Chr(34) & WCOD_BARRA & Chr(34)
            Print #1, "A295,130,0,3,1,1,N," & Chr(34) & WCOD_BARRA & " " & WCOD_AUX & Chr(34)   'Preço
    
            'Terceira etiqueta
            Print #1, "A550,0,0,3,2,2,N," & Chr(34) & WPRECO & Chr(34)   'Preço
            Print #1, "B570,50,0,1,2,5,72,N," & Chr(34) & WCOD_BARRA & Chr(34)
            Print #1, "A570,130,0,3,1,1,N," & Chr(34) & WCOD_BARRA & " " & WCOD_AUX & Chr(34)   'Preço
    
            Print #1, "P" & WQTDE
            
        'End If
    Next I

    Close #1
    
End Sub
