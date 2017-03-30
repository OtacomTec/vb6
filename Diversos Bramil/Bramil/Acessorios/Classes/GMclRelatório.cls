VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GM_clRelatório"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Enum GM_enAlinhamento
    GM_enEsquerdo = 0
    GM_enCentro = 1
    GM_enDireito = 2
    
End Enum
Private mvarMargemSuperior As Long 'local copy
Private mvarMargemEsquerda As Long 'local copy
Private mvarPx As Long 'local copy
Private mvarPy As Long 'local copy
Private mvarNúmeroDaPágina As Long 'local copy
Private mvarLarguraDaImpressão As Long 'local copy
Private mvarFonte As String 'local copy
Public Sub Abortar()
    Printer.KillDoc
End Sub

Public Property Let Fonte(ByVal vData As String)
    mvarFonte = vData
End Property

Public Property Get Fonte() As String
    Fonte = mvarFonte
End Property

Public Property Let LarguraDaImpressão(ByVal vData As Long)
    mvarLarguraDaImpressão = vData
End Property

Public Property Get LarguraDaImpressão() As Long
    LarguraDaImpressão = mvarLarguraDaImpressão
End Property
Public Property Get NúmeroDaPágina() As Long
    NúmeroDaPágina = mvarNúmeroDaPágina
End Property

Public Property Let Py(ByVal vData As Long)
    mvarPy = vData
End Property

Public Property Get Py() As Long
    Py = mvarPy
End Property

Public Property Let Px(ByVal vData As Long)
    mvarPx = vData
End Property

Public Property Get Px() As Long
    Px = mvarPx
End Property

Public Property Let MargemEsquerda(ByVal vData As Long)
    mvarMargemEsquerda = vData
End Property

Public Property Get MargemEsquerda() As Long
    MargemEsquerda = mvarMargemEsquerda
End Property

Public Property Let MargemSuperior(ByVal vData As Long)
    mvarMargemSuperior = vData
End Property

Public Property Get MargemSuperior() As Long
    MargemSuperior = mvarMargemSuperior
End Property

Sub CabeçalhoPágina(NomeDoRelatório_E As String, _
                         DataDoRelatório_E As Date, _
                         Optional NomeDoGrupo_C As String = "Grupo Mil", _
                         Optional TítuloDoRelatório_C As String = "Relatório", _
                         Optional SubTítuloDoRelatório_C As String)
    
    
    mvarPy = mvarMargemSuperior
    mvarPx = mvarMargemEsquerda
    Printer.FontName = mvarFonte
    lsAlturaDaLinha = 240
    
    '>>>>>>>>>>>>Linha 1 do Cabeçalho
    Printer.CurrentX = mvarMargemEsquerda
    Printer.CurrentY = mvarMargemSuperior
        Printer.Print "Versão"
        
    lsLarguraDoTexto = Printer.TextWidth(NomeDoGrupo_C)
    mvarPx = (4850 - (lsLarguraDoTexto / 2)) + mvarMargemEsquerda
    Printer.CurrentX = mvarPx
    Printer.CurrentY = mvarPy
        Printer.Print NomeDoGrupo_C
    
    mvarPx = mvarMargemEsquerda + 9000
    Printer.CurrentX = mvarPx
    Printer.CurrentY = mvarPy
        Printer.Print "Página " & NúmeroDaPágina
        
    '>>>>>>>>>Linha 2 do Cabeçalho
    mvarPy = mvarMargemSuperior + lsAlturaDaLinha
    mvarPx = mvarMargemEsquerda
    Printer.CurrentX = mvarPx
    Printer.CurrentY = mvarPy
        Printer.Print NomeDoRelatório_E
        
    lsLarguraDoTexto = Printer.TextWidth(TítuloDoRelatório_C)
    mvarPx = 4850 - (lsLarguraDoTexto / 2) + mvarMargemEsquerda
    Printer.CurrentX = mvarPx
    Printer.CurrentY = mvarPy
        Printer.Print TítuloDoRelatório_C
    
    mvarPx = mvarMargemEsquerda + 9000
    Printer.CurrentX = mvarPx
    Printer.CurrentY = mvarPy
        Printer.Print "Data: " & Format(Date, "dd/mm/yyyy")
        
    ' <<<<<<<<<<<<<
    
    '>>>>>>>>>Linha 3 do Cabeçalho
    mvarPy = mvarPy + lsAlturaDaLinha
    mvarPx = mvarMargemEsquerda
    Printer.CurrentX = mvarPx
    Printer.CurrentY = mvarPy
        Printer.Print Format(DataDoRelatório_E, "dd/mm/yy")
        
    lsLarguraDoTexto = Printer.TextWidth(SubTítuloDoRelatório_C)
    mvarPx = 4850 - (lsLarguraDoTexto / 2) + mvarMargemEsquerda
    Printer.CurrentX = mvarPx
    Printer.CurrentY = mvarPy
        Printer.Print SubTítuloDoRelatório_C
    
    mvarPx = mvarMargemEsquerda + 9000
    Printer.CurrentX = mvarPx
    Printer.CurrentY = mvarPy
        Printer.Print "Hora: "; Format(Time, "hh:mm:ss")
    ' <<<<<<<<<<<<<
        
    '>>>>>>>>>Linha 4 do Cabeçalho
    mvarPy = mvarPy + 300
    Printer.CurrentX = mvarMargemEsquerda
    Printer.CurrentY = mvarPy
    Printer.Line (mvarMargemEsquerda, mvarPy)- _
                 (mvarLarguraDaImpressão + mvarMargemEsquerda, mvarPy)
    '<<<<<<<<<<
End Sub

Function X()
       
    Dim lsLinhaAtual As Single
    Dim lsAlturaDaLinha As Single
    Dim lsLargura As Single
    
    lsAlturaDaLinha = 250
    
    Printer.Font.Name = "Draft 12cpi"
    Printer.Font.Size = 10
    '>>>>>>>>>Linha 1 do Cabeçalho


    'Printer.EndDoc
    'Exit Sub
        
    '>>>>>>>>>Linha 5 do Cabeçalho
    lsLinhaAtual = lsLinhaAtual + 150
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda
    Printer.Print "Status do Estoque: " & LabelControle.Caption
    
    lsLargura = Printer.TextWidth("Código de Barra: " & TextdCodBarraProdutotPrdBar.Text)
    lsLargura = 11300 - lsLargura
    Printer.CurrentX = lsLargura + mvarMargemEsquerda: Printer.CurrentY = lsLinhaAtual
    Printer.Print "Código de Barra: " & TextdCodBarraProdutotPrdBar.Text
    '<<<<<<<<<<<<<<
    
    '>>>>>>>>>Linha 6 do Cabeçalho
    lsLinhaAtual = lsLinhaAtual + lsAlturaDaLinha
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda
    Printer.Print "Estoque Calculado:" & TextEstoqueCalculado
    
    lsLinhaAtual = lsLinhaAtual + lsAlturaDaLinha
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda
    Printer.Print "Estoque Checado:" & TextEstoqueChecado.Text
    '<<<<<<<<<<<<<<<<<<<<<<<
    
    
    '>>>>>>>>>>>> Linha 7
    lsLinhaAtual = lsLinhaAtual + lsAlturaDaLinha
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda
    Printer.Print "Data"
    
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda + 1000
    Printer.Print "Tipo Operação"
    
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda + 4000
    Printer.Print "Entrada"
    
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda + 5000
    Printer.Print "Saída"
    
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda + 6000
    Printer.Print "Est Calc"
    
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda + 7000
    Printer.Print "Prev Calc"
    
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda + 8000
    Printer.Print "Est Chk"
    
    Printer.CurrentY = lsLinhaAtual
    Printer.CurrentX = mvarMargemEsquerda + 9000
    Printer.Print "Prev Chk"
    '<<<<<<<<<<<<<<<<<<<<<
    
    '>>>>>>>>>Linha 8 do Cabeçalho
    lsLinhaAtual = lsLinhaAtual + 300
    Printer.CurrentX = mvarMargemEsquerda: Printer.CurrentY = lsLinhaAtual
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + 11300, Printer.CurrentY)
    '<<<<<<<<<<
    
       
    'lsAlturaDaLinha = 200
    Dim i
    For i = 1 To Me.MSFlexGrid1.Rows - 1
        lsLinhaAtual = lsLinhaAtual + lsAlturaDaLinha
        Printer.CurrentY = lsLinhaAtual
        Printer.CurrentX = mvarMargemEsquerda
        Printer.Print Me.MSFlexGrid1.TextMatrix(i, 1)
        
        Printer.CurrentY = lsLinhaAtual
        Printer.CurrentX = mvarMargemEsquerda + 1000
        Printer.Print Me.MSFlexGrid1.TextMatrix(i, 2)
        
        Printer.CurrentY = lsLinhaAtual
        Printer.CurrentX = mvarMargemEsquerda + 4000
        Printer.Print Me.MSFlexGrid1.TextMatrix(i, 3)
        
        Printer.CurrentY = lsLinhaAtual
        Printer.CurrentX = mvarMargemEsquerda + 5000
        Printer.Print Me.MSFlexGrid1.TextMatrix(i, 4)
        
        Printer.CurrentY = lsLinhaAtual
        Printer.CurrentX = mvarMargemEsquerda + 6000
        Printer.Print Me.MSFlexGrid1.TextMatrix(i, 5)
        
        Printer.CurrentY = lsLinhaAtual
        Printer.CurrentX = mvarMargemEsquerda + 7000
        Printer.Print Me.MSFlexGrid1.TextMatrix(i, 6)
        
        Printer.CurrentY = lsLinhaAtual
        Printer.CurrentX = mvarMargemEsquerda + 8000
        Printer.Print Me.MSFlexGrid1.TextMatrix(i, 7)
        
        Printer.CurrentY = lsLinhaAtual
        Printer.CurrentX = mvarMargemEsquerda + 9000
        Printer.Print Me.MSFlexGrid1.TextMatrix(i, 8)
    Next i
    Printer.EndDoc
End Function

Sub Imprime(Texto As String, _
            Optional PosiçãoY As Long = Py, _
            Optional PosiçãoX As Long = Px, _
            Optional Alinhamento As GM_enAlinhamento = GM_enDireito)
    
    mvarPy = Py
    mvarPx = Px
    Printer.CurrentY = mvarPy
    Select Case Alinhamento
        Case GM_enEsquerdo
            Printer.CurrentX = mvarPx
            Printer.Print Texto
        Case GM_enCentro
            LarguraDoTexto = Printer.TextWidth(Texto)
            mvarPx = mvarPx - (LarguraDoTexto / 2)
            Printer.CurrentX = mvarPx
            Printer.Print Texto
            
        Case GM_enDireito
            LarguraDoTexto = Printer.TextWidth(Texto)
            mvarPx = mvarPx - LarguraDoTexto
            Printer.CurrentX = mvarPx
            Printer.Print Texto
    End Select
    
    
End Sub

Private Sub Class_Initialize()
    mvarMargemEsquerda = 300
    mvarMargemSuperior = 400
    mvarLarguraDaImpressão = 11200
    mvarNúmeroDaPágina = 1
    mvarFonte = "Draft 12cpi"
    
End Sub
Function NovaPágina()
    Printer.NewPage
    mvarNúmeroDaPágina = mvarNúmeroDaPágina + 1
End Function
Function FimDaImpressão()
    Printer.EndDoc
End Function
