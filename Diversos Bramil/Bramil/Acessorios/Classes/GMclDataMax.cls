VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMclDataMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'PROGRAMADOR: Luis Henrique Borges
'
'Esta Classe deve ser utilizada com:
'Classe GMclINI
'DataMax.ini

Private Const GMct_strTRANSF_COMM = "9600,N,8,1"


Private SOH As String
Private STX As String
Private ESC As String
Private CR As String


Public Enum DataMax_Enum
    GMen_milímetros = 0
    GMen_polegadas = 1
End Enum

Private Type EtiquetaCódigoDataMax
    LarguraDoFormulário As Integer
    LarguraDaEtiqueta As Integer
    EtiquetasPorLinha As Byte
    MargemEsquerda As Byte
    EspaçoHorizEntreEtiquetas As Byte
    EspaçoVertiEntreEtiquetas As Byte
    LarguraDaImpressão As Byte
    AlturaDaImpressão As Byte
    InícioDaImpressãoX1 As Byte
    InícioDaImpressãoX2 As Byte
    InícioDaImpressãoX3 As Byte
End Type
Private EtiquetaCódigo As EtiquetaCódigoDataMax

Private mvarMsComm As Object 'local copy
Private mvarPortaCom As Byte 'local copy
Private mvarCancelar As Boolean 'local copy
Private mvarTransferência As String 'local copy
Private mvarUnidadeDeMedida As DataMax_Enum 'local copy
'local variable(s) to hold property value(s)
Private mvarEtiqCódLargura As Integer 'local copy
Private mvarEtiqCódAltura As Integer 'local copy
Private mvarEtiqCódFormLagura As Integer 'local copy
Private mvarEtiqCódEspaçoHorizontal As Integer 'local copy
Private mvarEtiqCódEspaçoVertical As Integer 'local copy
Private mvarEtiqCódPosição1 As Integer 'local copy
Private mvarEtiqCódPosição2 As Integer 'local copy
Private mvarEtiqCódPosição3 As Integer 'local copy
Private mvarEtiqCódAjuste As Integer 'local copy
Private mvarEtiqCódMargemEsquerda As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarEtiqGonLargura As Integer 'local copy
Private mvarEtiqGonAltura As Integer 'local copy
Private mvarEtiqGonImpressãoX As Integer 'local copy
Private mvarEtiqGonImpressãoY As Integer 'local copy

Public Function Pausa(Segundos As Integer)
    Dim lbHora As Byte
    Dim liMinutos As Integer
    Dim liSegundos As Integer
    
    liSegundos = Segundos
    
    If liSegundos > 3599 Then           'Calcula Hora
        'Calcula a hora cheia
        lbHora = liSegundos \ 3600
    End If
    liSegundos = liSegundos Mod 3600    'Segundos restantes
    
    If liSegundos > 59 Then             'Calcula Minutos
        'Calcula a minutos cheio
        liMinutos = liSegundos \ 60
    End If
    liSegundos = liSegundos Mod 60      'Segundos Restantes
        
    ldtPausa = CDate(Trim(Format(lbHora, "00:")) & _
                     Trim(Format(liMinutos, "00:")) & _
                     Trim(Format(liSegundos, "00")))
    
    ldtIncial = CDate(Format(Time, "hh:mm:ss"))
    Do While ldtCtr <= ldtPausa
        DoEvents
        ldtCtr = CDate(Format(Time - ldtIncial, "hh:mm:ss"))
        Debug.Print ldtCtr
    Loop
End Function

Public Function EstadoDaPorta() As Boolean
    On Error GoTo ErroDePorta
    MsComm.PortOpen = True
    EstadoDaPorta = True
    
    Exit Function
ErroDePorta:
    EstadoDaPorta = False
    
End Function
Public Function FecharPorta(Optional Modo As Boolean = True)
    
    MsComm.PortOpen = Modo

End Function

Public Property Let EtiqGonImpressãoY(ByVal vData As Integer)
    mvarEtiqGonImpressãoY = vData
End Property
Public Property Get EtiqGonImpressãoY() As Integer
    EtiqGonImpressãoY = mvarEtiqGonImpressãoY
End Property

Public Property Let EtiqGonImpressãoX(ByVal vData As Integer)
    mvarEtiqGonImpressãoX = vData
End Property
Public Property Get EtiqGonImpressãoX() As Integer
    EtiqGonImpressãoX = mvarEtiqGonImpressãoX
End Property

Public Property Let EtiqGonAltura(ByVal vData As Integer)
    mvarEtiqGonAltura = vData
End Property
Public Property Get EtiqGonAltura() As Integer
    EtiqGonAltura = mvarEtiqGonAltura
End Property

Public Property Let EtiqGonLargura(ByVal vData As Integer)
    mvarEtiqGonLargura = vData
End Property
Public Property Get EtiqGonLargura() As Integer
    EtiqGonLargura = mvarEtiqGonLargura
End Property


Public Property Let EtiqCódMargemEsquerda(ByVal vData As Integer)
    mvarEtiqCódMargemEsquerda = vData
End Property

Public Property Get EtiqCódMargemEsquerda() As Integer
    EtiqCódMargemEsquerda = mvarEtiqCódMargemEsquerda
End Property

Public Property Let EtiqCódAjuste(ByVal vData As Integer)
    mvarEtiqCódAjuste = vData
End Property
Public Property Get EtiqCódAjuste() As Integer
    EtiqCódAjuste = mvarEtiqCódAjuste
End Property

Public Property Let EtiqCódPosição3(ByVal vData As Integer)
    mvarEtiqCódPosição3 = vData
End Property
Public Property Get EtiqCódPosição3() As Integer
    EtiqCódPosição3 = mvarEtiqCódPosição3
End Property

Public Property Let EtiqCódPosição2(ByVal vData As Integer)
    mvarEtiqCódPosição2 = vData
End Property
Public Property Get EtiqCódPosição2() As Integer
    EtiqCódPosição2 = mvarEtiqCódPosição2
End Property

Public Property Let EtiqCódPosição1(ByVal vData As Integer)
    mvarEtiqCódPosição1 = vData
End Property
Public Property Get EtiqCódPosição1() As Integer
    EtiqCódPosição1 = mvarEtiqCódPosição1
End Property

Public Property Let EtiqCódEspaçoVertical(ByVal vData As Integer)
    mvarEtiqCódEspaçoVertical = vData
End Property
Public Property Get EtiqCódEspaçoVertical() As Integer
    EtiqCódEspaçoVertical = mvarEtiqCódEspaçoVertical
End Property

Public Property Let EtiqCódEspaçoHorizontal(ByVal vData As Integer)
    mvarEtiqCódEspaçoHorizontal = vData
End Property
Public Property Get EtiqCódEspaçoHorizontal() As Integer
    EtiqCódEspaçoHorizontal = mvarEtiqCódEspaçoHorizontal
End Property

Public Property Let EtiqCódFormLagura(ByVal vData As Integer)
    mvarEtiqCódFormLagura = vData
End Property
Public Property Get EtiqCódFormLagura() As Integer
    EtiqCódFormLagura = mvarEtiqCódFormLagura
End Property



Public Property Let EtiqCódAltura(ByVal vData As Integer)
    mvarEtiqCódAltura = vData
End Property
Public Property Get EtiqCódAltura() As Integer
    EtiqCódAltura = mvarEtiqCódAltura
End Property

Public Property Let EtiqCódLargura(ByVal vData As Integer)
    mvarEtiqCódLargura = vData
End Property
Public Property Get EtiqCódLargura() As Integer
    EtiqCódLargura = mvarEtiqCódLargura
End Property

Public Property Let UnidadeDeMedida(ByVal vData As DataMax_Enum)
    mvarUnidadeDeMedida = vData
    If mvarUnidadeDeMedida = GMen_milímetros Then EnviarComando STX & "m"
    If mvarUnidadeDeMedida = GMen_polegadas Then EnviarComando STX & "n"
End Property
Public Property Get UnidadeDeMedida() As DataMax_Enum
    UnidadeDeMedida = mvarUnidadeDeMedida
End Property

Public Property Let Transferência(ByVal vData As String)
    mvarTransferência = vData
    MsComm.Settings = vData
End Property
Public Property Get Transferência() As String
    Transferência = mvarTransferência
End Property


Public Function ConfigurarDataMax()
    Dim lclINI As GMclINI
    Set lclINI = New GMclINI
    'mvarMsComm
    mvarPortaCom = Val(lclINI.INILê(App.Path & "\DataMax.ini", "DataMax", "Porta"))
    mvarTransferência = lclINI.INILê(App.Path & "\DataMax.ini", "DataMax", "Transferência")
    mvarUnidadeDeMedida = Val(lclINI.INILê(App.Path & "\DataMax.ini", "DataMax", "Unidade de Medida"))
        
    mvarEtiqCódLargura = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Largura"))
    mvarEtiqCódAltura = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Altura"))
    mvarEtiqCódFormLagura = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Form Largura"))
    mvarEtiqCódMargemEsquerda = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Margem Esquerda"))
    mvarEtiqCódEspaçoHorizontal = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Espaço Horizontal"))
    mvarEtiqCódEspaçoVertical = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Espaço Vertical"))
    mvarEtiqCódPosição1 = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Posição 1"))
    mvarEtiqCódPosição2 = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Posição 2"))
    mvarEtiqCódPosição3 = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Posição 3"))
    mvarEtiqCódAjuste = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Código", "Ajuste"))
    
    mvarEtiqGonLargura = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Gôndola", "Largura"))
    mvarEtiqGonAltura = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Gôndola", "Altura"))
    mvarEtiqGonImpressãoX = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Gôndola", "Posição X"))
    mvarEtiqGonImpressãoY = Val(lclINI.INILê(App.Path & "\DataMax.ini", "Etiqueta de Gôndola", "Posição Y"))
End Function

Private Function RetornaParâmetro(strValor As String, Procura As String) As String

    
End Function


Public Sub EnviarComando(Comando As String)
    MsComm.Output = Comando
End Sub

Private Sub ConfigurarUnidadeDeMedida(Optional Unidade As DataMax_Enum = GMen_milímetros)
    Select Case Unidade
        Case 0
            MsComm.Output = GMct_STX & "m"
        Case 1
            MsComm.Output = GMct_STX & "n"
    End Select
End Sub

Function Cancelar()
    MsComm.Output = SOH & "C" & CR
    
End Function
Function AvançaEtiqueta(Optional Quantidade As Integer = 1)
    PreparaImpressão
    MsComm.Output = STX & "E" & Format(Quantidade, "0000") & CR
    ImprimeTexto "  "
    FinalizaImpressão
End Function

Public Property Let PortaCom(ByVal vData As Byte)
    mvarPortaCom = vData
End Property

Public Property Get PortaCom() As Byte
    PortaCom = mvarPortaCom
End Property

Public Property Set MsComm(ByVal vData As Object)
    Set mvarMsComm = vData
    
End Property


Public Property Get MsComm() As Object
    Set MsComm = mvarMsComm
End Property

Private Sub Class_Initialize()
    SOH = Chr$(1)
    STX = Chr$(2)
    ESC = Chr$(27)
    CR = Chr$(13)  'vbCrLf
    ConfigurarDataMax
End Sub
Function PreparaImpressão(Optional Cópias As Integer = "1")
    On Error GoTo ErroComm
        
    If MsComm.PortOpen = False Then
          'alterado 27/06/2001
        MsComm.Settings = mvarTransferência   '"9600,N,8,1"
        MsComm.CommPort = PortaCom 'usando a porta passada pelo parametro
        MsComm.InputLen = 0
        MsComm.PortOpen = True
    End If
    
    MsComm.Output = STX & "L" & CR
'    MsComm.Output = "H07" & CR
    MsComm.Output = "D11" & CR
    MsComm.Output = "Q" & Trim(Format(Cópias, "0000")) & CR
    UnidadeDeMedida = mvarUnidadeDeMedida
    
    Exit Function
    
ErroComm:
    Select Case Err
        Case 8005
            'msgbox "Utilizar este Painel de Controle.", vbInformation,
            Resume Next
    End Select
End Function
Function ImprimeTexto(Texto As String, _
                      Optional Orientação As String = "1", _
                      Optional Fonte As String = "5", _
                      Optional MultipicadorDeLargura = "10", _
                      Optional Ignore = "000", _
                      Optional PosiçãoY As String = "0000", _
                      Optional PosiçãoX As String = "0000")
    
    MsComm.Output = Orientação & _
                     Fonte & _
                     MultipicadorDeLargura & _
                     Ignore & _
                     PosiçãoY & _
                     PosiçãoX & _
                     Texto & CR
                     
End Function
Function FinalizaImpressão(Optional FechaPorta As Boolean = False)
    MsComm.Output = "E"
    If FechaPorta Then
        MsComm.PortOpen = False
    End If
    
End Function
Function ImprimeBarra(CódigoDeBarra As String, _
                      Optional Orienteção As String = "1", _
                      Optional PadrãoDoCódigoDeBarra As String = "F", _
                      Optional LarguraDaBarra As String = "62", _
                      Optional AlturaDaBarra As String = "020", _
                      Optional PosiçãoY As String = "0000", _
                      Optional PosiçãoX As String = "0000")
                           
    MsComm.Output = Orienteção & _
                     PadrãoDoCódigoDeBarra & _
                     LarguraDaBarra & _
                     AlturaDaBarra & _
                     PosiçãoY & _
                     PosiçãoX & _
                     CódigoDeBarra & CR
                     
End Function

Function ImprimirEtiquetaDeGôndola(Linha1 As String, _
                                   Linha2 As String, _
                                   Localização As String, _
                                   Valor As String, _
                                   CódigoDeBarra As String)
    P1 = Format(mvarEtiqGonImpressãoX * 10, "0000")
    P2 = Format(mvarEtiqGonImpressãoY * 10, "0000")
    
    
    
    ImprimeTexto Linha1, , , , , "" & Format(P2 + 170, "0000") & "", "" & Format(P1, "0000") & ""
    ImprimeTexto Linha2, , , , , "" & Format(P2 + 100, "0000") & "", "" & Format(P1, "0000") & ""
    ImprimeTexto Localização, , 3, , , "" & Format(P2 + 100, "0000") & "", "" & Format(P1 + 690, "0000") & ""
    ImprimeTexto Right("      " + Trim(Valor), 6), , 6, "21", , "" & Format(P2, "0000") & "", "" & Format(P1 + 400, "0000") & ""
    ImprimeTexto "R$", , 5, "11", , "" & Format(P2, "0000") & "", "" & Format(P1 + 280, "0000") & ""
    ImprimeBarra Format(Trim(CódigoDeBarra), "0000000000000"), , , , "069", (P2), (P1)
                
End Function
Function ImprimirEtiquetaDeCódigo(Linha1 As String, _
                             Linha2 As String, _
                             Linha3 As String, _
                             CódigoDeBarra As String)
    Dim P1
    Dim P2
    Dim P3
    
    P1 = Format((mvarEtiqCódMargemEsquerda + mvarEtiqCódPosição1) * 10, "0000")
    P2 = Format((mvarEtiqCódMargemEsquerda + _
                 mvarEtiqCódLargura + _
                 mvarEtiqCódEspaçoHorizontal + _
                 mvarEtiqCódPosição2) * 10, "0000")
    P3 = Format((mvarEtiqCódMargemEsquerda + _
                 (2 * mvarEtiqCódLargura) + _
                 (2 * mvarEtiqCódEspaçoHorizontal) + _
                 mvarEtiqCódPosição3) * 10, "0000")
    
    ImprimeTexto UCase(Linha1), , "2", "10", , "0170", (P1)
    ImprimeTexto UCase(Linha2), , "2", "10", , "0140", (P1)
    ImprimeTexto UCase(Linha3), , "2", "10", , "0110", (P1)
    ImprimeBarra Format(CódigoDeBarra, "0000000000000"), , , "62", "069", "0010", (P1)

    ImprimeTexto UCase(Linha1), , "2", "10", , "0170", (P2)
    ImprimeTexto UCase(Linha2), , "2", "10", , "0140", (P2)
    ImprimeTexto UCase(Linha3), , "2", "10", , "0110", (P2)
    ImprimeBarra Format(CódigoDeBarra, "0000000000000"), , , "62", "069", "0010", (P2)
            
    ImprimeTexto UCase(Linha1), , "2", "10", , "0170", (P3)
    ImprimeTexto UCase(Linha2), , "2", "10", , "0140", (P3)
    ImprimeTexto UCase(Linha3), , "2", "10", , "0110", (P3)
    ImprimeBarra Format(CódigoDeBarra, "0000000000000"), , , "02", "069", "0010", (P3)
            
End Function

