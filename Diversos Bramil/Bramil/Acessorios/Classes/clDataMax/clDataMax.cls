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
'Esta Classe deve ser utilizada com:
'Classe clINI
'DataMax.ini

Private Const GMct_strTRANSF_COMM = "9600,N,8,1"

Private SOH As String
Private STX As String
Private ESC As String
Private CR As String


Public Enum DataMax_Enum
    GMen_mil�metros = 0
    GMen_polegadas = 1
End Enum

Private Type EtiquetaC�digoDataMax
    LarguraDoFormul�rio As Integer
    LarguraDaEtiqueta As Integer
    EtiquetasPorLinha As Byte
    MargemEsquerda As Byte
    Espa�oHorizEntreEtiquetas As Byte
    Espa�oVertiEntreEtiquetas As Byte
    LarguraDaImpress�o As Byte
    AlturaDaImpress�o As Byte
    In�cioDaImpress�oX1 As Byte
    In�cioDaImpress�oX2 As Byte
    In�cioDaImpress�oX3 As Byte
End Type
Private EtiquetaC�digo As EtiquetaC�digoDataMax

Private mvarMsComm As Object 'local copy
Private mvarPortaCom As Byte 'local copy
Private mvarCancelar As Boolean 'local copy
Private mvarTransfer�ncia As String 'local copy
Private mvarUnidadeDeMedida As DataMax_Enum 'local copy
'local variable(s) to hold property value(s)
Private mvarEtiqC�dLargura As Integer 'local copy
Private mvarEtiqC�dAltura As Integer 'local copy
Private mvarEtiqC�dFormLagura As Integer 'local copy
Private mvarEtiqC�dEspa�oHorizontal As Integer 'local copy
Private mvarEtiqC�dEspa�oVertical As Integer 'local copy
Private mvarEtiqC�dPosi��o1 As Integer 'local copy
Private mvarEtiqC�dPosi��o2 As Integer 'local copy
Private mvarEtiqC�dPosi��o3 As Integer 'local copy
Private mvarEtiqC�dAjuste As Integer 'local copy
Private mvarEtiqC�dMargemEsquerda As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarEtiqGonLargura As Integer 'local copy
Private mvarEtiqGonAltura As Integer 'local copy
Private mvarEtiqGonImpress�oX As Integer 'local copy
Private mvarEtiqGonImpress�oY As Integer 'local copy
Public Property Let EtiqGonImpress�oY(ByVal vData As Integer)
    mvarEtiqGonImpress�oY = vData
End Property
Public Property Get EtiqGonImpress�oY() As Integer
    EtiqGonImpress�oY = mvarEtiqGonImpress�oY
End Property

Public Property Let EtiqGonImpress�oX(ByVal vData As Integer)
    mvarEtiqGonImpress�oX = vData
End Property
Public Property Get EtiqGonImpress�oX() As Integer
    EtiqGonImpress�oX = mvarEtiqGonImpress�oX
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


Public Property Let EtiqC�dMargemEsquerda(ByVal vData As Integer)
    mvarEtiqC�dMargemEsquerda = vData
End Property

Public Property Get EtiqC�dMargemEsquerda() As Integer
    EtiqC�dMargemEsquerda = mvarEtiqC�dMargemEsquerda
End Property

Public Property Let EtiqC�dAjuste(ByVal vData As Integer)
    mvarEtiqC�dAjuste = vData
End Property
Public Property Get EtiqC�dAjuste() As Integer
    EtiqC�dAjuste = mvarEtiqC�dAjuste
End Property

Public Property Let EtiqC�dPosi��o3(ByVal vData As Integer)
    mvarEtiqC�dPosi��o3 = vData
End Property
Public Property Get EtiqC�dPosi��o3() As Integer
    EtiqC�dPosi��o3 = mvarEtiqC�dPosi��o3
End Property

Public Property Let EtiqC�dPosi��o2(ByVal vData As Integer)
    mvarEtiqC�dPosi��o2 = vData
End Property
Public Property Get EtiqC�dPosi��o2() As Integer
    EtiqC�dPosi��o2 = mvarEtiqC�dPosi��o2
End Property

Public Property Let EtiqC�dPosi��o1(ByVal vData As Integer)
    mvarEtiqC�dPosi��o1 = vData
End Property
Public Property Get EtiqC�dPosi��o1() As Integer
    EtiqC�dPosi��o1 = mvarEtiqC�dPosi��o1
End Property

Public Property Let EtiqC�dEspa�oVertical(ByVal vData As Integer)
    mvarEtiqC�dEspa�oVertical = vData
End Property
Public Property Get EtiqC�dEspa�oVertical() As Integer
    EtiqC�dEspa�oVertical = mvarEtiqC�dEspa�oVertical
End Property

Public Property Let EtiqC�dEspa�oHorizontal(ByVal vData As Integer)
    mvarEtiqC�dEspa�oHorizontal = vData
End Property
Public Property Get EtiqC�dEspa�oHorizontal() As Integer
    EtiqC�dEspa�oHorizontal = mvarEtiqC�dEspa�oHorizontal
End Property

Public Property Let EtiqC�dFormLagura(ByVal vData As Integer)
    mvarEtiqC�dFormLagura = vData
End Property
Public Property Get EtiqC�dFormLagura() As Integer
    EtiqC�dFormLagura = mvarEtiqC�dFormLagura
End Property



Public Property Let EtiqC�dAltura(ByVal vData As Integer)
    mvarEtiqC�dAltura = vData
End Property
Public Property Get EtiqC�dAltura() As Integer
    EtiqC�dAltura = mvarEtiqC�dAltura
End Property

Public Property Let EtiqC�dLargura(ByVal vData As Integer)
    mvarEtiqC�dLargura = vData
End Property
Public Property Get EtiqC�dLargura() As Integer
    EtiqC�dLargura = mvarEtiqC�dLargura
End Property

Public Property Let UnidadeDeMedida(ByVal vData As DataMax_Enum)
    mvarUnidadeDeMedida = vData
    If mvarUnidadeDeMedida = GMen_mil�metros Then EnviarComando STX & "m"
    If mvarUnidadeDeMedida = GMen_polegadas Then EnviarComando STX & "n"
End Property
Public Property Get UnidadeDeMedida() As DataMax_Enum
    UnidadeDeMedida = mvarUnidadeDeMedida
End Property

Public Property Let Transfer�ncia(ByVal vData As String)
    mvarTransfer�ncia = vData
    MsComm.Settings = vData
End Property
Public Property Get Transfer�ncia() As String
    Transfer�ncia = mvarTransfer�ncia
End Property


Public Function ConfigurarDataMax()
    Dim lclINI As clINI
    Set lclINI = New clINI
    
    mvarPortaCom = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "DataMax", "Porta"))
    mvarTransfer�ncia = lclINI.INIL�(App.Path & "\DataMax.ini", "DataMax", "Transfer�ncia")
    mvarUnidadeDeMedida = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "DataMax", "Unidade de Medida"))
        
    mvarEtiqC�dLargura = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Largura"))
    mvarEtiqC�dAltura = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Altura"))
    mvarEtiqC�dFormLagura = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Form Largura"))
    mvarEtiqC�dMargemEsquerda = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Margem Esquerda"))
    mvarEtiqC�dEspa�oHorizontal = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Espa�o Horizontal"))
    mvarEtiqC�dEspa�oVertical = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Espa�o Vertical"))
    mvarEtiqC�dPosi��o1 = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Posi��o 1"))
    mvarEtiqC�dPosi��o2 = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Posi��o 2"))
    mvarEtiqC�dPosi��o3 = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Posi��o 3"))
    mvarEtiqC�dAjuste = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de C�digo", "Ajuste"))
    
    mvarEtiqGonLargura = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de G�ndola", "Largura"))
    mvarEtiqGonAltura = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de G�ndola", "Altura"))
    mvarEtiqGonImpress�oX = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de G�ndola", "Posi��o X"))
    mvarEtiqGonImpress�oY = Val(lclINI.INIL�(App.Path & "\DataMax.ini", "Etiqueta de G�ndola", "Posi��o Y"))
End Function

Private Function RetornaPar�metro(strValor As String, Procura As String) As String

    
End Function


Public Sub EnviarComando(Comando As String)
    MsComm.Output = Comando
End Sub

Private Sub ConfigurarUnidadeDeMedida(Optional Unidade As DataMax_Enum = GMen_mil�metros)
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
Function Avan�aEtiqueta(Optional Quantidade As Integer = 1)
    PreparaImpress�o
    MsComm.Output = STX & "E" & Format(Quantidade, "0000") & CR
    ImprimeTexto "  "
    FinalizaImpress�o
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
Function PreparaImpress�o(Optional C�pias As Integer = "1")

    If MsComm.PortOpen = False Then
        MsComm.CommPort = PortaCom 'usando a porta passada pelo parametro
        MsComm.Settings = mvarTransfer�ncia   '"9600,N,8,1"
        MsComm.InputLen = 0
        MsComm.PortOpen = True
    End If
    
    MsComm.Output = STX & "L" & CR
'    MsComm.Output = "H07" & CR
    MsComm.Output = "D11" & CR
    MsComm.Output = "Q" & Trim(Format(C�pias, "0000")) & CR

End Function
Function ImprimeTexto(Texto As String, _
                      Optional Orienta��o As String = "1", _
                      Optional Fonte As String = "5", _
                      Optional MultipicadorDeLargura = "10", _
                      Optional Ignore = "000", _
                      Optional Posi��oY As String = "0000", _
                      Optional Posi��oX As String = "0000")
    MsComm.Output = Orienta��o & _
                     Fonte & _
                     MultipicadorDeLargura & _
                     Ignore & _
                     Posi��oY & _
                     Posi��oX & _
                     Texto & CR
End Function
Function FinalizaImpress�o(Optional FechaPorta As Boolean = False)
    MsComm.Output = "E"
    If FechaPorta Then
        MsComm.PortOpen = False
    End If
End Function
Function ImprimeBarra(C�digoDeBarra As String, _
                      Optional Oriente��o As String = "1", _
                      Optional Padr�oDoC�digoDeBarra As String = "F", _
                      Optional LarguraDaBarra As String = "62", _
                      Optional AlturaDaBarra As String = "020", _
                      Optional Posi��oY As String = "0000", _
                      Optional Posi��oX As String = "0000")
                           
    MsComm.Output = Oriente��o & _
                     Padr�oDoC�digoDeBarra & _
                     LarguraDaBarra & _
                     AlturaDaBarra & _
                     Posi��oY & _
                     Posi��oX & _
                     C�digoDeBarra & CR
                     
End Function

Function ImprimirEtiquetaDeG�ndola(Linha1 As String, _
                                   Linha2 As String, _
                                   Localiza��o As String, _
                                   Valor As String, _
                                   C�digoDeBarra As String)
    P1 = Format(mvarEtiqGonImpress�oX * 10, "0000")
    P2 = Format(mvarEtiqGonImpress�oY * 10, "0000")
    
    ImprimeTexto Linha1, , , , , "" & Format(P2 + 170, "0000") & "", "" & Format(P1, "0000") & ""
    ImprimeTexto Linha2, , , , , "" & Format(P2 + 100, "0000") & "", "" & Format(P1, "0000") & ""
    ImprimeTexto Localiza��o, , 3, , , "" & Format(P2 + 100, "0000") & "", "" & Format(P1 + 690, "0000") & ""
    ImprimeTexto Right("      " + Trim(Valor), 6), , 6, "21", , "" & Format(P2, "0000") & "", "" & Format(P1 + 400, "0000") & ""
    ImprimeTexto "R$", , 5, "11", , "" & Format(P2, "0000") & "", "" & Format(P1 + 280, "0000") & ""
    ImprimeBarra Format(Trim(C�digoDeBarra), "0000000000000"), , , , "069", (P2), (P1)
                
End Function
Function ImprimirEtiquetaDeC�digo(Linha1 As String, _
                             Linha2 As String, _
                             Linha3 As String, _
                             C�digoDeBarra As String)
    Dim P1
    Dim P2
    Dim P3
    
    P1 = Format((mvarEtiqC�dMargemEsquerda + mvarEtiqC�dPosi��o1) * 10, "0000")
    P2 = Format((mvarEtiqC�dMargemEsquerda + _
                 mvarEtiqC�dLargura + _
                 mvarEtiqC�dEspa�oHorizontal + _
                 mvarEtiqC�dPosi��o2) * 10, "0000")
    P3 = Format((mvarEtiqC�dMargemEsquerda + _
                 (2 * mvarEtiqC�dLargura) + _
                 (2 * mvarEtiqC�dEspa�oHorizontal) + _
                 mvarEtiqC�dPosi��o3) * 10, "0000")
    
    ImprimeTexto UCase(Linha1), , "2", "10", , "0180", (P1)
    ImprimeTexto UCase(Linha2), , "2", "10", , "0150", (P1)
    ImprimeTexto UCase(Linha3), , "2", "10", , "0120", (P1)
    ImprimeBarra Format(C�digoDeBarra, "0000000000000"), , , "62", "069", "0020", (P1)

    ImprimeTexto UCase(Linha1), , "2", "10", , "0180", (P2)
    ImprimeTexto UCase(Linha2), , "2", "10", , "0150", (P2)
    ImprimeTexto UCase(Linha3), , "2", "10", , "0120", (P2)
    ImprimeBarra Format(C�digoDeBarra, "0000000000000"), , , "62", "069", "0020", (P2)
            
    ImprimeTexto UCase(Linha1), , "2", "10", , "0180", (P3)
    ImprimeTexto UCase(Linha2), , "2", "10", , "0150", (P3)
    ImprimeTexto UCase(Linha3), , "2", "10", , "0120", (P3)
    ImprimeBarra Format(C�digoDeBarra, "0000000000000"), , , "02", "069", "0020", (P3)
            
End Function

