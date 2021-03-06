VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clVisualMixProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Type tpVMixProdutoDetalhe
    TipoDeRegistro              As String * 1   '(Sempre "D")
    TipoDeOpera��o              As String * 1         '1
    TipoDeA��o                  As String * 1   '(tabela 02)
    Situa��o                    As String * 1   '(tabela 03)
    C�digoAutoma��o             As String * 13  '13
    D�gitoAutoma��o             As String * 1   '1
    C�digoInterno               As String * 10  '10
    D�gitoInterno               As String * 1   '1
    Descri��oCompleta           As String * 35
    Descri��oReduzida           As String * 15
    C�digoVasilhame             As String * 13  '13
    D�gitoVasilhame             As String * 1   '1
    Pre�oDeVenda                As String * 11  '11-2
    Pre�oDeCusto                As String * 11  '11-2 SEM ICMS
    Validade                    As String * 4   '4
    Al�quotaPDV                 As String * 2   '2 (tabela 04)
    Al�quotaNotaFiscal          As String * 2   '2 (tabela 05)
    C�digoDaEmbalagem           As String * 2   '2
    Descri��oDaEmbalagem        As String * 3
    QuantidadeDaEmbalagem       As String * 7   '7-3
    PesoVari�vel                As String * 1
    VendeQuantidadeFracion�ria  As String * 1
    TipoDeEtiquetaDeG�ndola     As String * 2   '2
    QuantidadeDeEtiquetaG�ndola As String * 3   '3
    TipoDeEtiquetaDeProduto     As String * 2   '2
    DataDeIn�cioDaPromo��o      As String * 8   '8
    DataDeFimDaPromo��o         As String * 8   '8
    Pre�oDeVendaDaPromo��o      As String * 11  '11-2
    TipoDaPromo��o              As String * 3   '3
    Mercadol�gico1              As String * 5   '5
    Mercadol�gico2              As String * 5   '5
    Mercadol�gico3              As String * 5   '5
    Mercadol�gico4              As String * 5   '5
    Mercadol�gico5              As String * 5   '5
    C�digoDoFornecedor          As String * 8   '8
    C�digoDaFaixaDePre�o        As String * 4   '4
    MargemTe�rica               As String * 5   '5-2
    SituacaoTribut�ria          As String * 2   '2
End Type

Public Enum enVMX_TipoDeA��o
    VMXTDA_Imediato = 0
    VMXTDA_DiaSeguinte = 1
End Enum

Public Enum enVMX_TipoDeOpera��o
    VMXTDO_Inclus�oAltera��o = 0
    VMXTDO_Exclus�oDoC�digoDeAutoma��o = 1
    VMXTDO_Exclus�oDaEmbalagem = 2
    VMXTDO_Exclus�oDoProduto = 3
End Enum
Public Enum enVMX_Situa��o
    VMXST_Normal
    VMXST_Descontinuado
    VMXST_Exclu�do
End Enum

Public Enum CMRSTipoDeArquivo
    CMRS_ArqProdutos = 0
    CMRS_ArqVendedor = 1
    CMRS_ArqDepartamento = 2
End Enum

Public Enum enVMX_Al�quotaPDV
    VMXALQPDV_18 = 1
    VMXALQPDV_25 = 2
    VMXALQPDV_12 = 3
    VMXALQPDV_07 = 4
    VMXALQPDV_Isento = 5
    VMXALQPDV_N�oTributado = 6
    VMXALQPDV_RetidoNaFonte = 7
    VMXALQPDV_17 = 8
    VMXALQPDV_18ComRedu��o = 9
    VMXALQPDV_12ComRedu��o = 10
End Enum

Private mvarC�digoAutoma��o As Double 'local copy
Private mvarNomeDoProduto As String 'local copy
Private mvarC�digoDepartamento As String 'local copy
Private mvarC�digoTributa��o As String 'local copy
Private mvarTipoTributa��o As String 'local copy
Private mvarDecimais As String 'local copy
Private mvarLegenda As String 'local copy
Private mvarPre�o As String 'local copy
Private mvarNomeDoArquivo As String 'local copy
Private Linha As tpVMixProdutoDetalhe
Private mvarCanal As Integer 'local copy
Private mvarTipoDeRegistro As String 'local copy
Private mvarTipoDeA��o As enVMX_TipoDeA��o 'local copy
Private mvarTipoDeOpera��o As enVMX_TipoDeOpera��o
Private mvarProdutoDetalhe As tpVMixProdutoDetalhe
Private mvarSitua��o As enVMX_Situa��o
Private mvarD�gitoAutoma��o As Byte 'local copy
Private mvarC�digoInterno As Double 'local copy
Private mvarD�gitoC�digoInterno As Byte 'local copy
Private mvarDescri��oCompleta As String 'local copy
Private mvarDescri��oReduzida As String 'local copy
Private mvarC�digoVasilhame As Double 'local copy
Private mvarD�gitoVasilhame As Byte 'local copy
Private mvarPre�oDeVenda As Double 'local copy
Private mvarPre�oDeCusto As Double 'local copy
Private mvarValidade As Integer 'local copy
Private mvarAl�quotaPDV As Byte
Private mvarAl�quotaNotaFiscal As Byte 'local copy
Private mvarC�digoDaEmbalagem As Byte 'local copy
Private mvarDescri��oDaEmbalagem As String 'local copy
Private mvarQuantidadeDaEmbalagem As Double 'local copy
Private mvarPesoVari�vel As String 'local copy
Private mvarVendeQtdeFracion�ria As String 'local copy
Private mvarTipoDeEtiqueteDeG�ndola As Byte 'local copy
Private mvarQuantidadeDeEtiquetaDeG�ndola As Integer 'local copy
Private mvarTipoDeEtiquetaDeProduto As Byte 'local copy
Private mvarDataIn�cioPromo��o As Date 'local copy
Private mvarDataFimPromo��o As Date 'local copy
Private mvarPre�oDeVendaDaPromo��o As Double 'local copy
Private mvarTipoDaPromo��o As Integer 'local copy
Private mvarMercadol�gico1 As Long 'local copy
Private mvarMercadol�gico2 As Long 'local copy
Private mvarMercadol�gico3 As Long 'local copy
Private mvarMercadol�gico4 As Long 'local copy
Private mvarMercadol�gico5 As Long 'local copy
Private mvarC�digoDoFornecedor As Long 'local copy
Private mvarC�digoDaFaixaDePre�o As Integer 'local copy
Private mvarMargemTe�rica As Long 'local copy
Private mvarSitua��oTribut�ria As Byte 'local copy
'local variable(s) to hold property value(s)
Private mvarTotalDeRegistros As Long 'local copy

Public Property Let Situa��o(ByVal vData As enVMX_Situa��o)
    mvarSitua��o = vData
    Select Case vData
        Case enVMX_Situa��o.VMXST_Normal: mvarProdutoDetalhe.Situa��o = "N"
        Case enVMX_Situa��o.VMXST_Descontinuado: mvarProdutoDetalhe.Situa��o = "D"
        Case enVMX_Situa��o.VMXST_Exclu�do: mvarProdutoDetalhe.Situa��o = "E"
    End Select
    
End Property


Public Property Get Situa��o() As enVMX_Situa��o
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Situa��o
    Situa��o = mvarSitua��o
End Property



Public Property Let TotalDeRegistros(ByVal vData As Long)
    mvarTotalDeRegistros = vData
End Property

Public Property Get TotalDeRegistros() As Long
    TotalDeRegistros = mvarTotalDeRegistros
End Property

Public Property Let Situa��oTribut�ria(ByVal vData As Byte)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 2 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00"
    vDataTemp = Format(Trim(vDataTemp), "00")
    mvarProdutoDetalhe.SituacaoTribut�ria = vDataTemp

    mvarSitua��oTribut�ria = vData
End Property
Public Property Get Situa��oTribut�ria() As Byte
    Situa��oTribut�ria = mvarSitua��oTribut�ria
End Property

Public Property Let MargemTe�rica(ByVal vData As Long)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 999.99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData * 100
    End If
    
    vPosDecimal = InStr(vDataTemp, ",")
    If vPosDecimal = 0 Then vPosDecimal = InStr(vDataTemp, ".")
    If vPosDecimal <> 0 Then vDataTemp = Replace(Replace(vDataTemp, ",", Empty), ".", Empty)
    
    If Len(Trim(vDataTemp)) > 5 Then vDataTemp = 0
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00000"
    vDataTemp = Format(vDataTemp, "00000")
    mvarProdutoDetalhe.MargemTe�rica = vDataTemp
    
    mvarMargemTe�rica = vData
End Property
Public Property Get MargemTe�rica() As Long
    MargemTe�rica = mvarMargemTe�rica
End Property

Public Property Let C�digoDaFaixaDePre�o(ByVal vData As Integer)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 9999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 4 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "0000"
    vDataTemp = Format(Trim(vDataTemp), "0000")
    mvarProdutoDetalhe.C�digoDaFaixaDePre�o = vDataTemp
    
    mvarC�digoDaFaixaDePre�o = vData
End Property
Public Property Get C�digoDaFaixaDePre�o() As Integer
    C�digoDaFaixaDePre�o = mvarC�digoDaFaixaDePre�o
End Property

Public Property Let C�digoDoFornecedor(ByVal vData As Long)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99999999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 8 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00000000"
    vDataTemp = Format(Trim(vDataTemp), "00000000")
    mvarProdutoDetalhe.C�digoDoFornecedor = vDataTemp
    
    mvarC�digoDoFornecedor = vData
End Property
Public Property Get C�digoDoFornecedor() As Long
    C�digoDoFornecedor = mvarC�digoDoFornecedor
End Property


Public Property Let Mercadol�gico5(ByVal vData As Long)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 5 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00000"
    vDataTemp = Format(Trim(vDataTemp), "00000")
    mvarProdutoDetalhe.Mercadol�gico5 = vDataTemp
    
    mvarMercadol�gico5 = vData
End Property
Public Property Get Mercadol�gico5() As Long
    Mercadol�gico5 = mvarMercadol�gico5
End Property



Public Property Let Mercadol�gico4(ByVal vData As Long)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 5 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00000"
    vDataTemp = Format(Trim(vDataTemp), "00000")
    mvarProdutoDetalhe.Mercadol�gico4 = vDataTemp

    mvarMercadol�gico4 = vData
End Property
Public Property Get Mercadol�gico4() As Long
    Mercadol�gico4 = mvarMercadol�gico4
End Property

Public Property Let Mercadol�gico3(ByVal vData As Long)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 5 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00000"
    vDataTemp = Format(Trim(vDataTemp), "00000")
    mvarProdutoDetalhe.Mercadol�gico3 = vDataTemp

    mvarMercadol�gico3 = vData
End Property
Public Property Get Mercadol�gico3() As Long
    Mercadol�gico3 = mvarMercadol�gico3
End Property



Public Property Let Mercadol�gico2(ByVal vData As Long)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 5 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00000"
    vDataTemp = Format(Trim(vDataTemp), "00000")
    mvarProdutoDetalhe.Mercadol�gico2 = vDataTemp

    mvarMercadol�gico2 = vData
End Property
Public Property Get Mercadol�gico2() As Long
    Mercadol�gico2 = mvarMercadol�gico2
End Property


Public Property Let Mercadol�gico1(ByVal vData As Long)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 5 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00000"
    vDataTemp = Format(Trim(vDataTemp), "00000")
    mvarProdutoDetalhe.Mercadol�gico1 = vDataTemp
    
    mvarMercadol�gico1 = vData
End Property
Public Property Get Mercadol�gico1() As Long
    Mercadol�gico1 = mvarMercadol�gico1
End Property



Public Property Let TipoDaPromo��o(ByVal vData As Integer)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 3 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "000"
    vDataTemp = Format(Trim(vDataTemp), "000")
    mvarProdutoDetalhe.TipoDaPromo��o = vDataTemp

    mvarTipoDaPromo��o = vData
End Property
Public Property Get TipoDaPromo��o() As Integer
    TipoDaPromo��o = mvarTipoDaPromo��o
End Property

Public Property Let Pre�oDeVendaDaPromo��o(ByVal vData As Double)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 999999999.99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData * 100
    End If
    vPosDecimal = InStr(vDataTemp, ",")
    If vPosDecimal = 0 Then vPosDecimal = InStr(vDataTemp, ".")
    If vPosDecimal <> 0 Then vDataTemp = Replace(Replace(vDataTemp, ",", Empty), ".", Empty)
    
    If Len(Trim(vDataTemp)) > 11 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00000000000"
    vDataTemp = Format(Trim(vDataTemp), "00000000000")
    mvarProdutoDetalhe.Pre�oDeVendaDaPromo��o = vDataTemp

    mvarPre�oDeVendaDaPromo��o = vData
End Property
Public Property Get Pre�oDeVendaDaPromo��o() As Double
    Pre�oDeVendaDaPromo��o = mvarPre�oDeVendaDaPromo��o
End Property



Public Property Let DataFimPromo��o(ByVal vData As Date)
    Dim vDataTemp As String
    If vData = "0" Then
        vDataTemp = "00000000"
    Else
        vDataTemp = Format(vData, "ddmmyyyy")
    End If
    mvarProdutoDetalhe.DataDeFimDaPromo��o = vDataTemp
    mvarDataFimPromo��o = vData
End Property
Public Property Get DataFimPromo��o() As Date
    DataFimPromo��o = mvarDataFimPromo��o
End Property

Public Property Let DataIn�cioPromo��o(ByVal vData As Date)
    Dim vDataTemp As String
    If vData = "0" Then
        vDataTemp = "00000000"
    Else
        vDataTemp = Format(vData, "ddmmyyyy")
    End If
    mvarProdutoDetalhe.DataDeIn�cioDaPromo��o = vDataTemp
    
    mvarDataIn�cioPromo��o = vData
End Property
Public Property Get DataIn�cioPromo��o() As Date
    DataIn�cioPromo��o = mvarDataIn�cioPromo��o
End Property


Public Property Let TipoDeEtiquetaDeProduto(ByVal vData As Byte)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 2 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00"
    vDataTemp = Format(Trim(vDataTemp), "00")
    mvarProdutoDetalhe.TipoDeEtiquetaDeProduto = vDataTemp

    mvarTipoDeEtiquetaDeProduto = vData
End Property
Public Property Get TipoDeEtiquetaDeProduto() As Byte
    TipoDeEtiquetaDeProduto = mvarTipoDeEtiquetaDeProduto
End Property

Public Property Let QuantidadeDeEtiquetaDeG�ndola(ByVal vData As Integer)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 3 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "000"
    vDataTemp = Format(Trim(vDataTemp), "000")
    mvarProdutoDetalhe.QuantidadeDeEtiquetaG�ndola = vDataTemp
    
    mvarQuantidadeDeEtiquetaDeG�ndola = vData
End Property
Public Property Get QuantidadeDeEtiquetaDeG�ndola() As Integer
    QuantidadeDeEtiquetaDeG�ndola = mvarQuantidadeDeEtiquetaDeG�ndola
End Property



Public Property Let TipoDeEtiqueteDeG�ndola(ByVal vData As Byte)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 2 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00"
    vDataTemp = Format(Trim(vDataTemp), "00")
    mvarProdutoDetalhe.TipoDeEtiquetaDeG�ndola = vDataTemp

    mvarTipoDeEtiqueteDeG�ndola = vData
End Property
Public Property Get TipoDeEtiqueteDeG�ndola() As Byte
    TipoDeEtiqueteDeG�ndola = mvarTipoDeEtiqueteDeG�ndola
End Property

Public Property Let VendeQtdeFracion�ria(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = Trim(vData)
    If Len(vDataTemp) > 1 Then vDataTemp = Mid(vDataTemp, 1, 1)
    mvarProdutoDetalhe.VendeQuantidadeFracion�ria = vDataTemp
    mvarVendeQtdeFracion�ria = vData
End Property
Public Property Get VendeQtdeFracion�ria() As String
    VendeQtdeFracion�ria = mvarVendeQtdeFracion�ria
End Property


Public Property Let PesoVari�vel(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = Trim(vData)
    If Len(vDataTemp) > 1 Then vDataTemp = Mid(vDataTemp, 1, 1)
    mvarProdutoDetalhe.PesoVari�vel = vDataTemp
    mvarPesoVari�vel = vData
End Property
Public Property Get PesoVari�vel() As String
    PesoVari�vel = mvarPesoVari�vel
End Property



Public Property Let QuantidadeDaEmbalagem(ByVal vData As Double)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 9999.999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData * 1000
    End If
    
    vPosDecimal = InStr(vDataTemp, ",")
    If vPosDecimal = 0 Then vPosDecimal = InStr(vDataTemp, ".")
    If vPosDecimal <> 0 Then vDataTemp = Replace(Replace(vDataTemp, ",", Empty), ".", Empty)
    
    If Len(Trim(vDataTemp)) > 7 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "0000000"
    vDataTemp = Format(Trim(vDataTemp), "0000000")
    mvarProdutoDetalhe.QuantidadeDaEmbalagem = vDataTemp
    mvarQuantidadeDaEmbalagem = vData
End Property
Public Property Get QuantidadeDaEmbalagem() As Double
    QuantidadeDaEmbalagem = mvarQuantidadeDaEmbalagem
End Property



Public Property Let Descri��oDaEmbalagem(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = Trim(vData)
    If Len(vDataTemp) > 3 Then vDataTemp = Mid(vDataTemp, 1, 3)
    mvarProdutoDetalhe.Descri��oDaEmbalagem = vDataTemp
    mvarDescri��oDaEmbalagem = vData
    
End Property
Public Property Get Descri��oDaEmbalagem() As String
    Descri��oDaEmbalagem = mvarDescri��oDaEmbalagem
End Property



Public Property Let C�digoDaEmbalagem(ByVal vData As Byte)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 2 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00"
    vDataTemp = Format(Trim(vDataTemp), "00")
    mvarProdutoDetalhe.C�digoDaEmbalagem = vDataTemp
    mvarC�digoDaEmbalagem = vData
End Property
Public Property Get C�digoDaEmbalagem() As Byte
    C�digoDaEmbalagem = mvarC�digoDaEmbalagem
End Property



Public Property Let Al�quotaNotaFiscal(ByVal vData As Byte)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    If Len(Trim(vDataTemp)) > 2 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00"
    vDataTemp = Format(Trim(vDataTemp), "00")
    mvarProdutoDetalhe.Al�quotaNotaFiscal = vDataTemp
    mvarAl�quotaNotaFiscal = vData
End Property
Public Property Get Al�quotaNotaFiscal() As Byte
    Al�quotaNotaFiscal = mvarAl�quotaNotaFiscal
End Property


Public Property Let Al�quotaPDV(ByVal vData As Byte)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    
    If vData > 99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    
    If Len(Trim(vDataTemp)) > 2 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00"
    vDataTemp = Format(Trim(vDataTemp), "00")
    mvarProdutoDetalhe.Al�quotaPDV = vDataTemp
    mvarAl�quotaPDV = vData
End Property
Public Property Get Al�quotaPDV() As Byte
    Al�quotaPDV = mvarAl�quotaPDV
End Property



Public Property Let Validade(ByVal vData As Integer)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 9999 Then
        vDataTemp = 0
    Else
        vDataTemp = vData
    End If
    
    vPosDecimal = InStr(vDataTemp, ",")
    If vPosDecimal = 0 Then vPosDecimal = InStr(vDataTemp, ".")
    If vPosDecimal <> 0 Then vDataTemp = Replace(Replace(vDataTemp, ",", Empty), ".", Empty)
    
    If Len(Trim(vDataTemp)) > 4 Then vDataTemp = Empty
    If Trim(vDataTemp) = 0 Then vDataTemp = "0000"
    vDataTemp = Format(Trim(vDataTemp), "0000")
    mvarProdutoDetalhe.Validade = vDataTemp

    mvarValidade = vData
End Property

Public Property Get Validade() As Integer
    Validade = mvarValidade
End Property



Public Property Let Pre�oDeCusto(ByVal vData As Double)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    If vData > 999999999.99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData * 100
    End If
    
    vPosDecimal = InStr(vDataTemp, ",")
    If vPosDecimal = 0 Then vPosDecimal = InStr(vDataTemp, ".")
    If vPosDecimal <> 0 Then vDataTemp = Replace(Replace(vDataTemp, ",", Empty), ".", Empty)
    
    If Len(Trim(vDataTemp)) > 11 Then vDataTemp = Empty
    If Trim(vDataTemp) = 0 Then vDataTemp = "00000000000"
    vDataTemp = Format(Trim(vDataTemp), "00000000000")
    mvarProdutoDetalhe.Pre�oDeCusto = vDataTemp
    mvarPre�oDeCusto = vData
End Property

Public Property Get Pre�oDeCusto() As Double
    Pre�oDeCusto = mvarPre�oDeCusto
End Property



Public Property Let Pre�oDeVenda(ByVal vData As Double)
    Dim vDataTemp As String
    Dim vPosDecimal As Byte
    
    If vData > 999999999.99 Then
        vDataTemp = 0
    Else
        vDataTemp = vData * 100
    End If
    
    vPosDecimal = InStr(vDataTemp, ",")
    If vPosDecimal = 0 Then vPosDecimal = InStr(vDataTemp, ".")
    If vPosDecimal <> 0 Then vDataTemp = Replace(Replace(vDataTemp, ",", Empty), ".", Empty)
    
    If Len(Trim(vDataTemp)) > 11 Then vDataTemp = Empty
    If IsNumeric(Trim(vDataTemp)) = False Then vDataTemp = "00000000000"
    vDataTemp = Format(Trim(vDataTemp), "00000000000")
    mvarProdutoDetalhe.Pre�oDeVenda = vDataTemp
    mvarPre�oDeVenda = vData
End Property
Public Property Get Pre�oDeVenda() As Double
    Pre�oDeVenda = mvarPre�oDeVenda
End Property

Public Property Get D�gitoVasilhame() As Byte
    D�gitoVasilhame = mvarD�gitoVasilhame
End Property

Public Property Let C�digoVasilhame(ByVal vData As Double)
    Dim vDataTemp As String
    vDataTemp = vData
    If Len(Trim(vDataTemp)) > 13 Then vDataTemp = Empty
    If vDataTemp = 0 Then vDataTemp = "0000000000000"
    vDataTemp = Format(Trim(vDataTemp), "0000000000000")
    mvarD�gitoVasilhame = Right(vDataTemp, 1)
    mvarC�digoVasilhame = vData
    mvarProdutoDetalhe.C�digoVasilhame = "0" & Mid(vDataTemp, 1, 12)
    mvarProdutoDetalhe.D�gitoVasilhame = mvarD�gitoVasilhame
    
End Property
Public Property Get C�digoVasilhame() As Double
    C�digoVasilhame = mvarC�digoVasilhame
End Property

Public Property Let Descri��oReduzida(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = Trim(vData)
    
    If Len(vDataTemp) > 15 Then vDataTemp = Mid(vDataTemp, 1, 15)
    mvarDescri��oReduzida = vData
    mvarProdutoDetalhe.Descri��oReduzida = vDataTemp
End Property
Public Property Get Descri��oReduzida() As String
    Descri��oReduzida = mvarDescri��oReduzida
End Property

Public Property Let Descri��oCompleta(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = vData
    If Len(Trim(vDataTemp)) > 35 Then vData = Mid(vDataTemp, 1, 35)
    mvarDescri��oCompleta = vData
    mvarProdutoDetalhe.Descri��oCompleta = vDataTemp
End Property
Public Property Get Descri��oCompleta() As String
    Descri��oCompleta = mvarDescri��oCompleta
End Property

Public Property Get D�gitoC�digoInterno() As Byte
    D�gitoC�digoInterno = mvarD�gitoC�digoInterno
End Property

Public Property Let C�digoInterno(ByVal vData As Double)
    Dim vDataTemp As String
    vDataTemp = vData
    If Len(Trim(vDataTemp)) > 10 Then vDataTemp = Empty
    If Trim(vDataTemp) = 0 Then vDataTemp = "0000000000"
    vDataTemp = Format(Trim(vDataTemp), "0000000000")
    
    mvarC�digoInterno = vData
    'mvarD�gitoC�digoInterno = Right(vDataTemp, 1)
    'mvarProdutoDetalhe.C�digoInterno = "0" & Mid(vDataTemp, 1, 9)
    'mvarProdutoDetalhe.D�gitoInterno = mvarD�gitoC�digoInterno
    mvarD�gitoC�digoInterno = 0 '
    mvarProdutoDetalhe.C�digoInterno = vDataTemp
    mvarProdutoDetalhe.D�gitoInterno = 0
End Property

Public Property Get C�digoInterno() As Double
    C�digoInterno = mvarC�digoInterno
End Property

Public Property Get D�gitoAutoma��o() As Byte
    D�gitoAutoma��o = mvarD�gitoAutoma��o
End Property

Friend Property Let ProdutoDetalhe(ByRef vData As tpVMixProdutoDetalhe)
    mvarProdutoDetalhe = vData
End Property
Friend Property Get ProdutoDetalhe() As tpVMixProdutoDetalhe
    ProdutoDetalhe = mvarProdutoDetalhe
End Property

Friend Property Let TipoDeOpera��o(ByVal vData As enVMX_TipoDeOpera��o)
    mvarTipoDeOpera��o = vData
    mvarProdutoDetalhe.TipoDeOpera��o = vData
    
End Property
Friend Property Get TipoDeOpera��o() As enVMX_TipoDeOpera��o
    TipoDeOpera��o = mvarTipoDeOpera��o
End Property

Friend Property Let TipoDeA��o(ByVal vData As enVMX_TipoDeA��o)
    mvarTipoDeA��o = vData
    Select Case vData
        Case enVMX_TipoDeA��o.VMXTDA_Imediato: mvarProdutoDetalhe.TipoDeA��o = "I"
        Case enVMX_TipoDeA��o.VMXTDA_DiaSeguinte: mvarProdutoDetalhe.TipoDeA��o = "D"
    End Select
End Property

Friend Property Get TipoDeA��o() As enVMX_TipoDeA��o
    TipoDeA��o = mvarTipoDeA��o
End Property

Public Property Let TipoDeRegistro(ByVal vData As String)
    If Len(Trim(vData)) > 1 Then vData = ""
    mvarTipoDeRegistro = Trim(vData)
    mvarProdutoDetalhe.TipoDeRegistro = mvarTipoDeRegistro
End Property
Public Property Get TipoDeRegistro() As String
    TipoDeRegistro = mvarTipoDeRegistro
End Property

Public Property Let Canal(ByVal vData As Integer)
    If Len(vData) > 255 Or Len(vData) < 1 Then vData = 0
    If vData = 0 Then vData = FreeFile
    mvarCanal = vData
End Property
Public Property Get Canal() As Integer
    Canal = mvarCanal
End Property

Public Property Let NomeDoArquivo(ByVal vData As String)
    mvarNomeDoArquivo = vData
End Property
Public Property Get NomeDoArquivo() As String
    NomeDoArquivo = mvarNomeDoArquivo
End Property

Public Property Let C�digoAutoma��o(ByVal vData As Double)
    Dim vDataTemp As String
    vDataTemp = vData
    If Len(Trim(vDataTemp)) > 13 Then
        If Len(Trim(vDataTemp)) = 14 Then
            mvarProdutoDetalhe.C�digoAutoma��o = Mid(vDataTemp, 1, 13)
            mvarProdutoDetalhe.D�gitoAutoma��o = Right(vDataTemp, 1)
        Else
            vDataTemp = 0
            mvarProdutoDetalhe.C�digoAutoma��o = "0000000000000"
            mvarProdutoDetalhe.D�gitoAutoma��o = "0"
        End If
    Else
        If Trim(vDataTemp) = 0 Then vDataTemp = "0000000000000"
        vDataTemp = Format(vDataTemp, "0000000000000")
        mvarProdutoDetalhe.C�digoAutoma��o = "0" & Mid(vDataTemp, 1, 12)
        mvarD�gitoAutoma��o = Right(vDataTemp, 1)
        mvarProdutoDetalhe.D�gitoAutoma��o = mvarD�gitoAutoma��o
    End If
    mvarD�gitoAutoma��o = Right(vDataTemp, 1)
    mvarC�digoAutoma��o = vData
    
End Property
Public Property Get C�digoAutoma��o() As Double
    C�digoAutoma��o = mvarC�digoAutoma��o
End Property


Function IniciarImpress�o(NomeDoArquivo As String, Optional DataDoMovimento As Date, _
                          Optional Canal As Byte)
    Dim strDataDoMovimento As String
    If DataDoMovimento = Empty Then DataDoMovimento = Date
    strDataDoMovimento = Format(DataDoMovimento, "ddmmyyyy")
    
    If Canal = Empty Then
        Canal = FreeFile
    End If
    mvarCanal = Canal
    'Close #Canal
    Open NomeDoArquivo For Output As #Canal
    Print #mvarCanal, "H" & strDataDoMovimento & String(209, " ")
    mvarTotalDeRegistros = 0
End Function
Public Function FinalizarImpress�o(Optional Canal As Byte, Optional QuantidadeTotalDeRegistros As Long)
    Dim strQuantidadeTotalDeRegistros As String
    If QuantidadeTotalDeRegistros = Empty Or QuantidadeTotalDeRegistros = 0 Then
        strQuantidadeTotalDeRegistros = Format(mvarTotalDeRegistros, "00000000")
    Else
        strQuantidadeTotalDeRegistros = Format(QuantidadeTotalDeRegistros, "00000000")
    End If

    Print #mvarCanal, "R" & strQuantidadeTotalDeRegistros & String(209, " ")
    Close #mvarCanal
End Function

Public Function Imprimir(Optional ByVal Canal As Byte)
    If Canal = Empty Then
        Canal = mvarCanal
    End If
    Dim Linha As String
    
    Linha = mvarProdutoDetalhe.TipoDeRegistro & _
            mvarProdutoDetalhe.TipoDeOpera��o & _
            mvarProdutoDetalhe.TipoDeA��o & _
            mvarProdutoDetalhe.Situa��o & _
            mvarProdutoDetalhe.C�digoAutoma��o & _
            mvarProdutoDetalhe.D�gitoAutoma��o & _
            mvarProdutoDetalhe.C�digoInterno & _
            mvarProdutoDetalhe.D�gitoInterno & _
            mvarProdutoDetalhe.Descri��oCompleta & _
            mvarProdutoDetalhe.Descri��oReduzida & _
            mvarProdutoDetalhe.C�digoVasilhame & _
            mvarProdutoDetalhe.D�gitoVasilhame & _
            mvarProdutoDetalhe.Pre�oDeVenda & _
            mvarProdutoDetalhe.Pre�oDeCusto & _
            mvarProdutoDetalhe.Validade & _
            mvarProdutoDetalhe.Al�quotaPDV & _
            mvarProdutoDetalhe.Al�quotaNotaFiscal & _
            mvarProdutoDetalhe.C�digoDaEmbalagem & _
            mvarProdutoDetalhe.Descri��oDaEmbalagem & _
            mvarProdutoDetalhe.QuantidadeDaEmbalagem & _
            mvarProdutoDetalhe.PesoVari�vel & _
            mvarProdutoDetalhe.VendeQuantidadeFracion�ria & _
            mvarProdutoDetalhe.TipoDeEtiquetaDeG�ndola & _
            mvarProdutoDetalhe.QuantidadeDeEtiquetaG�ndola & _
            mvarProdutoDetalhe.TipoDeEtiquetaDeProduto
    Linha = Linha & _
            mvarProdutoDetalhe.DataDeIn�cioDaPromo��o & _
            mvarProdutoDetalhe.DataDeFimDaPromo��o & _
            mvarProdutoDetalhe.Pre�oDeVendaDaPromo��o & _
            mvarProdutoDetalhe.TipoDaPromo��o & _
            mvarProdutoDetalhe.Mercadol�gico1 & _
            mvarProdutoDetalhe.Mercadol�gico2 & _
            mvarProdutoDetalhe.Mercadol�gico3 & _
            mvarProdutoDetalhe.Mercadol�gico4 & _
            mvarProdutoDetalhe.Mercadol�gico5 & _
            mvarProdutoDetalhe.C�digoDoFornecedor & _
            mvarProdutoDetalhe.C�digoDaFaixaDePre�o & _
            mvarProdutoDetalhe.MargemTe�rica & _
            mvarProdutoDetalhe.SituacaoTribut�ria
    Print #mvarCanal, Linha
    mvarTotalDeRegistros = mvarTotalDeRegistros + 1
End Function

