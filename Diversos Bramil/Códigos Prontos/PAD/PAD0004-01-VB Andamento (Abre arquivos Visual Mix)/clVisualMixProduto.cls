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
    TipoDeOperação              As String * 1         '1
    TipoDeAção                  As String * 1   '(tabela 02)
    Situação                    As String * 1   '(tabela 03)
    CódigoAutomação             As String * 13  '13
    DígitoAutomação             As String * 1   '1
    CódigoInterno               As String * 10  '10
    DígitoInterno               As String * 1   '1
    DescriçãoCompleta           As String * 35
    DescriçãoReduzida           As String * 15
    CódigoVasilhame             As String * 13  '13
    DígitoVasilhame             As String * 1   '1
    PreçoDeVenda                As String * 11  '11-2
    PreçoDeCusto                As String * 11  '11-2 SEM ICMS
    Validade                    As String * 4   '4
    AlíquotaPDV                 As String * 2   '2 (tabela 04)
    AlíquotaNotaFiscal          As String * 2   '2 (tabela 05)
    CódigoDaEmbalagem           As String * 2   '2
    DescriçãoDaEmbalagem        As String * 3
    QuantidadeDaEmbalagem       As String * 7   '7-3
    PesoVariável                As String * 1
    VendeQuantidadeFracionária  As String * 1
    TipoDeEtiquetaDeGôndola     As String * 2   '2
    QuantidadeDeEtiquetaGôndola As String * 3   '3
    TipoDeEtiquetaDeProduto     As String * 2   '2
    DataDeInícioDaPromoção      As String * 8   '8
    DataDeFimDaPromoção         As String * 8   '8
    PreçoDeVendaDaPromoção      As String * 11  '11-2
    TipoDaPromoção              As String * 3   '3
    Mercadológico1              As String * 5   '5
    Mercadológico2              As String * 5   '5
    Mercadológico3              As String * 5   '5
    Mercadológico4              As String * 5   '5
    Mercadológico5              As String * 5   '5
    CódigoDoFornecedor          As String * 8   '8
    CódigoDaFaixaDePreço        As String * 4   '4
    MargemTeórica               As String * 5   '5-2
    SituacaoTributária          As String * 2   '2
End Type

Public Enum enVMX_TipoDeAção
    VMXTDA_Imediato = 0
    VMXTDA_DiaSeguinte = 1
End Enum

Public Enum enVMX_TipoDeOperação
    VMXTDO_InclusãoAlteração = 0
    VMXTDO_ExclusãoDoCódigoDeAutomação = 1
    VMXTDO_ExclusãoDaEmbalagem = 2
    VMXTDO_ExclusãoDoProduto = 3
End Enum
Public Enum enVMX_Situação
    VMXST_Normal
    VMXST_Descontinuado
    VMXST_Excluído
End Enum

Public Enum CMRSTipoDeArquivo
    CMRS_ArqProdutos = 0
    CMRS_ArqVendedor = 1
    CMRS_ArqDepartamento = 2
End Enum

Public Enum enVMX_AlíquotaPDV
    VMXALQPDV_18 = 1
    VMXALQPDV_25 = 2
    VMXALQPDV_12 = 3
    VMXALQPDV_07 = 4
    VMXALQPDV_Isento = 5
    VMXALQPDV_NãoTributado = 6
    VMXALQPDV_RetidoNaFonte = 7
    VMXALQPDV_17 = 8
    VMXALQPDV_18ComRedução = 9
    VMXALQPDV_12ComRedução = 10
End Enum

Private mvarCódigoAutomação As Double 'local copy
Private mvarNomeDoProduto As String 'local copy
Private mvarCódigoDepartamento As String 'local copy
Private mvarCódigoTributação As String 'local copy
Private mvarTipoTributação As String 'local copy
Private mvarDecimais As String 'local copy
Private mvarLegenda As String 'local copy
Private mvarPreço As String 'local copy
Private mvarNomeDoArquivo As String 'local copy
Private Linha As tpVMixProdutoDetalhe
Private mvarCanal As Integer 'local copy
Private mvarTipoDeRegistro As String 'local copy
Private mvarTipoDeAção As enVMX_TipoDeAção 'local copy
Private mvarTipoDeOperação As enVMX_TipoDeOperação
Private mvarProdutoDetalhe As tpVMixProdutoDetalhe
Private mvarSituação As enVMX_Situação
Private mvarDígitoAutomação As Byte 'local copy
Private mvarCódigoInterno As Double 'local copy
Private mvarDígitoCódigoInterno As Byte 'local copy
Private mvarDescriçãoCompleta As String 'local copy
Private mvarDescriçãoReduzida As String 'local copy
Private mvarCódigoVasilhame As Double 'local copy
Private mvarDígitoVasilhame As Byte 'local copy
Private mvarPreçoDeVenda As Double 'local copy
Private mvarPreçoDeCusto As Double 'local copy
Private mvarValidade As Integer 'local copy
Private mvarAlíquotaPDV As Byte
Private mvarAlíquotaNotaFiscal As Byte 'local copy
Private mvarCódigoDaEmbalagem As Byte 'local copy
Private mvarDescriçãoDaEmbalagem As String 'local copy
Private mvarQuantidadeDaEmbalagem As Double 'local copy
Private mvarPesoVariável As String 'local copy
Private mvarVendeQtdeFracionária As String 'local copy
Private mvarTipoDeEtiqueteDeGôndola As Byte 'local copy
Private mvarQuantidadeDeEtiquetaDeGôndola As Integer 'local copy
Private mvarTipoDeEtiquetaDeProduto As Byte 'local copy
Private mvarDataInícioPromoção As Date 'local copy
Private mvarDataFimPromoção As Date 'local copy
Private mvarPreçoDeVendaDaPromoção As Double 'local copy
Private mvarTipoDaPromoção As Integer 'local copy
Private mvarMercadológico1 As Long 'local copy
Private mvarMercadológico2 As Long 'local copy
Private mvarMercadológico3 As Long 'local copy
Private mvarMercadológico4 As Long 'local copy
Private mvarMercadológico5 As Long 'local copy
Private mvarCódigoDoFornecedor As Long 'local copy
Private mvarCódigoDaFaixaDePreço As Integer 'local copy
Private mvarMargemTeórica As Long 'local copy
Private mvarSituaçãoTributária As Byte 'local copy
'local variable(s) to hold property value(s)
Private mvarTotalDeRegistros As Long 'local copy

Public Property Let Situação(ByVal vData As enVMX_Situação)
    mvarSituação = vData
    Select Case vData
        Case enVMX_Situação.VMXST_Normal: mvarProdutoDetalhe.Situação = "N"
        Case enVMX_Situação.VMXST_Descontinuado: mvarProdutoDetalhe.Situação = "D"
        Case enVMX_Situação.VMXST_Excluído: mvarProdutoDetalhe.Situação = "E"
    End Select
    
End Property


Public Property Get Situação() As enVMX_Situação
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Situação
    Situação = mvarSituação
End Property



Public Property Let TotalDeRegistros(ByVal vData As Long)
    mvarTotalDeRegistros = vData
End Property

Public Property Get TotalDeRegistros() As Long
    TotalDeRegistros = mvarTotalDeRegistros
End Property

Public Property Let SituaçãoTributária(ByVal vData As Byte)
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
    mvarProdutoDetalhe.SituacaoTributária = vDataTemp

    mvarSituaçãoTributária = vData
End Property
Public Property Get SituaçãoTributária() As Byte
    SituaçãoTributária = mvarSituaçãoTributária
End Property

Public Property Let MargemTeórica(ByVal vData As Long)
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
    mvarProdutoDetalhe.MargemTeórica = vDataTemp
    
    mvarMargemTeórica = vData
End Property
Public Property Get MargemTeórica() As Long
    MargemTeórica = mvarMargemTeórica
End Property

Public Property Let CódigoDaFaixaDePreço(ByVal vData As Integer)
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
    mvarProdutoDetalhe.CódigoDaFaixaDePreço = vDataTemp
    
    mvarCódigoDaFaixaDePreço = vData
End Property
Public Property Get CódigoDaFaixaDePreço() As Integer
    CódigoDaFaixaDePreço = mvarCódigoDaFaixaDePreço
End Property

Public Property Let CódigoDoFornecedor(ByVal vData As Long)
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
    mvarProdutoDetalhe.CódigoDoFornecedor = vDataTemp
    
    mvarCódigoDoFornecedor = vData
End Property
Public Property Get CódigoDoFornecedor() As Long
    CódigoDoFornecedor = mvarCódigoDoFornecedor
End Property


Public Property Let Mercadológico5(ByVal vData As Long)
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
    mvarProdutoDetalhe.Mercadológico5 = vDataTemp
    
    mvarMercadológico5 = vData
End Property
Public Property Get Mercadológico5() As Long
    Mercadológico5 = mvarMercadológico5
End Property



Public Property Let Mercadológico4(ByVal vData As Long)
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
    mvarProdutoDetalhe.Mercadológico4 = vDataTemp

    mvarMercadológico4 = vData
End Property
Public Property Get Mercadológico4() As Long
    Mercadológico4 = mvarMercadológico4
End Property

Public Property Let Mercadológico3(ByVal vData As Long)
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
    mvarProdutoDetalhe.Mercadológico3 = vDataTemp

    mvarMercadológico3 = vData
End Property
Public Property Get Mercadológico3() As Long
    Mercadológico3 = mvarMercadológico3
End Property



Public Property Let Mercadológico2(ByVal vData As Long)
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
    mvarProdutoDetalhe.Mercadológico2 = vDataTemp

    mvarMercadológico2 = vData
End Property
Public Property Get Mercadológico2() As Long
    Mercadológico2 = mvarMercadológico2
End Property


Public Property Let Mercadológico1(ByVal vData As Long)
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
    mvarProdutoDetalhe.Mercadológico1 = vDataTemp
    
    mvarMercadológico1 = vData
End Property
Public Property Get Mercadológico1() As Long
    Mercadológico1 = mvarMercadológico1
End Property



Public Property Let TipoDaPromoção(ByVal vData As Integer)
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
    mvarProdutoDetalhe.TipoDaPromoção = vDataTemp

    mvarTipoDaPromoção = vData
End Property
Public Property Get TipoDaPromoção() As Integer
    TipoDaPromoção = mvarTipoDaPromoção
End Property

Public Property Let PreçoDeVendaDaPromoção(ByVal vData As Double)
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
    mvarProdutoDetalhe.PreçoDeVendaDaPromoção = vDataTemp

    mvarPreçoDeVendaDaPromoção = vData
End Property
Public Property Get PreçoDeVendaDaPromoção() As Double
    PreçoDeVendaDaPromoção = mvarPreçoDeVendaDaPromoção
End Property



Public Property Let DataFimPromoção(ByVal vData As Date)
    Dim vDataTemp As String
    If vData = "0" Then
        vDataTemp = "00000000"
    Else
        vDataTemp = Format(vData, "ddmmyyyy")
    End If
    mvarProdutoDetalhe.DataDeFimDaPromoção = vDataTemp
    mvarDataFimPromoção = vData
End Property
Public Property Get DataFimPromoção() As Date
    DataFimPromoção = mvarDataFimPromoção
End Property

Public Property Let DataInícioPromoção(ByVal vData As Date)
    Dim vDataTemp As String
    If vData = "0" Then
        vDataTemp = "00000000"
    Else
        vDataTemp = Format(vData, "ddmmyyyy")
    End If
    mvarProdutoDetalhe.DataDeInícioDaPromoção = vDataTemp
    
    mvarDataInícioPromoção = vData
End Property
Public Property Get DataInícioPromoção() As Date
    DataInícioPromoção = mvarDataInícioPromoção
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

Public Property Let QuantidadeDeEtiquetaDeGôndola(ByVal vData As Integer)
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
    mvarProdutoDetalhe.QuantidadeDeEtiquetaGôndola = vDataTemp
    
    mvarQuantidadeDeEtiquetaDeGôndola = vData
End Property
Public Property Get QuantidadeDeEtiquetaDeGôndola() As Integer
    QuantidadeDeEtiquetaDeGôndola = mvarQuantidadeDeEtiquetaDeGôndola
End Property



Public Property Let TipoDeEtiqueteDeGôndola(ByVal vData As Byte)
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
    mvarProdutoDetalhe.TipoDeEtiquetaDeGôndola = vDataTemp

    mvarTipoDeEtiqueteDeGôndola = vData
End Property
Public Property Get TipoDeEtiqueteDeGôndola() As Byte
    TipoDeEtiqueteDeGôndola = mvarTipoDeEtiqueteDeGôndola
End Property

Public Property Let VendeQtdeFracionária(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = Trim(vData)
    If Len(vDataTemp) > 1 Then vDataTemp = Mid(vDataTemp, 1, 1)
    mvarProdutoDetalhe.VendeQuantidadeFracionária = vDataTemp
    mvarVendeQtdeFracionária = vData
End Property
Public Property Get VendeQtdeFracionária() As String
    VendeQtdeFracionária = mvarVendeQtdeFracionária
End Property


Public Property Let PesoVariável(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = Trim(vData)
    If Len(vDataTemp) > 1 Then vDataTemp = Mid(vDataTemp, 1, 1)
    mvarProdutoDetalhe.PesoVariável = vDataTemp
    mvarPesoVariável = vData
End Property
Public Property Get PesoVariável() As String
    PesoVariável = mvarPesoVariável
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



Public Property Let DescriçãoDaEmbalagem(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = Trim(vData)
    If Len(vDataTemp) > 3 Then vDataTemp = Mid(vDataTemp, 1, 3)
    mvarProdutoDetalhe.DescriçãoDaEmbalagem = vDataTemp
    mvarDescriçãoDaEmbalagem = vData
    
End Property
Public Property Get DescriçãoDaEmbalagem() As String
    DescriçãoDaEmbalagem = mvarDescriçãoDaEmbalagem
End Property



Public Property Let CódigoDaEmbalagem(ByVal vData As Byte)
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
    mvarProdutoDetalhe.CódigoDaEmbalagem = vDataTemp
    mvarCódigoDaEmbalagem = vData
End Property
Public Property Get CódigoDaEmbalagem() As Byte
    CódigoDaEmbalagem = mvarCódigoDaEmbalagem
End Property



Public Property Let AlíquotaNotaFiscal(ByVal vData As Byte)
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
    mvarProdutoDetalhe.AlíquotaNotaFiscal = vDataTemp
    mvarAlíquotaNotaFiscal = vData
End Property
Public Property Get AlíquotaNotaFiscal() As Byte
    AlíquotaNotaFiscal = mvarAlíquotaNotaFiscal
End Property


Public Property Let AlíquotaPDV(ByVal vData As Byte)
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
    mvarProdutoDetalhe.AlíquotaPDV = vDataTemp
    mvarAlíquotaPDV = vData
End Property
Public Property Get AlíquotaPDV() As Byte
    AlíquotaPDV = mvarAlíquotaPDV
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



Public Property Let PreçoDeCusto(ByVal vData As Double)
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
    mvarProdutoDetalhe.PreçoDeCusto = vDataTemp
    mvarPreçoDeCusto = vData
End Property

Public Property Get PreçoDeCusto() As Double
    PreçoDeCusto = mvarPreçoDeCusto
End Property



Public Property Let PreçoDeVenda(ByVal vData As Double)
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
    mvarProdutoDetalhe.PreçoDeVenda = vDataTemp
    mvarPreçoDeVenda = vData
End Property
Public Property Get PreçoDeVenda() As Double
    PreçoDeVenda = mvarPreçoDeVenda
End Property

Public Property Get DígitoVasilhame() As Byte
    DígitoVasilhame = mvarDígitoVasilhame
End Property

Public Property Let CódigoVasilhame(ByVal vData As Double)
    Dim vDataTemp As String
    vDataTemp = vData
    If Len(Trim(vDataTemp)) > 13 Then vDataTemp = Empty
    If vDataTemp = 0 Then vDataTemp = "0000000000000"
    vDataTemp = Format(Trim(vDataTemp), "0000000000000")
    mvarDígitoVasilhame = Right(vDataTemp, 1)
    mvarCódigoVasilhame = vData
    mvarProdutoDetalhe.CódigoVasilhame = "0" & Mid(vDataTemp, 1, 12)
    mvarProdutoDetalhe.DígitoVasilhame = mvarDígitoVasilhame
    
End Property
Public Property Get CódigoVasilhame() As Double
    CódigoVasilhame = mvarCódigoVasilhame
End Property

Public Property Let DescriçãoReduzida(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = Trim(vData)
    
    If Len(vDataTemp) > 15 Then vDataTemp = Mid(vDataTemp, 1, 15)
    mvarDescriçãoReduzida = vData
    mvarProdutoDetalhe.DescriçãoReduzida = vDataTemp
End Property
Public Property Get DescriçãoReduzida() As String
    DescriçãoReduzida = mvarDescriçãoReduzida
End Property

Public Property Let DescriçãoCompleta(ByVal vData As String)
    Dim vDataTemp As String
    vDataTemp = vData
    If Len(Trim(vDataTemp)) > 35 Then vData = Mid(vDataTemp, 1, 35)
    mvarDescriçãoCompleta = vData
    mvarProdutoDetalhe.DescriçãoCompleta = vDataTemp
End Property
Public Property Get DescriçãoCompleta() As String
    DescriçãoCompleta = mvarDescriçãoCompleta
End Property

Public Property Get DígitoCódigoInterno() As Byte
    DígitoCódigoInterno = mvarDígitoCódigoInterno
End Property

Public Property Let CódigoInterno(ByVal vData As Double)
    Dim vDataTemp As String
    vDataTemp = vData
    If Len(Trim(vDataTemp)) > 10 Then vDataTemp = Empty
    If Trim(vDataTemp) = 0 Then vDataTemp = "0000000000"
    vDataTemp = Format(Trim(vDataTemp), "0000000000")
    
    mvarCódigoInterno = vData
    'mvarDígitoCódigoInterno = Right(vDataTemp, 1)
    'mvarProdutoDetalhe.CódigoInterno = "0" & Mid(vDataTemp, 1, 9)
    'mvarProdutoDetalhe.DígitoInterno = mvarDígitoCódigoInterno
    mvarDígitoCódigoInterno = 0 '
    mvarProdutoDetalhe.CódigoInterno = vDataTemp
    mvarProdutoDetalhe.DígitoInterno = 0
End Property

Public Property Get CódigoInterno() As Double
    CódigoInterno = mvarCódigoInterno
End Property

Public Property Get DígitoAutomação() As Byte
    DígitoAutomação = mvarDígitoAutomação
End Property

Friend Property Let ProdutoDetalhe(ByRef vData As tpVMixProdutoDetalhe)
    mvarProdutoDetalhe = vData
End Property
Friend Property Get ProdutoDetalhe() As tpVMixProdutoDetalhe
    ProdutoDetalhe = mvarProdutoDetalhe
End Property

Friend Property Let TipoDeOperação(ByVal vData As enVMX_TipoDeOperação)
    mvarTipoDeOperação = vData
    mvarProdutoDetalhe.TipoDeOperação = vData
    
End Property
Friend Property Get TipoDeOperação() As enVMX_TipoDeOperação
    TipoDeOperação = mvarTipoDeOperação
End Property

Friend Property Let TipoDeAção(ByVal vData As enVMX_TipoDeAção)
    mvarTipoDeAção = vData
    Select Case vData
        Case enVMX_TipoDeAção.VMXTDA_Imediato: mvarProdutoDetalhe.TipoDeAção = "I"
        Case enVMX_TipoDeAção.VMXTDA_DiaSeguinte: mvarProdutoDetalhe.TipoDeAção = "D"
    End Select
End Property

Friend Property Get TipoDeAção() As enVMX_TipoDeAção
    TipoDeAção = mvarTipoDeAção
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

Public Property Let CódigoAutomação(ByVal vData As Double)
    Dim vDataTemp As String
    vDataTemp = vData
    If Len(Trim(vDataTemp)) > 13 Then
        If Len(Trim(vDataTemp)) = 14 Then
            mvarProdutoDetalhe.CódigoAutomação = Mid(vDataTemp, 1, 13)
            mvarProdutoDetalhe.DígitoAutomação = Right(vDataTemp, 1)
        Else
            vDataTemp = 0
            mvarProdutoDetalhe.CódigoAutomação = "0000000000000"
            mvarProdutoDetalhe.DígitoAutomação = "0"
        End If
    Else
        If Trim(vDataTemp) = 0 Then vDataTemp = "0000000000000"
        vDataTemp = Format(vDataTemp, "0000000000000")
        mvarProdutoDetalhe.CódigoAutomação = "0" & Mid(vDataTemp, 1, 12)
        mvarDígitoAutomação = Right(vDataTemp, 1)
        mvarProdutoDetalhe.DígitoAutomação = mvarDígitoAutomação
    End If
    mvarDígitoAutomação = Right(vDataTemp, 1)
    mvarCódigoAutomação = vData
    
End Property
Public Property Get CódigoAutomação() As Double
    CódigoAutomação = mvarCódigoAutomação
End Property


Function IniciarImpressão(NomeDoArquivo As String, Optional DataDoMovimento As Date, _
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
Public Function FinalizarImpressão(Optional Canal As Byte, Optional QuantidadeTotalDeRegistros As Long)
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
            mvarProdutoDetalhe.TipoDeOperação & _
            mvarProdutoDetalhe.TipoDeAção & _
            mvarProdutoDetalhe.Situação & _
            mvarProdutoDetalhe.CódigoAutomação & _
            mvarProdutoDetalhe.DígitoAutomação & _
            mvarProdutoDetalhe.CódigoInterno & _
            mvarProdutoDetalhe.DígitoInterno & _
            mvarProdutoDetalhe.DescriçãoCompleta & _
            mvarProdutoDetalhe.DescriçãoReduzida & _
            mvarProdutoDetalhe.CódigoVasilhame & _
            mvarProdutoDetalhe.DígitoVasilhame & _
            mvarProdutoDetalhe.PreçoDeVenda & _
            mvarProdutoDetalhe.PreçoDeCusto & _
            mvarProdutoDetalhe.Validade & _
            mvarProdutoDetalhe.AlíquotaPDV & _
            mvarProdutoDetalhe.AlíquotaNotaFiscal & _
            mvarProdutoDetalhe.CódigoDaEmbalagem & _
            mvarProdutoDetalhe.DescriçãoDaEmbalagem & _
            mvarProdutoDetalhe.QuantidadeDaEmbalagem & _
            mvarProdutoDetalhe.PesoVariável & _
            mvarProdutoDetalhe.VendeQuantidadeFracionária & _
            mvarProdutoDetalhe.TipoDeEtiquetaDeGôndola & _
            mvarProdutoDetalhe.QuantidadeDeEtiquetaGôndola & _
            mvarProdutoDetalhe.TipoDeEtiquetaDeProduto
    Linha = Linha & _
            mvarProdutoDetalhe.DataDeInícioDaPromoção & _
            mvarProdutoDetalhe.DataDeFimDaPromoção & _
            mvarProdutoDetalhe.PreçoDeVendaDaPromoção & _
            mvarProdutoDetalhe.TipoDaPromoção & _
            mvarProdutoDetalhe.Mercadológico1 & _
            mvarProdutoDetalhe.Mercadológico2 & _
            mvarProdutoDetalhe.Mercadológico3 & _
            mvarProdutoDetalhe.Mercadológico4 & _
            mvarProdutoDetalhe.Mercadológico5 & _
            mvarProdutoDetalhe.CódigoDoFornecedor & _
            mvarProdutoDetalhe.CódigoDaFaixaDePreço & _
            mvarProdutoDetalhe.MargemTeórica & _
            mvarProdutoDetalhe.SituacaoTributária
    Print #mvarCanal, Linha
    mvarTotalDeRegistros = mvarTotalDeRegistros + 1
End Function

