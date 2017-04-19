Attribute VB_Name = "mdlMainZebra"
Option Explicit

Public Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Public Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long

Dim xImp As Printer

'Varivaies do sistema
Public Const gblStrNomeModulo = "ZEBRA TLP 2844"
Public gblStrTipoPesqDescNaVenda                As String
Public Const gblStrTemControleDeVencDosProdutos = "S"
Public forStrCodUltimaCategoriaPrdExp           As String
Public dbBanco                                  As Database
'Variaveis que realizam a customização da ARGOX
Public lngFatorEscuridao                        As Long
Public strOrientacaoImpressao                   As String
Public strEtiquetasPorLinha                     As String
Public strCaractersLoja                         As String
Public strCaractersProdutos                     As String
'----------------------------------------------

Sub Main()
    
    On Error GoTo erro
    
    Call sbAbreBanco
    'Call sbCarregaConfiguracoes
    
    gblStrTipoPesqDescNaVenda = "Q"
    frmZebraTlp2844.Show vbModal
    'frmArgox214plus.Show vbModal
        
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number, ".Main")
End Sub

Private Sub sbAbreBanco()
    
    If VBA.Dir(App.Path & "\..\BancoDeDados\ADM.mdb") <> "" Then
        Set dbBanco = OpenDatabase(App.Path & "\..\BancoDeDados\ADM.mdb", False, False)
    Else
        VBA.MsgBox "Banco de dados do sistema não encontrado, a inicialização do sistema será abortada.", vbExclamation, gblStrNomeModulo
        End
    End If
    
End Sub

Public Sub sbDescricaoDeErro(ByVal IntCodErro As Integer, _
                             Optional ByVal PontoDoErro As String)
    
    VBA.MsgBox "Erro: " & IntCodErro & " Erro em: " & Err.Source & " descrição: " & Err.Description & " Ponto do erro: " & PontoDoErro, vbCritical, gblStrNomeModulo
    
End Sub

Public Function checkNull(strTexto As Variant) As Variant 'ME 15mar03
    
    If VBA.IsNull(strTexto) Then
        checkNull = ""
    Else
        checkNull = strTexto
    End If
    
End Function

Public Function fcVarChekValor(strTexto As Variant) As Variant 'ME 17mar03
    
    If VBA.IsNull(strTexto) Or strTexto = "" Then
        fcVarChekValor = VBA.CDbl("0")
    Else
        If VBA.IsNumeric(strTexto) Then
            fcVarChekValor = VBA.CDbl(strTexto)
        End If
    End If
    
End Function

Public Sub CarregaComboComGruposDeProdutos(ByVal Combo As ComboBox)
    
    On Error GoTo erro
    
    Dim rsTabela As Recordset
    Dim clsGrupo As New clsGrupo
    
    Combo.Clear
    
    Set rsTabela = clsGrupo.GruposProdutos_Consultar
    Do While Not rsTabela.EOF
        Combo.AddItem rsTabela!grucDescricao & "-" & rsTabela!grucCodigo
        rsTabela.MoveNext
    Loop
    
    rsTabela.Close
    Set rsTabela = Nothing
    Set clsGrupo = Nothing
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number)
End Sub

Public Function fcProcuraEmComboRetornandoListIndex(Optional ByVal Combo As ComboBox, _
                                                    Optional ByVal sItemProcurado As String, _
                                                    Optional ByVal iTamanhoNoCombo As Integer = 1, _
                                                    Optional ByVal Direcao As String = "I=Inicio do combo, F=Final do combo") As Integer
    'ME 15mar03
    
    Dim i As Integer
    
    fcProcuraEmComboRetornandoListIndex = -1
    
    For i = 0 To Combo.ListCount
        If VBA.Trim(Combo.List(i)) <> "" Then
            If Direcao = "I" Then
                If VBA.Mid(Combo.List(i), 1, iTamanhoNoCombo) = sItemProcurado Then
                    fcProcuraEmComboRetornandoListIndex = i
                    Exit For
                End If
            ElseIf Direcao = "F" Then
                If VBA.Right(Combo.List(i), iTamanhoNoCombo) = sItemProcurado Then
                    fcProcuraEmComboRetornandoListIndex = i
                    Exit For
                End If
            End If
        End If
    Next i
    
End Function

Public Function fcRetornaDadosDoProduto(ByVal Codigo As String, _
                                        Optional ByVal TipoDoDado As String = "NOME=Nome, VL_UNIT=Valor unitario, ET_ATUAL=Estoque atual, QT_RUIM_ATUAL=Quantidade ruim atual, VL_CMP=Valor de compra, UND_VND=Unidade de venda, GRUPO=Grupo do produto, COD_COMPL=Codigo completo, ET_INICIAL=Estoque Inicial, ST=Situação, AL=Aliquota, ") As String
    
    On Error GoTo erro
    
    Dim rsTabela        As Recordset
    Dim clsProdutos     As New clsProdutos
        
    If VBA.Trim(Codigo) <> "" Then
        Set rsTabela = clsProdutos.Produtos_Consultar(Codigo)
        If Not rsTabela.EOF Then
            Select Case TipoDoDado
                Case "NOME": fcRetornaDadosDoProduto = checkNull(rsTabela!prdcDescricao)
                Case "VL_UNIT": fcRetornaDadosDoProduto = fcVarChekValor(rsTabela!prdnValorAvista)
                Case "ET_ATUAL": fcRetornaDadosDoProduto = fcVarChekValor(rsTabela!prdnEstoqueAtual)
                Case "QT_RUIM_ATUAL": fcRetornaDadosDoProduto = fcVarChekValor(rsTabela!prdnRuim)
                Case "VL_CMP": fcRetornaDadosDoProduto = fcVarChekValor(rsTabela!prdnPrecoCompra)
                Case "UND_VND": fcRetornaDadosDoProduto = checkNull(rsTabela!prdcUndVnd)
                Case "GRUPO": fcRetornaDadosDoProduto = checkNull(rsTabela!grucCodigo)
                Case "COD_COMPL": fcRetornaDadosDoProduto = checkNull(rsTabela!prdcCodigo)
                Case "ET_INICIAL": fcRetornaDadosDoProduto = fcVarChekValor(rsTabela!prdnEstoqueInicial)
                Case "ST": fcRetornaDadosDoProduto = checkNull(rsTabela!prdcSituacaoTributaria)     'Situação tributaria
                Case "AL": fcRetornaDadosDoProduto = checkNull(rsTabela!prdcAliquota)               'Aliquota
            End Select
        Else
            fcRetornaDadosDoProduto = ""
        End If
        rsTabela.Close
    End If
    
    Set rsTabela = Nothing
    Set clsProdutos = Nothing
    
    Exit Function
erro:
    Call sbDescricaoDeErro(Err.Number)
End Function

Public Sub sbCarregaConfiguracoes()
    
    On Error GoTo erro
    
    lngFatorEscuridao = fcVarChekValor(mdlTrataRegistroWindows.fcRetornaValorDoRegistroDoWindows("ARGOX_ADM", "CONFIG", "FATOR_ESCURIDAO"))
    If lngFatorEscuridao = 0 Then lngFatorEscuridao = 6
    strOrientacaoImpressao = mdlTrataRegistroWindows.fcRetornaValorDoRegistroDoWindows("ARGOX_ADM", "CONFIG", "ORIENTACAO")
    If VBA.Trim(strOrientacaoImpressao) = "" Then strOrientacaoImpressao = "RETRATO"
    strEtiquetasPorLinha = mdlTrataRegistroWindows.fcRetornaValorDoRegistroDoWindows("ARGOX_ADM", "CONFIG", "ETIQUETAS_LINHA")
    If VBA.Trim(strEtiquetasPorLinha) = "" Then strEtiquetasPorLinha = "03"
    
    strCaractersLoja = mdlTrataRegistroWindows.fcRetornaValorDoRegistroDoWindows("ARGOX_ADM", "CONFIG", "CARACTERS_LOJA")
    If VBA.Trim(strCaractersLoja) = "" Then strCaractersLoja = "20"
    strCaractersProdutos = mdlTrataRegistroWindows.fcRetornaValorDoRegistroDoWindows("ARGOX_ADM", "CONFIG", "CARACTERS_PRODUTOS")
    If VBA.Trim(strCaractersProdutos) = "" Then strCaractersProdutos = "20"
    
    Exit Sub
erro:
    Call sbDescricaoDeErro(Err.Number)
End Sub
