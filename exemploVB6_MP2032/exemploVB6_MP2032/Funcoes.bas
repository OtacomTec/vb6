Attribute VB_Name = "Module1"
' Declara��o das Fun��es da DLL

Public Declare Function AcionaGuilhotina Lib "mp2032.dll" (ByVal Modo As Integer) As Integer
Public Declare Function AutenticaDoc Lib "mp2032.dll" (ByVal BufTras As String, ByVal Tempo As Integer) As Integer
Public Declare Function BematechTX Lib "mp2032.dll" (ByVal BufTrans As String) As Integer
Public Declare Function CaracterGrafico Lib "mp2032.dll" (ByVal Buffer As String, ByVal TamBuffer As Integer) As Integer
Public Declare Function ComandoTX Lib "mp2032.dll" (ByVal BufTrans As String, ByVal TamBufTrans As Integer) As Integer
Public Declare Function ConfiguraModeloImpressora Lib "mp2032.dll" (ByVal ModeloImpressora As Integer) As Integer
Public Declare Function ConfiguraTamanhoExtrato Lib "mp2032.dll" (ByVal NumeroLinhas As Integer) As Integer
Public Declare Function DocumentInserted Lib "mp2032.dll" () As Integer
Public Declare Function EsperaImpressao Lib "mp2032.dll" () As Integer
Public Declare Function FechaPorta Lib "mp2032.dll" () As Integer
Public Declare Function FormataTX Lib "mp2032.dll" (ByVal BufTras As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal expand As Integer, ByVal enfat As Integer) As Integer
Public Declare Function HabilitaEsperaImpressao Lib "mp2032.dll" (ByVal Flag As Integer) As Integer
Public Declare Function HabilitaExtratoLongo Lib "mp2032.dll" (ByVal Flag As Integer) As Integer
Public Declare Function HabilitaPresenterRetratil Lib "mp2032.dll" (ByVal Flag As Integer) As Integer
Public Declare Function IniciaPorta Lib "mp2032.dll" (ByVal iPorta As String) As Integer
Public Declare Function Le_Status Lib "mp2032.dll" () As Integer
Public Declare Function Le_Status_Gaveta Lib "mp2032.dll" () As Integer
Public Declare Function ProgramaPresenterRetratil Lib "mp2032.dll" (ByVal Tempo As Integer) As Integer
Public Declare Function Status_Porta Lib "mp2032.dll" () As Integer
Public Declare Function VerificaPapelPresenter Lib "mp2032.dll" () As Integer
'fun��o para configura��o dos c�digos de barras
Public Declare Function ConfiguraCodigoBarras Lib "mp2032.dll" (ByVal Altura As Integer, ByVal Largura As Integer, ByVal PosicaoCaracteres As Integer, ByVal Fonte As Integer, ByVal Margem As Integer) As Integer


'fun��es para impress�o do bitmap
Public Declare Function ImprimeBmpEspecial Lib "mp2032.dll" (ByVal FileName As String, ByVal xScale As Integer, _
                                                            ByVal yScale As Integer, ByVal angle As Integer) As Integer
                                                            
                                                            
Public Declare Function ImprimeBitmap Lib "mp2032.dll" (ByVal FileName As String, ByVal mode As Integer) As Integer

Public Declare Function AjustaLarguraPapel Lib "mp2032.dll" (ByVal width As Integer) As Integer
Public Declare Function SelectDithering Lib "mp2032.dll" (ByVal algorithm As Integer) As Integer


'fun��es para impress�o dos c�digos de barras
Public Declare Function ImprimeCodigoBarrasUPCA Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasUPCE Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasEAN13 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasEAN8 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODE39 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODE93 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODE128 Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasITF Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODABAR Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasISBN Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasMSI Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasPLESSEY Lib "mp2032.dll" (ByVal Codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasPDF417 Lib "mp2032.dll" (ByVal NivelCorrecaoErros As Integer, ByVal Altura As Integer, ByVal Largura As Integer, ByVal Colunas As Integer, ByVal Codigo As String) As Integer





' Fun��o que realiza a tradu��o do Software

Function TraduzCaption(iFlag As Integer)

  ' Se iFlag for igual a 0 (zero), a tradu��o
  ' ser� para o Portugu�s, sen�o Ingl�s

  If iFlag = 0 Then
  
     frmPrincipal.Caption = "Aplicativo de teste usando a API de comunica��o e o driver de spooler"
     ' Tradu��o da Aba da API para o Portugu�s
  
     frmPrincipal.SSTab1.TabCaption(0) = "Usando a API"
     frmPrincipal.Frame8.Caption = "Porta de Comunica��o"
     frmPrincipal.Frame9.Caption = "Modelo da Impressora"
     frmPrincipal.Text1.Text = "Digite o texto aqui."
     frmPrincipal.cmdAcentos.Visible = True
     frmPrincipal.cmdAcentos.Caption = "Caracteres A&centuados"
     
     'frame modos de impress�o
     frmPrincipal.Frame2.Caption = "Modos de Impress�o"
     frmPrincipal.Option3.Caption = "Normal"
     frmPrincipal.Option4.Caption = "Elite"
     frmPrincipal.Option5.Caption = "Condensado"
     
     'frame modos de formata��o
     frmPrincipal.Frame3.Caption = "Modos de Formata��o"
     frmPrincipal.Check1.Caption = "Negrito"
     frmPrincipal.Check2.Caption = "Sublinhado"
     frmPrincipal.Check3.Caption = "It�lico"
     frmPrincipal.Check4.Caption = "Expandido"
     
     'bot�es de impress�o de texto
     frmPrincipal.cmdImprimeTextoSemFormatacao.Caption = "Imprime te&xto sem formata��o"
     frmPrincipal.cmdImprimeTextoComFormatacao.Caption = "Imprime texto com &formata��o"
     frmPrincipal.cmdTesteTextoFormatado.Caption = "Teste Texto Formatado"
     
     'frame programa��o do presenter
     frmPrincipal.Frame4.Caption = "Programa��o do Presenter"
     frmPrincipal.Label1.Caption = "segundo(s)"
     frmPrincipal.Label3.Caption = "Tempo de retra��o:"
     frmPrincipal.Text2.Text = "5"
     frmPrincipal.cmdProgramarPresenter.Caption = "&Programar"
     If frmPrincipal.cmdHabilitarPresenter.Caption = "Disable" Then
        frmPrincipal.cmdHabilitarPresenter.Caption = "D&esabilitar"
     Else
        frmPrincipal.cmdHabilitarPresenter.Caption = "&Habilitar"
     End If
     'frame tamanho do extrato
     frmPrincipal.Frame5.Caption = "Tamanho do extrato"
     frmPrincipal.Label4.Caption = "N�mero de Linhas:"
     frmPrincipal.Text3.Text = "90"
     frmPrincipal.cmdProgramarExtrato.Caption = "Programar"
     frmPrincipal.cmdHabilitarExtrato.Caption = "Habilitar"
     
     'frame Status da impressora
     frmPrincipal.Frame6.Caption = "Status da Impressora"
     frmPrincipal.cmdStatusImpressora.Caption = "Status da Impressora"
     
     'Outros bot�es
     frmPrincipal.cmdImprimirCaracterGrafico.Caption = "Imprimir Caracter Gr�fico"
     frmPrincipal.cmdCortarPapel.Caption = "Corte Total do Papel"
     frmPrincipal.cmdCortarParcial.Caption = "Corte Parcial do Papel"
     frmPrincipal.cmdAguardarImpressaoTexto.Caption = "Aguardar Impress�o Texto"
     frmPrincipal.cmdVerificarPapelPresenter.Caption = "Verificar Papel no Presenter"
     frmPrincipal.cmdSair.Caption = "Sair"
     
     'Frame autentica��o de documentos
     frmPrincipal.Frame10.Caption = "Autentica��o de Documentos"
     frmPrincipal.Label2.Caption = "Texto:"
     frmPrincipal.Label9.Caption = "seg."
     frmPrincipal.Text7.Text = "Teste de Autentica��o"
     frmPrincipal.cmdVerificaDocInserido.Caption = "Verificar Documento Inserido"
     frmPrincipal.cmdAutenticacao.Caption = "Autenticar Documento"
     
     ' Tradu��o da Aba da Usando o Driver de Spooler para o Portugu�s
  
     frmPrincipal.SSTab1.TabCaption(2) = "Usando o Driver de Spooler"
     frmPrincipal.Label5.Caption = "Entre com seu Texto:"
     frmPrincipal.Text5.Text = "Bematech Solu��es"
     frmPrincipal.cmdModificarFonte.Caption = "Modificar Fonte"
     frmPrincipal.cmdImprimir.Caption = "Imprimir"
     frmPrincipal.Label6.Caption = "Imprimir Figura"
     frmPrincipal.cmdImprimirFigura.Caption = "Imprimir"
     frmPrincipal.Frame7.Caption = "Informa��es"
     frmPrincipal.Label7.Caption = "Impressoras:"
     
     ' Tradu��o da Aba da Usando Codigo de Barras com a API
     frmPrincipal.SSTab1.TabCaption(1) = "Usando C�digo de Barras com a API"
     frmPrincipal.frmCodigoBarras.Caption = "Escolha o C�digo de Barras"
     frmPrincipal.frmLarguraBarras.Caption = "Largura das Barras"
     frmPrincipal.optFinas.Caption = "Finas"
     frmPrincipal.optMedias.Caption = "M�dias (default)"
     frmPrincipal.optGrossas.Caption = "Grossas"
     frmPrincipal.frmPosicaoCaracter.Caption = "Posi��o dos Caracteres"
     frmPrincipal.optAcima.Caption = "Acima do c�digo"
     frmPrincipal.optAbaixo.Caption = "Abaixo do c�digo"
     frmPrincipal.optAcimaAbaixo.Caption = "Acima e abaixo do c�digo"
     frmPrincipal.optNaoImprime.Caption = "N�o imprime os caracteres"
     frmPrincipal.frmFonte.Caption = "Fonte"
     frmPrincipal.optNormal.Caption = "Normal"
     frmPrincipal.optCondensada.Caption = "Condensada"
     frmPrincipal.frmCodigo.Caption = "C�digo"
     frmPrincipal.cmdImprimirCodBarras.Caption = "Imprimir"
     frmPrincipal.lbImprimirCodigo.Caption = "Imprimir C�digo de Barras"
     
     If frmPrincipal.cmdLigarSensorPoucoPapel.Caption = "Enable Low Paper Sensor" Then
        frmPrincipal.cmdLigarSensorPoucoPapel.Caption = "Ligar Sensor de Pouco Papel"
     Else
        frmPrincipal.cmdLigarSensorPoucoPapel.Caption = "Desligar Sensor de Pouco Papel"
     End If
     
    ' tradu��o da aba impress�o de bitmap
    frmPrincipal.SSTab1.TabCaption(3) = "Impress�o de Bitmap"
    frmPrincipal.LabelBmpFile.Caption = "Nome do arquivo"
    frmPrincipal.Frame13.Caption = "Orienta��o do papel"
    frmPrincipal.RetratoBtn.Caption = "Retrato"
    frmPrincipal.PaisagemBtn.Caption = "Paisagem"
    frmPrincipal.Label14.Caption = "Bitola do papel"
    frmPrincipal.Frame14.Caption = "Redimensionar"
    frmPrincipal.Label10.Caption = "Altura"
    frmPrincipal.Label11.Caption = "Largura"
    frmPrincipal.AjustaBtn.Caption = "Ajusta na largura do papel"
    frmPrincipal.Frame15.Caption = "Girar"
    frmPrincipal.Label13.Caption = "Graus"
    frmPrincipal.Command3.Caption = "Imprimir"
    frmPrincipal.Command2.Caption = "Imprimir"
     
  Else
  
     frmPrincipal.Caption = "Test application using the API of comunication and spooler driver"
    
     ' Tradu��o da Aba da API para o Ingl�s
  
     frmPrincipal.SSTab1.TabCaption(0) = "Using the API"
     
     'frame porta de comunica��o
     frmPrincipal.Frame8.Caption = "Communication Port"
     
     'frame modelo de impressora
     frmPrincipal.Frame9.Caption = "Printer Model"
     
     frmPrincipal.Text1.Text = "Type the text here."
     frmPrincipal.cmdAcentos.Visible = False
     
     'frame modos de impress�o
     frmPrincipal.Frame2.Caption = "Printing Modes"
     frmPrincipal.Option3.Caption = "Normal"
     frmPrincipal.Option4.Caption = "Elite"
     frmPrincipal.Option5.Caption = "Condensed"
     
     'frame modos de formata��o
     frmPrincipal.Frame3.Caption = "Formatting Modes"
     frmPrincipal.Check1.Caption = "Bold"
     frmPrincipal.Check2.Caption = "Underlined"
     frmPrincipal.Check3.Caption = "Italic"
     frmPrincipal.Check4.Caption = "Expanded"
     
     'bot�es de impress�o de texto
     frmPrincipal.cmdImprimeTextoSemFormatacao.Caption = "Prints Te&xt Without Formatting"
     frmPrincipal.cmdImprimeTextoComFormatacao.Caption = "Prints Text With &Formatting"
     frmPrincipal.cmdTesteTextoFormatado.Caption = "T&est Formatted Text"
     
     'frame programa��o do presenter
     frmPrincipal.Frame4.Caption = "Presenter Programming"
     frmPrincipal.Label1.Caption = "second(s)"
     frmPrincipal.Label3.Caption = "Retracting Time:"
     frmPrincipal.Text2.Text = "5"
     frmPrincipal.cmdProgramarPresenter.Caption = "Program"
     If frmPrincipal.cmdHabilitarPresenter.Caption = "&Habilitar" Then
        frmPrincipal.cmdHabilitarPresenter.Caption = "Enable"
     Else
        frmPrincipal.cmdHabilitarPresenter.Caption = "Disable"
     End If
     
     'frame tamanho do extrato
     frmPrincipal.Frame5.Caption = "Coupon Size"
     frmPrincipal.Label4.Caption = "Number of Lines:"
     frmPrincipal.Text3.Text = "90"
     frmPrincipal.cmdProgramarExtrato.Caption = "Program"
     frmPrincipal.cmdHabilitarExtrato.Caption = "Enable"
     
     'frame status da impressora
     frmPrincipal.Frame6.Caption = "Printer Status"
     frmPrincipal.cmdStatusImpressora.Caption = "Printer Status"
     
     'outros bot�es
     frmPrincipal.cmdImprimirCaracterGrafico.Caption = "Print Graphic Caracter"
     frmPrincipal.cmdCortarPapel.Caption = "Full Paper Cut"
     frmPrincipal.cmdCortarParcial.Caption = "Partial Paper Cut"
     frmPrincipal.cmdAguardarImpressaoTexto.Caption = "Wait Print Text"
     frmPrincipal.cmdVerificarPapelPresenter.Caption = "Check Paper in Presenter"
     frmPrincipal.cmdSair.Caption = "Exit"
     
     'Frame autentica��o de documentos
     frmPrincipal.Frame10.Caption = "Document Authentication"
     frmPrincipal.Label2.Caption = "Text:"
     frmPrincipal.Label9.Caption = "sec."
     frmPrincipal.Text7.Text = "Authentication Test"
     frmPrincipal.cmdVerificaDocInserido.Caption = "Verify Inserted Document"
     frmPrincipal.cmdAutenticacao.Caption = "Validate Document"
     
     
     ' Tradu��o da Aba da Usando o Driver de Spooler para o Ingl�s
  
     frmPrincipal.SSTab1.TabCaption(2) = "Using the Spooler Driver"
     frmPrincipal.Label5.Caption = "Enter your text:"
     frmPrincipal.Text5.Text = "Bematech Solutions"
     frmPrincipal.cmdModificarFonte.Caption = "Change Font"
     frmPrincipal.cmdImprimir.Caption = "Print"
     frmPrincipal.Label6.Caption = "Image Printing"
     frmPrincipal.cmdImprimirFigura.Caption = "Print"
     frmPrincipal.Frame7.Caption = "Information"
     frmPrincipal.Label7.Caption = "Printers:"
     
     ' Tradu��o da Aba da Usando Codigo de Barras com a API
     frmPrincipal.SSTab1.TabCaption(1) = "API Barcode printing with the API"
     frmPrincipal.frmCodigoBarras.Caption = "Choose the barcode"
     frmPrincipal.frmLarguraBarras.Caption = "Bar width"
     frmPrincipal.optFinas.Caption = "Thin"
     frmPrincipal.optMedias.Caption = "Medium (default)"
     frmPrincipal.optGrossas.Caption = "Thick"
     frmPrincipal.frmPosicaoCaracter.Caption = "Character position"
     frmPrincipal.optAcima.Caption = "Top of barcode"
     frmPrincipal.optAbaixo.Caption = "bottom of barcode"
     frmPrincipal.optAcimaAbaixo.Caption = "Top and bottom of barcode"
     frmPrincipal.optNaoImprime.Caption = "No character printing"
     frmPrincipal.frmFonte.Caption = "Font"
     frmPrincipal.optNormal.Caption = "Normal"
     frmPrincipal.optCondensada.Caption = "Condensed"
     frmPrincipal.frmCodigo.Caption = "Code"
     frmPrincipal.cmdImprimirCodBarras.Caption = "Print"
     frmPrincipal.lbImprimirCodigo.Caption = "Print barcode"
     
     If frmPrincipal.cmdLigarSensorPoucoPapel.Caption = "Ligar Sensor de Pouco Papel" Then
        frmPrincipal.cmdLigarSensorPoucoPapel.Caption = "Enable Low Paper Sensor"
     Else
        frmPrincipal.cmdLigarSensorPoucoPapel.Caption = "Disable Low Paper Sensor"
     End If
    
    ' tradu��o da aba impress�o de bitmap
     frmPrincipal.SSTab1.TabCaption(3) = "Bitmap printing"
     frmPrincipal.LabelBmpFile.Caption = "File Name"
     frmPrincipal.Frame13.Caption = "Paper Layout"
     frmPrincipal.RetratoBtn.Caption = "Portrait"
     frmPrincipal.PaisagemBtn.Caption = "Landscape"
     frmPrincipal.Label14.Caption = "Paper Width"
     frmPrincipal.Frame14.Caption = "Transform"
     frmPrincipal.Label10.Caption = "Height"
     frmPrincipal.Label11.Caption = "Width"
     frmPrincipal.AjustaBtn.Caption = "Fit on page width"
     frmPrincipal.Frame15.Caption = "Rotate"
     frmPrincipal.Label13.Caption = "Degrees"
     frmPrincipal.Command3.Caption = "Print"
     frmPrincipal.Command2.Caption = "Print"
     
     
    
    
  End If
End Function

