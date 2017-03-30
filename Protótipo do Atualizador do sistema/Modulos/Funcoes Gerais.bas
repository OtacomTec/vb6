Attribute VB_Name = "Funcoes_Gerais"
'*******************************************************************************************
'Módulo............................: Nenhum
'Conexão...........................: Nenhuma
'Formulário........................: Funcoes_Gerais
'Objetivo do formulário............: Funcoes Gerais
'Programação.......................: Marcos Baião
'Data..............................: 14/03/2000
'*******************************************************************************************

Public Const piStrTitulo_Menu = "Only Tech"

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    'API do Windows
    Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    'Declare Function SetWindowPos Lib "user32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Dimensao) As Long
    Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
    Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

    'Constantes para a API
    Public Const apiSempre_No_Topo = -1
    Public Const apiPular_Para_O_Topo = -2
    Public Const apiSem_Dimensionar = &H1
    Public Const apiSem_Mover = &H2
    Public Const apiSem_Ativar = &H10
    Public Const apiEstilo_Da_Janela = &H40
    Public Const apiProximo_hWnd = 2
    Public Const apiAtivar = &H10
    Public Const apiFechar_Janela = &H10
    Public Const apiNormal = 1

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Type tForm_Aberto
    hwnd As Long
    Caption As String
End Type
Public vetFormularios() As tForm_Aberto ' Vetor que recebe a identificação dos formulários abertos


Public intCod_Titulo_Receber As Integer
Public curVal_Titulo_Recdeber As Currency
Public intContrato_Vencido As Integer
Public intOrcamento_Vencido As Integer
Public strMenu_Nome As String
Public intContrato As Integer
Public intJanela_Contrato As Integer
Public strDigitos As String


Public intEmpresa As Integer
'Esta variavel ira armazenar qual interface de contrato esta aberta
'numero 1 = cliente_estatico, 2 = cliente_dinamico, 3 = fornecedor_estatico, 4 = fornecedor_dinamico
Public intNivel As Integer ' Nível do usuário logado
Public intID_Usuario_Logado As Integer ' ID do usuário logado
Public intId_Form As Integer
Public strUsuario_Informacao As String

'Variáveis de segurança
Public booConsultar As Boolean
Public booIncluir As Boolean
Public booAlterar As Boolean
Public booExcluir As Boolean

'Variável que habilita a passagem da tacla Alt dos forms para o Ambiente
Public booCeder_Alt As Boolean

'UserTypes usadas pelas APIs
Public Type Dimensao
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type Registro
    hwnd As Long
    Caption As String
    Size As Dimensao
End Type

Private Const piLonTamanho_Bloco As Long = 16384
Private Const piStrArquivo_Temp As String = "\~imagem.tmp"


'Mensagens padrão de impressão
Public Const piStrMsgContinuo80 As String = "Este relatório é configurado para impressão em formulário contínuo de 80 colunas. Deseja continuar? "
Public Const piStrMsgContinuo132 As String = "Este relatório é configurado para impressão em formulário contínuo de 132 colunas. Deseja continuar? "
Public Const piStrMsgA4Matricial As String = "Este relatório é configurado para impressão em folha solta tamanho A4 impressora matricial. Deseja continuar?"
Public Const piStrMsgA4JatoTinta As String = "Este relatório é configurado para impressão em folha solta tamanho A4 em impressora a jato de tinta. Deseja continuar?"


Private Const PM_REMOVE As Long = &H1
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal LpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Public Function Impressora_Padrao() As Printer

Dim strBuffer As String * 254
Dim iRetValue As Long
Dim strDefaultPrinterInfo As String
Dim tblDefaultPrinterInfo() As String
Dim objPrinter As Printer

' pega as informacoes da impressora padrao
  iRetValue = GetProfileString("windows", "device", ",,,", strBuffer, 254)
  strDefaultPrinterInfo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
  tblDefaultPrinterInfo = Split(strDefaultPrinterInfo, ",")
  For Each objPrinter In Printers
        If objPrinter.DeviceName = tblDefaultPrinterInfo(0) Then
          ' se achou a impressora padrao entao sai
          Exit For
        End If
   Next
   
   ' se nao achou retrona nothing
  If objPrinter.DeviceName <> tblDefaultPrinterInfo(0) Then
      Set objPrinter = Nothing
  End If
  
  Set Impressora_Padrao = objPrinter
  
End Function
Public Function Arredonde(ByVal Valor As Single, Optional ByVal Precisao As Integer = 0) As Single
    Dim sinTemp As Single

    Valor = Valor * (10 ^ Precisao) ' valor vezes dez elevado à precisão
    sinTemp = Int(Valor)

    If (Valor - sinTemp) * 10 Then
        Valor = sinTemp + 1
    End If

    Arredonde = Valor / (10 ^ Precisao)
End Function

Public Function FecharAplicativo(ByVal CaptionJanela As String) As Boolean
    '------------------------------------------------------------------
    'Efetua uma chamada à função FindWindow seguida da PostMessage
    'São duas funções de API do Windows declaradas em General
    'Declarations, bem como suas constantes.
    'FindWindow retorna um número que identifica a tela no ambiente
    'Windows
    'PostMessage manda uma função/ação para a tal janela
    '------------------------------------------------------------------
    On Error GoTo Erro

    Const WinClose = &H10
    Dim lonWinWnd As Long
    lonWinWnd = FindWindow(vbNullString, CaptionJanela)
    If lonWinWnd <> 0 Then
        PostMessage lonWinWnd, WinClose, 0&, 0&
    End If
    Exit Function

Erro:

    'MEXER Call Erro.Erro("FecharAplicativo")'MEXER
    Resume Next
End Function

'Public Function Empresa_Padrao_Inicio(DataCombo_Original As DataCombo, Optional strCod_Empresa As Variant) As String
'    Dim strSql As String
'
'    On Error GoTo Erro
'
'    strSql = strSql & "SELECT DFnome_fantasia, DFCod_Empresa FROM TBempresa ORDER BY DFnome_fantasia"
'
'    Set TBrecordset = New ADODB.Recordset
'        TBrecordset.CursorLocation = adUseClient
'        TBrecordset.Open strSql, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
'
'    Set DataCombo_Original.RowSource = TBrecordset
'        DataCombo_Original.ListField = "DFnome_fantasia"
'        DataCombo_Original.BoundColumn = "DFCod_Empresa"
'
'    Set TBrecordset = Nothing
'
'    DataCombo_Original.BoundText = intEmpresa
'
'    If Not IsMissing(strCod_Empresa) Then
'        strCod_Empresa = intEmpresa
'    End If
'
'    Exit Function
'
'Erro:
'    Call Erro.Erro("Empresa_Padrao_Inicio")
''-----------------------------------------------------------------------------------
''Esta função deverá ser usada no Load do Form, indicando um Data Combo e a variável
''String previamente cadastrada no General Declarations, a qual conterá o código da
''Empresa. Serve para trazer a empresa padrão do usuário.
''-----------------------------------------------------------------------------------
'End Function

Public Function Formata_Hora(Horas As Integer, Minutos As Integer) As String
'******************************************************************************
'Sistema...........................: Only Tech
'Módulo............................: Funções Gerais
'Objetivo do módulo................: Dados dois números inteiros, ditos horas
'                                    e minutos, realiza formatação nessa hora
'                                    de modo que os minutos nunca seja maiores
'                                    que 59, retornando um horário no formato (x)x:xx
'                                    Ex.: 30 e 95 retorna 31:35
'Desenvolvimento...................: Marcos Baião
'Data..............................: 17/07/2000
'******************************************************************************
    Dim MinutosString As String
    If Minutos >= 60 Then
        Inteiro = Int(Minutos / 60)
        Resto = Minutos Mod 60
        Horas = Horas + Inteiro
        Minutos = Resto
    End If
    If Minutos <= 9 Then
        MinutosString = "0" & CStr(Minutos)
    Else
        MinutosString = CStr(Minutos)
    End If
    Formata_Hora = CStr(Horas) & ":" & CStr(MinutosString)
End Function

Public Function Soma_Horas(Hora1 As String, Hora2 As String) As String
'******************************************************************************
'Sistema...........................: Only Tech
'Módulo............................: Funções Gerais
'Objetivo do módulo................: Dados duas strings, ditas horas,
'                                    no formato hh:mm, realiza a soma dessas horas
'                                    retornando um horário no formato (x)x:xx
'Desenvolvimento...................: Marcos Baião
'Data..............................: 18/07/2000
'Observações.......................: Faz chamada à função Formata_Hora
'******************************************************************************
    Dim HorasDaHora1 As Integer
    Dim MinutosDaHora1 As Integer
    Dim HorasDaHora2 As Integer
    Dim MinutosDaHora2 As Integer
    Dim HorasDoTotal As Integer
    Dim MinutosDoTotal As Integer
    HorasDaHora1 = Mid(Hora1, 1, InStr(1, Hora1, ":") - 1)
    MinutosDaHora1 = Mid(Hora1, InStr(1, Hora1, ":") + 1, 2)
    HorasDaHora2 = Mid(Hora2, 1, InStr(1, Hora2, ":") - 1)
    MinutosDaHora2 = Mid(Hora2, InStr(1, Hora2, ":") + 1, 2)
    MinutosDoTotal = MinutosDaHora1 + MinutosDaHora2
    HorasDoTotal = HorasDaHora1 + HorasDaHora2
    Soma_Horas = Formata_Hora(HorasDoTotal, MinutosDoTotal)
End Function

Public Function Abrir_App(Form_Caption As String, Form_Path As String, Optional Formulario As Form) As Long

    '------------------------------------------------------------------
    'Executa o programa definido na variável FormPath por APIs do
    'Windows, declaradas do General Declarations, e a posiciona
    'de forma a não ultrapassar a interface do Only Tech.
    'Caso esse programa já esteja carregado, apenas ativa-o
    'FindWindow retorna um número que identifica a janela no ambiente
    'Windows
    'ShowWindowPos reposiciona a janela
    '------------------------------------------------------------------
    On Error GoTo Erro
    DoEvents

    Const No_Topo = -2
    Const Exibir_Janela = &H40
    Const Nao_Redimensionar = &H1
    Dim hwnd As Long

    hwnd = FindWindow(vbNullString, Form_Caption)
    If hwnd = 0 Then
        Shell Form_Path, vbNormalFocus
        'Acrescenta 50 twips à coordenada y do form para que não apareca em cima
        'do Only Tech.exe
        hwnd = FindWindow(vbNullString, Form_Caption)


        If Not Formulario Is Nothing Then
            SetWindowPos hwnd, No_Topo, Formulario.CurrentX, Formulario.CurrentY + 50, Formulario.Height, Formulario.Width, Exibir_Janela + Nao_Redimensionar
        Else
            SetWindowPos hwnd, No_Topo, frmPrincipal.CurrentX, frmPrincipal.CurrentY + 50, frmPrincipal.Height, frmPrincipal.Width, Exibir_Janela + Nao_Redimensionar
        End If

       '-------------------------------------------------------
       'Rotina responsável por registrar os formulários abertos
       'em um arquivo txt para o controle do frmPricipal

        Dim regForm_Aberto As Registro
        Dim regContagem As Registro

        regForm_Aberto.Caption = Form_Caption
        regForm_Aberto.hwnd = hwnd

        Open "c:\Only Tech\formularios.drg" For Random As #1

            'faz uma contagem até chegar ao último registro
            Do Until EOF(1) Or LOF(1) = 0
                Get #1, , regContagem
            Loop

            Put #1, Seek(1), regForm_Aberto
        Close #1

       '-------------------------------------------------------

    Else
        'O Form está carregado... Ativá-lo
        If Not Formulario Is Nothing Then
            SetWindowPos hwnd, No_Topo, Formulario.CurrentX, Formulario.CurrentY + 50, Formulario.Height, Formulario.Width, Exibir_Janela + Nao_Redimensionar
        Else
            SetWindowPos hwnd, No_Topo, frmPrincipal.CurrentX, frmPrincipal.CurrentY + 50, frmPrincipal.Height, frmPrincipal.Width, Exibir_Janela + Nao_Redimensionar
        End If
    End If

    Abrir_App = hwnd

    Exit Function
Erro:
    'MEXER Call Erro.Erro("Abrir_App")
End Function

'Public Function Mover_Guia(SSTab As SSTab, KeyCode As Integer, Shift As Integer)
'    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
'        KeyCode = 0
'        SSTab.SetFocus
'        SendKeys "^{TAB}"
'    ElseIf Shift = (vbCtrlMask + vbShiftMask) And KeyCode = vbKeyTab Then
'        KeyCode = 0
'        SSTab.SetFocus
'        SendKeys "^+{TAB}"
'    End If
'End Function

Public Function Grava_Moeda(ByVal Valor As Variant) As Variant
    
    Dim strInteiro As String
    Dim strDecimal As String
    Dim strRetorno As String

    If Valor = Empty Or Valor = " " Or Valor = "" Then
        Grava_Moeda = 0
        Exit Function
    End If

    strRetorno = Format(Valor, "#0.0000;-#0.0000")
    strDecimal = Mid(strRetorno, (InStr(1, strRetorno, ",") + 1))
    strInteiro = Mid(strRetorno, 1, (InStr(1, strRetorno, ",") - 1))

    Grava_Moeda = strInteiro & "." & strDecimal
End Function

Public Function Grava_Decimal(ByVal Valor As Variant) As Variant

    Dim strInteiro As String
    Dim strDecimal As String
    Dim strRetorno As String

    If Valor = Empty Then
        Grava_Decimal = 0
        Exit Function
    End If

    strRetorno = Format(Valor, "##0.00;-##0.00")
    strDecimal = Mid(strRetorno, (InStr(1, strRetorno, ",") + 1))
    strInteiro = Mid(strRetorno, 1, (InStr(1, strRetorno, ",") - 1))

    Grava_Decimal = strInteiro & "." & strDecimal

End Function

Public Function Verifica_Numero(Nome_Campo As String, Nome_Tabela As String, Nome_textbox As TextBox) As Boolean
    If Nome_textbox.Text = Empty Then
        Exit Function
    End If

    Dim SQL As String

    On Error GoTo Erro
    SQL = ""
    SQL = SQL & "SELECT " & Nome_Campo & " "
    SQL = SQL & "FROM " & Nome_Tabela & " "
    SQL = SQL & "WHERE " & Nome_Campo & " "
    SQL = SQL & "= " & Nome_textbox.Text & ""

    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open SQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText

    If Val(TBrecordset(Nome_Campo)) = Val(Nome_textbox.Text) Then
        MsgBox "Registro já existente", vbCritical, "Only Tech"
        Nome_textbox.Text = Empty
        Nome_textbox.SetFocus
        Verifica_Numero = True
    Else
        Verifica_Numero = False
    End If
    Exit Function

Erro:
    If Err.Number = 3021 Then
        Verifica_Numero = False
        Exit Function
    Else
        'MEXER Call Erro.Erro("Verifica_Numero")
    End If
    Resume Next

End Function
Public Function Localiza_ID(Nome_Campo_ID As String, Nome_Campo_Codigo As String, Valor_Campo_Codigo As String, Nome_Tabela As String, Aplicacao As String, Form As Object, Optional Banco As String, Optional Nome_Campo_Codigo_Empresa As String, Optional Valor_Campo_Codigo_Empresa As Integer, Optional Nome_Extra_Comparacao_Campo As String, Optional Valor_Extra_Comparacao_Campo As String, Optional Nome_Extra_Comparacao_Campo2 As String, Optional Valor_Extra_Comparacao_Campo2 As String, Optional Nome_Data_Source As String) As String

    Dim SQL As String
    Dim rstID As New ADODB.Recordset
    Dim conexao_id As New DLLConexao_Sistema.Conexao

    SQL = Empty
    SQL = "SELECT " & Nome_Campo_ID & " " & _
          "FROM " & Nome_Tabela & " " & _
          "WHERE " & Nome_Campo_Codigo & " = '" & Valor_Campo_Codigo & "' "

    If Nome_Campo_Codigo_Empresa <> Empty Then
       SQL = SQL + "AND " & Nome_Campo_Codigo_Empresa & " = " & Valor_Campo_Codigo_Empresa & " "
    End If
    
    If Nome_Extra_Comparacao_Campo <> Empty Then
       SQL = SQL + "AND " & Nome_Extra_Comparacao_Campo & " = '" & Valor_Extra_Comparacao_Campo & "'"
    End If
    
    If Nome_Extra_Comparacao_Campo2 <> Empty Then
       SQL = SQL + "AND " & Nome_Extra_Comparacao_Campo2 & " = '" & Valor_Extra_Comparacao_Campo2 & "'"
    End If
    
    'Indicando o banco à conectar-se
    conexao_id.Data_Source = Nome_Data_Source
    
    conexao_id.Initial_Catalog = Banco
    
    conexao_id.Abrir_conexao (Aplicacao)

    rstID.CursorLocation = adUseClient
    rstID.Open SQL, conexao_id.CNconexao, adOpenStatic, adLockReadOnly

    Localiza_ID = rstID.Fields(0)
    
    Set rstID = Nothing
    
    conexao_id.Fechar_conexao

End Function
Public Function Localiza_Proximo_Codigo(Campo_Para_Selecionar As String, Campo_Empresa As String, Valor_Campo_Empresa As Integer, Nome_Tabela As String, Aplicacao As String, Form As Object, Optional Banco As String) As String

'    On Error GoTo Erro

    Dim SQL As String
    Dim rstID As New ADODB.Recordset
    Dim conexao_id As New DLLConexao_Sistema.Conexao

    SQL = Empty
    SQL = "SELECT " & Campo_Para_Selecionar & " " & _
          "FROM " & Nome_Tabela & " " & _
          "WHERE " & Campo_Empresa & " = " & Valor_Campo_Empresa & " "

    'Indicando o banco à conectar-se
    conexao_id.Initial_Catalog = Banco

    conexao_id.Abrir_conexao (Aplicacao)

    rstID.CursorLocation = adUseClient
    rstID.Open SQL, conexao_id.CNconexao, adOpenStatic, adLockReadOnly

    Localiza_Proximo_Codigo = rstID.Fields(0)


'    Exit Function

'Erro:
 '   Call Erro.Erro(Form, Aplicacao, "Localiza_Proximo_Codigo")

End Function


Public Function Verifica_Texto(Nome_Campo As String, Nome_Tabela As String, Nome_textbox As TextBox) As Boolean
    If Nome_textbox = Empty Then
        Exit Function
    End If

    Dim SQL As String

    On Error GoTo Erro
    
    SQL = ""
    SQL = SQL & "SELECT " & Nome_Campo & " "
    SQL = SQL & "FROM " & Nome_Tabela & " "
    SQL = SQL & "WHERE " & Nome_Campo & " "
    SQL = SQL & "= '" & Nome_textbox & "'"

    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open SQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText

    If TBrecordset(Nome_Campo) = Nome_textbox Then
        MsgBox "Registro já existente", vbCritical, "Only Tech"
        Nome_textbox = Empty
        Nome_textbox.SetFocus
        Verifica_Texto = True
    Else
        Verifica_Texto = False
    End If
    Exit Function
Erro:
    If Err.Number = 3021 Then
        Verifica_Texto = False
        Exit Function
    Else
        'MEXER Call Erro.Erro("Verifica_Texto")
    End If
    Resume Next
End Function


Public Function Ler_Imagem(strTabela As String, strCampo_Imagem As String, strCampo_Codigo As String, strValor_Codigo As String) As IPictureDisp
    On Error GoTo Erro
    Dim adrTemp As ADODB.Recordset
    Dim strSql As String

    Dim bytDados() As Byte
    Dim varTemp As Variant
    Dim f As Long


    '----------------------------------------------------------------------------
    'abrindo Recordset
    strSql = _
        "SELECT " & strCampo_Imagem & " FROM " & strTabela & " " & _
        "WHERE " & strCampo_Codigo & " = " & strValor_Codigo

    Set adrTemp = New ADODB.Recordset
        adrTemp.CursorLocation = adUseClient
        adrTemp.Open strSql, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
    '----------------------------------------------------------------------------


    On Error GoTo Erro_Lendo_Imagem
    f = FreeFile()
    Open App.Path & piStrArquivo_Temp For Binary As #f

        'As informações são retornadas do banco em blocos pela função GetChunk
        Do
            varTemp = adrTemp(strCampo_Imagem).GetChunk(piLonTamanho_Bloco)
  
            If IsNull(varTemp) Then
                Set Ler_Imagem = Nothing
                Close #f
                Exit Function
            End If

            bytDados = varTemp
            Put #f, , bytDados
        Loop While LenB(varTemp) = piLonTamanho_Bloco

    Close #f

    Set Ler_Imagem = LoadPicture(App.Path & piStrArquivo_Temp)
    Kill App.Path & piStrArquivo_Temp


    Exit Function

Erro_Lendo_Imagem:
    Close #f
    Exit Function
Erro:
    'MEXER Call Erro.Erro("Ler_Imagem")

'-------------------------------------------------------------------------
'Carrega uma imagem do banco
'Parâmetros:
'  • strTabela: recebe o nome da tabela
'  • strCampo_Imagem: recebe o nome do campo que receberá a imagem
'  • strCampo_Codigo: o nome do campo chave na tabela
'  • strValor_Codigo: o valor do campo chave
'-------------------------------------------------------------------------
End Function

Public Sub Salvar_Imagem(Imagem As IPictureDisp, strTabela As String, strCampo_Imagem As String, strCampo_Codigo As String, strValor_Codigo As String)
    On Error GoTo Erro
    Dim adrTemp As ADODB.Recordset
    Dim strSql As String

    Dim bytDados() As Byte
    Dim varTemp As Variant
    Dim lonTamanho_Arquivo As Long
    Dim lonBytes_Lidos As Long

    Const Limite_Bloco As Long = 1048576 'o máximo que cada bloco pode receber
   

    'abrindo Recordset
    strSql = _
        "SELECT " & strCampo_Imagem & " FROM " & strTabela & " " & _
        "WHERE " & strCampo_Codigo & " = " & strValor_Codigo

    Set adrTemp = New ADODB.Recordset
        adrTemp.CursorLocation = adUseClient
        adrTemp.Open strSql, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
    


    If Imagem Is Nothing Then
        adrTemp(strCampo_Imagem).Value = Null
        adrTemp.Update
        adrTemp.Close
        Set adrTemp = Nothing
        Exit Sub

    ElseIf Imagem = 0 Then
        adrTemp(strCampo_Imagem).Value = Null
        adrTemp.Update
        adrTemp.Close
        Set adrTemp = Nothing
        Exit Sub
    End If

    'salvando a imagem em um arquivo temporário
    SavePicture Imagem, App.Path & piStrArquivo_Temp

    On Error GoTo Erro_Lendo_Imagem
    Open App.Path & piStrArquivo_Temp For Binary As #1
        lonTamanho_Arquivo = LOF(1)

        If lonTamanho_Arquivo > Limite_Bloco Then

            Do While lonTamanho_Arquivo <> lonBytes_Lidos
                If lonTamanho_Arquivo - lonBytes_Lidos < piLonTamanho_Bloco Then
                    bytDados = InputB(lonTamanho_Arquivo - lonBytes_Lidos, 1)
                    lonBytes_Lidos = lonTamanho_Arquivo

                Else
                    bytDados = InputB(piLonTamanho_Bloco, 1)
                    lonBytes_Lidos = lonBytes_Lidos + piLonTamanho_Bloco
                End If

                adrTemp(strCampo_Imagem).AppendChunk bytDados
            Loop

        Else
            bytDados = InputB(lonTamanho_Arquivo, 1)
            adrTemp(strCampo_Imagem).Value = bytDados
        End If
    Close #f

    'salvando os dados
        adrTemp.Update
        adrTemp.Close
    Set adrTemp = Nothing

    Kill App.Path & piStrArquivo_Temp


    Exit Sub

Erro_Lendo_Imagem:
    Close #1
    Exit Sub
Erro:
    'MEXER Call Erro.Erro("Salvar_Imagem")

'-------------------------------------------------------------------------
'Salva uma imagem no banco de dados
'Parâmetros:
'  • Imagem: recebe a propriedade imagem de um objeto
'        Expl. Picture1.Picture
'  • strTabela: recebe o nome da tabela
'  • strCampo_Imagem: recebe o nome do campo que receberá a imagem
'  • strCampo_Codigo: o nome do campo chave na tabela
'  • strValor_Codigo: o valor do campo chave
'-------------------------------------------------------------------------
End Sub

Public Sub Imprime_Listagem(Nome_do_form As Form, Altura_Atual As Integer, Nome_do_Relatorio As String, strSQL_Impressao As String, Optional Mensagem As String = piStrMsgContinuo80)
    '******************************************************************************
    'Módulo............................: Funcoes_Gerais
    'Procedimento/Função...............: Imprime_Listagem
    'Objeto/classe correspondente......: -
    'Objetivo:.........................: Imprime uma listagem dos registros existentes no
    '                                    grid da guia de listagem nas telas de cadastro
    '                                    padrão
    'Desenvolvimento...................: Marcos Baião
    'Data de criação...................: 16/03/2001
    'Utilização........................: Call Funcoes_Gerais.Imprime_Listagem(v1, v2, v3, v4, v5, v6)
    'Parâmetros de entrada.............: Variável v1 é o nome do form que chama o
    '                                    procedimento, v2 é a altura normal do form,
    '                                    v3 é o nome do relatório incluindo sua extensão,    '
    '                                    v4 é uma cláusula SQL construída no procedimento
    '                                    Reposição que chama o Imprime_Listagem,
    '                                    v5 é a constante relativa à mensagem de impressão
    '                                    a ser exibida para o usuário (veja
    '                                    Funcoes_Gerais.General ou
    '                                    Padrao_de_desenvolvimento_processa.doc) e
    '                                    v6 é opcional e caso não seja informada o default

    '                                    é uma mensagem de impressão em matricial 80 colunas
    'Saída.............................: Dispara a impresão do relatório
    'Observações.......................:
    '******************************************************************************
    On Error GoTo Mnpld_cmdImprimir_Click
    Dim strSenha_Admin As String
    Dim strServidor As String
    Dim strSQLEmpresa As String
    Dim adrEmpresa As ADODB.Recordset
    Dim I As Integer
    Dim adrTBTemp As ADODB.Recordset
    Dim bTabela_Vazia As Boolean

    bTabela_Vazia = False
        
    If Mensagem <> Empty Then
        If MsgBox(Mensagem, vbInformation + vbYesNo, "Only Tech") = vbNo Then
            Exit Sub
        End If
    End If

    'Comentado por causa da mercedes
    'Nome_do_form.ProgressBar1.Visible = True
    'Nome_do_form.ProgressBar1.Max = 100
    'Nome_do_form.Height = Altura_Atual + 270

    DoEvents
    Screen.MousePointer = 11
    'For I = 0 To 25
        'Nome_do_form.ProgressBar1.Value = I
    'Next

    'Verifica se a tabela existe e apaga a tabela
    Conexao.Execute "If Exists(SELECT * FROM sysObjects WHERE id = Object_id('dbo.TBTemp_Impressao'))Begin DROP TABLE TBTemp_Impressao End "

    'strSQLEmpresa = "SELECT DFnome_fantasia FROM TBempresa WHERE DFcod_empresa = " & intEmpresa & ""
    'Call Banco_Dados.SQLgeral(strSQLEmpresa, adrEmpresa)

    'Gera uma adr utilizando a variável strSQLImpressao
    'Grava a tabela temporária
    Conexao.Execute strSQL_Impressao

    Call conexao_banco.SQLgeral("SELECT * FROM TBTemp_Impressao ", adrTBTemp)

    If adrTBTemp.RecordCount = 0 Then
        bTabela_Vazia = True
    End If

    If bTabela_Vazia = True Then
        Screen.MousePointer = 0
        Beep
        MsgBox "Não há registros que satisfaçam a elaboração do relatório.", vbInformation, "Only Tech"
        Nome_do_form.ProgressBar1.Value = 0
        Nome_do_form.ProgressBar1.Visible = False
        Nome_do_form.Height = Altura_Atual
        Exit Sub
    End If

    For I = 26 To 50
        Nome_do_form.ProgressBar1.Value = I
    Next

    'Depois de criado o executável tirar o comentário abaixo
    'Nome_do_form.crpImpressao_Listagem.ReportFileName = App.Path & "\" & Nome_do_Relatorio
    Nome_do_form.crpImpressao_Listagem.ReportFileName = "F:\Projetos\Structure\Rpt\" & Nome_do_Relatorio
    'Nome_do_form.crpImpressao_Listagem.ReportFileName = "C:\Balancas\" & Nome_do_Relatorio 'Maquina do Pedrinho

    For I = 51 To 100
        Nome_do_form.ProgressBar1.Value = I
    Next

    On Error Resume Next
        Nome_do_form.txtConsulta.SetFocus
    On Error GoTo Mnpld_cmdImprimir_Click

    Nome_do_form.crpImpressao_Listagem.WindowTitle = Nome_do_form.Caption
    Nome_do_form.crpImpressao_Listagem.WindowState = crptMaximized
    Nome_do_form.crpImpressao_Listagem.DiscardSavedData = True


    Nome_do_form.crpImpressao_Listagem.Action = 1

    'Verifica se a tabela existe e apaga a tabela
    Conexao.Execute "If Exists(SELECT * FROM sysObjects WHERE id = Object_id('dbo.TBTemp_Impressao'))Begin DROP TABLE TBTemp_Impressao End "

    Nome_do_form.ProgressBar1.Visible = False
    Nome_do_form.Height = Altura_Atual
    DoEvents

    Screen.MousePointer = 0
    Nome_do_form.ProgressBar1.Value = 0

    Exit Sub

Mnpld_cmdImprimir_Click:
    If Err.Number = 20507 Then
        Beep
        MsgBox "Não foi possível encontrar o arquivo de impressão.", vbCritical, "Only Tech"
    Else
        'MEXER Call Erro.Erro("Imprime_Listagem")
    End If


    Screen.MousePointer = 0
    Nome_do_form.ProgressBar1.Value = 0
    Nome_do_form.ProgressBar1.Visible = False
    If Nome_do_form.WindowState <> vbMinimized And Nome_do_form.WindowState <> vbMaximized Then
        Nome_do_form.Height = Altura_Atual
    End If
    DoEvents

End Sub

'Public Sub Controlar_Guia(Nome_Form As Form, Nome_SSTab As SSTab, Optional ByRef intBotao As Integer = -1, Optional booControlar_Botoes As Boolean = True)
    '******************************************************************************
    'Módulo............................: Funcoes_Gerais
    'Procedimento/Função...............: Controlar_Guias
    'Objeto/classe correspondente......: -
    'Objetivo:.........................: Cotrola o TabStop, TabIndex e o Foco de todos os
    '                                    controles da guia selecionada, bem como de guias
    '                                    que necessitem de compartilhamento.
    'Desenvolvimento...................: José Braga
    'Data de criação...................: 12/04/2001
    'Utilização........................: Call Funcoes_Gerais.Controlar_Guias(Nome_Form, Nome_SSTab)
    'Parâmetros de entrada.............: Nome_Form é o nome do formulário corrente (Caso
    '                                    preferir poderá ser usado "Me" p/ se referir ao Form),
    '                                    Nome_SSTab é o nome do controle SSTab usado.
    'Saída.............................: A chamada a esta função deve ser feita no evento
    '                                    Click do controle SSTab.
    'Data da última manutenção.........: 29/01/2002
    'Manutenção executada por..........: Marcos Baião
    'Observações.......................: Foi alterado alguns pontos do codigo, para que pudesse
    '                                    ser inserido uma guia dentro de outra.
    '
    'Orientações para implementação:...: 1. Colocar a chamada da função no evento click da
    '                                       tab
    '                                    2. Atualizar a propriedade tag de cada controle
    '                                       com o índice da guia à qual ele pertence
    '                                    3. Controles que não devem receber o foco
    '                                       permanecem com a tag vazia
    '                                    4. Colocar no primeiro controle de cada guia
    '                                       a instrução
    '                                       Se nome_da_tab.Tab <> 0 então
    '                                           nome_da_tab.tab = n
    '                                       senão
    '                                           Call nome_da_tab_click
    '                                       fim-se
    '                                       onde n é o
    '                                       índice da guia
    '                                    5. Caso o primeiro controle da guia possa não
    '                                       receber o foco, o item 4 deve ser feito também
    '                                       para o segundo controle.
    '                                    6. Colocar no último controle de cada guia instrução
    '                                       do item 4
    '                                    7. Caso o último controle da guia possa não receber
    '                                       foco, o item 4 deve ser feito também
    '                                       para o penúltimo controle.
    '                                    8. Os itens 4 a 7 devem ser analisados para cada
    '                                       interface.
    '                                    9. Em guias onde existe uma continuidade de
    '                                       tabulação deve ser feita a seguinte
    '                                       identificação na tag:
    '                                       x-y;z
    '                                       onde x = índice da guia onde se encontra o
    '                                       controle,
    '                                       y = índice da guia onde inicia-se a continuidade
    '                                       da tabulação,
    '                                       z = índice da guia onde termina continuidade
    '                                       da tabulação.
    '                                    10.Em interfaces onde a propriedade tag é utilizada
    '                                       o valor armazenado na tag deve ser armazenado
    '                                       em variável
    '
    '******************************************************************************

'    Dim Indice As Integer, I As Integer
'    Dim booControle_Aceito As Boolean
'    Dim booTabIndex_Aceito As Boolean
'    Dim X As Integer, Y As Integer
'    Dim Objeto As Variant
'    Dim intValor_Botao_Atual As Integer
'
'
'    If intBotao = -1 Then
'        intValor_Botao_Atual = intValor_Botao ' utilizando a variável global
'    Else
'        intValor_Botao_Atual = intBotao ' utilizando o parâmetro
'    End If
'
'
'    ReDim intvetor(0) As Integer
'    intvetor(0) = -1
'
'    'Dá um tempo p/ que a tab seja montada completamente antes do controle receber foco ----
'    If intValor_Botao_Atual = 1 Or intValor_Botao_Atual = 5 Or Nome_SSTab.Caption = "Listagem" Then
'        DoEvents
'    End If
'    '---------------------------------------------------------------------------------------
'
'    On Error Resume Next
'
'    'Prende o foco na guia ativa -----------------------------------------------------------
'    For Each Objeto In Nome_Form.Controls
'        If Objeto.Tag <> Empty Then
'            If Objeto.Name = "adbEmpresa" Then
'                Objeto.TabStop = False
'            Else
'                If InStr(1, Objeto.Tag, "-") <> 0 Then
'                    If Nome_SSTab.Tab >= Mid(Objeto.Tag, (InStr(1, Objeto.Tag, "-") + 1), ((InStr(1, Objeto.Tag, ";") - 1) - InStr(1, Objeto.Tag, "-"))) And Nome_SSTab.Tab <= Mid(Objeto.Tag, (InStr(1, Objeto.Tag, ";") + 1)) Then
'                        If TypeOf Objeto Is OptionButton Then
'                            Objeto.TabStop = Objeto.Value
'                        Else
'                            Objeto.TabStop = True
'                        End If
'                    Else
'                        Objeto.TabStop = False
'                    End If
'                Else
'                    If Nome_SSTab.Tab = Objeto.Tag Then
'                        If TypeOf Objeto Is OptionButton Then
'                            Objeto.TabStop = Objeto.Value
'                        Else
'                            Objeto.TabStop = True
'                        End If
'                    Else
'                        Objeto.TabStop = False
'                    End If
'                End If
'            End If
'
'            'Adiciona cada TabIndex achado ao vetor
'            If Objeto.TabStop = True Then
'                booTabIndex_Aceito = False
'                'Verifica se o controle pertence à guia selecionada para gravar o TabIndex
'                If InStr(1, Objeto.Tag, "-") <> 0 Then
'                    If Mid(Objeto.Tag, 1, (InStr(1, Objeto.Tag, "-") - 1)) = Nome_SSTab.Tab Then
'                        booTabIndex_Aceito = True
'                    End If
'                Else
'                    If Objeto.Tag = Nome_SSTab.Tab Then
'                        booTabIndex_Aceito = True
'                    End If
'                End If
'                '---------------------------------------------------------------------------
'                If booTabIndex_Aceito = True Then
'                    If intvetor(LBound(intvetor)) <> -1 Then
'                        ReDim Preserve intvetor(UBound(intvetor) + 1)
'                    End If
'                    intvetor(UBound(intvetor)) = Objeto.TabIndex
'                End If
'            End If
'            '-------------------------------------------------------------------------------
'        End If
'    Next Objeto
'    '---------------------------------------------------------------------------------------
'
'    'Ordena TabIndex dentro do vetor (Sort de Bolha) ---------------------------------------
'    Do While Indice < UBound(intvetor)
'        If intvetor(Indice) > intvetor(Indice + 1) Then
'            X = intvetor(Indice)
'            Y = intvetor(Indice + 1)
'            intvetor(Indice) = Y
'            intvetor(Indice + 1) = X
'            Indice = 0
'        Else
'            Indice = Indice + 1
'        End If
'    Loop
'    '---------------------------------------------------------------------------------------
'
'    If Nome_Form.ActiveControl.Name = Nome_SSTab.Name Then
'        Indice = 0
'        Err.Number = 0
'        Do While I <= Nome_Form.Count
'            If Nome_Form.Controls(I).Tag <> "" Then
'                'Verifica se o controle pertence à guia selecionada ------------------------
'                If InStr(1, Nome_Form.Controls(I).Tag, "-") <> 0 Then
'                    If Mid(Nome_Form.Controls(I).Tag, 1, (InStr(1, Nome_Form.Controls(I).Tag, "-") - 1)) = Nome_SSTab.Tab Then
'                        booControle_Aceito = True
'                    Else
'                        booControle_Aceito = False
'                    End If
'                Else
'                    If Nome_Form.Controls(I).Tag = Nome_SSTab.Tab Then
'                        booControle_Aceito = True
'                    Else
'                        booControle_Aceito = False
'                    End If
'                End If
'                '---------------------------------------------------------------------------
'
'                'Caso o controle pertença à guia, continua com as verificações --------------
'                If booControle_Aceito = True Then
'                    If Nome_Form.Controls(I).TabIndex = intvetor(Indice) Then
'                        If Err.Number = 0 Then
'                            If Nome_Form.Controls(I).TabStop = True And Nome_Form.Controls(I).Visible = True And Nome_Form.Controls(I).Enabled = True Then
'                                If intValor_Botao_Atual = 1 Or intValor_Botao_Atual = 5 Then
'                                    If booControlar_Botoes = True Then
'                                        Nome_Form.Controls(I).SetFocus
'                                    End If
'                                Else
'                                    If Nome_SSTab.Caption = "Listagem" Then
'                                        Nome_Form.Controls(I).SetFocus
'                                    Else
'                                        Nome_Form.cmdIncluir.SetFocus
'                                        If Err.Number <> 0 Then
'                                            DoEvents
'                                            Nome_Form.Controls(I).SetFocus
'                                        End If
'                                    End If
'                                End If
'                                Exit Do
'                            Else
'                                If Indice < UBound(intvetor) Then
'                                    I = 0
'                                    Indice = Indice + 1
'                                Else
'                                    Exit Do
'                                End If
'                            End If
'                        Else
'                            If I < Nome_Form.Count Then
'                                I = I + 1
'                                Err.Number = 0
'                            Else
'                                I = 0
'                                Err.Number = 0
'                                Indice = Indice + 1
'                            End If
'                        End If
'                    Else
'                        If I <= Nome_Form.Count Then
'                            I = I + 1
'                        End If
'                    End If
'                Else
'                    I = I + 1
'                End If
'            Else
'                I = I + 1
'            End If
'        Loop
'    End If
'    '---------------------------------------------------------------------------------------
'
'    'Controle dos botões -------------------------------------------------------------------
'    If booControlar_Botoes = True Then
'        If Nome_SSTab.Tab = (Nome_SSTab.Tabs - 1) Then
'            If intValor_Botao_Atual = 1 Or intValor_Botao_Atual = 5 Then
'                Nome_Form.cmdCancelar.Enabled = True
'            Else
'                Nome_Form.cmdCancelar.Enabled = False
'            End If
'            Nome_Form.cmdIncluir.Enabled = booIncluir And (intValor_Botao_Atual <> 1)
'            Nome_Form.cmdAlterar.Enabled = booAlterar
'            Nome_Form.cmdConfirmar.Enabled = False
'            Nome_Form.cmdExcluir.Enabled = booExcluir And (intValor_Botao_Atual <> 5)
'        Else
'            If intValor_Botao_Atual = 1 Or intValor_Botao_Atual = 5 Then
'                Nome_Form.cmdIncluir.Enabled = False
'                Nome_Form.cmdConfirmar.Enabled = (booIncluir And intValor_Botao_Atual = 1) Or (booAlterar And intValor_Botao_Atual = 5)
'                Nome_Form.cmdCancelar.Enabled = True
'                Nome_Form.cmdAlterar.Enabled = False
'                Nome_Form.cmdExcluir.Enabled = False
'            Else
'                Nome_Form.cmdIncluir.Enabled = booIncluir
'                Nome_Form.cmdConfirmar.Enabled = False
'                Nome_Form.cmdCancelar.Enabled = False
'                Nome_Form.cmdAlterar.Enabled = False
'                Nome_Form.cmdExcluir.Enabled = False
'            End If
'        End If
'    End If
'    '---------------------------------------------------------------------------------------
'
'
'    If intBotao = -1 Then
'        intBotao = intValor_Botao_Atual ' utilizando o parâmetro
'    Else
'        intValor_Botao = intValor_Botao_Atual ' utilizando a variável global
'    End If
'
'
'
'End Sub
'
Public Function Formatar_Valor(Valor As String, Optional Formato As String = "#,##0.00") As String
    If IsNumeric(Valor) Then
        Formatar_Valor = Format(Valor, Formato)
    Else
        Formatar_Valor = Empty
    End If
End Function

Public Function Verifica_Apostrofo(strValores As String) As String
'******************************************************************************
'Sistema...........................: Only Tech
'Módulo............................:
'Procedimento/Função...............: Verifica_Apostrofo
'Objetivo:.........................: Substituir os Apostrofos cadastrados pelo usuário,
'                                    pelo comando CHAR(39) do SQL.
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 27/04/2001
'Observaçãoes......................:
'******************************************************************************
    On Error GoTo Erro
    Dim strTexto_anterior As String, strTexto_posterior As String
    Dim strTexto_corrigido As String
    Dim strValores_originais As String
    Dim intLetra_anterior As Integer, intLetra_posterior As Integer
    Dim intPosicao_apostrofo As Integer
    Dim intCont As Integer, intQuant_apostrofos As Integer, intLoops As Integer
    
    If InStr(1, strValores, "'") = 0 Then
        Verifica_Apostrofo = strValores
        Exit Function
    End If

    Do
        intPosicao_apostrofo = InStr(intPosicao_apostrofo + 1, strValores, "'")
        intCont = intCont + 1

    Loop While InStr(intPosicao_apostrofo + 1, strValores, "'") <> 0

    intPosicao_apostrofo = 0
    strValores = LTrim(RTrim(strValores))
    intPosicao_apostrofo = InStr(1, strValores, "'")

    If intPosicao_apostrofo = 1 Then
        strValores = "'" & LTrim(Mid(strValores, 2, Len(strValores)))
        intPosicao_apostrofo = InStr(2, strValores, "'")
    End If

    strValores_originais = strValores
    intQuant_apostrofos = 1

    Do While intQuant_apostrofos < intCont

        strValores = RTrim(Mid(strValores, 1, intPosicao_apostrofo - 1)) & "'" & LTrim(Mid(strValores, intPosicao_apostrofo + 1, Len(strValores) - intPosicao_apostrofo))
        intQuant_apostrofos = intQuant_apostrofos + 1
        intPosicao_apostrofo = 1

        For intLoops = 1 To intQuant_apostrofos
            intPosicao_apostrofo = InStr(intPosicao_apostrofo + 1, strValores, "'")
        Next

    Loop
   
    intPosicao_apostrofo = 0

    intPosicao_apostrofo = InStr(1, strValores, "'")
    intLetra_anterior = intPosicao_apostrofo - 1
    intLetra_posterior = intPosicao_apostrofo + 1
         
    Do While intPosicao_apostrofo <> 0

        If intPosicao_apostrofo = 1 Or intPosicao_apostrofo = Len(strValores) Then

            intPosicao_apostrofo = InStr(intPosicao_apostrofo + 1, strValores, "'")
            intLetra_anterior = intPosicao_apostrofo - 1
            intLetra_posterior = intPosicao_apostrofo + 1

        ElseIf Mid(strValores, intLetra_anterior, 1) = "," Or Mid(strValores, intLetra_posterior, 1) = "," Or Mid(strValores, intLetra_anterior, 1) = Space(1) Or Mid(strValores, intLetra_posterior, 1) = Space(1) Then

            intPosicao_apostrofo = InStr(intPosicao_apostrofo + 1, strValores, "'")
            intLetra_anterior = intPosicao_apostrofo - 1
            intLetra_posterior = intPosicao_apostrofo + 1

        Else

            strTexto_anterior = Left(strValores, intPosicao_apostrofo - 1)
            strTexto_posterior = Right(strValores, Len(strValores) - intPosicao_apostrofo)
            strValores = strTexto_anterior & "' + CHAR(39) + '" & strTexto_posterior

            intPosicao_apostrofo = InStr(intPosicao_apostrofo + 16, strValores, "'")
            intLetra_anterior = intPosicao_apostrofo - 1
            intLetra_posterior = intPosicao_apostrofo + 1

        End If

    Loop

    Verifica_Apostrofo = strValores

    Exit Function
Erro:
    'MEXER Erro.Erro ("Verifica_Apostrofo")
    Resume Next
End Function

Public Sub Verifica_Tecla(KeyCode As Integer, Form As Form, Shift As Integer) ', Optional Guia As SSTab)
'******************************************************************************
'Sistema...........................: Only Tech
'Módulo............................:
'Procedimento/Função...............: Verifica_Tecla
'Objetivo:.........................: Fazer verificação de Teclas para que se possa
'                                    por exemplo mandar uma mensagem apertando uma
'                                    tecla e para substituir o KeyPress do Form.
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 02/05/2001
'Observaçãoes......................: Alterei a parte do código que ativa o menu para
'                                    simplificar o código.
'
'******************************************************************************
    'Antiga função de Ativar menu
    On Error GoTo Erro
    Dim Objeto As Variant
    Dim booStatus As Boolean
    

    ' esta rotina trata os atalhos para os botões do formuláriao
    ' caso um dos atalhos seja executado o parâmetro Status retorna true, portanto,
    ' não será mais necessário continuar a verificação de teclas
    Call Botoes.Atalhos(Form, KeyCode, Shift, booStatus)
    If booStatus = True Then
        Exit Sub
    End If
    

    ' Ativando o menu
    If Shift = vbAltMask Then
        If KeyCode >= vbKeyA And KeyCode <= vbKeyZ Then
            On Error Resume Next

            AppActivate "Only Tech"
            If Err.Number = 0 Then
                SendKeys "%{" & Chr(KeyCode) & "}"
            End If

            Exit Sub
        End If
    End If


    'Verifica se a tecla F12 foi acionada
    If KeyCode = vbKeyF12 Then
        MsgBox "Número da última versão do executável " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "Only Tech"
        KeyCode = 0
    'Verifica se a tecla F2 foi acionada
    ElseIf KeyCode = vbKeyF2 And Not Form Is Nothing Then
        For Each Objeto In Form.Controls
            If Objeto.Name = "adbEmpresa" Then
'                If Guia.Name <> Empty Then
'                    Guia.Tab = Left(Form.adbEmpresa.Tag, 1)
                    If Form.adbEmpresa.Enabled Then
                        Form.adbEmpresa.SetFocus
                    End If
'                    Exit For
'                End If
'                Form.adbEmpresa.SetFocus
            End If
        Next Objeto
        KeyCode = 0
    'Verifica se a tecla Enter foi acionada
    ElseIf KeyCode = vbKeyReturn Then
        If Form.KeyPreview = True Then
            If Shift <> vbCtrlMask Then
                SendKeys "{TAB}"
            End If
        End If
        KeyCode = 0
    'Verifica se a tecla ESC foi acionada
    ElseIf KeyCode = vbKeyEscape Then
        Unload Form
    End If

    Exit Sub

Erro:
    'MEXER Call Erro.Erro("Verifica_Tecla")
    Resume Next
End Sub


Public Sub Verifica_Estado(Nome_Objeto As Variant)
'******************************************************************************
'Sistema...........................: Only Tech
'Módulo............................: Funcoes_Gerais
'Procedimento/Função...............: Verifica_Estado
'Objetivo:.........................: Faz a validação dos estados
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 19/06/2001
'Observaçãoes......................:
'******************************************************************************
    On Error GoTo Erro

    If Len(Nome_Objeto.Text) < 2 And Len(Nome_Objeto.Text) <> 0 Then
        MsgBox "Para cadastrar uma unidade federativa são necessárias 2 letras", vbInformation, "Only Tech"
        Nome_Objeto.Text = Empty
        Nome_Objeto.SetFocus
        Exit Sub
    End If

    If Len(Nome_Objeto.Text) > 2 Then
        MsgBox "Para cadastrar uma unidade federativa são necessárias 2 letras", vbInformation, "Only Tech"
        Nome_Objeto.Text = Empty
        Nome_Objeto.SetFocus
        Exit Sub
    End If


    Nome_Objeto.Text = UCase(Nome_Objeto.Text)

    If Nome_Objeto.Text <> "SC" And Nome_Objeto.Text <> "PR" And Nome_Objeto.Text <> "RS" And _
       Nome_Objeto.Text <> "MG" And Nome_Objeto.Text <> "RJ" And Nome_Objeto.Text <> "SP" And _
       Nome_Objeto.Text <> "ES" And Nome_Objeto.Text <> "MS" And Nome_Objeto.Text <> "MT" And _
       Nome_Objeto.Text <> "GO" And Nome_Objeto.Text <> "DF" And Nome_Objeto.Text <> "AC" And _
       Nome_Objeto.Text <> "RR" And Nome_Objeto.Text <> "PA" And Nome_Objeto.Text <> "RO" And _
       Nome_Objeto.Text <> "TO" And Nome_Objeto.Text <> "MA" And Nome_Objeto.Text <> "SE" And _
       Nome_Objeto.Text <> "AP" And Nome_Objeto.Text <> "BA" And Nome_Objeto.Text <> "PE" And _
       Nome_Objeto.Text <> "CE" And Nome_Objeto.Text <> "AL" And Nome_Objeto.Text <> "PB" And _
       Nome_Objeto.Text <> "PI" And Nome_Objeto.Text <> "RN" And Nome_Objeto.Text <> "AM" And _
       Nome_Objeto.Text <> "EX" And Len(Nome_Objeto.Text) <> 0 Then
       MsgBox "Não é nenhum estado brasileiro", vbInformation, "Only Tech"
       Nome_Objeto.Text = Empty
    End If

    Exit Sub
Erro:
    'MEXER Call Erro.Erro("Verifica_Estado")
    Resume Next
End Sub



Function Extrair_Numeros(Valor As String) As String
'******************************************************************************
'Sistema...........................: Only Tech
'Módulo............................: Funções_Gerais
'Procedimento/Função...............: Extrair_Numeros
'Objetivo:.........................: Extrai todos os numeros de uma string sem caracteres outros
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 22/11/2001
'Data da última manutenção.........: 22/11/2001
'Manutenção executada por..........:
'Observações.......................:
'******************************************************************************
    On Error GoTo Erro
    Dim I As Integer
    For I = 1 To Len(Valor)
        If Mid(Valor, I, 1) Like "#" Then
            Extrair_Numeros = Extrair_Numeros & Mid(Valor, I, 1)
        End If
    Next I

    Exit Function
Erro:
    'MEXER Call Erro.Erro("Extrair_Numeros")
    Resume Next
End Function

Public Function Valida_PIS(ByVal strPIS_PASEP As String) As Boolean
'******************************************************************************
'Sistema...........................: Only Tech
'Módulo............................: Funções_Gerais
'Procedimento/Função...............: Valida_PIS
'Objetivo:.........................: VALIDAR O NUMERO DO PIS INFORMADO
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 14/01/2002
'Data da última manutenção.........:
'Manutenção executada por..........:
'Observações.......................:
'******************************************************************************
    Dim intCont As Integer
    Dim lonSoma As Long
    Dim strTabela As String

    Valida_PIS = False
    strTabela = "3298765432"

    If Val(strPIS_PASEP) = 0 Then
        Exit Function
    ElseIf Len(strPIS_PASEP) <> 11 Then
        Exit Function
    End If

    For intCont = 1 To 10
        lonSoma = lonSoma + Val(Mid(strPIS_PASEP, intCont, 1)) * Val(Mid(strTabela, intCont, 1))
    Next
    If IIf(Int(lonSoma Mod 11) <> 0, 11 - Int(lonSoma Mod 11), Int(lonSoma Mod 11)) <> Val(Mid(strPIS_PASEP, 11, 1)) Then
        Exit Function
    End If

    Valida_PIS = True
End Function

'Function Gravar_log(strData As String, strTipo As String, strUsuario As String, strDescricao As String, strEstacao As String, strPrograma As String, strEvento As String)
'
'  Dim conexao_log As New DLLConexao_Sistema.conexao
'
'  On Error GoTo Erro
'
'    'T R A N S A Ç Ã O  1
'    If strData = Empty Then
'       MsgBox "Data não informada", vbCritical, "Only Tech"
'       Exit Function
'    End If
'    If strTipo = Empty Then
'       MsgBox "Tipo não informado", vbCritical, "Only Tech"
'       Exit Function
'    End If
'    If strUsuario = "" Then
'       MsgBox "Usuário não informado", vbCritical, "Only Tech"
'       Exit Function
'    End If
'    If strDescricao = "" Then
'       MsgBox "Descrição não informada", vbCritical, "Only Tech"
'       Exit Function
'    End If
'    If strEstacao = "" Then
'        MsgBox "Estação não informada", vbCritical, "Only Tech"
'        Exit Function
'    End If
'    If strPrograma = "" Then
'        MsgBox "Programa não informado", vbCritical, "Only Tech"
'        Exit Function
'    End If
'    If strEvento = "" Then
'        MsgBox "Evento não Informado", vbCritical, "Only Tech"
'    End If
'
'    On Error GoTo Erro
'
'    conexao_log.Initial_Catalog = "BDLog"
'    conexao_log.Abrir_conexao ("Otica")
'
'    'Indica o inicio da transação junto o banco
'    conexao_log.CNConexao.BeginTrans
'
'    strSql = ""
'    strSql = "INSERT INTO TBLog(IXData_TBLog,DFTipo_TBLog,IXUsuario_TBLog,DFDescricao_TBLog,DFEstacao_TBLog,DFPrograma_TBLog,DFEvento_TBlog) " & _
'             "SELECT '" & Format(strData, "yyyymmdd") & "'," & strTipo & ",'" & strUsuario & "','" & strDescricao & "','" & strEstacao & "','" & strPrograma & "','" & strEvento & "'"
'
'    'Gravando na tabela Log
'    conexao_log.CNConexao.Execute strSql
'
''    'F I M   T R A N S A Ç Ã O  1
''    '--------------------------------------------------------------------------------------------------------------'
'
''    'T R A N S A Ç Ã O  2
''
''    If strTipo = 3 Then
''        Dim ID_log As String
''        Dim Erro_log As Object
''
''        Set Erro_log = CreateObject("DLLGestor_mil.Erro_log")
''        'Verifica se o erro está cadastrado na tabela de erro
''        If Erro_log.Verifica_erro(Err.Number) = False Then
''            'Gravando na tabela de erro caso já não exista
''            Erro_log.numero_erro = Err.Number
''            Erro_log.Descricao = Err.Description
''            Erro_log.Gravar
''        End If
''
''        'Gravando na tabela de log_erro
''        ID_log = conexao_log.CNConexao.Execute("SELECT MAX(PKId_TBLog)as DFultimo  FROM TBLog").Fields("DFultimo")
''        strSQL = "INSERT INTO TBLog_erro(PFKId_TBLog,FKnumero_TBErro) SELECT " & ID_log & " , " & Err.Number & ""
''        conexao_log.CNConexao.Execute (strSQL)
''
''    End If
''
'    'Indica o sucesso da transação do banco
'     conexao_log.CNConexao.CommitTrans
'
'    'Fecha a conexão com o banco
'     conexao_log.Fechar_conexao
''
''    DoEvents
''
''    'F I M   T R A N S A Ç Ã O  2
''    '--------------------------------------------------------------------------------------------------------------'
'
'    Exit Function
'
'Erro:
'
'    'Indica o fracasso da transação do banco
'    conexao_log.CNConexao.RollbackTrans
'
'    'Fecha a conexão com o banco
'    conexao_log.Fechar_conexao
'
'    Call Erro.Erro("Funções Gerais")
'
'End Function

Function Abrir_relatorio_registro(Aplicacao As String, Form As Object, Optional Chave As String) As String

      On Error GoTo Erro

      Dim Registro As New DLLSystemManager.Registro
      Dim Caminho As String
      
      If Chave <> "" Then
         Caminho = Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\" & Chave & "", "Caminho")
      Else
         Caminho = Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\RPT", "Caminho")
      End If
      
      Abrir_relatorio_registro = Caminho
      
      Exit Function

Erro:
   Call Erro.Erro(Form, Aplicacao, "Funções Gerais")
   Exit Function

End Function
Function Abrir_figura_registro(Aplicacao As String, Form As Object) As String

      On Error GoTo Erro

      Dim Registro As New DLLSystemManager.Registro
      Dim Caminho As String

      Caminho = Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\IMG", "Caminho")

      Abrir_figura_registro = Caminho

      Exit Function

Erro:
   Call Erro.Erro(Form, Aplicacao, "Funções Gerais")
   Exit Function

End Function
Function Abrir_nome_cliente_registro(Aplicacao As String, Form As Object) As String

      On Error GoTo Erro

      Dim Registro As New DLLSystemManager.Registro
      Dim Caminho As String

      Caminho = Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\CLIENTE", "NOME")

      Abrir_nome_cliente_registro = Caminho

      Exit Function

Erro:
   Call Erro.Erro(Form, Aplicacao, "Funções Gerais")
   Exit Function

End Function

Function Acha_Resolucao(X As String, Y As String)

    Dim xTwips%, yTwips%, xPixels#, YPixels#
    
    xTwips = Screen.TwipsPerPixelX
    yTwips = Screen.TwipsPerPixelY
    YPixels = Screen.Height / yTwips
    xPixels = Screen.Width / xTwips
    
    X = Str$(xPixels)
    Y = Str$(YPixels)
    
End Function

Function Grava_String(Texto As String) As String
       
    Grava_String = Replace(Texto, "'", "''")
            
End Function

Public Function RetornaHoraDif(Hora_Inicio As String, Hora_Final As String) As String
      
    Dim strVetorHora_Inicial() As String
    Dim strVetorHora_Final() As String
    
    Dim strHora_Total As String
    Dim strMinuto_Total As String
    Dim strSegundo_Total As String
    
    strVetorHora_Inicial = Split(Hora_Inicio, ":")
    strVetorHora_Final = Split(Hora_Final, ":")
    
    If strVetorHora_Final(0) > strVetorHora_Inicial(0) Then
       strHora_Total = strVetorHora_Final(0) - strVetorHora_Inicial(0)
    ElseIf strVetorHora_Inicial(0) > strVetorHora_Final(0) Then
       strHora_Total = strVetorHora_Inicial(0) - strVetorHora_Final(0)
    Else
       strHora_Total = strVetorHora_Inicial(0) - strVetorHora_Final(0)
    End If
    
    If strVetorHora_Final(1) > strVetorHora_Inicial(1) Then
       strMinuto_Total = strVetorHora_Final(1) - strVetorHora_Inicial(1)
    ElseIf strVetorHora_Inicial(1) > strVetorHora_Final(1) Then
       strMinuto_Total = strVetorHora_Inicial(1) - strVetorHora_Final(1)
    Else
       strMinuto_Total = strVetorHora_Inicial(1) - strVetorHora_Final(1)
    End If
    
    If strVetorHora_Final(2) > strVetorHora_Inicial(2) Then
       strSegundo_Total = strVetorHora_Final(2) - strVetorHora_Inicial(2)
    ElseIf strVetorHora_Inicial(2) > strVetorHora_Final(2) Then
       strSegundo_Total = strVetorHora_Inicial(2) - strVetorHora_Final(2)
    Else
       strSegundo_Total = strVetorHora_Inicial(2) - strVetorHora_Final(2)
    End If
      
    RetornaHoraDif = strHora_Total & ":" & strMinuto_Total & ":" & strSegundo_Total
            
End Function

Public Function Valida_Trava_Sistema(Form As Form) As Boolean

    Dim intDia As String
    Dim intMes As String
    Dim intAno As String
    Dim strData As String
    Dim rstValida_trava As New ADODB.Recordset
    Dim strSql As String
    
    strSql = Empty
    strSql = "SELECT * FROM TBPedidos_validos"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstValida_trava, "Otica", Form
    
    If rstValida_trava.BOF = True And rstValida_trava.EOF = True Then
       MsgBox "Sistema impossibilitado de iniciar!Autentique a framework e tente novamente.", vbInformation, "Only Tech"
       End
    End If
    
    rstValida_trava.MoveFirst
    
    If rstValida_trava!DFBloqueado_TBPedidos_validos = 1 Then
       Valida_Trava_Sistema = True
       Exit Function
    End If
    
    '**************************************************************************************************
    'Montando o dia
    '**************************************************************************************************
    'Primeiro digito
    If rstValida_trava!DFLetra1_TBPedidos_validos = "A" Then
       intDia = 1
    ElseIf rstValida_trava!DFLetra1_TBPedidos_validos = "B" Then
       intDia = 2
    ElseIf rstValida_trava!DFLetra1_TBPedidos_validos = "C" Then
       intDia = 3
    ElseIf rstValida_trava!DFLetra1_TBPedidos_validos = "D" Then
       intDia = 4
    ElseIf rstValida_trava!DFLetra1_TBPedidos_validos = "E" Then
       intDia = 5
    ElseIf rstValida_trava!DFLetra1_TBPedidos_validos = "F" Then
       intDia = 6
    ElseIf rstValida_trava!DFLetra1_TBPedidos_validos = "G" Then
       intDia = 7
    ElseIf rstValida_trava!DFLetra1_TBPedidos_validos = "H" Then
       intDia = 8
    ElseIf rstValida_trava!DFLetra1_TBPedidos_validos = "I" Then
       intDia = 9
    ElseIf rstValida_trava!DFLetra1_TBPedidos_validos = "*" Then
       intDia = 0
    End If
   'Segundo digito
    If rstValida_trava!DFLetra2_TBPedidos_validos = "A" Then
       intDia = intDia & 1
    ElseIf rstValida_trava!DFLetra2_TBPedidos_validos = "B" Then
       intDia = intDia & 2
    ElseIf rstValida_trava!DFLetra2_TBPedidos_validos = "C" Then
       intDia = intDia & 3
    ElseIf rstValida_trava!DFLetra2_TBPedidos_validos = "D" Then
       intDia = intDia & 4
    ElseIf rstValida_trava!DFLetra2_TBPedidos_validos = "E" Then
       intDia = intDia & 5
    ElseIf rstValida_trava!DFLetra2_TBPedidos_validos = "F" Then
       intDia = intDia & 6
    ElseIf rstValida_trava!DFLetra2_TBPedidos_validos = "G" Then
       intDia = intDia & 7
    ElseIf rstValida_trava!DFLetra2_TBPedidos_validos = "H" Then
       intDia = intDia & 8
    ElseIf rstValida_trava!DFLetra2_TBPedidos_validos = "I" Then
       intDia = intDia & 9
    ElseIf rstValida_trava!DFLetra2_TBPedidos_validos = "*" Then
       intDia = intDia & 0
    End If

    '**************************************************************************************************
    'Montando o mes
    '**************************************************************************************************
    'Primeiro digito
    If rstValida_trava!DFLetra3_TBPedidos_validos = "A" Then
       intMes = 1
    ElseIf rstValida_trava!DFLetra3_TBPedidos_validos = "B" Then
       intMes = 2
    ElseIf rstValida_trava!DFLetra3_TBPedidos_validos = "C" Then
       intMes = 3
    ElseIf rstValida_trava!DFLetra3_TBPedidos_validos = "D" Then
       intMes = 4
    ElseIf rstValida_trava!DFLetra3_TBPedidos_validos = "E" Then
       intMes = 5
    ElseIf rstValida_trava!DFLetra3_TBPedidos_validos = "F" Then
       intMes = 6
    ElseIf rstValida_trava!DFLetra3_TBPedidos_validos = "G" Then
       intMes = 7
    ElseIf rstValida_trava!DFLetra3_TBPedidos_validos = "H" Then
       intMes = 8
    ElseIf rstValida_trava!DFLetra3_TBPedidos_validos = "I" Then
       intMes = 9
    ElseIf rstValida_trava!DFLetra3_TBPedidos_validos = "*" Then
       intMes = 0
    End If
    'Segundo digito
    If rstValida_trava!DFLetra4_TBPedidos_validos = "A" Then
       intMes = intMes & 1
    ElseIf rstValida_trava!DFLetra4_TBPedidos_validos = "B" Then
       intMes = intMes & 2
    ElseIf rstValida_trava!DFLetra4_TBPedidos_validos = "C" Then
       intMes = intMes & 3
    ElseIf rstValida_trava!DFLetra4_TBPedidos_validos = "D" Then
       intMes = intMes & 4
    ElseIf rstValida_trava!DFLetra4_TBPedidos_validos = "E" Then
       intMes = intMes & 5
    ElseIf rstValida_trava!DFLetra4_TBPedidos_validos = "F" Then
       intMes = intMes & 6
    ElseIf rstValida_trava!DFLetra4_TBPedidos_validos = "G" Then
       intMes = intMes & 7
    ElseIf rstValida_trava!DFLetra4_TBPedidos_validos = "H" Then
       intMes = intMes & 8
    ElseIf rstValida_trava!DFLetra4_TBPedidos_validos = "I" Then
       intMes = intMes & 9
    ElseIf rstValida_trava!DFLetra4_TBPedidos_validos = "*" Then
       intMes = intMes & 0
    End If
    
    '**************************************************************************************************
    'Montando o ano
    '**************************************************************************************************
    'Primeiro digito
    If rstValida_trava!DFLetra5_TBPedidos_validos = "A" Then
       intAno = 1
    ElseIf rstValida_trava!DFLetra5_TBPedidos_validos = "B" Then
       intAno = 2
    ElseIf rstValida_trava!DFLetra5_TBPedidos_validos = "C" Then
       intAno = 3
    ElseIf rstValida_trava!DFLetra5_TBPedidos_validos = "D" Then
       intAno = 4
    ElseIf rstValida_trava!DFLetra5_TBPedidos_validos = "E" Then
       intAno = 5
    ElseIf rstValida_trava!DFLetra5_TBPedidos_validos = "F" Then
       intAno = 6
    ElseIf rstValida_trava!DFLetra5_TBPedidos_validos = "G" Then
       intAno = 7
    ElseIf rstValida_trava!DFLetra5_TBPedidos_validos = "H" Then
       intAno = 8
    ElseIf rstValida_trava!DFLetra5_TBPedidos_validos = "I" Then
       intAno = 9
    ElseIf rstValida_trava!DFLetra5_TBPedidos_validos = "*" Then
       intAno = 0
    End If
    'Segundo digito
    If rstValida_trava!DFLetra6_TBPedidos_validos = "A" Then
       intAno = intAno & 1
    ElseIf rstValida_trava!DFLetra6_TBPedidos_validos = "B" Then
       intAno = intAno & 2
    ElseIf rstValida_trava!DFLetra6_TBPedidos_validos = "C" Then
       intAno = intAno & 3
    ElseIf rstValida_trava!DFLetra6_TBPedidos_validos = "D" Then
       intAno = intAno & 4
    ElseIf rstValida_trava!DFLetra6_TBPedidos_validos = "E" Then
       intAno = intAno & 5
    ElseIf rstValida_trava!DFLetra6_TBPedidos_validos = "F" Then
       intAno = intAno & 6
    ElseIf rstValida_trava!DFLetra6_TBPedidos_validos = "G" Then
       intAno = intAno & 7
    ElseIf rstValida_trava!DFLetra6_TBPedidos_validos = "H" Then
       intAno = intAno & 8
    ElseIf rstValida_trava!DFLetra6_TBPedidos_validos = "I" Then
       intAno = intAno & 9
    ElseIf rstValida_trava!DFLetra6_TBPedidos_validos = "*" Then
       intAno = intAno & 0
    End If
    
    'Concatenando e validando a data corrente
    strData = intDia & "/" & intMes & "/" & intAno
    
    If CDate(strData) < Format(Now, "DD/MM/YY") Then
       Valida_Trava_Sistema = True
       'Marcando o sistema como travado
       funcoes_banco.Alterar "TBPedidos_validos", "SET DFBloqueado_TBPedidos_validos = 1", "PKId_TBPedidos_validos", rstValida_trava!PKId_TBPedidos_validos, "Otica", Form
    Else
       Valida_Trava_Sistema = False
    End If
    
End Function
