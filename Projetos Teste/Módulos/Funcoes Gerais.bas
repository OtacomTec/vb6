Attribute VB_Name = "Funcoes_Gerais"
'*******************************************************************************************
'Módulo............................: Nenhum
'Conexão...........................: Nenhuma
'Formulário........................: Funcoes_Gerais
'Objetivo do formulário............: Funcoes Gerais
'Programação.......................: Marcos Baião
'Data..............................: 14/03/2000
'*******************************************************************************************

Public Const piStrTitulo_Menu = "Integrador"

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    'API do Windows
    Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    'Declare Function SetWindowPos Lib "user32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Dimensao) As Long
    Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
        
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
    hWnd As Long
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
    hWnd As Long
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
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long

Public Function Arredonde(ByVal Valor As Single, Optional ByVal Precisao As Integer = 0) As Single
    Dim sinTemp As Single
    
    Valor = Valor * (10 ^ Precisao) ' valor vezes dez elevado à precisão
    sinTemp = Int(Valor)
    
    If (Valor - sinTemp) * 10 Then
        Valor = sinTemp + 1
    End If
    
    Arredonde = Valor / (10 ^ Precisao)
End Function

Public Sub DoEventsSimples()
    Dim tMsg As MSG
    Dim tPoint As POINTAPI
    
    tMsg.pt = tPoint
    
    Do While PeekMessage(tMsg, 0, 0, 0, PM_REMOVE)
        Call TranslateMessage(tMsg)
        Call DispatchMessage(tMsg)
    Loop
    
    Sleep 0
End Sub


Public Sub Corrigir_Deposito_Item(Nome_Tabela As String, Optional Codigo_Deposito As Integer)
'******************************************************************************
'Módulo............................: Estoque
'Procedimento/Função...............: Corrigir_Deposito_Item
'Objetivo:.........................: Varrer a tabela TBDeposito_Item e converter
'                                    todos as unidades de itens que sejam diferentes
'                                    da unidade de compra para a própria unidade de
'                                    compra
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 01/02/2001
'Data da última manutenção.........:
'Manutenção executada por..........:
'Observações.......................: "Nome_Tabela" é o nome de uma tabela temporária
'                                    onde serão armazenados os registros provenientes
'                                    da operação de correção.
'                                    "Codigo"
'******************************************************************************
    On Error GoTo erro
    Dim Meu_Erro As String
    Dim strSQL As String
    Dim adrTemp_Deposito_Item As ADODB.Recordset
       
    '1.------ Constroi espelho de Deposito_Item com as duas
    '         unidades
    
    'Verifica existencia TBtemp, se existir apaga
    Conexao.Execute "IF EXISTS(SELECT NAME FROM sysObjects WHERE NAME = '" & Nome_Tabela & "') DROP TABLE " & Nome_Tabela & " "
    
    'Gerar TBtemp
    strSQL = _
        "SELECT TBDeposito_Item.DFId_Deposito_Item, " & _
               "TBdeposito.DFcod_deposito, " & _
               "TBitem_estoque.DFcod_item_estoque, " & _
               "TBitem_estoque.DFdescricao, " & _
               "TBunidade.DFcod_unidade AS DFcod_unidade, " & _
               "TBunidade_padrao.DFcod_unidade AS DFcod_unidade_padrao, " & _
               "TBdeposito_item.DFquantidade_estoque, " & _
               "TBdeposito_item.DFquantidade_estoque_reservado, " & _
               "(TBdeposito_item.DFquantidade_estoque - TBdeposito_item.DFquantidade_estoque_reservado) AS DFqtde_disponivel, " & _
               "TBunidade.DFfator_conversao, " & _
               "TBunidade_padrao.DFfator_conversao AS DFfator_conversao_padrao " & _
        "INTO " & Nome_Tabela & " " & _
        "FROM TBdeposito_item " & _
             "INNER JOIN TBdeposito " & _
                   "ON TBdeposito_item.DFcod_deposito = TBdeposito.DFcod_deposito " & _
             "INNER JOIN TBitem_estoque " & _
                   "ON TBdeposito_item.DFcod_item_estoque = TBitem_estoque.DFcod_item_estoque " & _
             "INNER JOIN TBunidade " & _
                   "ON TBdeposito_item.DFcod_unidade_armazenagem = TBunidade.DFcod_unidade " & _
             "INNER JOIN TBunidade AS TBunidade_padrao " & _
                   "ON TBitem_estoque.DFcod_unidade_compra = TBunidade_padrao.DFcod_unidade "
    If Codigo_Deposito <> 0 Then
        strSQL = strSQL & "WHERE TBdeposito_item.DFcod_deposito = " & Codigo_Deposito & " "
    End If
    strSQL = strSQL & "ORDER BY TBitem_estoque.DFcod_item_estoque "
    
    Conexao.Execute (strSQL)
    
    
    '2.------------ Compara as unidade e converte as quantidades
    '               para a unidade padrão
    
    Call Banco_Dados.SQLgeral("SELECT * FROM " & Nome_Tabela, adrTemp_Deposito_Item)
    If adrTemp_Deposito_Item.RecordCount <> 0 Then
        adrTemp_Deposito_Item.MoveFirst
        Do While Not adrTemp_Deposito_Item.EOF
            If adrTemp_Deposito_Item("DFcod_unidade") <> adrTemp_Deposito_Item("DFcod_unidade_padrao") Then
                adrTemp_Deposito_Item("DFquantidade_estoque") = (adrTemp_Deposito_Item("DFfator_conversao") / adrTemp_Deposito_Item("DFfator_conversao_padrao")) * adrTemp_Deposito_Item("DFquantidade_estoque")
                adrTemp_Deposito_Item("DFquantidade_estoque_reservado") = (adrTemp_Deposito_Item("DFfator_conversao") / adrTemp_Deposito_Item("DFfator_conversao_padrao")) * adrTemp_Deposito_Item("DFquantidade_estoque_reservado")
                adrTemp_Deposito_Item("DFqtde_disponivel") = (adrTemp_Deposito_Item("DFfator_conversao") / adrTemp_Deposito_Item("DFfator_conversao_padrao")) * adrTemp_Deposito_Item("DFqtde_disponivel")
                adrTemp_Deposito_Item("DFcod_unidade") = adrTemp_Deposito_Item("DFcod_unidade_padrao")
                adrTemp_Deposito_Item.Update
            End If
            adrTemp_Deposito_Item.MoveNext
        Loop
    End If
    Exit Sub
erro:
    Meu_Erro = Error
    Call erro.erro("Corrigir_Deposito_Item")
End Sub



Public Function FecharAplicativo(ByVal CaptionJanela As String) As Boolean
    '------------------------------------------------------------------
    'Efetua uma chamada à função FindWindow seguida da PostMessage
    'São duas funções de API do Windows declaradas em General
    'Declarations, bem como suas constantes.
    'FindWindow retorna um número que identifica a tela no ambiente
    'Windows
    'PostMessage manda uma função/ação para a tal janela
    '------------------------------------------------------------------
    On Error GoTo erro
    
    Const WinClose = &H10
    Dim lonWinWnd As Long
    lonWinWnd = FindWindow(vbNullString, CaptionJanela)
    If lonWinWnd <> 0 Then
        PostMessage lonWinWnd, WinClose, 0&, 0&
    End If
    Exit Function
    
erro:
    Call erro.erro("FecharAplicativo")
    Resume Next
End Function

Public Function Empresa_Padrao_Inicio(DataCombo_Original As DataCombo, Optional strCod_Empresa As Variant) As String
    Dim strSQL As String
    
    On Error GoTo erro

    strSQL = strSQL & "SELECT DFnome_fantasia, DFCod_Empresa FROM TBempresa ORDER BY DFnome_fantasia"
    
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open strSQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set DataCombo_Original.RowSource = TBrecordset
        DataCombo_Original.ListField = "DFnome_fantasia"
        DataCombo_Original.BoundColumn = "DFCod_Empresa"
    
    Set TBrecordset = Nothing
    
    DataCombo_Original.BoundText = intEmpresa
    
    If Not IsMissing(strCod_Empresa) Then
        strCod_Empresa = intEmpresa
    End If
    
    Exit Function

erro:
    Call erro.erro("Empresa_Padrao_Inicio")
'-----------------------------------------------------------------------------------
'Esta função deverá ser usada no Load do Form, indicando um Data Combo e a variável
'String previamente cadastrada no General Declarations, a qual conterá o código da
'Empresa. Serve para trazer a empresa padrão do usuário.
'-----------------------------------------------------------------------------------
End Function

Public Function Barra(Nome_Menu As Menu) As String
    If Nome_Menu.Checked = False Then
        Nome_Menu.Checked = True
    Else
        Nome_Menu.Checked = False
    End If
End Function

Public Function Formata_Hora(Horas As Integer, Minutos As Integer) As String
'******************************************************************************
'Sistema...........................: Integrador
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
'Sistema...........................: Integrador
'Módulo............................: Funções Gerais
'Objetivo do módulo................: Dados duas strings, ditas horas,
'                                    no formato hh:mm, realiza a soma dessas horas
'                                    retornando um horário no formato (x)x:xx
'Desenvolvimento...................: marcos Baião
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


Public Sub Codigo_Autonumeracao(Nome_Tabela As String, TextBox_Codigo As TextBox, Nome_Campo_Auto As String)
   
    On Error GoTo erro
    
    intcodigo = Conexao.Execute("SELECT (ISNULL(MAX(" & Nome_Campo_Auto & "), 0) + 1) AS DFproximo_codigo FROM " & Nome_Tabela).Fields("DFproximo_codigo")
    TextBox_Codigo = intcodigo
    
    Exit Sub
    
erro:
    Call erro.erro("Codigo_Autonumeracao")
    Resume Next
    
End Sub

Public Sub Codigo_Autonumeracao_Interno(Nome_Tabela As String, Codigo_Auto_Num As Variant, Nome_Campo_Auto As String)
        
    On Error GoTo erro
    
    intcodigo = Conexao.Execute("SELECT (ISNULL(MAX(" & Nome_Campo_Auto & "), 0) + 1) AS DFproximo_codigo FROM " & Nome_Tabela).Fields("DFproximo_codigo")
    Codigo_Auto_Num = intcodigo
    Exit Sub
erro:
    Call erro.erro("Codigo_Autonumeracao")
    Resume Next
    
End Sub

Public Function Abrir_App(Form_Caption As String, Form_Path As String, Optional Formulario As Form) As Long
    
    '------------------------------------------------------------------
    'Executa o programa definido na variável FormPath por APIs do
    'Windows, declaradas do General Declarations, e a posiciona
    'de forma a não ultrapassar a interface do Integrador.
    'Caso esse programa já esteja carregado, apenas ativa-o
    'FindWindow retorna um número que identifica a janela no ambiente
    'Windows
    'ShowWindowPos reposiciona a janela
    '------------------------------------------------------------------
    On Error GoTo erro
    DoEvents
    
    Const No_Topo = -2
    Const Exibir_Janela = &H40
    Const Nao_Redimensionar = &H1
    Dim hWnd As Long
    
    hWnd = FindWindow(vbNullString, Form_Caption)
    If hWnd = 0 Then
        Shell Form_Path, vbNormalFocus
        'Acrescenta 50 twips à coordenada y do form para que não apareca em cima
        'do Integrador.exe
        hWnd = FindWindow(vbNullString, Form_Caption)
        
        
        If Not Formulario Is Nothing Then
            SetWindowPos hWnd, No_Topo, Formulario.CurrentX, Formulario.CurrentY + 50, Formulario.Height, Formulario.Width, Exibir_Janela + Nao_Redimensionar
        Else
            SetWindowPos hWnd, No_Topo, frmPrincipal.CurrentX, frmPrincipal.CurrentY + 50, frmPrincipal.Height, frmPrincipal.Width, Exibir_Janela + Nao_Redimensionar
        End If
        
       '-------------------------------------------------------
        
       'rotina responsável por registrar os formulários abertos
       'em um arquivo txt para o controle do frmPricipal
        
        Dim regForm_Aberto As Registro
        Dim regContagem As Registro
        
        regForm_Aberto.Caption = Form_Caption
        regForm_Aberto.hWnd = hWnd
        
        Open "c:\Integrador\formularios.drg" For Random As #1
            
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
            SetWindowPos hWnd, No_Topo, Formulario.CurrentX, Formulario.CurrentY + 50, Formulario.Height, Formulario.Width, Exibir_Janela + Nao_Redimensionar
        Else
            SetWindowPos hWnd, No_Topo, frmPrincipal.CurrentX, frmPrincipal.CurrentY + 50, frmPrincipal.Height, frmPrincipal.Width, Exibir_Janela + Nao_Redimensionar
        End If
    End If
    
    Abrir_App = hWnd
    
    Exit Function
erro:
    Call erro.erro("Abrir_App")
End Function


Public Function MinimizarRestaurar(ByVal booMinimizar As Boolean, ByVal CaptionJanela As String) As Boolean
    '------------------------------------------------------------------
    'Efetua uma chamada à função FindWindow seguida da ShowWindow
    'São duas funções de API do Windows declaradas em General
    'Declarations, bem como suas constantes.
    'FindWindow retorna um número que identifica a janela no ambiente
    'Windows
    'ShowWindow apenas esconde essa janela se o Integrador estiver
    'minimizado
    '------------------------------------------------------------------
    
    Dim lonWinWnd As Long
    Dim intWinShow As Integer
    On Error GoTo erro
    lonWinWnd = FindWindow(vbNullString, CaptionJanela)
    intWinShow = IIf((booMinimizar = True), vbHide, vbNormalNoFocus)
    If lonWinWnd <> 0 Then
        DoEvents
        ShowWindow lonWinWnd, intWinShow
    End If
    Exit Function
    
erro:
    Call erro.erro("MinimizarRestaurar")
    Resume Next
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
    
    If Valor = Empty Then
        Grava_Moeda = 0
        Exit Function
    End If
    
    strRetorno = Format(Valor, "#0.0000;-#0.0000")
    strDecimal = Mid(strRetorno, (InStr(1, strRetorno, ",") + 1))
    strInteiro = Mid(strRetorno, 1, (InStr(1, strRetorno, ",") - 1))
    
    Grava_Moeda = strInteiro & "." & strDecimal
End Function

Public Function Verifica_Numero(Nome_Campo As String, Nome_Tabela As String, Nome_textbox As TextBox) As Boolean
    If Nome_textbox.Text = Empty Then
        Exit Function
    End If
    
    Dim SQL As String
    
    On Error GoTo erro
    SQL = ""
    SQL = SQL & "SELECT " & Nome_Campo & " "
    SQL = SQL & "FROM " & Nome_Tabela & " "
    SQL = SQL & "WHERE " & Nome_Campo & " "
    SQL = SQL & "= " & Nome_textbox.Text & ""
    
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open SQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
        
    If Val(TBrecordset(Nome_Campo)) = Val(Nome_textbox.Text) Then
        MsgBox "Registro já existente", vbCritical, "Integrador"
        Nome_textbox.Text = Empty
        Nome_textbox.SetFocus
        Verifica_Numero = True
    Else
        Verifica_Numero = False
    End If
    Exit Function
    
erro:
    If Err.Number = 3021 Then
        Verifica_Numero = False
        Exit Function
    Else
        Call erro.erro("Verifica_Numero")
    End If
    Resume Next

End Function

Public Function Verifica_Texto(Nome_Campo As String, Nome_Tabela As String, Nome_textbox As TextBox) As Boolean
    If Nome_textbox = Empty Then
        Exit Function
    End If
    
    Dim SQL As String
    
    On Error GoTo erro
    SQL = ""
    SQL = SQL & "SELECT " & Nome_Campo & " "
    SQL = SQL & "FROM " & Nome_Tabela & " "
    SQL = SQL & "WHERE " & Nome_Campo & " "
    SQL = SQL & "= '" & Nome_textbox & "'"
    
    Set TBrecordset = New ADODB.Recordset
        TBrecordset.CursorLocation = adUseClient
        TBrecordset.Open SQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
        
    If TBrecordset(Nome_Campo) = Nome_textbox Then
        MsgBox "Registro já existente", vbCritical, "Integrador"
        Nome_textbox = Empty
        Nome_textbox.SetFocus
        Verifica_Texto = True
    Else
        Verifica_Texto = False
    End If
    Exit Function
erro:
    If Err.Number = 3021 Then
        Verifica_Texto = False
        Exit Function
    Else
        Call erro.erro("Verifica_Texto")
    End If
    Resume Next
End Function


Public Function Ler_Imagem(strTabela As String, strCampo_Imagem As String, strCampo_Codigo As String, strValor_Codigo As String) As IPictureDisp
    On Error GoTo erro
    Dim adrTemp As ADODB.Recordset
    Dim strSQL As String
    
    Dim bytDados() As Byte
    Dim varTemp As Variant
    Dim f As Long
    
    
    '----------------------------------------------------------------------------
    'abrindo Recordset
    strSQL = _
        "SELECT " & strCampo_Imagem & " FROM " & strTabela & " " & _
        "WHERE " & strCampo_Codigo & " = " & strValor_Codigo
    
    Set adrTemp = New ADODB.Recordset
        adrTemp.CursorLocation = adUseClient
        adrTemp.Open strSQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
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
erro:
    Call erro.erro("Ler_Imagem")

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
    On Error GoTo erro
    Dim adrTemp As ADODB.Recordset
    Dim strSQL As String
    
    Dim bytDados() As Byte
    Dim varTemp As Variant
    Dim lonTamanho_Arquivo As Long
    Dim lonBytes_Lidos As Long

    Const Limite_Bloco As Long = 1048576 'o máximo que cada bloco pode receber
    
    
    'abrindo Recordset
    strSQL = _
        "SELECT " & strCampo_Imagem & " FROM " & strTabela & " " & _
        "WHERE " & strCampo_Codigo & " = " & strValor_Codigo
    
    Set adrTemp = New ADODB.Recordset
        adrTemp.CursorLocation = adUseClient
        adrTemp.Open strSQL, Conexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    
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
erro:
    Call erro.erro("Salvar_Imagem")
    
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
        If MsgBox(Mensagem, vbInformation + vbYesNo, "Integrador") = vbNo Then
            Exit Sub
        End If
    End If
    
    Nome_do_form.ProgressBar1.Visible = True
    Nome_do_form.ProgressBar1.Max = 100
    Nome_do_form.Height = Altura_Atual + 270
    
    DoEvents
    Screen.MousePointer = 11
    For I = 0 To 25
        Nome_do_form.ProgressBar1.Value = I
    Next
    
    'Verifica se a tabela existe e apaga a tabela
    Conexao.Execute "If Exists(SELECT * FROM sysObjects WHERE id = Object_id('dbo.TBTemp_Impressao'))Begin DROP TABLE TBTemp_Impressao End "
    
    'strSQLEmpresa = "SELECT DFnome_fantasia FROM TBempresa WHERE DFcod_empresa = " & intEmpresa & ""
    'Call Banco_Dados.SQLgeral(strSQLEmpresa, adrEmpresa)

    'Gera uma adr utilizando a variável strSQLImpressao
    'Grava a tabela temporária
    Conexao.Execute strSQL_Impressao
    
    Call Conexao_Banco.SQLgeral("SELECT * FROM TBTemp_Impressao ", adrTBTemp)
    
    If adrTBTemp.RecordCount = 0 Then
        bTabela_Vazia = True
    End If
    
    If bTabela_Vazia = True Then
        Screen.MousePointer = 0
        Beep
        MsgBox "Não há registros que satisfaçam a elaboração do relatório.", vbInformation, "Integrador"
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
        MsgBox "Não foi possível encontrar o arquivo de impressão.", vbCritical, "Integrador"
    Else
        Call erro.erro("Imprime_Listagem")
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
'Sistema...........................: Integrador
'Módulo............................:
'Procedimento/Função...............: Verifica_Apostrofo
'Objetivo:.........................: Substituir os Apostrofos cadastrados pelo usuário,
'                                    pelo comando CHAR(39) do SQL.
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 27/04/2001
'Observaçãoes......................:
'******************************************************************************
    On Error GoTo erro
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
erro:
    erro.erro ("Verifica_Apostrofo")
    Resume Next
End Function

Public Sub Verifica_Tecla(KeyCode As Integer, Form As Form, Shift As Integer) ', Optional Guia As SSTab)
'******************************************************************************
'Sistema...........................: Integrador
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
    On Error GoTo erro
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
            
            AppActivate "Integrador"
            If Err.Number = 0 Then
                SendKeys "%{" & Chr(KeyCode) & "}"
            End If
            
            Exit Sub
        End If
    End If
    
    
    'Verifica se a tecla F12 foi acionada
    If KeyCode = vbKeyF12 Then
        MsgBox "Número da última versão do executável " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "Integrador"
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

erro:
    Call erro.erro("Verifica_Tecla")
    Resume Next
End Sub

Public Sub Saldo_Bancario(Acao As String, ID_Conta As Integer, Valor As String, Data_Emissao As String, Historico As Integer, Optional DataGrid As DataGrid)
'******************************************************************************
'Sistema...........................: Integrador
'Módulo............................: Funcoes_Gerais
'Procedimento/Função...............: Saldo_Bancario
'Objetivo:.........................: Atualizar o Saldo mensal.
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 19/06/2001
'Observaçãoes......................:
'******************************************************************************
    On Error GoTo erro
    Dim strSQL As String
    Dim adrNatureza_Historico As ADODB.Recordset
    Dim adrSaldo_Bancario As ADODB.Recordset
    Const Incluindo_Saldo_Novamente = "Inclusao_Movimento"

    If Acao = "Inclusao_Movimento" Then

        strSQL = "SELECT DFnatureza FROM TBhistorico_padrao_movto_bancario WHERE DFcod_historico_movto_bancario = " & Historico
        Call Banco_Dados.SQLgeral(strSQL, adrNatureza_Historico)

        strSQL = "SELECT DFid_saldo_bancario_mensal FROM TBsaldo_bancario_mensal WHERE DFmes_ano_referencia = '" & Format(Data_Emissao, "MMyyyy") & "' AND DFid_conta = " & ID_Conta
        Call Banco_Dados.SQLgeral(strSQL, adrSaldo_Bancario)


        If adrSaldo_Bancario.RecordCount = 0 Then
        'ainda não existe saldo para o mês corrente, será necessário criar este registro

            If adrNatureza_Historico("DFnatureza").Value = "D" Then  'Débito
                Conexao.Execute "INSERT INTO TBsaldo_bancario_mensal VALUES(" & ID_Conta & ",'" & Format(Data_Emissao, "MMyyyy") & "',0," & Funcoes_Gerais.Grava_Moeda(Valor) & ",0) "
            Else 'Crédito
                Conexao.Execute "INSERT INTO TBsaldo_bancario_mensal VALUES(" & ID_Conta & ",'" & Format(Data_Emissao, "MMyyyy") & "',0,0," & Grava_Moeda(Valor) & ") "
            End If


        Else
        'já existe existe saldo para o mês corrente
        'será necessário excluir este valor e incluir
        'novamente com as novas configurações

            If adrNatureza_Historico("DFnatureza").Value = "D" Then 'Débito
                Conexao.Execute "UPDATE TBsaldo_bancario_mensal SET DFtotal_debito_mes = DFtotal_debito_mes + " & Grava_Moeda(Valor) & " WHERE DFid_saldo_bancario_mensal = " & adrSaldo_Bancario("DFid_saldo_bancario_mensal").Value
            Else 'Crédito
                Conexao.Execute "UPDATE TBsaldo_bancario_mensal SET DFtotal_credito_mes = DFtotal_credito_mes + " & Grava_Moeda(Valor) & " WHERE DFid_saldo_bancario_mensal = " & adrSaldo_Bancario("DFid_saldo_bancario_mensal").Value
            End If
        End If

    ElseIf Acao = "Alteracao_Movimento" Then
    'o valor que antes foi acrescentado no banco por este documento será retirado
    'e uma nova entrada a esta tabela será realizada

        strSQL = "SELECT DFnatureza FROM TBhistorico_padrao_movto_bancario WHERE DFcod_historico_movto_bancario = " & DataGrid.Columns("Código Histórico Padrão").Value
        Call Banco_Dados.SQLgeral(strSQL, adrNatureza_Historico)

        strSQL = "SELECT DFid_saldo_bancario_mensal FROM TBsaldo_bancario_mensal WHERE DFmes_ano_referencia = '" & Format(DataGrid.Columns("Emissão").Value, "MMyyyy") & "' AND DFid_conta = " & DataGrid.Columns("ID Conta").Value
        Call Banco_Dados.SQLgeral(strSQL, adrSaldo_Bancario)


        'Subtraindo o valor do movimento do saldo bancario...
        If adrNatureza_Historico("DFnatureza").Value = "D" Then 'Débito
            Conexao.Execute "UPDATE TBsaldo_bancario_mensal SET DFtotal_debito_mes = (DFtotal_debito_mes - " & Grava_Moeda(Valor) & ") WHERE DFid_saldo_bancario_mensal = " & adrSaldo_Bancario("DFid_saldo_bancario_mensal").Value
        Else 'Crédito
            Conexao.Execute "UPDATE TBsaldo_bancario_mensal SET DFtotal_credito_mes = (DFtotal_credito_mes - " & Grava_Moeda(Valor) & ") WHERE DFid_saldo_bancario_mensal = " & adrSaldo_Bancario("DFid_saldo_bancario_mensal").Value
        End If

        'Incluindo o valor do movimento a outro saldo bancario...
        Call Saldo_Bancario(Incluindo_Saldo_Novamente, ID_Conta, Valor, Data_Emissao, Historico, DataGrid)

    End If

    Exit Sub
erro:
    Call erro.erro("Saldo_Bancario")
End Sub

Public Sub Verifica_Estado(Nome_Objeto As Variant)
'******************************************************************************
'Sistema...........................: Integrador
'Módulo............................: Funcoes_Gerais
'Procedimento/Função...............: Verifica_Estado
'Objetivo:.........................: Faz a validação dos estados
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 19/06/2001
'Observaçãoes......................:
'******************************************************************************
    On Error GoTo erro
    
    If Len(Nome_Objeto.Text) < 2 And Len(Nome_Objeto.Text) <> 0 Then
        MsgBox "Para cadastrar uma unidade federativa são necessárias 2 letras", vbInformation, "Integrador"
        Nome_Objeto.Text = Empty
        Nome_Objeto.SetFocus
        Exit Sub
    End If
    
    If Len(Nome_Objeto.Text) > 2 Then
        MsgBox "Para cadastrar uma unidade federativa são necessárias 2 letras", vbInformation, "Integrador"
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
       MsgBox "Não é nenhum estado brasileiro", vbInformation, "Integrador"
       Nome_Objeto.Text = Empty
    End If
    
    Exit Sub
erro:
    Call erro.erro("Verifica_Estado")
    Resume Next
End Sub

Public Function Converte_Codigo_Id(strValor_Converter As String, strOpcao As String, strCampo_Id As String, strCampo_Cod As String, strTabela As String) As String
'******************************************************************************
'Sistema...........................: Integrador
'Módulo............................: Funções_Gerais
'Procedimento/Função...............: Converte_Codigo_Id
'Objetivo:.........................: Corverte um Id para Codigo e vice-versa.
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 10/05/2001
'Observações.......................: Esta função só poderá ser utilizada se assim como
'                                    o campo ID o campo Codigo seja exclusivo, isto é,
'                                    não se repita.
'******************************************************************************
    On Error GoTo erro
    
    Dim strSQL As String
    Dim adrValor_Convertido As ADODB.Recordset
    
    If strOpcao = "Id" Then
        strSQL = strSQL & "SELECT " & strCampo_Id & " From " & strTabela
        strSQL = strSQL & " WHERE " & strCampo_Cod & " = " & strValor_Converter
    
    ElseIf strOpcao = "Codigo" Then
        strSQL = strSQL & "SELECT " & strCampo_Cod & " From " & strTabela
        strSQL = strSQL & " WHERE " & strCampo_Id & " = " & strValor_Converter
    
    End If
    
    Call Banco_Dados.SQLgeral(strSQL, adrValor_Convertido)
    
    If adrValor_Convertido.RecordCount <> 0 Then
        adrValor_Convertido.MoveFirst
        
        If strOpcao = "Id" Then
            Converte_Codigo_Id = adrValor_Convertido(strCampo_Id)
        ElseIf strOpcao = "Codigo" Then
            Converte_Codigo_Id = adrValor_Convertido(strCampo_Cod)
        End If
        
    End If
    
    Exit Function
erro:
    Call erro.erro("Converte_Codigo_Id")
    Resume Next
End Function

Function Extrair_Numeros(Valor As String) As String
'******************************************************************************
'Sistema...........................: Integrador
'Módulo............................: Funções_Gerais
'Procedimento/Função...............: Extrair_Numeros
'Objetivo:.........................: Extrai todos os numeros de uma string sem caracteres outros
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 22/11/2001
'Data da última manutenção.........: 22/11/2001
'Manutenção executada por..........:
'Observações.......................:
'******************************************************************************
    On Error GoTo erro
    Dim I As Integer
    For I = 1 To Len(Valor)
        If Mid(Valor, I, 1) Like "#" Then
            Extrair_Numeros = Extrair_Numeros & Mid(Valor, I, 1)
        End If
    Next I
    
    Exit Function
erro:
    Call erro.erro("Extrair_Numeros")
    Resume Next
End Function

Public Function Valida_PIS(ByVal strPIS_PASEP As String) As Boolean
'******************************************************************************
'Sistema...........................: Integrador
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
Function Gravar_Log(programa_log As String, documento_log As String, acao_log As String, usuario As String, observacao_log As String) As String
'******************************************************************************]
'Sistema...........................: Integrador
'Módulo............................: Funções_Gerais
'Procedimento/Função...............: Gravar Log da aplicação
'Objetivo:.........................: Gravar inf. no banco de log
'Desenvolvimento...................: Marcos Baião
'Data de criação...................: 25/07/2002
'Data da última manutenção.........:   /  /
'Manutenção executada por..........:
'Observações.......................:
'******************************************************************************
    On Error GoTo erro
    
    Dim strSQL As String
    Dim txtcodigo As TextBox
    Dim intcodigo_auto As Integer
    
    Call Funcoes_Gerais.Codigo_Autonumeracao_Interno("TBlog", intcodigo_auto, "PKId_TBLog")
    
    strSQL = " "
    strSQL = "INSERT INTO TBlog(PKId_TBLog,IXData_TBLog,DFPrograma_TBLog,DFEvento_TBlog,IXUsuario_TBLog,DFDescricao_TBLog) " & _
             "SELECT " & intcodigo_auto & ",'" & Format(Date, "yyyymmdd") & "','" & programa_log & "','" & acao_log & "','" & usuario & "','" & observacao_log & "' "
             
    Conexao.Execute (strSQL)
    
    Exit Function
erro:
    Call erro.erro("Gravar_Log")
    Resume Next
    
End Function
