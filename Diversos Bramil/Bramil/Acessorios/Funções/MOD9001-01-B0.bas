Attribute VB_Name = "Module0"
Option Explicit

' Variáveis que são retornadas pela Rotina Inicial

Public pstrNomeUsuario As String
Public piCodEmpresatPar As Integer
Public pbIndMixEmpresatPar As Byte
Public pstrEstadoEmpresatPar As String
Public pstrSenhaBancoDadosBDCONFUS As String
Public pstrSenhaBancoDadostLoc As String
Public pstrCoordenadaNavegtLogin As String
Public pstrLocacaoIcoLogotipotLogin As String
Public pstrLocacaoAcessoriostLogin As String
Public piCodUsuariotLogin As Integer
Public plSeqLogtLogin As Long
Public pbCodEsquemaBancoDadostLogin As Byte
Public pdtDataRealtLogin  As Date
Public pbTabNivelUsuariotUsu As Byte
Public pbCodGrupoUsuariotUsu As Byte
Public piCodEstacaotLog As Integer
Public pstrCodPrograma As String
Public pstrCodFuncaotLogin As String
Public pstrDescrProgramatLogin As String
Public pbTabCtrLogProgramatLogin As Byte
Public pstrLocacaobdLog As String
Public pstrLocacaobdConfus As String
Public pstrNomeIdentFormtLogin As String
Public pstrNomeIdentReporttLogin As String
Public pstrFuncaoToolbar As String
Public pboAcessoSenhaArea As Boolean
Public pboAcessoSistema As Boolean
Public pboFlagIndexado As Boolean
Public pboCarregaDados As Boolean
Public pstrLocacaoProgramasRemtPar As String
Public pstrLocacaoProgramasLoctPar As String
Public pboNomeExecReduzidotPar As Boolean
Public piKeyAscii As Integer
Public pboCasasDecimais As Boolean
Public pboProcesso As Boolean

'Variáveis Diversas

Public pstrSql As String
Public pstrSqlAnt As String

Public Const pstrEmpresa = "Grupo Mil - D.I."

Public Const pstrUF = "AC-AL-AM-AP-BA-CE-DF-ES-GO-MA-MG-MS-MT-" & _
                      "PA-PB-PE-PI-PR-RJ-RN-RO-RR-RS-SC-SE-TO"

Public Const pstrUnidadesMedidas = "KG_MT_M2_L "

'Variáveis para geração do Log e controle de permissões

Public prstLogin As Recordset
Public prsSeleção As Recordset
Public prsFormularios As Recordset
Public pdbGMUSLOG As Database
Public pdbConfus As Database
Public pdbLog As Database
Public pWrkArea As Workspace
Public pWrkAreaLog As Workspace
'Public pstrLocacaobdGMS005 As String



Public plQuantidadeInc As Long
Public plQuantidadeAlt As Long
Public plQuantidadeExc As Long
Public plQuantidadeCon As Long
Public plQuantidadeChv As Long
Public plQuantidadePrt As Long
Public plQuantidadeAtu As Long
Public plQuantidadeExe As Long
Public plQuantidadeImp As Long
Public plQuantidadeExp As Long

Public piQuantidadeInc As Integer
Public piQuantidadeAlt As Integer
Public piQuantidadeExc As Integer
Public piQuantidadeCon As Integer
Public piQuantidadeChv As Integer
Public piQuantidadePrt As Integer
Public piQuantidadeAtu As Integer
Public piQuantidadeExe As Integer
Public piQuantidadeImp As Integer
Public piQuantidadeExp As Integer

Public pmMemorandoInc As String
Public pmMemorandoAlt As String
Public pmMemorandoExc As String
Public pmMemorandoCon As String
Public pmMemorandoChv As String
Public pmMemorandoPrt As String
Public pmMemorandoAtu As String
Public pmMemorandoExe As String
Public pmMemorandoImp As String
Public pmMemorandoExp As String

Public pboPermissaoInc As Boolean
Public pboPermissaoAlt As Boolean
Public pboPermissaoExc As Boolean
Public pboPermissaoCon As Boolean
Public pboPermissaoChv As Boolean
Public pboPermissaoPrt As Boolean
Public pboPermissaoAtu As Boolean
Public pboPermissaoExe As Boolean
Public pboPermissaoImp As Boolean
Public pboPermissaoExp As Boolean

Public Enum VBBase
    base1 = 0
    base10 = 1
    base100 = 2
    BaseOnze = 3
End Enum

' Funções da API 32 Bits
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendTBMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' Constantes da Toolbar
Private Const TBSTYLE_TRANSPARENT = &H8000
Private Const TBSTYLE_FLAT = &H800
Private Const WM_USER = &H400
Private Const TB_SETSTYLE = (WM_USER + 56)
Private Const TB_GETSTYLE = (WM_USER + 57)
Private Const TBSTYLE_LIST = &H1000
Private Const CCS_NODIVIDER = &H40

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259
Global Const KEY_ALL_ACCESS = &H3F
Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal _
hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags_ As Long) As Long
Global Const conHwndTopmost = -1
Global Const conHwndNoTopmost = -2
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If

        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:

    Resume QueryValueExExit

End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
' Description:
'   This Function will return the data field of a value
'
' Syntax:
'   Variable = QueryValue(Location, KeyName, ValueName)
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the key that the value is under (example: "Software\Microsoft\Windows\CurrentVersion\Explorer")
'
'   ValueName is the name of the value you want to access (example: "link")

       Dim lRetVal As Long         'result of the API functions
       Dim hKey As Long         'handle of opened key
       Dim vValue As Variant      'setting of queried value

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValue = vValue
       RegCloseKey (hKey)
End Function

Public Function pfboRotinasIniciais(lstrCodPrograma As String) As Boolean
    On Error GoTo Erro
    pfboRotinasIniciais = False
    
    pstrCodPrograma = lstrCodPrograma
    pstrSenhaBancoDadosBDCONFUS = Empty
    
    'Verifico se o arquivo de Senhas do CONFUS existe
    If Dir(App.Path & IIf(Right(App.Path, 1) = "\", Empty, "\") & "txtCONFUS") <> Empty Then
        'Pego Senha do BDCONFUS
        Open App.Path & IIf(Right(App.Path, 1) = "\", Empty, "\") & "txtCONFUS" For Input As #1
        Line Input #1, pstrSenhaBancoDadosBDCONFUS
        pstrSenhaBancoDadosBDCONFUS = fstrCript(pstrSenhaBancoDadosBDCONFUS)
        Close #1
    End If
    Call ppAbre_BDAcesso(pWrkArea, Empty, pdbGMUSLOG, "C:\InfoMil_Estacao\GmusLog.Dll")
    pstrSql = "SELECT tLogin.* FROM tLogin WHERE tLogin.strCodProgramatLogin = '" & lstrCodPrograma & "'"
    
    Set prstLogin = pfrsSelecao(pdbGMUSLOG, pstrSql)
    
    If prstLogin.RecordCount > 1 Then
        MsgBox "Falha no Sistema. Comunique-se com o D.I.", vbInformation, "pfboRotinasIniciais"
        prstLogin.Close
        pdbGMUSLOG.Close
        pfboRotinasIniciais = False
        Exit Function
    End If
    
        
    pboAcessoSistema = False
    If prstLogin.RecordCount = 0 Then
            FormAcessoSistema.Show 1
            If pboAcessoSistema = False Then
                pfboRotinasIniciais = False
                Exit Function
            End If
        
    ElseIf IsNull(prstLogin.Fields("strCoordenadaNavegtLogin")) Or prstLogin.Fields("strCoordenadaNavegtLogin") = Empty Then
'        If Trim(UCase(Command$)) <> "\MODULAR" Then
        FormAcessoSistema.Show 1
        If pboAcessoSistema = False Then
            pfboRotinasIniciais = False
            Exit Function
        End If
    End If
    
    If pboAcessoSistema = True Then
        pstrSql = "SELECT * FROM tLogin WHERE strCodProgramatLogin = '" & lstrCodPrograma & "'"
        Set prstLogin = pfrsSelecao(pdbGMUSLOG, pstrSql)
    End If
    
    pboAcessoSistema = True
    pstrCoordenadaNavegtLogin = IIf(IsNull(prstLogin.Fields("strCoordenadaNavegtLogin")), "00.00.00.00", prstLogin.Fields("strCoordenadaNavegtLogin"))
    pstrLocacaoIcoLogotipotLogin = IIf(IsNull(prstLogin.Fields("strLocacaoIcoLogotipoTLogin")), Empty, prstLogin.Fields("strLocacaoIcoLogotipoTLogin"))
    piCodUsuariotLogin = IIf(IsNull(prstLogin.Fields("iCodUsuariotLogin")), 0, prstLogin.Fields("iCodUsuariotLogin"))
    plSeqLogtLogin = IIf(IsNull(prstLogin.Fields("lSeqLogtLogin")), 0, prstLogin.Fields("lSeqLogtLogin"))
    pbCodEsquemaBancoDadostLogin = IIf(IsNull(prstLogin.Fields("bCodEsquemaBancoDadostLogin")), 1, prstLogin.Fields("bCodEsquemaBancoDadostLogin"))
    pstrCodFuncaotLogin = IIf(IsNull(prstLogin("strCodFuncaotLogin")), "0000000000", prstLogin("strCodFuncaotLogin"))
    pbTabCtrLogProgramatLogin = IIf(IsNull(prstLogin("bTabCtrLogProgramatLogin")) Or prstLogin("bTabCtrLogProgramatLogin") = 0, 1, prstLogin("bTabCtrLogProgramatLogin"))
    pstrLocacaoAcessoriostLogin = Trim(IIf(IsNull(prstLogin.Fields("strLocacaoAcessoriostLogin")), Empty, prstLogin.Fields("strLocacaoAcessoriostLogin")))
    pstrLocacaobdConfus = prstLogin.Fields("strLocacaobdConfustLogin")
    pdtDataRealtLogin = Format(prstLogin.Fields("dtDataRealtLogin"), "dd/mm/yyyy")

    ' Abro o arquivo BDConfus
    Call ppAbre_BDAcesso(pWrkArea, lstrCodPrograma, pdbConfus, pstrLocacaobdConfus, pstrSenhaBancoDadosBDCONFUS)
    ' Abro o tabela tUsuarios e pego o Nivel dele
    
    If piCodUsuariotLogin <> 8888 And piCodUsuariotLogin <> 9999 Then
        pstrSql = "SELECT tUsuarios.iCodUsuariotUsu, tUsuarios.bCodGrupoUsuariotUsu, tUsuarios.bTabNivelUsuariotUsu FROM tUsuarios WHERE tUsuarios.iCodUsuariotUsu = " & prstLogin.Fields("iCodUsuariotLogin")
        Set prsSeleção = pfrsSelecao(pdbConfus, pstrSql)
        If prsSeleção.EOF Then
            MsgBox "Não consigo localizar o usuário " & prsSeleção.Fields("iCodUsuariotLogin") & " em tUsuarios !", vbInformation, "pfboRotinasIniciais"
            prstLogin.Close
            pdbGMUSLOG.Close
            prsSeleção.Close
            pdbConfus.Close
            Exit Function
        End If
        pbTabNivelUsuariotUsu = prsSeleção.Fields("bTabNivelUsuariotUsu")
        pbCodGrupoUsuariotUsu = prsSeleção.Fields("bCodGrupoUsuariotUsu")
        
    Else
    
        pbTabNivelUsuariotUsu = 4
        pbCodGrupoUsuariotUsu = 1
        pstrCodFuncaotLogin = "111111111111"
    End If
    
    'Pegar Locacao do banco de dados bdLog
    
    pstrSql = "SELECT tLocacaoBancoDados.bCodEsquemaBancoDadostLoc, tLocacaoBancoDados.strCodBancoDadostLoc, tLocacaoBancoDados.strLocacaoBancoDadostLoc, tLocacaoBancoDados.bTabTipoBancoDadostLoc From tLocacaoBancoDados Where (((tLocacaoBancoDados.bCodEsquemaBancoDadostLoc) = " & Str(pbCodEsquemaBancoDadostLogin) & ") And ((tLocacaoBancoDados.strCodBancoDadostLoc) = 'bdLog') And ((tLocacaoBancoDados.bTabTipoBancoDadostLoc) = 1))"
    Set prsSeleção = pfrsSelecao(pdbConfus, pstrSql)
    
    If prsSeleção.EOF Then
        MsgBox "Não consigo pegar Locacao do bdLog em tLocacaoBancoDados !", vbInformation, "pfboRotinasIniciais"
        prstLogin.Close
        pdbGMUSLOG.Close
        prsSeleção.Close
        pdbConfus.Close
        Exit Function
    End If
    
    pstrNomeIdentFormtLogin = IIf(IsNull(prstLogin("strNomeIdentFormtLogin")), "Não Encontrado", prstLogin("strNomeIdentFormtLogin"))
    pstrNomeIdentReporttLogin = IIf(IsNull(prstLogin("strNomeIdentReporttLogin")), "Não Encontrado", prstLogin("strNomeIdentReporttLogin"))
    pstrDescrProgramatLogin = IIf(IsNull(prstLogin("strDescrProgramatLogin")), Empty, prstLogin("strDescrProgramatLogin"))
    pstrLocacaobdLog = prsSeleção.Fields("strLocacaoBancoDadostLoc")
    piCodEmpresatPar = 0
    pstrEstadoEmpresatPar = Empty
    
    pstrSql = "Select * from tParametros"
    Set prsSeleção = pfrsSelecao(pdbConfus, pstrSql)
    
    If Not prsSeleção.EOF Then
        pstrLocacaoProgramasRemtPar = prsSeleção("strLocacaoProgramasRemtPar")
        pstrLocacaoProgramasLoctPar = prsSeleção("strLocacaoProgramasLoctPar")
        pboNomeExecReduzidotPar = prsSeleção("boNomeExecReduzidotPar")
        piCodEmpresatPar = 0 & prsSeleção("iCodEmpresatPar")
        pstrEstadoEmpresatPar = Empty & prsSeleção("strEstadoEmpresatPar")
        pbIndMixEmpresatPar = 0 & prsSeleção("bIndMixEmpresatPar")
    End If
    
    prsSeleção.Close
    pdbConfus.Close
    
    'Abre o arquivo de Logs
    Call ppAbre_BDAcesso(pWrkAreaLog, Empty, pdbLog, pstrLocacaobdLog, pstrSenhaBancoDadosBDCONFUS)
    If plSeqLogtLogin > 0 Then
        'Pega o Codigo do terminal
        pstrSql = "SELECT iCodEstacaotLog From tLogAcesso Where lSeqLogtLog = " & plSeqLogtLogin
        Set prsSeleção = pfrsSelecao(pdbLog, pstrSql)
        If prsSeleção.EOF Then
            MsgBox "Não consigo encontrar o Codigo da estacao em tLogAcesso !", vbInformation, "pfboRotinasIniciais"
            prstLogin.Close
            pdbGMUSLOG.Close
            prsSeleção.Close
            pdbLog.Close
            Exit Function
        End If
        piCodEstacaotLog = prsSeleção.Fields("iCodEstacaotLog")
    Else
        pstrSql = "Insert INTO tLogAcesso (iCodUsuariotLog, iCodEstacaotLog, strCodProgramatLog, strCodFuncaotLog, dtDataInicialtLog, lHoraInicialtLog, bCodSituacaoLogtLog) Values (" & piCodUsuariotLogin & ", " & piCodEstacaotLog & ", '" & Trim(pstrCodPrograma) & "', '" & pstrCodFuncaotLogin & "', '" & Date & "', " & CLng(Hour(Time) & Format(Minute(Time), "00") & Format(Second(Time), "00")) & ", 1)"
        pdbLog.Execute pstrSql, dbFailOnError
        
        pstrSql = "Select lSeqLogtLog from tLogAcesso where icodUsuariotLog = " & piCodUsuariotLogin & " and strCodProgramatLog = '" & Trim(pstrCodPrograma) & "' Order by dtDataInicialtLog, lHoraInicialtLog"
        Set prsSeleção = pfrsSelecao(pdbLog, pstrSql)
        prsSeleção.MoveLast
        plSeqLogtLogin = prsSeleção("lSeqLogtLog")
    End If
    
    pdbLog.Close
    pfboRotinasIniciais = True
    
    Call ppDesmembraPermissao(IIf(IsNull(pstrCodFuncaotLogin), "00000000", pstrCodFuncaotLogin))
    
    prstLogin.Close
    
    pstrSql = "UPDATE tLogin SET strCodProgramatLogin = Null, strCoordenadaNavegtLogin = Null, strCodFuncaotLogin = Null, bTabCtrLogProgramatLogin = 0, lSeqLogtLogin = 0, strDescrProgramatLogin = Null WHERE strCodProgramatLogin = '" & Trim(pstrCodPrograma) & "'"
    pdbGMUSLOG.Execute pstrSql, dbFailOnError
    
    pdbGMUSLOG.Close
    
    plQuantidadeInc = 0
    plQuantidadeAlt = 0
    plQuantidadeExc = 0
    plQuantidadeCon = 0
    plQuantidadeChv = 0
    plQuantidadePrt = 0
    plQuantidadeAtu = 0
    plQuantidadeExe = 0
    plQuantidadeImp = 0
    plQuantidadeExp = 0
    
    pmMemorandoInc = Empty
    pmMemorandoAlt = Empty
    pmMemorandoExc = Empty
    pmMemorandoCon = Empty
    pmMemorandoChv = Empty
    pmMemorandoPrt = Empty
    pmMemorandoAtu = Empty
    pmMemorandoExe = Empty
    pmMemorandoImp = Empty
    pmMemorandoExp = Empty
    Exit Function
    
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "pfboRotinasIniciais"
    Err.Clear
End Function

Sub ppAbre_BDAcesso(mWrkArea As Workspace, mstrArea As String, mbdfile_name As Database, mstrdiretorio As String, Optional lstrSenhaBancoDadostLoc As String, Optional lboExclusivo As Boolean)
    On Error Resume Next
    
    Dim lstrCodBancoDados As String

    If IsNull(lboExclusivo) Then lboExclusivo = False
   
    If Dir(mstrdiretorio) = Empty Then
        MsgBox "Arquivo " & mstrdiretorio & " não Encontrado! Verifique...", vbInformation, "ppAbre_BDAcesso"
        Call ppRotinasFinais
        End
    End If
    
Abre:
    
    Err.Number = 3050
    
    Do While Err.Number = 3050
        Err.Number = 0
        If IsNull(mstrArea) Or mstrArea = Empty Then
            Set mbdfile_name = DBEngine.Workspaces(0).OpenDatabase(mstrdiretorio, lboExclusivo, False, ";pwd=" & lstrSenhaBancoDadostLoc)
        Else
            'area e para trabalhar com begintrans
            Set mWrkArea = DBEngine.CreateWorkspace(mstrArea, "Admin", Empty)
            DBEngine.Workspaces.Append mWrkArea
            Set mbdfile_name = mWrkArea.OpenDatabase(mstrdiretorio, lboExclusivo, False, ";pwd=" & lstrSenhaBancoDadostLoc)
            'Set mbdfile_name = DBEngine.Workspaces(mstrArea).OpenDatabase(mstrdiretorio, False, False, ";pwd=" & lstrSenha)
        End If
    Loop
    
    If Err.Number <> 0 And Err.Number <> 3050 Then GoTo Erro
    
    Exit Sub

Erro:
    If Err.Number = 3031 Then
        MsgBox "Senha de acesso ao " & mstrdiretorio & " inválida!" & Chr(13) & Chr(13) & "Favor entrar en contato com Depto Informática!", vbCritical, "ppAbre_BDAcesso"
    ElseIf Err.Number = 3356 Or Err.Number = 3045 Or Err.Number = 3049 Then
        If MsgBox("Outro usuário está usando o Banco de Dados em modo Exclusivo!" & vbCr & vbCr & "Favor verificar nas outras estações!" & vbCr & vbCr & "Deseja tentar novamente ?", vbCritical + vbYesNo, "ppAbre_BDAcesso") = vbYes Then
            GoTo Abre
        Else
            Err.Clear
            Call ppRotinasFinais
        End If
        
    ElseIf Err.Number = 3024 Then
        If MsgBox("Banco de dados não encontrado! Provavelmente em processo de verificação em outra estação!" & vbCr & vbCr & "Caso não esteja em processo de verificação, favor entrar em contato URGENTE com o Depto de Informática!" & vbCr & vbCr & Str(Err.Number) & ": " & Err.Description & vbCr & vbCr & "Deseja tentar novamente ?", vbCritical + vbYesNo, "ppAbre_BDAcesso") = vbYes Then
            GoTo Abre
        Else
            Err.Clear
            Call ppRotinasFinais
        End If
    Else
        MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "ppAbre_BDAcesso"
        Err.Clear
    End If
    End
End Sub

Public Sub ppDesmembraPermissao(lstrCodFuncao As String)
    pboPermissaoInc = IIf(Mid(lstrCodFuncao, 1, 1) = 1, True, False)
    pboPermissaoAlt = IIf(Mid(lstrCodFuncao, 2, 1) = 1, True, False)
    pboPermissaoExc = IIf(Mid(lstrCodFuncao, 3, 1) = 1, True, False)
    pboPermissaoCon = IIf(Mid(lstrCodFuncao, 4, 1) = 1, True, False)
    pboPermissaoChv = IIf(Mid(lstrCodFuncao, 5, 1) = 1, True, False)
    pboPermissaoPrt = IIf(Mid(lstrCodFuncao, 6, 1) = 1, True, False)
    pboPermissaoAtu = IIf(Mid(lstrCodFuncao, 7, 1) = 1, True, False)
    pboPermissaoExe = IIf(Mid(lstrCodFuncao, 8, 1) = 1, True, False)
    pboPermissaoImp = IIf(Mid(lstrCodFuncao, 9, 1) = 1, True, False)
    pboPermissaoExp = IIf(Mid(lstrCodFuncao, 10, 1) = 1, True, False)
End Sub

Public Sub ppValidaTeclaNumerico(KeyAscii, Optional lboValor As Boolean)
'-----------------------------------------------------------------------
'Função de validação da tecla pressionada em um campo
'numérico. Valores aceitos para KeyAscii:
' 8 = Tecla BackSpace
' 48 a 57 = Números
' 46 = ponto
' 44= vírgula
'-----------------------------------------------------------------------
    piKeyAscii = KeyAscii
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{TAB}": Exit Sub
    If lboValor = False Then If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 8 Or KeyAscii = 44 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Public Sub ppValidaTeclaTexto(KeyAscii As Integer, Optional lboEnter As Boolean)
'-----------------------------------------------------------------------
'Função de validação da tecla pressionada em um campo
'texto. Valores aceitos para KeyAscii:
' 8 = Tecla BackSpace
' 34 = "
' 39 = '
'-----------------------------------------------------------------------
    piKeyAscii = KeyAscii
    If KeyAscii = 13 And lboEnter = False Then SendKeys "{TAB}": Exit Sub
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Function pfrsSelecao(mfdbBanco As Database, mfstrsql As String, Optional mboRegistroEditavel As Boolean) As Recordset
    On Error GoTo Erro
    If mboRegistroEditavel Then
        Set pfrsSelecao = mfdbBanco.OpenRecordset(mfstrsql, dbOpenDynaset)
    Else
        Set pfrsSelecao = mfdbBanco.OpenRecordset(mfstrsql, dbOpenSnapshot)
    End If
    
    DBEngine.Idle dbRefreshCache

    Exit Function
    
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical + vbMsgBoxHelpButton, "pfrsSelecao", Err.HelpFile, Err.HelpContext
    Err.Clear
    'Resume ' Next
    
 End Function

Function pfboAtualizaRegistro(ldbBanco As Database, Optional lstrSql As String, Optional lrsRecordset As Recordset) As Boolean
    On Error GoTo ErroExecuteSql
    Dim db As Database
    
    pfboAtualizaRegistro = False
    
    If lstrSql = Empty Then
        lrsRecordset.Update
    Else
        'ldbBanco.Execute lstrSql, dbFailOnError
        'ldbBanco.Execute lstrSql, dbFailOnError
        ldbBanco.Execute lstrSql, dbFailOnError
    End If
    
    DBEngine.Idle dbFreeLocks
    
    pfboAtualizaRegistro = True
    
    Exit Function
    
ErroExecuteSql:
    Select Case Err.Number
        Case 3167, 3197, 3260
            Select Case pfbErros_Rede()
                Case 1, 2
                    Resume
                Case 4
                    pfboAtualizaRegistro = False
            End Select
        Case Else
            MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "`pfboAtualizaRegistro"
    End Select
    Resume
    Err.Clear
    
End Function

Function pfboEditaRegistro(lrsRecordset As Recordset, Optional lboTentativas As Boolean) As Boolean
    On Error GoTo ErroEditaRegistro
    Dim lbTentativas As Byte
    
    pfboEditaRegistro = False
    
    lrsRecordset.Edit
    
    pfboEditaRegistro = True
    
    'If FormAguarde.Visible = True Then ppAguarde FormAguarde, False, 0, 0
        
    Exit Function
    
ErroEditaRegistro:

    If lboTentativas And Err.Number = 3260 Then
        'Inicializa form aguarde...
     '   ppAguarde FormAguarde, True, 0, 0, "Aguarde, registro bloqueado..."
        lbTentativas = lbTentativas + 1
        If lbTentativas < 10 Then Resume
        'Finaliza form aguarde...
     '   ppAguarde FormAguarde, False, 0, 0
        lbTentativas = 0
    End If
    
    Select Case Err.Number
        Case 3167, 3197, 3260
            Select Case pfbErros_Rede()
                Case 1, 2
                    Resume
                Case 4
                    pfboEditaRegistro = False
            End Select
        Case Else
            MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "pfboEditaRegistro"
    End Select
    Err.Clear
End Function

Public Sub ppRotinasFinais()
    On Error GoTo Erro

    If pbTabCtrLogProgramatLogin > 0 Then
        Call ppAbre_BDAcesso(pWrkAreaLog, Empty, pdbLog, pstrLocacaobdLog, pstrSenhaBancoDadosBDCONFUS)
        pstrSql = "UPDATE tLogAcesso SET dtDataFinaltLog = #" & Format(Date, "mm/dd/yyyy") & "#, lHoraFinaltLog = " & CLng(Hour(Time) & Format(Minute(Time), "00") & Format(Second(Time), "00")) & ", bCodSituacaoLogtLog = 3, boAtualizaNavegadortLog = False WHERE lSeqLogtLog = " & plSeqLogtLogin

        pdbLog.Execute pstrSql, dbFailOnError
        
        If pbTabCtrLogProgramatLogin > 1 Then
            If pbTabCtrLogProgramatLogin <> 3 Then
                pmMemorandoInc = Empty
                pmMemorandoAlt = Empty
                pmMemorandoExc = Empty
                pmMemorandoCon = Empty
                pmMemorandoChv = Empty
                pmMemorandoPrt = Empty
                pmMemorandoAtu = Empty
                pmMemorandoExe = Empty
                pmMemorandoImp = Empty
                pmMemorandoExc = Empty
            End If
            
            If plQuantidadeInc > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",1," & plQuantidadeInc & ",'" & pmMemorandoInc & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If
            
            If plQuantidadeAlt > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",2," & plQuantidadeAlt & ",'" & pmMemorandoAlt & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If

            If plQuantidadeExc > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",3," & plQuantidadeExc & ",'" & pmMemorandoExc & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If
            
            If plQuantidadeCon > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",4," & plQuantidadeCon & ",'" & pmMemorandoCon & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If
            
            If plQuantidadeChv > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",5," & plQuantidadeChv & ",'" & pmMemorandoChv & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If
            
            If plQuantidadePrt > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",6," & plQuantidadePrt & ",'" & pmMemorandoPrt & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If
            
            If plQuantidadeAtu > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",7," & plQuantidadeAtu & ",'" & pmMemorandoAtu & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If

            If plQuantidadeExe > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",8," & plQuantidadeExe & ",'" & pmMemorandoExe & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If
            
            If plQuantidadeImp > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",9," & plQuantidadeImp & ",'" & pmMemorandoImp & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If

            If plQuantidadeExp > 0 Then
                pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & ",10," & plQuantidadeExp & ",'" & pmMemorandoExp & "')"
                pdbLog.Execute pstrSql, dbFailOnError
            End If
            
        End If
        
        pdbLog.Close
        
    End If
    
'    Call ppAbre_BDAcesso(EMPTY, pdbGMUSLOG, "C:\InfoMil_Estacao\GmusLog.Dll")
        
'    pstrSql = "UPDATE tLogin SET strCodProgramatLogin = Null, strCoordenadaNavegtLogin = Null, strCodFuncaotLogin = Null, bTabCtrLogProgramatLogin = 0, lSeqLogtLogin = 0, strDescrProgramatLogin = Null WHERE strCodProgramatLogin = '" & Trim(pstrCodPrograma) & "'"
    
'    pdbGMUSLOG.Execute pstrSql, dbFailOnError
    
'    pdbGMUSLOG.Close
    Exit Sub
    
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "ppRotinasFinGMS"
    Err.Clear
End Sub

Public Sub ppAtualiza_LogAcesso()
    On Error GoTo Erro
    Call ppAbre_BDAcesso(pWrkAreaLog, Empty, pdbLog, pstrLocacaobdLog, pstrSenhaBancoDadosBDCONFUS)
    pstrSql = "UPDATE tLogAcesso SET boAtualizaNavegadortLog = 1 Where bCodSituacaoLogtLog = 1"
    pdbLog.Execute pstrSql, dbFailOnError
    pdbLog.Close
    Exit Sub
    
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "ppAtualiza_LogAcesso"
    Err.Clear
End Sub

Public Sub ppCarregaPropriedadesForm(lstrNameForm As Form, Optional lstrHelpContext As HelpConstants)
    On Error GoTo Erro
    
    'Pegando o nome do usuario plugado
    Dim llBuffer As Long
    Dim lstrNome As String
    Dim llongw As Long
    llBuffer = 200
    lstrNome = String(llBuffer, 0)
    llongw = GetUserName(lstrNome, llBuffer)
    
    pstrNomeUsuario = Mid(lstrNome, 1, llBuffer - 1)
    
    If Dir(pstrLocacaoProgramasRemtPar & IIf(Right(pstrLocacaoProgramasRemtPar, 1) <> "\", "\", Empty) & "Infomil.HLP") <> Empty Then
        App.HelpFile = pstrLocacaoProgramasRemtPar & IIf(Right(pstrLocacaoProgramasRemtPar, 1) <> "\", "\", Empty) & "Infomil.HLP"
    ElseIf Dir(pstrLocacaoProgramasLoctPar & IIf(Right(pstrLocacaoProgramasLoctPar, 1) <> "\", "\", Empty) & "Infomil.HLP") <> Empty Then
        App.HelpFile = pstrLocacaoProgramasLoctPar & IIf(Right(pstrLocacaoProgramasLoctPar, 1) <> "\", "\", Empty) & "Infomil.HLP"
    Else
        App.HelpFile = App.Path & IIf(Right(App.Path, 1) <> "\", "\", Empty) & "Infomil.HLP"
    End If
    
    lstrNameForm.HelpContextID = lstrHelpContext
    lstrNameForm.Icon = IIf(Dir(pstrLocacaoIcoLogotipotLogin) <> Empty, LoadPicture(pstrLocacaoIcoLogotipotLogin), Empty)
    lstrNameForm.Left = (Screen.Width - lstrNameForm.Width) / 2
    lstrNameForm.Top = (Screen.Height - lstrNameForm.Height) / 2
    If lstrNameForm.Name = "FormPrincipal" Then lstrNameForm.Caption = pstrCoordenadaNavegtLogin & " - " & pstrNomeIdentFormtLogin & " - " & IIf(Trim(pstrDescrProgramatLogin) <> Empty, pstrDescrProgramatLogin, Trim(lstrNameForm.Caption)) Else lstrNameForm.Caption = Trim(lstrNameForm.Caption) & " - " & pstrFuncaoToolbar
    Exit Sub
    
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "ppCarregaPropriedadesForm"
    Err.Clear
End Sub

Public Sub ppGravaLogAnalitico(pstrDetalheTransacao As String, Optional pboArquivo As Boolean)
    On Error GoTo Erro
    If plSeqLogtLogin > 0 And pbTabCtrLogProgramatLogin = 3 Then
        Dim pbFuncao As Byte, plQuantidade As Integer
            
        If plQuantidadeInc > 0 Then pbFuncao = 1: plQuantidade = plQuantidadeInc
        If plQuantidadeAlt > 0 Then pbFuncao = 2: plQuantidade = plQuantidadeAlt
        If plQuantidadeExc > 0 Then pbFuncao = 3: plQuantidade = plQuantidadeExc
        If plQuantidadeCon > 0 Then pbFuncao = 4: plQuantidade = plQuantidadeCon
        If plQuantidadeChv > 0 Then pbFuncao = 5: plQuantidade = plQuantidadeChv
        If plQuantidadePrt > 0 Then pbFuncao = 6: plQuantidade = plQuantidadePrt
        If plQuantidadeAtu > 0 Then pbFuncao = 7: plQuantidade = plQuantidadeAtu
        If plQuantidadeExe > 0 Then pbFuncao = 8: plQuantidade = plQuantidadeExe
        If plQuantidadeImp > 0 Then pbFuncao = 9: plQuantidade = plQuantidadeImp
        If plQuantidadeExp > 0 Then pbFuncao = 10: plQuantidade = plQuantidadeExp

        If pboArquivo = False Then Call ppAbre_BDAcesso(pWrkAreaLog, Empty, pdbLog, pstrLocacaobdLog, pstrSenhaBancoDadosBDCONFUS)

        pstrSql = "Select * from tLogAcessoFuncao Where lSeqLogtLogFun = " & plSeqLogtLogin & " and bcodFuncaoExectLogFun = " & pbFuncao
        Set prsSeleção = pfrsSelecao(pdbLog, pstrSql)
        If prsSeleção.EOF Then
            pstrSql = "INSERT INTO tLogAcessoFuncao (lSeqLogtLogFun, bCodFuncaoExectLogFun, lQtdeTransacaotLogFun, mDetalheTransacaotLogFun) VALUES (" & plSeqLogtLogin & "," & pbFuncao & "," & plQuantidade & ",'" & pstrDetalheTransacao & "')"
            pdbLog.Execute pstrSql, dbFailOnError
        Else
            If InStr(1, prsSeleção("mDetalheTransacaotLogFun"), pstrDetalheTransacao, vbTextCompare) = 0 Then
                pstrSql = "UPDATE tLogAcessoFuncao Set lQtdeTransacaotLogFun = " & prsSeleção("lQtdeTransacaotLogFun") + plQuantidade & ", mDetalheTransacaotLogFun = '" & prsSeleção("mDetalheTransacaotLogFun") & " " & Trim(pstrDetalheTransacao) & "' Where lSeqLogtLogFun = " & plSeqLogtLogin & " and bcodFuncaoExectLogFun = " & pbFuncao
            Else
                pstrSql = "UPDATE tLogAcessoFuncao Set lQtdeTransacaotLogFun = " & prsSeleção("lQtdeTransacaotLogFun") + plQuantidade & " Where lSeqLogtLogFun = " & plSeqLogtLogin & " and bcodFuncaoExectLogFun = " & pbFuncao
            End If
            pdbLog.Execute pstrSql, dbFailOnError
        End If
        If pboArquivo = False Then pdbLog.Close
        plQuantidadeInc = 0: plQuantidadeAlt = 0
        plQuantidadeExc = 0: plQuantidadeCon = 0
        plQuantidadeChv = 0: plQuantidadePrt = 0
        plQuantidadeAtu = 0: plQuantidadeExe = 0
        plQuantidadeImp = 0: plQuantidadeExp = 0
    End If
    Exit Sub
    
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "ppGravaLogAnalitico"
    Err.Clear
End Sub

Function pfboCriaQueries(pdbBanco As Database, pqdNomeQuery As QueryDef, pstrComandoQuery As String, pstrVariavelQuery As String, Optional pboArquivo As Boolean) As Boolean
    'pdbBanco           ==> Variavel do Banco de Dados
    'pqdNomeQuery       ==> Variavel Query Def
    'pstrComandoQuery   ==> String SQL
    'pstrVariavelQuery  ==> Nome da Consulta
    'pboArquivo         ==> Caso o programa trabalhe com o bdLog aberto então deverá ser True para não abrir e fechar o bdLog
    
    Err.Clear: On Error Resume Next
    Dim liTentativas As Integer, llNumErro As Long, lstrTextoErro As String
    Dim lbTabLogAnterior As Byte
    
Inicio:
    pfboCriaQueries = True: liTentativas = 0: llNumErro = 0
    Do While liTentativas < 30000
        On Error Resume Next
        Set pqdNomeQuery = pdbBanco.CreateQueryDef(pstrVariavelQuery, pstrComandoQuery)
        liTentativas = liTentativas + 1
        If Err.Number <> 0 Then
            lstrTextoErro = IIf(InStr(1, lstrTextoErro, Str(Err.Number), 1) > 0, lstrTextoErro, lstrTextoErro & " " & Err.Number)
            llNumErro = llNumErro + 1
            If Err.Number = 3012 Then
                pdbBanco.QueryDefs.Delete pstrVariavelQuery
            End If
        Else
            liTentativas = 30001
        End If
        Err.Clear
    Loop
    If llNumErro > 0 Then
        plQuantidadeAtu = plQuantidadeAtu + llNumErro: lbTabLogAnterior = pbTabCtrLogProgramatLogin
        pbTabCtrLogProgramatLogin = 3: Call ppGravaLogAnalitico("Erros de Query: " & lstrTextoErro, pboArquivo)
        pbTabCtrLogProgramatLogin = lbTabLogAnterior
    End If
    
    If liTentativas = 30000 Then
        pfboCriaQueries = False
        If MsgBox("Não foi possivel criar uma consulta no Banco de Dados!" & vbCr & "Deseja tentar novamente ?", vbQuestion + vbYesNo + vbDefaultButton1, "ppCriaQueries") = vbYes Then GoTo Inicio
    End If
End Function

Function pfvRetornaValor(lvValor As Variant) As Variant
    If Not IsNumeric(lvValor) Then lvValor = 0
    pfvRetornaValor = CDbl(lvValor)
End Function

Public Function pfstrConverteQtde(pvQtde As Variant) As String
    'Função de conversão das quantidades
    If IsNull(pvQtde) Or Trim(pvQtde) = Empty Then pvQtde = 0
    pfstrConverteQtde = Format(pvQtde, "###,##0" & IIf(pboCasasDecimais, ".000", Empty))
End Function

Public Function pfboVerificaFracao(lvqtdeCheckDeposito As Variant) As Boolean
    'Verifica a fracao da quantidade
    Dim lcFracEmbMed As Double, lcFracMovto As Double, lstrDigitos As String, lsQtdeEmbUndMed As Single
    
    pfboVerificaFracao = False
    
    lsQtdeEmbUndMed = IIf(pboCasasDecimais, 1000, 1)
    
    lcFracEmbMed = CDbl("0," & pfvRetornaValor(lsQtdeEmbUndMed) - 1)
    lcFracMovto = CDbl(pfvRetornaValor(lvqtdeCheckDeposito) - Int(pfvRetornaValor(lvqtdeCheckDeposito)))
    
    If lcFracEmbMed = 0 Then Exit Function
    
    If InStr(1, lvqtdeCheckDeposito, ",") = 0 Then
        lstrDigitos = "0,0"
    Else
        lstrDigitos = "0," & Mid(lvqtdeCheckDeposito, InStr(1, lvqtdeCheckDeposito, ",") + 1)
    End If
    
    'Verifica se a quantidade de digitos digitado e maior q a quantidade da embalagem - 1
    If Len(lstrDigitos) > Len(Trim(CStr(lcFracEmbMed))) Then pfboVerificaFracao = True: Exit Function
    'Verifica se o valor da fração digitada e maior q a fração da embalagem - 1
    If Format(lcFracMovto, "#.#########") > Format(lcFracEmbMed, "#.#########") Then pfboVerificaFracao = True: Exit Function
    'If Len(Trim(CStr(lcFracMovto))) > Len(Trim(CStr(lcFracEmbMed))) Then pfboVerificaFracao = True
End Function

'FUNCAO pfstrExtenso

Function pfstrExtenso(Valor As Double) As String
    Dim C1 As Integer, C2 As Integer
    Dim u As Integer
    Dim M As Integer
    Dim VV As String
    Dim Centavos As String
    Dim Unidade As String
    Dim Milhar As String
VV = Format(Valor, "000000.00")
M = Mid(VV, 1, 3)
u = Mid(VV, 4, 3)
C1 = Mid(VV, InStr(VV, ",") + 1, 1) + "0"
C2 = Mid(VV, InStr(VV, ",") + 2, 1)
'Descricao dos Centavos
Centavos = Centena(C1 + C2)
If C1 + C2 = 0 Or C1 + C2 > 1 Then
    Centavos = Centavos + " Centavos"
Else
    Centavos = Centavos + " Centavo"
End If
Unidade = Centena(u)
 
If (u <= 1) And (M = 0) Then
    Unidade = Unidade + " Real"
Else
    Unidade = Unidade + " Reais"
End If
  
Milhar = Centena(M)
Milhar = Milhar + " mil "
 
pfstrExtenso = IIf(M > 0, Milhar, Empty)
 
If M > 0 Then
    pfstrExtenso = Trim(pfstrExtenso) + IIf(u = 0, " Reais ", " e " + Unidade)
Else
    pfstrExtenso = Trim(pfstrExtenso) + IIf(u = 0, " Reais ", Unidade)
End If
 
If M + u = 0 Then
    pfstrExtenso = pfstrExtenso + IIf(C1 + C2 = 0, Empty, Centavos)
Else
    pfstrExtenso = pfstrExtenso + IIf(C1 + C2 = 0, Empty, " e " + Centavos)
End If
 
End Function

Function Centena(i As Integer) As String
Dim VV As String
Dim u1 As Integer, u2 As Integer, u3 As Integer
Dim Pri As String, Seg As String, Ter As String
VV = Format(i, "000")
u1 = Mid(VV, 1, 1) + "00"
u2 = Mid(VV, 2, 1) + "0"
u3 = Mid(VV, 3, 1)

If i = 0 Then
    Centena = "Zero"
    Exit Function
End If

    If u1 = 100 And u2 + u3 > 0 Then
        Pri = "Cento"
    ElseIf u1 = 100 And u2 + u3 = 0 Then
        Pri = "Cem "
    ElseIf u1 > 100 Then
        Pri = Base(u1, base100)
    ElseIf u2 + u3 > 0 Then
        Pri = Empty
    End If
    
    If (u2 + u3 > 10 And u2 + u3 < 20) Then
        Seg = Base(u2 + u3, BaseOnze)
    ElseIf u2 > 0 Then
        Seg = Base(u2, base10)
    Else
        Seg = Empty
    End If
    
    If u2 + u3 > 10 And u2 + u3 < 20 Then
        Ter = Empty
    ElseIf u3 > 0 Then
        Ter = Base(u3, base1)
    Else
        Ter = Empty
    End If
    
    If (Trim(Pri) = Empty Or Trim(Pri) = "Cem") Then
        Pri = Pri
    ElseIf u2 + u3 = 0 Then
        Pri = Pri
    Else
        Pri = Pri + " e "
    End If
    
    If u2 + u3 > 10 And u2 + u3 < 20 Then
        Seg = Seg
    ElseIf (Seg = Empty) Or (u3 = 0) Then
        Seg = Seg
    Else
        Seg = Seg + " e "
    End If
    Centena = Pri + Seg + Ter
End Function

Function Base(i As Integer, TipoBase As VBBase) As String
Select Case TipoBase
    Case base1
        Select Case i
            Case 1
                Base = "Um"
            Case 2
                Base = "Dois"
            Case 3
                Base = "Três"
            Case 4
                Base = "Quatro"
            Case 5
                Base = "Cinco"
            Case 6
                Base = "Seis"
            Case 7
                Base = "Sete"
            Case 8
                Base = "Oito"
            Case 9
                Base = "Nove"
        End Select
    Case base10
        Select Case i
            Case 10
                Base = "Dez"
            Case 20
                Base = "Vinte"
            Case 30
                Base = "Trinta"
            Case 40
                Base = "Quarenta"
            Case 50
                Base = "Cinquenta"
            Case 60
                Base = "Sessenta"
            Case 70
                Base = "Setenta"
            Case 80
                Base = "Oitenta"
            Case 90
                Base = "Noventa"
        End Select
        
    Case base100
        Select Case i
            Case 100
                Base = "Cem"
            Case 200
                Base = "Duzentos"
            Case 300
                Base = "Trezentos"
            Case 400
                Base = "Quatrocentos"
            Case 500
                Base = "Quinhentos"
            Case 600
                Base = "Seicentos"
            Case 700
                Base = "Setecentos"
            Case 800
                Base = "Oitocentos"
            Case 900
                Base = "Novecentos"
        End Select
    Case BaseOnze
        Select Case i
            Case 11
                Base = "Onze"
            Case 12
                Base = "Doze"
            Case 13
                Base = "Treze"
            Case 14
                Base = "Quatorze"
            Case 15
                Base = "Quinze"
            Case 16
                Base = "Dezesseis"
            Case 17
                Base = "Dezessete"
            Case 18
                Base = "Dezoito"
            Case 19
                Base = "Dezenove"
        End Select
    End Select
End Function

'Fim da Funcao pfstrExtenso

Public Function fstrCript(lstrDados As String, Optional lboStatus As Boolean) As String

       Dim lbCaracter As Byte
       Dim lbIndex As Byte
       
       fstrCript = Empty
       
       Do While lbIndex < Len(lstrDados)
            lbIndex = lbIndex + 1
            
            fstrCript = fstrCript & Chr(Asc(Mid(lstrDados, lbIndex, 1)) + IIf(lboStatus, 3, -3))
       Loop

End Function

Public Sub ppAguarde(pstrNomeForm As Form, pboLiga As Boolean, Optional plProgress As Long, Optional plTotalProgress As Long, Optional pstrTitulo As String, Optional pstrLabel As String)
    Static llcounter As Long
    
    If plTotalProgress = 0 Then
        If pboLiga Then
            pstrNomeForm.LabelTitulo.Caption = IIf(pstrTitulo <> Empty, pstrTitulo, "Aguarde, selecionando registros...")
            pstrNomeForm.Show
            pstrNomeForm.Animation.Play
            pstrNomeForm.Refresh
        Else
            llcounter = 0
            Unload pstrNomeForm
            Exit Sub
        End If
    Else
        If plProgress = 0 Then
            pstrNomeForm.LabelTotal.Enabled = True
            pstrNomeForm.LabelTotal.Caption = IIf(pstrLabel = Empty, "Total Reg: ", pstrLabel) & plTotalProgress
            pstrNomeForm.CommandCancela.Enabled = True
            pstrNomeForm.LabelReg.Enabled = True
            pstrNomeForm.ProgressBar.Enabled = True
            pstrNomeForm.ProgressBar.Min = 0
            pstrNomeForm.ProgressBar.Max = plTotalProgress
            pstrNomeForm.ProgressBar.Value = 0
        End If
        If plTotalProgress >= 5 Then
            If plProgress >= llcounter Or plProgress = plTotalProgress Then
                llcounter = llcounter + (plTotalProgress / 40)
                
                pstrNomeForm.ProgressBar.Value = plProgress
                pstrNomeForm.LabelReg.Caption = "Registros Lidos: " & Format((plProgress * 100) / plTotalProgress, "#0") & " %"
                pstrNomeForm.Refresh
                
                If plProgress = plTotalProgress Then pstrNomeForm.Animation.Stop
            End If
        End If

    End If
End Sub

Public Function pfstrValidaString(pstrTexto As String) As String
    pfstrValidaString = pstrTexto
    If InStr(1, pstrTexto, "'", 1) = 0 And InStr(1, pstrTexto, """", 1) = 0 Then Exit Function
    Dim liContador As Integer
    For liContador = 1 To Len(pstrTexto)
        If InStr(liContador, pstrTexto, "'") = 0 And InStr(liContador, pstrTexto, """") = 0 Then Exit For
        If InStr(liContador, pstrTexto, "'") > 0 Then Mid(pstrTexto, InStr(liContador, pstrTexto, "'", 1)) = " "
        If InStr(liContador, pstrTexto, """") > 0 Then Mid(pstrTexto, InStr(liContador, pstrTexto, """", 1)) = " "
    Next
    pfstrValidaString = pstrTexto
End Function

Public Sub ppPreencheObjeto(lobjNomeObjeto As Object)
     lobjNomeObjeto.SelStart = 0
     lobjNomeObjeto.SelLength = Len(Trim(lobjNomeObjeto.Text))
End Sub

Public Sub ppError(lstrProcedure As String, lobObjeto As Object)
    MsgBox "Ocorreu o erro número : " & Err.Number & Chr(10) & Err.Description & Chr(10) & Err.Source & Chr(10) & App.Revision, 16, lstrProcedure
    Err.Clear
    lobObjeto.MousePointer = 0
End Sub

Public Function pfbErros_Rede() As Byte
    
    '1 - Tenta novamente sem informar ao usuário
    '2 - Tenta novamente informando ao usuário
    '3 - Aborta o processo
    
    pfbErros_Rede = 0
    
    If Err.Number = 0 Then Exit Function
    
    Select Case Err.Number
        Case 3167
            MsgBox "O registro foi deletado por outro usuário!"
            pfbErros_Rede = 3
        Case 3197
            pfbErros_Rede = 1
        Case 3260
            If MsgBox("Registro bloqueado por outra estação! Tenta Novamente? ", vbCritical + vbQuestion + vbYesNo, "pfbErros_Rede") = vbYes Then
                pfbErros_Rede = 2
            Else
                pfbErros_Rede = 4
            End If
        Case Else
            If MsgBox(Err.Description & vbCrLf & "Tenta Novamente? ", vbCritical + vbQuestion, "pfbErros_Rede") = vbYes Then
                pfbErros_Rede = 2
            Else
                pfbErros_Rede = 4
            End If
    End Select
            
'    Err.Clear
        
End Function
