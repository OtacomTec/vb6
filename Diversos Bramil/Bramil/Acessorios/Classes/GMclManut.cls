VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMclManut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Const BRM_MSG1 = "Ocorreu o erro número "


Private Enum brmConst
    brmCriarCampo = 0
    brmExcluirCampo = 1
    brmAlterarCampo = 2
    brmAlterarPropCampo = 3
    
    brmCriarBanco = 0
    brmRepararBanco = 1
    brmCompactarBanco = 2
    
    brmArquivoCopiar = 0
    brmArquivoRenomear = 1
    brmArquivoExcluir = 2
    brmArquivoMover = 4
End Enum

Private mStrLocacao

Private mvarDAOListaIndex() As Variant
Private mvarMensagemDeErro As String 'local copy
Public Property Get MensagemDeErro() As String
    MensagemDeErro = mvarMensagemDeErro
End Property

Public Function DAOCriarConsulta(BaseDeDados As String, _
                                 NomeDaConsulta As String, _
                                 StringSQL As String, _
                                 Optional Usuário As String, _
                                 Optional Senha As String, _
                                 Optional Exclusivo As Boolean = False) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    
    If Trim(BaseDeDados) = "" _
        Or Trim(StringSQL) = "" _
        Or Trim(NomeDaConsulta) = "" Then Err.Raise 449

    Dim bd As Database
    Dim NovaConsulta As QueryDef
    
    Senha = "PWD=" & Senha & ";"
    Usuário = "UID=" & Usuário & ";"
        
    'Set bd = OpenDatabase(BaseDeDados)
    Set bd = OpenDatabase(BaseDeDados, Exclusivo, False, ";" & Senha & Usuário)
    
    Set NovaConsulta = bd.CreateQueryDef()
    NovaConsulta.Name = NomeDaConsulta
    NovaConsulta.SQL = StringSQL
    bd.QueryDefs.Append NovaConsulta
    bd.Close
       
ErroCls:
    If Err.Number = 0 Then
        DAOCriarConsulta = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAOCriarConsulta - " & Err.Description
    End If
End Function

Public Function DAOExecutarSQL(BaseDeDados As String, _
                               StringSQL As String, _
                               Optional Usuário As String, _
                               Optional Senha As String, _
                               Optional Opções As DAO.RecordsetOptionEnum = dbFailOnError, _
                               Optional Exclusivo As Boolean = False) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    
    If Trim(BaseDeDados) = "" _
        Or Trim(StringSQL) = "" Then Err.Raise 449        'erro Argumento não opcional
    
    Senha = "PWD=" & Senha & ";"
    Usuário = "UID=" & Usuário & ";"
            
    Dim wrk As Workspace
    Set wrk = Workspaces(0)
    Dim bd As Database
    Set bd = wrk.OpenDatabase(BaseDeDados, Exclusivo, False, ";" & Senha & Usuário)
    
    wrk.BeginTrans
    bd.Execute StringSQL, Opções
       
    wrk.CommitTrans
    
ErroCls:
    If Err.Number = 0 Then
        DAOExecutarSQL = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAOExecutarSQL - " & Err.Description
    End If
End Function
Public Function Executar(Comando As String, _
                         Optional Janela As VbAppWinStyle = vbHide) As Boolean
    Err.Clear
    
    If Trim(Comando) = "" Then Err.Raise 449       'erro Argumento não opcional
    
    rtn = Shell(Commando, Janela)
    
    If Err.Number = 0 Then
        Executar = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.Executar - " & Err.Description
    End If
End Function
Public Property Get DAOListaIndex() As Variant
    DAOListaIndex = mvarDAOListaIndex
    'esta propriedade é alimentada pelo método DAOListarIndex
End Property

Private Function DAOListarIndex(BaseDeDados As String, NomeDaTabela As String) As Boolean
    Dim db As Database
    Dim rs As Recordset
    On Error GoTo ErroCls
    Err.Clear
    If Trim(BaseDeDados) = "" _
        Or Trim(NomeDaTabela) = "" Then Err.Raise 449        'erro Argumento não opcional
    
    
    Set db = OpenDatabase(BaseDeDados)
    
    Dim td As TableDef
    Dim ind As Index
    Dim f As Field
    
    Set td = db.TableDefs(NomeDaTabela)
    ReDim mvarDAOListaIndex(0)
  
    For Each ind In td.Indexes
        ReDim Preserve mvarDAOListaIndex(UBound(mvarDAOListaIndex()) + 1)
        mvarDAOListaIndex(UBound(mvarDAOListaIndex())) = ind.Name
        'Utiliza o for abaixo para listar os campos da indexação
        'For Each f In ind.Fields
        '    ColunaIndexada= f.Name
        'Next
    Next

ErroCls:
    If Err.Number = 0 Then
        DAOListarIndex = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAOListarIndex - " & Err.Description
    End If


End Function

Public Function DAOCriarIndex(BaseDeDados As String, _
                             NomeDaTabela As String, _
                             NomeDaIndex As String, _
                             ColunasIndexadas As Variant) As Boolean
                             'O argumento ColunasIndexadas deverá ser passado como uma matriz
                             'com os seu valores ordenados a partir do menor índice 1
                             'Esta argumento carregará valores string para os nomes das coluas
                             'que serão utilizadas para indexação.
    On Error GoTo ErroCls
    Err.Clear
    If Trim(BaseDeDados) = "" _
        Or Trim(NomeDaTabela) = "" _
        Or Trim(NomeDaIndex) = "" Then Err.Raise 449 'erro Argumento não opcional
    
    Dim db As Database
    Set db = OpenDatabase(BaseDeDados)
    Dim td As TableDef
    Dim ind As Index
    Dim f As Field
        
    Set td = db.TableDefs(NomeDaTabela)
    Set ind = td.CreateIndex(NomeDaIndex)

    For i = 1 To UBound(ColunasIndexadas())
        ind.Fields.Append td.CreateField(ColunasIndexadas())
        ind.Primary = Chave
    Next i
    td.Indexes.Append ind
    
ErroCls:
    If Err.Number = 0 Then
        DAOCriarIndex = True
        rtn = DAOListarIndex(BaseDeDados, NomeDaTabela)
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAOCriarIndex - " & Err.Description
    End If
     
End Function

Public Function ADOCriarBanco(BaseDeDados As String, Optional Provedor As String) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    If Trim(BaseDeDados) = "" Then Err.Raise 449       'erro Argumento não opcional
       
    Dim cat As New ADOX.Catalog
    ADOCriarBanco = False
    Provedor = IIf(Provider = "", "Provider=Microsoft.Jet.OLEDB.4.0;", "Provider=" & Provedor & ";")
    cat.Create Provider & "Data Source=" & BaseDeDados & ";"
    
ErroCls:
    If Err.Number = 0 Then
        ADOCriarBanco = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.ADOCriarBanco - " & Err.Description
    End If
End Function
Public Function ADOCriarTabela(FonteDeDadosMDB As String, NovaTabela As String) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    
    Dim tbl As New ADOX.Table
    Dim cat As New ADOX.Catalog
    ADOCriarTabela = False
    FonteDeDadosMDB = IIf(Right(FonteDeDadosMDB, 1) <> ";", "Data Source=" & FonteDeDadosMDB & ";", "Data Source=" & FonteDeDadosMDB)
    
    'conecta com o banco de dados
    cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & FonteDeDados
    tbl.Name = NovaTabela
    cat.Tables.Append tbl
    
ErroCls:
    If Err.Number = 0 Then
        ADOCriarTabela = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.ADOCriarTabela - " & Err.Description
    End If
End Function

Public Function DAOCriarBanco(BaseDeDados As String, _
                             Optional Sobrescrever As Boolean = False, _
                             Optional Usuário As String, _
                             Optional Senha As String) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    If Trim(BaseDeDados) = "" Then Err.Raise 449       'erro Argumento não opcional
    
    Senha = "pwd=" & Senha & ";"
    Usuário = "uid=" & Usuário & ";"
    If Trim(BaseDeDados) = "" Then Err.Raise 449 'erro Argumento não opcional
        
    
    If Dir(BaseDeDados) <> "" Then
        If Sobrescrever = True Then
            Kill BaseDeDados
            Set bd = DAO.CreateDatabase(BaseDeDados, dbLangGeneral & ";" & Senha & Usuário)
        End If
    Else
        Set bd = DAO.CreateDatabase(BaseDeDados, dbLangGeneral & ";" & Senha & Usuário)
    End If
    
ErroCls:
    If Err.Number = 0 Then
        DAOCriarBanco = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAOCriarBanco - " & Err.Description
    End If
End Function
Public Function DAORepararBanco(BaseDeDados As String) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    If Trim(BaseDeDados) = "" Then Err.Raise 449         'erro Argumento não opcional
       
    DAO.DBEngine.RepairDatabase BaseDeDados
    
ErroCls:
    If Err.Number = 0 Then
        DAORepararBanco = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAORepararBanco - " & Err.Description
    End If
End Function
Public Function DAOCompactarBanco(BaseDeDados As String, Optional Senha As String) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    If Trim(BaseDeDados) = "" Then Err.Raise 449         'erro Argumento não opcional
    
    Dim strHora As String
    Dim strData As String
    
    If Dir(BaseDeDados) <> "" Then
        rtn = InStrRev(BaseDeDados, "\")
        If rtn = 0 Then
            BaseDeDadosCompactado = BaseDeDados
            CaminhoDoMDB = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
        Else
            BaseDeDadosCompactado = Mid(BaseDeDados, rtn + 1, Len(BaseDeDados) - rtn)
            CaminhoDoMDB = Mid(BaseDeDados, 1, rtn)
        End If
        strHora = Replace(Time, ":", "")
        strData = Format(Date, "yyyymmdd")
               
        'inicia cópia de segurança
        FileCopy BaseDeDados, CaminhoDoMDB & "Cópia " & strData & "-" & strHora & " " & BaseDeDadosCompactado
               
        If Dir(CaminhoDoMDB & "Compactado " & BaseDeDadosCompactado) <> "" Then
            Kill CaminhoDoMDB & "Compactado " & BaseDeDadosCompactado
        End If
        DAO.DBEngine.CompactDatabase BaseDeDados, CaminhoDoMDB & "Compactado " & BaseDeDadosCompactado, Senha & dbLangGeneral
                
        If Err.Number = 0 Then
            Kill BaseDeDados
            If Err.Number = 0 Then FileCopy CaminhoDoMDB & "Compactado " & BaseDeDadosCompactado, BaseDeDados
                If Err.Number = 0 Then
                    If Err.Number = 0 Then DAOCompactarBanco = True
                    Kill CaminhoDoMDB & "Compactado " & BaseDeDadosCompactado
                    Kill CaminhoDoMDB & "Cópia " & strData & "-" & strHora & " " & BaseDeDadosCompactado
                End If
            End If
        End If
        
ErroCls:
    If Err.Number = 0 Then
        DAOCompactarBanco = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAOCompactarBanco - " & Err.Description
    End If
End Function
Public Function DAOCriarColuna(BaseDeDados As String, _
                               NomeDaTabela As String, _
                               NovaColuna As String, _
                               Optional TipoDaColuna As DAO.DataTypeEnum = DAO.dbText, _
                               Optional TamDaColuna As Byte = 255, _
                               Optional Requerido As Boolean = False, _
                               Optional Usuário As String, _
                               Optional Senha As String, _
                               Optional ModoExclusivo As Boolean = False) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    If Trim(BaseDeDados) = "" _
        Or Trim(NomeDaTabela) = "" _
        Or Trim(NovaColuna) = "" Then Err.Raise 449        'erro Argumento não opcional
    
    Senha = "pwd=" & Senha & ";"
    Usuário = "uid=" & Usuário & ";"
    Dim db As Database
    Set db = OpenDatabase(BaseDeDados, ModoExclusivo, , Usuário & Senha)
    Dim td As TableDef
    Dim f As Field

    Set td = db.TableDefs(NomeDaTabela)
    If TipoDaColuna = dbText Then
        Set f = td.CreateField(NovaColuna)
        f.Type = dbText
        f.Size = TamDaColuna
        f.AllowZeroLength = True
    Else
        Set f = td.CreateField(NovaColuna, TipoDaColuna)
    End If
    f.Required = Requerido
   ' f.AllowZeroLength = True
    td.Fields.Append f

    db.Close
    
ErroCls:
    If Err.Number = 0 Then
        DAOCriarColuna = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAOCriarColuna - " & Err.Description
    End If
End Function
Public Function DAOExcluirColuna(BaseDeDados As String, _
                               NomeDaTabela As String, _
                               NovaColuna As String) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    If Trim(BaseDeDados) = "" _
        Or Trim(NomeDaTabela) = "" _
        Or Trim(NovaColuna) = "" Then Err.Raise 449        'erro Argumento não opcional
    
    Dim db As Database
    Set db = OpenDatabase(BaseDeDados, True)
    Dim td As TableDef
    Dim f As Field
    
    Set td = db.TableDefs(NomeDaTabela)
    td.Fields.Delete (NovaColuna)

    db.Close
    
ErroCls:
    If Err.Number = 0 Then
        DAOExcluirColuna = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAOExcluirColuna - " & Err.Description
    End If
End Function
Public Function DAOCriarTabela(BaseDeDados As String, _
                               NomeDaTabela As String, _
                               NomeDaColuna As String, _
                               Optional TipoDaColuna As DAO.DataTypeEnum = DAO.dbText, _
                               Optional TamDaColuna As Byte = 255, _
                               Optional Usuário As String, _
                               Optional Senha As String, _
                               Optional ModoExclusivo As Boolean = False) As Boolean
    On Error GoTo ErroCls
    Err.Clear
    If Trim(BaseDeDados) = "" _
        Or Trim(NomeDaTabela) = "" _
        Or Trim(NomeDaColuna) = "" Then Err.Raise 449     'erro Argumento não opcional
    
    Senha = "PWR=" & Senha & ";"
    Usuário = "UID=" & Usuário & ";"
    
    Dim db As Database
    Set db = OpenDatabase(BaseDeDados, ModoExclusivo, , Usuário & Senha)
    Dim td As TableDef
    Dim f As Field

    Set td = New TableDef
    Set f = New Field
    f.Name = NomeDaColuna                                    'define nome do campo
    f.Type = TipoDaColuna                                    'define o tipo do campo
    If TipoDaColuna = DAO.dbText Then f.Size = TamDaColuna    'define o tamanho do campo no caso de campo texto
    td.Name = NomeDaTabela                                  'define o nome da tabela
    td.Fields.Append f                                      'adiciona o campo a tabela
    db.TableDefs.Append td
    db.Close
    
ErroCls:
    If Err.Number = 0 Then
        DAOCriarTabela = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.DAOCriarTabela - " & Err.Description
    End If
End Function

Function ExecutarComando(Commando As String, Optional Janela As VbAppWinStyle = vbMinimizedNoFocus)
    On Error GoTo ErroCls
    Err.Clear
    If Trim(Comando) = "" Then Err.Raise 449        'erro Argumento não opcional
        
    Shell Commando, Janela
    
ErroCls:
    If Err.Number = 0 Then
        ExecutarComando = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "clManut.ExecutarComando - " & Err.Description
    End If
End Function

