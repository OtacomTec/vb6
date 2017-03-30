VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMclDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------------------
'Codigo Programa: GMclDIR
'Descr.Programa.: Classe de Listagem de Aquivos e Pastas
'Analista.......:
'Programador....: Luis Henrique Borges
'Data Criação...: 00 Abr 2001
'Data Alteração.:
'--------------------------------------------------------------------------------------------
Private Const BRM_MSG1 = "Ocorreu o erro número "
Private mstrDir() As String
Private mstrDirOcul() As String
Private mstrDirSis() As String

Private mstrArq() As String
Private mstrArqLeitura() As String
Private mstrArqOculto() As String
Private mstrArqSistema() As String
Private mstrArqLeitOcul() As String
Private mstrArqLeitSist() As String
Private mstrArqSistOcul() As String
Private mstrArqSistLeitOcul() As String
Private mstrCuringa() As String

Private mstrVolume As String
Private mstrCaminho As String

Private miTotal As Variant
Private miTotalArq As Variant
Private miTotalArqNormal As Variant
Private miTotalPastasTodas As Variant
Private miTotalPastas As Variant
Private miTotalPastasNormais As Variant

Private strPastaCorrente As String
'local variable(s) to hold property value(s)
Private mvarMensagemDeErro As String 'local copy

Public Property Get MensagemDeErro() As String
    MensagemDeErro = mvarMensagemDeErro
End Property

Public Function ExcluirArquivo(strCaminhodoArquivo As String) As Boolean
    Err.Clear
    On Error GoTo ErroCls
    If Trim(strCaminhodoArquivo) = "" Then Err.Raise 449
       
    Kill strCaminhodoArquivo
    
ErroCls:

    If Err.Number = 0 Then
        ExcluirArquivo = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "GMclDir.CopiarArquivo - " & Err.Description
    End If
    
End Function

Public Function RenomearArquivo(strNomeAtual As String, _
                           strNomeNovo As String, _
                           Optional strPasta As String) As Boolean
    Err.Clear
    On Error GoTo ErroCls
    
    If Trim(strNomeAtual) = "" _
       Or Trim(strNomeNovo) = "" _
       Or Trim(strNomeAtual) = Trim(strNomeNovo) Then Err.Raise 449
       

    If strPasta = "" Then strPasta = App.Path
    
    If ValidaNome(strNomeAtual) = False Then Exit Function
    If ValidaNome(strNomeNovo) = False Then Exit Function
    
    If Right(strPasta, 1) <> "\" Then strPasta = strPasta & "\"
    If Dir(strPasta & strNomeAtual) = "" Or Dir(strPasta & strNomeNome) = "" Then Exit Function
   
    FileCopy strPasta & strNomeAtual, strPasta & strNomeNovo
     
ErroCls:
    If Err.Number = 0 Then
        RenomearArquivo = True
        Kill strPasta & strNomeAtual
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "GMclDir.CopiarArquivo - " & Err.Description
    End If
    
End Function

Public Function MoverArquivo(strCaminhoOrigem As String, _
                             strPastaDestino As String, _
                    Optional strNomeDestino As String, _
                    Optional boSobregravar As Boolean = False, _
                    Optional boMsgBox As Boolean = False) As Boolean
    Err.Clear
    On Error GoTo ErroCls
    
    If Trim(strCaminhoOrigem) = "" _
       Or Trim(strPastaDestino) = "" Then Err.Raise 449
       
    
    
    If Dir(strCaminhoOrigem) = "" Or Right(strCaminhoOrigem, 1) = "\" Then Exit Function
    If strNomeDestino <> "" Then
        If Right(strNomeDestino, 1) = "\" Then Exit Function
    End If
    
    If strPastaDestino <> "" Then
        If Not Right(strPastaDestino, 1) = "\" Then strPastaDestino = strPastaDestino & "\"
    End If
    
    
    rtn = InStrRev(strCaminhoOrigem, "\")
    strPasta = Mid(strCaminhoOrigem, 1, rtn)
    strArquivo = Mid(strCaminhoOrigem, rtn + 1, Len(strCaminhoOrigem) - rtn)
    
    Select Case boSobregravar
        Case True
            If strNomeDestino = "" Then
                FileCopy strCaminhoOrigem, strPastaDestino & strArquivo
            Else
                FileCopy strCaminhoOrigem, strPastaDestino & strNomeDestino
            End If
        Case False
            If strNomeDestino = "" Then
                If Dir(strPastaDestino & strArquivo) = "" Then
                    FileCopy strCaminhoOrigem, strPastaDestino & strArquivo
                Else
                    If boMsgBox = True Then
                        msg = "O Arquivo " & strArquivo & " "
                        msg = msg & "já existe na pasta de destino." & Chr(10)
                        msg = msg & "Deseja Sobregravar?"
                        rtn = MsgBox(msg, vbYesNo, "GMclDir")
                        If rtn = 6 Then
                            FileCopy strCaminhoOrigem, strPastaDestino & strArquivo
                        Else
                            mvarMensagemDeErro = "O Arquivo não foi movido. Operação Abortada pelo usuário."
                            Exit Function
                        End If
                    End If
                End If
            Else
                FileCopy strCaminhoOrigem, strPastaDestino & strNomeDestino
            End If
    End Select
    
ErroCls:
    If Err.Number = 0 Then
        MoverArquivo = True
        Kill strCaminhoOrigem
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "GMclDir.CopiarArquivo - " & Err.Description
    End If
End Function

Public Function CopiarArquivo(strOrigem As String, _
                     Optional boSobregravar As Boolean = False, _
                     Optional strDestino As String, _
                     Optional boMsgBox As Boolean = False) As Boolean
    'Este método copia um arquivo.
    'strOrigem      -> Caminho completo do arquivo de origem (ex 'c:\pasta\...\teste.txt')
    'strDestino     -> Caminho completo do arquivo a ser copiado (ex: 'c:\outrapasta\...\NovoArquivo.txt')
    'boSobregravar  -> True = sobregrava o arquivo sem perguntar <> False = pergunta se boMsgbox=true
    'boMsgbox       -> True = se boSobregravar=false pergunta antes <> se False não pergunta nada
    Err.Clear
    On Error GoTo ErroCls
    If Trim(strOrigem) = "" Then Err.Raise 449
    
    Dim rtn
                                    'VERIFICAR SE O DIRETÓRIO É VÁLIDO
    If Dir(strOrigem) = "" Or Right(strOrigem, 1) = "\" Or Right(strDestino, 1) = "\" Then
        Exit Function
    End If

     If Trim(strOrigem) = "" Then Err.Raise 449
    
    If strDestino = "" Then        'Sem pasta de destino deduz a mesma pasta
        rtn = InStrRev(strOrigem, "\")
        strPasta = Mid(strOrigem, 1, rtn)
        strArquivo = Mid(strOrigem, rtn + 1, Len(strOrigem) - rtn)
        i = 1
        Do While copia = False
            nr = Format(i, "##000")
            rtnstr = Dir(strPasta & "Cópia " & nr & " de " & strArquivo)
            If rtnstr = "" Then
                FileCopy strOrigem, strPasta & "Cópia " & nr & " de " & strArquivo
                copia = True
            End If
            i = i + 1
        Loop
    Else
        Select Case boSobregravar
            Case True
                FileCopy strOrigem, strDestino
            Case False
                If Dir(strDestino) = "" Then
                    FileCopy strOrigem, strDestino
                Else
                    If boMsgBox = True Then
                        msg = "O Arquivo " & Mid(strDestino, InStrRev(strDestino, "\") + 1, Len(strDestino) - (InStrRev(strDestino, "\"))) & " "
                        msg = msg & "já existe na pasta de destino." & Chr(10)
                        msg = msg & "Deseja Sobregravar?"
                        rtn = MsgBox(msg, vbYesNo, "GMclDir")
                        If rtn = 6 Then
                            FileCopy strOrigem, strDestino
                        End If
                    Else
                        FileCopy strOrigem, strDestino
                    End If
                End If
        End Select
    End If

ErroCls:
    If Err.Number = 0 Then
        CopiarArquivo = True
        mvarMensagemDeErro = ""
    Else
        mvarMensagemDeErro = BRM_MSG1 & Str(Err.Number) & " - " & "GMclDir.CopiarArquivo - " & Err.Description
    End If
End Function

Public Function MeuDir(Diretório As String) As String
    On Error GoTo ErroDocDir
    ReDim mstrDir(0)
    ReDim mstrDirOcul(0)
    ReDim mstrDirSis(0)
    
    ReDim mstrArq(0)
    ReDim mstrArqLeitura(0)
    ReDim mstrArqOculto(0)
    ReDim mstrArqSistema(0)
    ReDim mstrArqLeitOcul(0)
    ReDim mstrArqLeitSist(0)
    ReDim mstrArqSistOcul(0)
    ReDim mstrArqSistLeitOcul(0)
    ReDim mstrCuringa(0)
    
    Mypath = Diretório
    mstrCaminho = Mypath
    
    mstrVolume = Dir(Mypath, vbVolume)
        AchouBarra = False
    If Right(Mypath, 1) <> "\" Then
        NomeItem = Dir(Mypath)
        Do While NomeItem <> ""
            
            Do While AchouBarra = False 'a consulta é por arquivos em diretório corrente
                X = InStr(p + 1, Mypath, "\")
                If X <> 0 Then
                    p = X
                Else
                    AchouBarra = True
                End If
            Loop
            MyPath1 = Mid(Mypath, 1, p) & NomeItem
            mstrCaminho = Mid(Mypath, 1, p)
            If (GetAttr(MyPath1) And vbNormal) = vbNormal Then
                ReDim Preserve mstrCuringa(UBound(mstrCuringa()) + 1)
                mstrCuringa(UBound(mstrCuringa())) = NomeItem
            End If
           NomeItem = Dir  '(MyPath)
            If ÚltimoItem <> NomeItem Then
                ÚltimoItem = NomeItem
            Else
                Exit Function
            End If
             
        Loop
    End If
    
    'Define as Pastas ou Diretórios
    NomeItem = Dir(Mypath, vbDirectory)
    Do While NomeItem <> ""
        If NomeItem <> "." And NomeItem <> ".." Then
            If (GetAttr(Mypath & NomeItem) And vbDirectory) = vbDirectory Then
                ReDim Preserve mstrDir(UBound(mstrDir()) + 1)
                mstrDir(UBound(mstrDir())) = NomeItem
            End If
        End If
        NomeItem = Dir   ' Obtém a próxima entrada.
    Loop
    
    'Localiza Pastas Ocultas
    NomeItem = Dir(Mypath, vbDirectory + vbHidden)
    Do While NomeItem <> ""
        If NomeItem <> "." And NomeItem <> ".." Then
            If (GetAttr(Mypath & NomeItem) And vbDirectory + vbHidden) = vbDirectory + vbHidden Then
                ReDim Preserve mstrDirOcul(UBound(mstrDirOcul()) + 1)
                mstrDirOcul(UBound(mstrDirOcul())) = NomeItem
            End If
        End If
        NomeItem = Dir   ' Obtém a próxima entrada.
    Loop
    
    'Localiza Pastas de Sistema
    NomeItem = Dir(Mypath, vbDirectory + vbSystem)
    Do While NomeItem <> ""
        If NomeItem <> "." And NomeItem <> ".." Then
            If (GetAttr(Mypath & NomeItem) And vbDirectory + vbSystem) = vbDirectory + vbSystem Then
                ReDim Preserve mstrDirSis(UBound(mstrDirSis()) + 1)
                mstrDirSis(UBound(mstrDirSis())) = NomeItem
            End If
        End If
        NomeItem = Dir   ' Obtém a próxima entrada.
    Loop
    
    
    
    'Define os arquivos com atributo normal
    NomeItem = Dir(Mypath, vbNormal)
    Do While NomeItem <> ""
        If (GetAttr(Mypath & NomeItem) And vbNormal) = vbNormal Then
            ReDim Preserve mstrArq(UBound(mstrArq()) + 1)
            mstrArq(UBound(mstrArq())) = NomeItem
        End If
        NomeItem = Dir
    Loop
       
    
    NomeItem = Dir(Mypath, vbReadOnly) 'Somente Leitura
    Do While NomeItem <> ""
        If (GetAttr(Mypath & NomeItem) And vbReadOnly) = vbReadOnly Then
            ReDim Preserve mstrArqLeitura(UBound(mstrArqLeitura()) + 1)
            mstrArqLeitura(UBound(mstrArqLeitura())) = NomeItem
        End If
        NomeItem = Dir
    Loop
        
    NomeItem = Dir(Mypath, vbHidden) 'Somente Ocultos
    Do While NomeItem <> ""
        If (GetAttr(Mypath & NomeItem) And vbHidden) = vbHidden Then
            ReDim Preserve mstrArqOculto(UBound(mstrArqOculto()) + 1)
            mstrArqOculto(UBound(mstrArqOculto())) = NomeItem
        End If
        NomeItem = Dir
    Loop
            
    NomeItem = Dir(Mypath, vbSystem) 'Somente Sistema
    Do While NomeItem <> ""
        If (GetAttr(Mypath & NomeItem) And vbSystem) = vbSystem Then
            ReDim Preserve mstrArqSistema(UBound(mstrArqSistema()) + 1)
            mstrArqSistema(UBound(mstrArqSistema())) = NomeItem
        End If
        NomeItem = Dir
    Loop
    
    NomeItem = Dir(Mypath, vbReadOnly + vbHidden) 'Somente Leitura e Oculto
    Do While NomeItem <> ""
        If (GetAttr(Mypath & NomeItem) And vbReadOnly + vbHidden) = vbReadOnly + vbHidden Then
            ReDim Preserve mstrArqLeitOcul(UBound(mstrArqLeitOcul()) + 1)
            mstrArqLeitOcul(UBound(mstrArqLeitOcul())) = NomeItem
        End If
        NomeItem = Dir
    Loop
        
    NomeItem = Dir(Mypath, vbReadOnly + vbSystem) 'Somente Leitura e Oculto
    Do While NomeItem <> ""
        If (GetAttr(Mypath & NomeItem) And vbReadOnly + vbSystem) = vbReadOnly + vbSystem Then
            ReDim Preserve mstrArqLeitSist(UBound(mstrArqLeitSist()) + 1)
            mstrArqLeitSist(UBound(mstrArqLeitSist())) = NomeItem
        End If
        NomeItem = Dir
    Loop
        
    NomeItem = Dir(Mypath, vbSystem + vbHidden)   'Somente sistema e Oculto
    Do While NomeItem <> ""
        If (GetAttr(Mypath & NomeItem) And vbSystem + vbHidden) = vbSystem + vbHidden Then
            ReDim Preserve mstrArqSistOcul(UBound(mstrArqSistOcul()) + 1)
            mstrArqSistOcul(UBound(mstrArqSistOcul())) = NomeItem
        End If
        NomeItem = Dir
    Loop
    
    NomeItem = Dir(Mypath, vbSystem + vbHidden + vbReadOnly)   'Somente Leitura e Oculto
    Do While NomeItem <> ""
        If (GetAttr(Mypath & NomeItem) And vbSystem + vbHidden + vbReadOnly) = vbSystem + vbHidden + vbReadOnly Then
            ReDim Preserve mstrArqSistLeitOcul(UBound(mstrArqSistLeitOcul()) + 1)
            mstrArqSistLeitOcul(UBound(mstrArqSistLeitOcul())) = NomeItem
        End If
        NomeItem = Dir
    Loop
    
    
    '''''''' totais
    miTotalArq = UBound(mstrArq()) + _
               UBound(mstrArqLeitura()) + _
               (UBound(mstrArqOculto()) - UBound(mstrArqLeitOcul())) + _
               UBound(mstrArqSistema()) + _
               UBound(mstrArqLeitOcul()) + _
               UBound(mstrArqLeitSist()) + _
               UBound(mstrArqSistLeitOcul())
               
    miTotalArqNormal = UBound(mstrArq())
    
    miTotalPastas = UBound(mstrDir()) + _
                         UBound(mstrDirOcul()) + _
                         UBound(mstrDirSis())
    
    miTotalPastasNormais = UBound(mstrDir())
               
    miTotal = miTotalArq + miTotalPastas
    Exit Function
    
ErroDocDir:
    Select Case Err
        Case 53
            Exit Function
        Case 52
            Exit Function
        
    End Select
    Resume
    
End Function

Public Property Get ResulPesq() As Variant
    ResulPesq = mstrCuringa
End Property

Public Property Get TotalPastas() As Variant
    TotalPastas = miTotalPastas
End Property
Public Property Get TotalArq() As Variant
    TotalArq = miTotalArq
End Property
Public Property Get Total() As Variant
    Total = miTotal
End Property
Public Property Get TotalPastasSemAtributo() As Variant
    TotalPastasSemAtributo = miTotalPastasNormais
End Property
Public Property Get Caminho() As Variant
    Caminho = mstrCaminho
End Property
Public Property Get TotalArqNormal() As Variant
    TotalArqNormal = miTotalArqNormal
End Property
Public Property Get Arquivo() As Variant
    Arquivo = mstrArq
End Property
Public Property Get Diretório() As Variant
    Diretório = mstrDir
End Property
Public Property Get ArquivoLeitura() As Variant
    ArquivoLeitura = mstrArqLeitura
End Property

Public Property Get ArquivoOculto() As Variant
    ArquivoOculto = mstrArqOculto
End Property

Public Property Get ArquivoSistema() As Variant
    ArquivoSistema = mstrArqSistema
End Property

Public Property Get ArquivoLeitOcul() As Variant
    ArquivoLeitOcul = mstrArqLeitOcul
End Property

Public Property Get ArquivoLeitSist() As Variant
    ArquivoLeitSist = mstrArqLeitSist
End Property
Public Property Get ArquivoSistOcul() As Variant
    ArquivoSistOcul = mstrArqSistOcul
End Property

Public Property Get ArquivoSistLeitOcul() As Variant
    ArquivoSistLeitOcul = mstrArqSistLeitOcul
End Property

Public Property Get DiretórioOculto() As Variant
    DiretórioOculto = mstrDirOcul
End Property

Public Property Get DiretórioSistema() As Variant
    DiretórioSistema = mstrDirSis
End Property
Public Property Get Volume() As Variant
    Volume = mstrVolume
End Property

Private Function ValidaNome(strNome As String) As Boolean
    ValidaNome = False
    If InStr(strNome, "/") <> 0 Then Exit Function
    If InStr(strNome, "\") <> 0 Then Exit Function
    If InStr(strNome, ":") <> 0 Then Exit Function
    If InStr(strNome, "?") <> 0 Then Exit Function
    If InStr(strNome, "*") <> 0 Then Exit Function
        
    ValidaNome = True
End Function

