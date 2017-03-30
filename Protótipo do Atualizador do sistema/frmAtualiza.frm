VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Atualiza Estação Onlytech - Beta"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   5355
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDestino 
      Enabled         =   0   'False
      Height          =   360
      Left            =   150
      TabIndex        =   8
      Top             =   2100
      Width           =   7125
   End
   Begin VB.CommandButton cmdAtualizar 
      Caption         =   "Atualizar"
      Height          =   525
      Left            =   6030
      TabIndex        =   6
      Top             =   4680
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Height          =   1980
      Left            =   120
      TabIndex        =   5
      Top             =   2610
      Width           =   7155
   End
   Begin VB.TextBox txtAtualizando 
      Enabled         =   0   'False
      Height          =   360
      Left            =   150
      TabIndex        =   4
      Top             =   1320
      Width           =   7125
   End
   Begin VB.TextBox txtCaminho_Aplicacao 
      Enabled         =   0   'False
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   7125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Destino . . . "
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Atualizando arquivos . . ."
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   990
      Width           =   2130
   End
   Begin VB.Label lblINI_OK 
      AutoSize        =   -1  'True
      Caption         =   "Caminho da Aplicação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   4500
      TabIndex        =   2
      Top             =   210
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Caminho da Aplicação"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   1890
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FO_DELETE = &H3
Private Const GCT_INVALID = &H0
Private Const GCT_LFNCHAR = &H1
Private Const GCT_SEPARATOR = &H8
Private Const GCT_SHORTCHAR = &H2
Private Const GCT_WILD = &H4

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathFindOnPath Lib "shlwapi.dll" Alias "PathFindOnPathA" (ByVal pszPath As String, ByVal ppszOtherDirs As String) As Boolean
Private Declare Function PathGetCharType Lib "shlwapi.dll" Alias "PathGetCharTypeA" (ByVal ch As Byte) As Long
Private Declare Function PathGetDriveNumber Lib "shlwapi.dll" Alias "PathGetDriveNumberA" (ByVal pszPath As String) As Long

Public Caminho_Aplicacao As String
Public strLinha As String

Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Function Abrir_nome_cliente_registro(Aplicacao As String, Form As Object) As String

      On Error GoTo erro

      Dim Registro As New DLLSystemManager.Registro

      Caminho_Aplicacao = Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\RPT", "Caminho")
      Caminho_Aplicacao = Mid(Caminho_Aplicacao, 1, Len(Caminho_Aplicacao) - 3)

      Exit Function

erro:

   Call erro.erro(Form, Aplicacao, "Funções Gerais")
   Exit Function

End Function

Private Sub cmdAtualizar_Click()

    ' Abrindo arquivo ONLYTECH.INI
    intNome_Arquivo = FreeFile
    Open Caminho_Aplicacao + "onlytech.ini" For Input As #1
    
    ' Loop do arquivo INI
    Do While Not EOF(1)
    
        Line Input #1, strLinha
        
        ' Testando o final da arquivo ONLYTECH.INI
        If Trim(strLinha) = "" Then
            Close #1
            Screen.MousePointer = vbDefault
            MsgBox "Atualização finalizada com sucesso", vbInformation
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        path = Mid(strLinha, 1, Len(strLinha) - 6)
        SearchStr = Mid(strLinha, Len(strLinha) - 4, 5)
        
        ' Localizando os arquivos a serem atualizados
        
        Dim FileName As String ' Walking filename variable...
        Dim DirName As String ' SubDirectory Name
        Dim dirNames() As String ' Buffer for directory name entries
        Dim nDir As Integer ' Number of directories in this path
        Dim I As Integer ' For-loop counter...
        Dim hSearch As Long ' Search Handle
        Dim WFD As WIN32_FIND_DATA
        Dim Cont As Integer
        Dim strDestino As String
        Dim Ft1 As FILETIME
        Dim dtData_Origem As FILETIME
        Dim dtData_Destino As FILETIME
        Dim SysTime As SYSTEMTIME
        Dim lngHandle As Long
        Dim dblData_Origem As Double
        Dim dblData_Destino As Double
        
        If Right(path, 1) <> "\" Then path = path & "\"
            ' Search for subdirectories.
            nDir = 0
            ReDim dirNames(nDir)
            Cont = True
            hSearch = FindFirstFile(path & "*", WFD)
            ' Walk through this directory and sum file sizes.
            hSearch = FindFirstFile(path & SearchStr, WFD)
            Cont = True
            If hSearch <> INVALID_HANDLE_VALUE Then
                While Cont
                    FileName = StripNulls(WFD.cFileName)
                    If (FileName <> ".") And (FileName <> "..") Then
                        FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                        FileCount = FileCount + 1
                        txtAtualizando.Text = path & FileName
                        ' Montando a string com o caminho para atualizar
                        Select Case Mid(SearchStr, 3, 3)
                                Case "EXE"
                                    strDestino = Caminho_Aplicacao + "EXE\" + FileName
                                Case "RPT"
                                    strDestino = Caminho_Aplicacao + "RPT\" + FileName
                                Case "JPG"
                                    strDestino = Caminho_Aplicacao + "IMG\" + FileName
                                Case "GIF"
                                    strDestino = Caminho_Aplicacao + "IMG\" + FileName
                                Case "BMP"
                                    strDestino = Caminho_Aplicacao + "IMG\" + FileName
                        End Select
                        ' Verificando se o arquivo existe na estação
                        If CBool(PathFileExists(strDestino)) = False Then
                            txtDestino.Text = strDestino
                            CopyFile txtAtualizando.Text, strDestino, 0
                            ' Carregando o grid com os arquivos q foram atualizados
                            List1.AddItem strDestino
                        Else
                            ' Data dos arquivos de origem
                            lngHandle = CreateFile(txtAtualizando.Text, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
                            GetFileTime lngHandle, Ft1, Ft1, dtData_Origem
                            FileTimeToLocalFileTime dtData_Origem, Ft1
                            FileTimeToSystemTime Ft1, SysTime
                            dblData_Origem = CDbl(LTrim(Str$(SysTime.wYear)) + Replace(Str$(SysTime.wMonth), " ", "0") + Replace(Str$(SysTime.wDay), " ", "0"))
                            CloseHandle lngHandle ' axl
                            ' Data dos arquivos de destino
                            lngHandle = CreateFile(strDestino, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
                            CloseHandle lngHandle ' axl
                            GetFileTime lngHandle, Ft1, Ft1, dtData_Destino
                            FileTimeToLocalFileTime dtData_Destino, Ft1
                            FileTimeToSystemTime Ft1, SysTime
                            dblData_Destino = CDbl(LTrim(Str$(SysTime.wYear)) + Replace(Str$(SysTime.wMonth), " ", "0") + Replace(Str$(SysTime.wDay), " ", "0"))
                            If CDate(strData_Destino) < CDate(strData_Origem) Then
                                CopyFile txtAtualizando.Text, strDestino, 0
                                List1.AddItem strDestino
                            End If
                        End If
                        DoEvents
                        txtAtualizando.Text = Empty
                        txtDestino.Text = Empty
                        strDestino = Empty
                    End If
                    Cont = FindNextFile(hSearch, WFD) ' Get next file
                Wend
                Cont = FindClose(hSearch)
            End If
        ' If there are sub-directories...
        If nDir > 0 Then
            ' Recursively walk into them...
            For I = 0 To nDir - 1
                FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(I) & "\", SearchStr, FileCount, DirCount)
            Next I
        End If
    Loop
    
End Sub
Private Sub Form_Load()

    Call Abrir_nome_cliente_registro("Otica", Me)
    
    txtCaminho_Aplicacao.Text = Caminho_Aplicacao
    
    If CBool(PathFileExists(Caminho_Aplicacao + "onlytech.ini")) = False Then
        lblINI_OK.Caption = "Arquivo .INI não encontrado"
        MsgBox "Arquivo ONLYTECH.INI não encontrado na pasta de instalação", vbCritical
        End
    End If
    
    lblINI_OK.Caption = "Arquivo .INI OK"
    
    List1.Clear

End Sub
