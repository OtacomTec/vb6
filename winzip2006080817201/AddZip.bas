Attribute VB_Name = "modAddZip"
Option Explicit
' Visual Basic declares file for
'
'     azip32.dll    addZIP 32-bit compression library
'

' Function declarations
Declare Function addZIP Lib "azip32.dll" () As Integer
Declare Function addZIP_Abort Lib "azip32.dll" (ByVal bFlag As Integer) As Integer
Declare Function addZIP_ArchiveName Lib "azip32.dll" (ByVal lpStr As String) As Integer
Declare Function addZIP_BuildSFX Lib "azip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addZIP_ClearAttributes Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_Comment Lib "azip32.dll" (ByVal lpStr As String) As Integer
Declare Function addZIP_Delete Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_DeleteComment Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_DisplayComment Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_Encrypt Lib "azip32.dll" (ByVal lpStr As String) As Integer
Declare Function addZIP_Exclude Lib "azip32.dll" (ByVal lpStr As String) As Integer
Declare Function addZIP_ExcludeListFile Lib "azip32.dll" (ByVal lpStr As String) As Integer
Declare Function addZIP_GetLastError Lib "azip32.dll" () As Integer
Declare Function addZIP_GetLastWarning Lib "azip32.dll" () As Integer
Declare Function addZIP_Include Lib "azip32.dll" (ByVal lpStr As String) As Integer
Declare Function addZIP_IncludeArchive Lib "azip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addZIP_IncludeDirectoryEntries Lib "azip32.dll" (ByVal flag As Integer) As Integer
Declare Function addZIP_IncludeFilesNewer Lib "azip32.dll" (ByVal DateVal As String) As Integer
Declare Function addZIP_IncludeFilesOlder Lib "azip32.dll" (ByVal DateVal As String) As Integer
Declare Function addZIP_IncludeHidden Lib "azip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addZIP_IncludeListFile Lib "azip32.dll" (ByVal lpStr As String) As Integer
Declare Function addZIP_IncludeReadOnly Lib "azip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addZIP_IncludeSystem Lib "azip32.dll" (ByVal iFlag As Integer) As Integer
Declare Sub addZIP_Initialise Lib "azip32.dll" ()
Declare Function addZIP_InstallCallback Lib "azip32.dll" (ByVal cbFunction As Long) As Integer
Declare Function addZIP_Overwrite Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_Recurse Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_Register Lib "azip32.dll" (ByVal lpStr As String, ByVal Uint32 As Long) As Integer
Declare Function addZIP_SaveAttributes Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_SaveRelativeTo Lib "azip32.dll" (ByVal szPath As String) As Integer
Declare Function addZIP_SaveStructure Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_SetArchiveDate Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_SetCompressionLevel Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_SetParentWindowHandle Lib "azip32.dll" (ByVal hwnd As Long) As Integer
Declare Function addZIP_SetTempDrive Lib "azip32.dll" (ByVal lpStr As String) As Integer
Declare Function addZIP_SetWindowHandle Lib "azip32.dll" (ByVal hwnd As Long) As Integer
Declare Function addZIP_Span Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_Store Lib "azip32.dll" (ByVal lpStr As String) As Integer
Declare Function addZIP_UseLFN Lib "azip32.dll" (ByVal Int16 As Integer) As Integer
Declare Function addZIP_View Lib "azip32.dll" (ByVal Int16 As Integer) As Integer

' Visual Basic declares file for
'
'     aunzip32.dll  addUNZIP 32-bit decompression library
'

Declare Function addUNZIP Lib "aunzip32.dll" () As Long
Declare Function addUNZIP_Abort Lib "aunzip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addUNZIP_ArchiveName Lib "aunzip32.dll" (ByVal filename As String) As Integer
Declare Function addUNZIP_Decrypt Lib "aunzip32.dll" (ByVal cPassword As String) As Integer
Declare Function addUNZIP_DisplayComment Lib "aunzip32.dll" (ByVal bFlag As Integer) As Integer
Declare Function addUNZIP_Exclude Lib "aunzip32.dll" (ByVal files As String) As Integer
Declare Function addUNZIP_ExcludeListFile Lib "aunzip32.dll" (ByVal cFile As String) As Integer
Declare Function addUNZIP_ExtractTo Lib "aunzip32.dll" (ByVal cPath As String) As Integer
Declare Function addUNZIP_Freshen Lib "aunzip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addUNZIP_GetLastError Lib "aunzip32.dll" () As Integer
Declare Function addUNZIP_GetLastWarning Lib "aunzip32.dll" () As Integer
Declare Function addUNZIP_Include Lib "aunzip32.dll" (ByVal files As String) As Integer
Declare Function addUNZIP_IncludeListFile Lib "aunzip32.dll" (ByVal cFile As String) As Integer
Declare Sub addUNZIP_Initialise Lib "aunzip32.dll" ()
Declare Function addUNZIP_InstallCallback Lib "aunzip32.dll" (ByVal fn As Long) As Integer
Declare Function addUNZIP_Overwrite Lib "aunzip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addUNZIP_Register Lib "aunzip32.dll" (ByVal cName As String, ByVal iNumber As Long) As Integer
Declare Function addUNZIP_ResetDefaults Lib "aunzip32.dll" ()
Declare Function addUNZIP_RestoreAttributes Lib "aunzip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addUNZIP_RestoreStructure Lib "aunzip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addUNZIP_SetParentWindowHandle Lib "aunzip32.dll" (ByVal hwnd As Long) As Integer
Declare Function addUNZIP_SetWindowHandle Lib "aunzip32.dll" (ByVal hwnd As Long) As Integer
Declare Function addUNZIP_Test Lib "aunzip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addUNZIP_ToMemory Lib "aunzip32.dll" (ByVal lpStr As String, ByVal Uint32 As Long) As Integer
Declare Function addUNZIP_Update Lib "aunzip32.dll" (ByVal iFlag As Integer) As Integer
Declare Function addUNZIP_View Lib "aunzip32.dll" (ByVal bFlag As Integer) As Integer


' Visual Basic constants file for
'
'     azip32.dll        addZIP 32-bit compression library
'     aunzip32.dll      addUNZIP 32-bit compression library
'

' Function declarations
'  constants for addZIP_SetCompressionLevel(...)

Global Const azCOMPRESSION_MAXIMUM = &H3
Global Const azCOMPRESSION_MINIMUM = &H1
Global Const azCOMPRESSION_NONE = &H0
Global Const azCOMPRESSION_NORMAL = &H2

' constants for addZIP_SaveStructure(...)
Global Const azSTRUCTURE_ABSOLUTE = &H2
Global Const azSTRUCTURE_NONE = &H0
Global Const azSTRUCTURE_RELATIVE = &H1

' constants for addZIP_Overwrite(...)
' constants for addUNZIP_Overwrite(...)
Global Const azOVERWRITE_ALL = &HB
Global Const azOVERWRITE_NONE = &HC
Global Const azOVERWRITE_QUERY = &HA

' constants for addZIP_SetArchiveDate()
Global Const DATE_NEWEST = &H3
Global Const DATE_OLDEST = &H2
Global Const DATE_ORIGINAL = &H0
Global Const DATE_TODAY = &H1

' constants for addZIP_IncludeXXX attribute functions
Global Const azNEVER = &H0       ' files must never have this attribute set
Global Const azALWAYS = &HFF ' files may or may not have this attribute set
Global Const azYES = &H1         ' files must always have this attribute set

'  constants for addZIP_ClearAttributes(...)
' constants for addUNZIP_RestoreAttributes(...)
Global Const azATTR_NONE = 0
Global Const azATTR_READONLY = 1
Global Const azATTR_HIDDEN = 2
Global Const azATTR_SYSTEM = 4
Global Const azATTR_ARCHIVE = 32
Global Const azATTR_ALL = 39

' constants used in messages to identify libraries
Global Const azLIBRARY_ADDZIP = 0
Global Const azLIBRARY_ADDUNZIP = 1

' 'messages' used to provide information to the calling program
Global Const AM_SEARCHING = &HA
Global Const AM_ZIPCOMMENT = &HB
Global Const AM_ZIPPING = &HC
Global Const AM_ZIPPED = &HD
Global Const AM_UNZIPPING = &HE
Global Const AM_UNZIPPED = &HF
Global Const AM_TESTING = &H10
Global Const AM_TESTED = &H11
Global Const AM_DELETING = &H12
Global Const AM_DELETED = &H13
Global Const AM_DISKCHANGE = &H14
Global Const AM_VIEW = &H15
Global Const AM_ERROR = &H16
Global Const AM_WARNING = &H17
Global Const AM_QUERYOVERWRITE = &H18
Global Const AM_COPYING = &H19
Global Const AM_COPIED = &H1A
Global Const AM_ABORT = &HFF

' Constants for whether file is encrypted or not in AM_VIEW
Global Const azFT_ENCRYPTED = &H1
Global Const azFT_NOT_ENCRYPTED = &H0

' Constants for whether file is text or binary in AM_VIEW
Global Const azFT_BINARY = &H1
Global Const azFT_TEXT = &H0

' Constants for compression method in AM_VIEW
Global Const azCM_DEFLATED_FAST = &H52
Global Const azCM_DEFLATED_MAXIMUM = &H51
Global Const azCM_DEFLATED_NORMAL = &H50
Global Const azCM_DEFLATED_SUPERFAST = &H53
Global Const azCM_IMPLODED = &H3C
Global Const azCM_NONE = &H0
Global Const azCM_REDUCED_1 = &H14
Global Const azCM_REDUCED_2 = &H1E
Global Const azCM_REDUCED_3 = &H28
Global Const azCM_REDUCED_4 = &H32
Global Const azCM_SHRUNK = &HA
Global Const azCM_TOKENISED = &H46
Global Const azCM_UNKNOWN = &HFF

' Constants used in returning from a AM_QUERYOVERWRITE message
Global Const azOW_NO = &H2
Global Const azOW_NO_TO_ALL = &H3
Global Const azOW_YES = &H0
Global Const azOW_YES_TO_ALL = &H1
' Apenas para retorno das funções
Public Z As Integer


'-----------------------
'Funções do addZIP
'-----------------------

Function GetAction(cFrom As String) As Integer
    GetAction = Val(GetPiece(cFrom, "|", 2))
End Function

Function GetFileCompressedSize(cFrom As String) As Long
    GetFileCompressedSize = Val(GetPiece(cFrom, "|", 6))
End Function

Function GetFileCompressionRatio(cFrom As String) As Integer
    GetFileCompressionRatio = Val(GetPiece(cFrom, "|", 7))
End Function

Function GetFileName(cFrom As String) As String
    GetFileName = GetPiece(cFrom, "|", 4)
End Function

Function GetFileOriginalSize(cFrom As String) As Long
    GetFileOriginalSize = Val(GetPiece(cFrom, "|", 5))
End Function

Function GetPercentComplete(cFrom As String) As Integer
    GetPercentComplete = Val(GetPiece(cFrom, "|", 7))
End Function

Function GetPiece(from As String, delim As String, Index As Integer) As String
'Tipo de ação retornada pelo arquivo ou compactação
    Dim Temp$
    Dim Count As Integer
    Dim Where As Integer
    
    Temp$ = from & delim
    Where = InStr(Temp$, delim)
    Count = 0
    Do While (Where > 0)
        Count = Count + 1
        If (Count = Index) Then
            GetPiece = Left$(Temp$, Where - 1)
            Exit Function
        End If
        Temp$ = Right$(Temp$, Len(Temp$) - Where)
        Where = InStr(Temp$, delim)
    Loop
    If (Count = 0) Then
        GetPiece = from
    Else
        GetPiece = ""
    End If
End Function

'---------------------------------------------------------------------
'Rotinas para o addZIP
'---------------------------------------------------------------------
Sub Compacta(cArqCompactado As String, cArq As String)
    'Compacta um ou mais arquivos no formato WinZip
    Z = addZIP_SetCompressionLevel(azCOMPRESSION_MAXIMUM)
    Z = addZIP_SaveStructure(azCM_NONE) 'StoreFullPathName - azSTRUCTURE_ABSOLUTE
    Z = addZIP_Include(cArq)
    Z = addZIP_ArchiveName(cArqCompactado)
    'Z = addZIP_Delete(DeletarOrig)
    Z = addZIP()
End Sub
          
Sub DesCompacta(cArqCompactado As String, cNomeArq As String, ExtrairPara As String, MontaDir As Boolean)
'Descompacta um ou mais arquivos no formato WinZip
    Z = addUNZIP_Overwrite(azOVERWRITE_ALL)
    Z = addUNZIP_ArchiveName(cArqCompactado)
    Z = addUNZIP_Include(cNomeArq)
    Z = addUNZIP_ExtractTo(ExtrairPara)
    Z = addUNZIP_RestoreStructure(MontaDir)
    Z = addUNZIP()
End Sub

Sub ListaConteudoArquivo(cArquivo As String)
'Lista o conteudo de um arquivo zipado.
    Z = addZIP_ArchiveName(cArquivo)
    Z = addZIP_View(True)
    Z = addZIP()
End Sub

Sub InicializaZip(F As Form, TextoZip As Control)
    On Error GoTo InializaZipError
    ' Inicializa as bibliotecas do addZIP
    ' É necessário um form e um TextBox
     addZIP_Initialise
     addUNZIP_Initialise
     Z = addZIP_SetParentWindowHandle(F.hwnd)
     Z = addUNZIP_SetParentWindowHandle(F.hwnd)
     Z = addZIP_SetWindowHandle(TextoZip.hwnd)
     Z = addUNZIP_SetWindowHandle(TextoZip.hwnd)
    Exit Sub
InializaZipError:
    MsgBox "Erro inicializando bibliotecas ZIP"
End Sub

Function TipoAção(nTipo As Long) As String
Select Case nTipo
    Case AM_SEARCHING: TipoAção = "Procurando"
    Case AM_ZIPCOMMENT: TipoAção = "Comentário"
    Case AM_ZIPPING: TipoAção = "Zipando"
    Case AM_ZIPPED: TipoAção = "Zipado"
    Case AM_UNZIPPING: TipoAção = "Deszipando"
    Case AM_UNZIPPED: TipoAção = "Deszipado"
    Case AM_TESTING: TipoAção = "Testando"
    Case AM_TESTED: TipoAção = "Testado"
    Case AM_DELETING: TipoAção = "Deletando"
    Case AM_DELETED: TipoAção = "Deletado"
    Case AM_DISKCHANGE: TipoAção = "Troca Disco"
    Case AM_VIEW: TipoAção = "Visualizar"
    Case AM_ERROR: TipoAção = "Erro"
    Case AM_WARNING: TipoAção = "Aviso"
    Case AM_QUERYOVERWRITE: TipoAção = "Sobrescrever"
    Case AM_COPYING: TipoAção = "Copiando"
    Case AM_COPIED: TipoAção = "Copiado"
    Case AM_ABORT: TipoAção = "Abortando"
    Case Else
        TipoAção = "-=Desconhecido=-"
End Select

End Function
