VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUnZipMil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum ZNívelDeMensagem
    All = 0
    Less = 1
    NoMessages = 2
End Enum
Public Enum ZExtractType
    Extract = 0
    ListContents = 1
End Enum
Public Enum ZPrivilege
    Ignore = 0
    ACL = 1
    Privileges = 2
End Enum

Private miExtrairRecente     As Integer  ' 1 = Extract Only Newer, Else 0
Private miSpaceUnderScore  As Integer  ' 1 = Convert Space To Underscore, Else 0
Private miProntoRegravar  As Integer  ' 1 = Prompt To Overwrite Required, Else 0
Private miQuiet            As ZNívelDeMensagem  ' 2 = No Messages, 1 = Less, 0 = All
Private miWriteStdOut      As Integer  ' 1 = Write To Stdout, Else 0
Private miTestarZip          As Integer  ' 1 = Test Zip File, Else 0
Private miListarExtração      As ZExtractType  ' 0 = Extract, 1 = List Contents
Private miExtrairSóRecentes As Integer  ' 1 = Extract Only Newer, Else 0
Private miExibirComentários   As Integer  ' 1 = Display Zip File Comment, Else 0
Private miHonorDirectories As Integer  ' 1 = Honor Directories, Else 0
Private miSobregravarArquivos   As Integer  ' 1 = Overwrite Files, Else 0
Private miConverterCR_CRLF   As Integer  ' 1 = Convert CR To CRLF, Else 0
Private miVerbose          As Integer  ' 1 = Zip Info Verbose
Private miSensibilidadeCaracter  As Integer  ' 1 = Case Insensitivity, 0 = Case Sensitivity
Private miPrivilege        As ZPrivilege  ' 1 = ACL, 2 = Privileges, Else 0
Private msNomeDoZip      As String   ' The Zip File Name
Private msPastaExtração       As String   ' Extraction Directory, Null If Current Directory

Public Property Get ExtrairRecente() As Boolean
    ExtrairRecente = miExtrairRecente = 1
End Property

Public Property Let ExtrairRecente(ByVal bExtrairRecente As Boolean)
    miExtrairRecente = IIf(bExtrairRecente, 1, 0)
End Property

Public Property Get SpaceToUnderScore() As Boolean
    SpaceToUnderScore = miSpaceUnderScore = 1
End Property

Public Property Let SpaceToUnderScore(ByVal bConvert As Boolean)
    miSpaceUnderScore = IIf(bConvert, 1, 0)
End Property

Public Property Get ProntoRegravar() As Boolean
    ProntoRegravar = miProntoRegravar = 1
End Property

Public Property Let ProntoRegravar(ByVal bPrompt As Boolean)
    miProntoRegravar = IIf(bPrompt, 1, 0)
End Property

Public Property Get NívelDeMensagem() As ZNívelDeMensagem
    NívelDeMensagem = miQuiet
End Property

Public Property Let NívelDeMensagem(ByVal iLevel As ZNívelDeMensagem)
    miQuiet = iLevel
End Property

Public Property Get WriteToStdOut() As Boolean
    WriteToStdOut = miWriteStdOut = 1
End Property

Public Property Let WriteToStdOut(ByVal bWrite As Boolean)
    miWriteStdOut = IIf(bWrite, 1, 0)
End Property

Public Property Get TestarZip() As Boolean
    TestarZip = miTestarZip = 1
End Property

Public Property Let TestarZip(ByVal bTest As Boolean)
    miTestarZip = IIf(bTest, 1, 0)
End Property

Public Property Get ListarExtração() As ZExtractType
    ListarExtração = miListarExtração
End Property

Public Property Let ListarExtração(ByVal zExType As ZExtractType)
    miListarExtração = zExType
End Property

Public Property Get ExtrairSóRecentes() As Boolean
    ExtrairSóRecentes = miExtrairSóRecentes = 1
End Property

Public Property Let ExtrairSóRecentes(ByVal bOnlyNewer As Boolean)
    miExtrairSóRecentes = IIf(bOnlyNewer, 1, 0)
End Property

Public Property Get ExibirComentários() As Boolean
    ExibirComentários = miExibirComentários = 1
End Property

Public Property Let ExibirComentários(ByVal bDisplay As Boolean)
    miExibirComentários = IIf(bDisplay, 1, 0)
End Property

Public Property Get HonorDirectories() As Boolean
    HonorDirectories = miHonorDirectories = 1
End Property

Public Property Let HonorDirectories(ByVal bHonor As Boolean)
    miHonorDirectories = IIf(bHonor, 1, 0)
End Property

Public Property Get SobregravarArquivos() As Boolean
    SobregravarArquivos = miSobregravarArquivos = 1
End Property

Public Property Let SobregravarArquivos(ByVal bOverWrite As Boolean)
    miSobregravarArquivos = IIf(bOverWrite, 1, 0)
End Property

Public Property Get ConverterCRparaCRLF() As Boolean
    ConverterCRparaCRLF = miConverterCR_CRLF = 1
End Property

Public Property Let ConverterCRparaCRLF(ByVal bConvert As Boolean)
    miConverterCR_CRLF = IIf(bConvert, 1, 0)
End Property

Public Property Get Verbose() As Boolean
    Verbose = miVerbose = 1
End Property

Public Property Let Verbose(ByVal bVerbose As Boolean)
    miVerbose = IIf(bVerbose, 1, 0)
End Property

Public Property Get SensívelCaracter() As Boolean
    SensívelCaracter = miSensibilidadeCaracter = 1
End Property

Public Property Let SensívelCaracter(ByVal bSensívelCaracter As Boolean)
    miSensibilidadeCaracter = IIf(bSensívelCaracter, 1, 0)
End Property

Public Property Get Privilege() As ZPrivilege
    Privilege = miPrivilege
End Property

Public Property Let Privilege(ByVal zPriv As ZPrivilege)
    miPrivilege = zPriv
End Property

Public Property Get NomeDoZip() As String
    NomeDoZip = msNomeDoZip
End Property

Public Property Let NomeDoZip(ByVal sNomeDoZip As String)
    msNomeDoZip = sNomeDoZip
End Property

Public Property Get PastaExtração() As String
    PastaExtração = msPastaExtração
End Property

Public Property Let PastaExtração(ByVal sPastaExtração As String)
    msPastaExtração = sPastaExtração
End Property

Public Function UnZipar(Optional sNomeDoZip As String, _
    Optional sPastaExtração As String) As Long
    
On Error GoTo vbErrorHandler

    Dim lRet As Long
    
    If Len(sNomeDoZip) > 0 Then
        msNomeDoZip = sNomeDoZip
    End If
    
    If Len(sPastaExtração) > 0 Then
        msPastaExtração = sPastaExtração
    End If
    
    
    lRet = VBUnzip(msNomeDoZip, msPastaExtração, miExtrairRecente, _
        miSpaceUnderScore, miProntoRegravar, CInt(miQuiet), _
        miWriteStdOut, miTestarZip, CInt(miListarExtração), _
        miExtrairSóRecentes, miExibirComentários, miHonorDirectories, _
        miSobregravarArquivos, miConverterCR_CRLF, miVerbose, _
        miSensibilidadeCaracter, CInt(miPrivilege))
    
    UnZipar = lRet
    
    Exit Function

vbErrorHandler:
    Err.Raise Err.Number, "UnZipMil::UnZipar", Err.Description

End Function

Private Sub Class_Initialize()
    miExtrairRecente = 0
    miSpaceUnderScore = 0
    miProntoRegravar = 0
    miQuiet = NoMessages
    miWriteStdOut = 0
    miTestarZip = 0
    miListarExtração = Extract
    miExtrairSóRecentes = 0
    miExibirComentários = 0
    miHonorDirectories = 1
    miSobregravarArquivos = 1
    miConverterCR_CRLF = 0
    miVerbose = 0
    miSensibilidadeCaracter = 1
    miPrivilege = Ignore
End Sub

Public Function ÚltimaMensagem() As String
    ÚltimaMensagem = msOutput
End Function



