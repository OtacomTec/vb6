VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZipMil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum ZTranslate
    CRLFtoLF = 1
    LFtoCRLF = 2
End Enum
Private mCole��o As Collection
Private miIncluiSubPastas As Integer
Private msNomeDoZip As String
Private miEncript As Integer
Private miSistema As Integer
Private msPastaRaiz As String
Private miVerbose As Integer
Private miQuiet As Integer
Private miTranslateCRLF As ZTranslate
Private miUpdateZip As Integer

Private Sub Class_Initialize()
    Set mCole��o = New Collection    'Inicializa a Cole��o
'
' We have to add in a dummy file into the collection because
' the Zip routines fall over otherwise.
'
' I think this is a bug, but it's not documented anywhere
' on the InfoZip website.
'
' The Zip process *always* fails on the first file,
' regardless of whether it's a valid file or not!
'
    mCole��o.Add "querty", "querty"
    miEncript = 0
    miSistema = 0
    msPastaRaiz = "\"
    miQuiet = 0
    miUpdateZip = 0
    
End Sub

Private Sub Class_Terminate()
    Set mCole��o = Nothing
End Sub

Public Property Get IncluiSubPastas() As Boolean
    IncluiSubPastas = miIncluiSubPastas = 1
End Property

Public Property Let IncluiSubPastas(ByVal bRecurse As Boolean)
    miIncluiSubPastas = IIf(bRecurse, 1, 0)
End Property

Public Property Get NomeDoZip() As String
    NomeDoZip = msNomeDoZip
End Property

Public Property Let NomeDoZip(ByVal sNomeDoZip As String)
    msNomeDoZip = sNomeDoZip '& vbNullChar
End Property

Public Property Get Encriptar() As Boolean
    Encriptar = miEncript = 1
End Property

Public Property Let Encriptar(ByVal bEncript As Boolean)
    miEncript = IIf(bEncript, 1, 0)
End Property

Public Property Get IncluiArquivosDeSistema() As Boolean
    IncluiArquivosDeSistema = miSistema = 1
End Property

Public Property Let IncluiArquivosDeSistema(ByVal bIncluir As Boolean)
    miSistema = IIf(bIncluir, 1, 0)
End Property

Public Property Get N�meroArquivosNoZip() As Long
    If mCole��o Is Nothing Then
        N�meroArquivosNoZip = 0
    Else
        N�meroArquivosNoZip = mCole��o.Count - 1
    End If
End Property
Public Property Get PastaRaiz() As String
    PastaRaiz = msPastaRaiz
End Property

Public Property Let PastaRaiz(ByVal sPastaRz As String)
    msPastaRaiz = sPastaRz ' & vbNullChar
End Property

Public Property Get AtualizarZip() As Boolean
    AtualizarZip = miUpdateZip = 1
End Property

Public Property Let AtualizarZip(ByVal bAtualizar As Boolean)
    miUpdateZip = IIf(bAtualizar, 1, 0)
End Property

Public Function AdicionarArquivo(ByVal sNomeDoArquivo As String)
    Dim lCount As Long
    Dim sFile As String
    On Error Resume Next
    
    sFile = mCole��o.Item(sNomeDoArquivo)
    If Len(sFile) = 0 Then
        Err.Clear
        On Error GoTo 0
        mCole��o.Add sNomeDoArquivo, sNomeDoArquivo
    Else
        On Error GoTo 0
        Err.Raise vbObjectError + 2001, "CGZip::AdicionarArquivo", "File is already in Zip List"
    End If
    
End Function

Public Function RemoverArquivo(ByVal sNomeDoArquivo As String)
    Dim lCount As Long
    Dim sFile As String
    On Error Resume Next
    
    sFile = mCole��o.Item(sNomeDoArquivo)
    If Len(sFile) = 0 Then
        Err.Raise vbObjectError + 2002, "CGZip::RemoverArquivo", "File is not in Zip List"
    Else
        mCole��o.Remove sNomeDoArquivo
    End If
    
End Function

Public Function Zipar() As Long
    Dim zFileArray As ZIPnames
    Dim sNomeDoArquivo As Variant
    Dim lFileCount As Long
    Dim iIgnorePath As Integer
    Dim iRecurse As Integer
    On Error GoTo vbErrorHandler
    
    lFileCount = 0
    For Each sNomeDoArquivo In mCole��o
        zFileArray.s(lFileCount) = sNomeDoArquivo
        lFileCount = lFileCount + 1
    Next
    Zipar = VBZip(CInt(lFileCount), msNomeDoZip, _
        zFileArray, iIgnorePath, _
        miIncluiSubPastas, miUpdateZip, _
        0, msPastaRaiz)
    Exit Function

vbErrorHandler:
    Zipar = -99
    Err.Raise Err.Number, "cZipMil:Zipar", Err.Description

End Function

Public Function �ltimaMensagem() As String
    �ltimaMensagem = msOutput
End Function
