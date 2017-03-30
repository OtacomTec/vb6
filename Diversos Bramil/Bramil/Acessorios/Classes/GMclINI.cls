VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMclINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'PROGRAMADOR: Luis Henrique Borges
'
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Ret As String

Public Function INIGrava(NomeDoArquivo As String, Seção As String, Chave As String, Text As String)
    WritePrivateProfileString Seção, Chave, Text, NomeDoArquivo
End Function

Public Function INILê(NomeDoArquivo As String, Seção As String, Chave As String)
    Ret = Space$(255)
    RetLen = GetPrivateProfileString(Seção, Chave, "", Ret, Len(Ret), NomeDoArquivo)
    Ret = Left$(Ret, RetLen)
    INILê = Ret
End Function

