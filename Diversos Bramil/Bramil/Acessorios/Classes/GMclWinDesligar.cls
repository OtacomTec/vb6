VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMclWinDesligar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Declaração para a Função Desligar
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Dim rVal As Long
'---------------------------------

Public Enum WIN_DesligarWindows
    Win_LogOff = 0
    Win_Desligar = 1
    Win_Reiniciar = 2
    Win_Force = 4
End Enum
Public Function Desligar(Optional OpçãoDesligamento As WIN_DesligarWindows = Win_Desligar)
    rVal = ExitWindowsEx(OpçãoDesligamento, 0&)      'Shutdown...
End Function

