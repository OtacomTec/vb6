VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Dim rVal As Long
Private Enum GMen_DesligarWindows
    Win_LogOff = 0
    Win_Desligar = 1
    Win_Reiniciar = 2
    Win_Force = 4
End Enum
Function Desligar(Optional OpçãoDesligamento As GMen_DesligarWindows = 1)
    rVal = ExitWindowsEx(OpçãoDesligamento, 0&)      'Shutdown...
End Function

