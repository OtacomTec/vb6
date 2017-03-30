Attribute VB_Name = "Module1"
'Para que a DLL funcione, copiei-a para dentro da pasta ..\windows\system

Public Declare Function ActiveDll Lib "Online.dll" () As Long
Public Declare Function DeactiveDll Lib "Online.dll" () As Long

Public Declare Function InsertTerminal Lib "Online.dll" (ByVal Terminal As Long) As Long
Public Declare Function DeleteTerminal Lib "Online.dll" (ByVal Terminal As Long) As Long
Public Declare Function EnableTerminal Lib "Online.dll" (ByVal Terminal As Long) As Long
Public Declare Function DisableTerminal Lib "Online.dll" (ByVal Terminal As Long) As Long

Public Declare Function SetPoolingIntervalTime Lib "Online.dll" (ByVal IntervalTime As Long) As Long
Public Declare Function SetTerminalResponseTime Lib "Online.dll" (ByVal Time As Long) As Long

Public Declare Function SetComm Lib "Online.dll" (ByVal CommPort As Long) As Long
Public Declare Function SetBaudRate Lib "Online.dll" (ByVal Baudrate As Long) As Long

Public Declare Function SetTerminalTimeOut Lib "Online.dll" (ByVal TerminalTimeOut As String) As Long
Public Declare Function SetConditionAfterTimeOut Lib "Online.dll" (ByVal TerminalCondition As String) As Long

Public Declare Function OpenComm Lib "Online.dll" () As Long
Public Declare Function CloseComm Lib "Online.dll" () As Long

Public Declare Function StartPooling Lib "Online.dll" () As Long
Public Declare Function StopPooling Lib "Online.dll" () As Long

Public Declare Function SetDateTime Lib "Online.dll" (ByVal TerminalCurrentDateTime As String) As Long
Public Declare Function SendMessage Lib "Online.dll" (ByVal TerminalTimeMessagePersonalMessage As String) As Long

Public Declare Function Question Lib "Online.dll" () As String
Public Declare Function Answer Lib "Online.dll" (ByVal TerminalBadgePositionStatusTimeMessagePersonalMessage As String) As Long




