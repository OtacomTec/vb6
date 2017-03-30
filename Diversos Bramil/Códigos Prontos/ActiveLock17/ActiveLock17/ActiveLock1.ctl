VERSION 5.00
Begin VB.UserControl ActiveLock 
   CanGetFocus     =   0   'False
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ActiveLock1.ctx":0000
   PropertyPages   =   "ActiveLock1.ctx":030A
   ScaleHeight     =   465
   ScaleWidth      =   480
   ToolboxBitmap   =   "ActiveLock1.ctx":035E
End
Attribute VB_Name = "ActiveLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' Author: Nelson Ferraz
' Date  : 1998-2002

Public Enum THastType 'Added 19 Apr 2002 --- Francois de Wet
  htSHA1AA1 = 0
  htSHA1AA2 = 1
  htMD5AA1 = 2
  htMD5AA2 = 3
  htMD5AB1 = 4
  htMD5AB2 = 5
End Enum

Private mstr_Salt As String

Private mstr_SoftwareName As String
Private mi_LiberationKeyLength As Integer
Private mi_SoftwareCodeLength As Integer

Private mb_LockToHardDrive As Boolean
Private mb_LockToWindowsSerial As Boolean
Private mb_LockToRandomNumber As Boolean
Private mb_LockToComputerName As Boolean

Private mstr_RegistryPath As String
Private ht_HashAlgorithm As THastType

Private Const SOFTWARECODE_MAXLength = 16
Private Const LIBERATIONKEY_MAXLength = 16
Private Const SOFTWARENAME_MAXLength = 40
Private Const SALT_MAXLength = 2

Public Event Registration(WasSuccessful As Boolean)
Attribute Registration.VB_Description = "This event is raised when the LiberationKey is changed."
Attribute Registration.VB_MemberFlags = "40"
Public Event Transference(WasSuccessful As Boolean, LiberationKey As String)
Attribute Transference.VB_Description = "This event is raised after the Transfer method is called. If succesful, it returns the LiberationKey that will unlock the other computer."
Attribute Transference.VB_MemberFlags = "40"
Public Event InvalidDate()


Public Property Let HashAlgorithm(ByVal vData As THastType)
    ht_HashAlgorithm = vData
    PropertyChanged "HashAlgorithm"
End Property

Public Property Get HashAlgorithm() As THastType
    HashAlgorithm = ht_HashAlgorithm
End Property

Private Function Hash(strHashThis As String) As String
  ' Allow different hash types

  Select Case ht_HashAlgorithm
    Case htSHA1AA1: Hash = SHA1AA1Hash(strHashThis)
    Case htSHA1AA2: Hash = SHA1AA2Hash(strHashThis)
    Case htMD5AA1: Hash = MD5AA1Hash(strHashThis)
    Case htMD5AA2: Hash = MD5AA2Hash(strHashThis)
    Case htMD5AB1: Hash = MD5AB1Hash(strHashThis)
    Case htMD5AB2: Hash = MD5AB2Hash(strHashThis)
    Case Else: Hash = SHA1AA1Hash(strHashThis) ' Default type
  End Select

End Function

Public Sub Reset()
Attribute Reset.VB_Description = "Resets ActiveLock's properties."
  ' Reset ActiveLock properties
  
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    SaveSetting mstr_SoftwareName, "ActiveLock", "RandomKey", ""
    SaveSetting mstr_SoftwareName, "ActiveLock", "LiberationKey", ""
    SaveSetting mstr_SoftwareName, "ActiveLock", "Counter", ""
    SaveSetting mstr_SoftwareName, "ActiveLock", "InitialDate", ""
    SaveSetting mstr_SoftwareName, "ActiveLock", "LastRunDate", ""
    PropertyChanged "SoftwareCode"
    PropertyChanged "LiberationKey"
    PropertyChanged "Counter"
    PropertyChanged "InitialDate"
    PropertyChanged "LastRunDate"
  End If
  ' TO-DO: Remove registry key
End Sub

Public Sub Transfer(OtherSoftwareCode As String)
Attribute Transfer.VB_Description = "Transfers the current license to another computer.\r\nThis method requires that you've set the LockToRandomNumber property, since the SoftwareCode is supposed to change."
  ' Transfer license to another computer

  ' Pre-requisites: RegisteredUser = True And mb_LockToRandomNumber = True
  If Not (mb_LockToRandomNumber) Or Not (RegisteredUser) Then
    RaiseEvent Transference(False, "")
    Exit Sub
  End If
  
  ' 1. This computer's license is cancelled
  Dim strRnd As Long
  Randomize
  strRnd = CStr(CLng(Rnd(1) * 2147483647))
  SaveSetting mstr_SoftwareName, "ActiveLock", "RandomKey", strRnd
  
  ' 2. A new license is generated to the other computer
  Dim strHash As String
  strHash = Hash(OtherSoftwareCode & mstr_SoftwareName) ' 123
  strHash = Left(strHash, LIBERATIONKEY_MAXLength)
  
  ' 3. Raise ActiveLock_Transference event
  RaiseEvent Transference(True, strHash)
End Sub

Public Sub Register(LiberationKey As String)
Attribute Register.VB_Description = "Changes the  LiberationKey."
  ' Stores the liberation code in the registry
  
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    LiberationKey = Left(LiberationKey, LIBERATIONKEY_MAXLength)
    SaveSetting mstr_SoftwareName, "ActiveLock", "LiberationKey", LiberationKey
    PropertyChanged "LiberationKey"
  End If
  RaiseEvent Registration(RegisteredUser) ' True or False
End Sub

Public Property Let LiberationKey(ByVal vData As String)
Attribute LiberationKey.VB_Description = "This is the key that will unlock your software. The user must set the correct value in order to have the software registered.\r\nIt's important to note that each SoftwareCode needs a different LiberationKey."
Attribute LiberationKey.VB_ProcData.VB_Invoke_PropertyPut = ";Misc"
Attribute LiberationKey.VB_MemberFlags = "440"
  ' For compatibility purposes only. Use method Register instead.
  Register (vData)
End Property

Public Property Let LockToComputerName(ByVal vData As Boolean)
  mb_LockToComputerName = vData
  PropertyChanged "LockToComputerName"
End Property

Public Property Get LockToComputerName() As Boolean
Attribute LockToComputerName.VB_Description = "Set this property to make the SoftwareCode depend on the computer name."
Attribute LockToComputerName.VB_ProcData.VB_Invoke_Property = "PropertyPage4;Behavior"
  LockToComputerName = mb_LockToComputerName
End Property

Public Property Let LockToHardDrive(ByVal vData As Boolean)
  mb_LockToHardDrive = vData
  PropertyChanged "LockToHardDrive"
End Property

Public Property Get LockToHardDrive() As Boolean
Attribute LockToHardDrive.VB_Description = "Set this property to make the SoftwareCode depend on the hard drive serial number."
Attribute LockToHardDrive.VB_ProcData.VB_Invoke_Property = "PropertyPage4;Behavior"
  LockToHardDrive = mb_LockToHardDrive
End Property

Public Property Let LockToWindowsSerial(ByVal vData As Boolean)
  mb_LockToWindowsSerial = vData
  PropertyChanged "LockToWindowsSerial"
End Property

Public Property Get LockToWindowsSerial() As Boolean
Attribute LockToWindowsSerial.VB_Description = "Set this property to make the SoftwareCode depend on the Windows serial number."
Attribute LockToWindowsSerial.VB_ProcData.VB_Invoke_Property = "PropertyPage4;Behavior"
  LockToWindowsSerial = mb_LockToWindowsSerial
End Property

Public Property Let LockToRandomNumber(ByVal vData As Boolean)
  mb_LockToRandomNumber = vData
  PropertyChanged "LockToRandomNumber"
End Property

Public Property Get LockToRandomNumber() As Boolean
Attribute LockToRandomNumber.VB_Description = "Set this property to make the SoftwareCode depend on an unique random number. \r\nThis is necessary if you want that the SoftwareCode changes when you reset ActiveLock."
Attribute LockToRandomNumber.VB_ProcData.VB_Invoke_Property = "PropertyPage4;Behavior"
  LockToRandomNumber = mb_LockToRandomNumber
End Property

Public Property Let RegistryPath(ByVal vData As String)
  If Right(vData, 1) <> "\" Then vData = vData & "\"
  mstr_RegistryPath = vData
  gstrRegistryPath = mstr_RegistryPath
  PropertyChanged "RegistryPath"
End Property

Public Property Get RegistryPath() As String
Attribute RegistryPath.VB_Description = "Returns/sets the path in the registry where the ActiveLock information is stored."
Attribute RegistryPath.VB_ProcData.VB_Invoke_Property = "PropertyPage2;Text"
  RegistryPath = mstr_RegistryPath
End Property

Public Property Get SoftwareCode() As String
Attribute SoftwareCode.VB_Description = "Returns the code that uniquely identifies each installation of your software. Read-only.\r\n"
Attribute SoftwareCode.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute SoftwareCode.VB_MemberFlags = "400"
  ' The SoftwareCode is a hash based on several
  ' properties (depending on the LockTo... values)
  
  Dim strParam As String
  Dim strCode As String
  
  If mb_LockToRandomNumber Then strParam = strParam & RandomNumber
  If mb_LockToWindowsSerial Then strParam = strParam & WindowsProductKey
  If mb_LockToHardDrive Then strParam = strParam & DriveSerial("C")
  If mb_LockToComputerName Then strParam = strParam & ComputerName
  
  ' If no parameters selected, use default ones
  If strParam = "" Then strParam = RandomNumber & WindowsProductKey
  
  strCode = Hash(strParam)
  ' Return the first characters (depends on SoftwareCodeLength)
  SoftwareCode = Left(strCode, mi_SoftwareCodeLength)
  PropertyChanged "SoftwareCode"
End Property

Private Sub UserControl_InitProperties()
  ' Set module variables (default values)
  
  mstr_SoftwareName = "YourAppName"
  mi_LiberationKeyLength = 16
  mi_SoftwareCodeLength = 6
  mb_LockToWindowsSerial = True
  mb_LockToRandomNumber = True ' Required by Transference method
  mstr_RegistryPath = "ActiveLock\"
  
  ' Set global variables (default values)
  gstrRegistryPath = mstr_RegistryPath
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Get module variables from PropBag
  
  mstr_SoftwareName = PropBag.ReadProperty("SoftwareName")
  mi_LiberationKeyLength = PropBag.ReadProperty("LiberationKeyLength")
  mi_SoftwareCodeLength = PropBag.ReadProperty("SoftwareCodeLength")
  mb_LockToHardDrive = PropBag.ReadProperty("LockToHardDrive")
  mb_LockToWindowsSerial = PropBag.ReadProperty("LockToWindowsSerial")
  mb_LockToRandomNumber = PropBag.ReadProperty("LockToRandomNumber")
  mb_LockToComputerName = PropBag.ReadProperty("LockToComputerName")
  mstr_RegistryPath = PropBag.ReadProperty("RegistryPath")
  mstr_Salt = PropBag.ReadProperty("Salt")
  ht_HashAlgorithm = PropBag.ReadProperty("HashAlgorithm")

  ' Set global variables
  gstrRegistryPath = mstr_RegistryPath
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  ' Write properties into PropBag
  PropBag.WriteProperty "SoftwareName", mstr_SoftwareName
  PropBag.WriteProperty "LiberationKeyLength", mi_LiberationKeyLength
  PropBag.WriteProperty "SoftwareCodeLength", mi_SoftwareCodeLength
  PropBag.WriteProperty "LockToHardDrive", mb_LockToHardDrive
  PropBag.WriteProperty "LockToWindowsSerial", mb_LockToWindowsSerial
  PropBag.WriteProperty "LockToRandomNumber", mb_LockToRandomNumber
  PropBag.WriteProperty "LockToComputerName", mb_LockToComputerName
  PropBag.WriteProperty "RegistryPath", mstr_RegistryPath
  PropBag.WriteProperty "Salt", mstr_Salt
  PropBag.WriteProperty "HashAlgorithm", ht_HashAlgorithm
End Sub

Public Property Get Counter() As Long
Attribute Counter.VB_Description = "Returns the number of times this software has been run in the user's machine. Read-only."
Attribute Counter.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute Counter.VB_MemberFlags = "400"
  ' Set and returns Counter + 1 everytime it is called
  
  Dim varCounter
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    varCounter = GetSetting(SoftwareName, "ActiveLock", "Counter")
    If varCounter = "" Then varCounter = 0
    varCounter = varCounter + 1
    SaveSetting SoftwareName, "ActiveLock", "Counter", varCounter
  End If
  Counter = varCounter
  PropertyChanged "Counter"
End Property

Public Property Get UsedDays() As Long
Attribute UsedDays.VB_Description = "Returns the number of days since the first run date. Read-only."
Attribute UsedDays.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute UsedDays.VB_MemberFlags = "400"
  ' Count the number of days since first run time
  
  Dim varInitialDate As Variant
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    varInitialDate = GetSetting(mstr_SoftwareName, "ActiveLock", "InitialDate")
    If varInitialDate = "" Then
      ' If it's not in the registry
      varInitialDate = Now
      SaveSetting mstr_SoftwareName, "ActiveLock", "InitialDate", varInitialDate
    End If
    ' Return the number of days since initial date
    UsedDays = CLng(DateDiff("d", varInitialDate, Now))
    PropertyChanged "UsedDays"
    ' See if the user have changed the time settings
    If LastRunDate > Now Then RaiseEvent InvalidDate
  End If
End Property

Public Property Get RegisteredUser() As Boolean
Attribute RegisteredUser.VB_Description = "Returns a boolean value indicating if the user is correctly registered. Read-only."
Attribute RegisteredUser.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute RegisteredUser.VB_MemberFlags = "400"
  ' Compares LiberationKey with Hash(SoftwareCode & SoftwareName)
  
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    Dim strHash As String, strKey As String
    strHash = Hash(SoftwareCode & mstr_SoftwareName) ' 123
    strKey = GetSetting(mstr_SoftwareName, "ActiveLock", "LiberationKey")
    If UCase(Left(strHash, mi_LiberationKeyLength)) = UCase(strKey) Then
        RegisteredUser = True
    Else
        RegisteredUser = False
    End If
    PropertyChanged "RegisteredUser"
  End If
End Property

Public Property Let SoftwareName(ByVal vData As String)
  mstr_SoftwareName = Left(vData, SOFTWARENAME_MAXLength)
  PropertyChanged "SoftwareName"
End Property

Public Property Get SoftwareName() As String
Attribute SoftwareName.VB_Description = "IMPORTANT ! Set this property to identify your software."
Attribute SoftwareName.VB_ProcData.VB_Invoke_Property = "PropertyPage1;Text"
  SoftwareName = mstr_SoftwareName
End Property

Public Property Let Salt(ByVal vData As String)
  mstr_Salt = Left(vData, SALT_MAXLength)
  PropertyChanged "Salt"
End Property

Public Property Get Salt() As String
  Salt = mstr_Salt
End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
  frmAbout.Show 1
End Sub

Private Sub UserControl_Paint()
  UserControl.Line (15, 15)-(450, 15), &HFFFFFF
  UserControl.Line (450, 15)-(450, 450), &H808080
  UserControl.Line (450, 450)-(15, 450), &H808080
  UserControl.Line (15, 450)-(15, 15), &HFFFFFF
End Sub

Private Sub UserControl_Resize()
  UserControl.Height = 465
  UserControl.Width = 480
End Sub

Public Property Get LiberationKeyLength() As Integer
Attribute LiberationKeyLength.VB_Description = "Returns/sets the length of the LiberationKey."
Attribute LiberationKeyLength.VB_ProcData.VB_Invoke_Property = "PropertyPage3;Behavior"
  LiberationKeyLength = mi_LiberationKeyLength
End Property

Public Property Let LiberationKeyLength(ByVal vNewValue As Integer)
  If vNewValue > LIBERATIONKEY_MAXLength Then vNewValue = LIBERATIONKEY_MAXLength
  If vNewValue < 1 Then vNewValue = 1
  mi_LiberationKeyLength = vNewValue
  PropertyChanged "LiberationKeyLength"
End Property

Public Property Get SoftwareCodeLength() As Integer
Attribute SoftwareCodeLength.VB_Description = "Returns/sets the length of the SoftwareCode."
Attribute SoftwareCodeLength.VB_ProcData.VB_Invoke_Property = "PropertyPage3;Behavior"
  SoftwareCodeLength = mi_SoftwareCodeLength
End Property

Public Property Let SoftwareCodeLength(ByVal vNewValue As Integer)
  If vNewValue > SOFTWARECODE_MAXLength Then vNewValue = SOFTWARECODE_MAXLength
  If vNewValue < 1 Then vNewValue = 1
  mi_SoftwareCodeLength = vNewValue
  PropertyChanged "SoftwareCodeLength"
End Property

Private Function RandomNumber() As String
  ' Generates a random number, which is stored in the registry
  ' (This random number will be always the same, for a given SoftwareName)
  
  Dim strRnd As String
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    strRnd = GetSetting(mstr_SoftwareName, "ActiveLock", "RandomKey")
    If strRnd = "" Then
      ' If there isn't a random number, generate one
      Randomize
      strRnd = CStr(CLng(Rnd(1) * 2147483647))
      SaveSetting mstr_SoftwareName, "ActiveLock", "RandomKey", strRnd
    End If
    ' If there is a random number, return it
    RandomNumber = strRnd
  End If
End Function

Public Property Get LastRunDate() As Date
Attribute LastRunDate.VB_Description = "Returns the last date when the program was executed. Read-only."
Attribute LastRunDate.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute LastRunDate.VB_MemberFlags = "440"
  ' Returns last run date. If LastRunDate < Now(), user may be cheating
  
  Dim varLastRunDate As Variant
  Dim Agora As Date
  Agora = Now
  If mstr_SoftwareName = "" Then
    SoftwareNameError
  Else
    varLastRunDate = GetSetting(mstr_SoftwareName, "ActiveLock", "LastRunDate", Agora)
    If Not IsDate(varLastRunDate) Then varLastRunDate = Agora
    ' Store current date only if LastRunDate < Now()
    If varLastRunDate <= Agora Then SaveSetting mstr_SoftwareName, "ActiveLock", "LastRunDate", Agora
    LastRunDate = varLastRunDate
    PropertyChanged "LastRunDate"
  End If
End Property

'' THE END ''
