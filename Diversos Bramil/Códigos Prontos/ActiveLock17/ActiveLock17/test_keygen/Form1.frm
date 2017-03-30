VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveLock 1.7 KeyGenerator"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select HashAlgorithm"
      Height          =   2175
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generate Key"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AB2"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AB1"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AA2"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "MD5AA1"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "SHA1AA2"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optHashAlgorithm 
         Caption         =   "SHA1AA1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties and LiberationKey"
      Height          =   2175
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   4335
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "SoftwareName:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblLiberationKeyLenght 
         Alignment       =   1  'Right Justify
         Caption         =   "LiberationKeyLength:"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SoftwareCode:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "LiberationKey:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Label lblPanel 
      Caption         =   "lblPanel"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   3000
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Form1.frx":0442
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":0884
      Height          =   495
      Left            =   600
      TabIndex        =   17
      Top             =   2400
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************************** _
  * This KeyGenerator for Activelock17b is made by                      * _
  * Mads Moen                                                           * _
  * My e-mailadress: mads@industri.no                                   * _
  * My homepage: http://www.industri.no                                 * _
  * Code is made from Mr. Nelson Correa de Toledo Ferraz`s posting to   * _
  * http://groups.yahoo.com/group/activelock Sat Apr 20, 2002           * _
  * Feel free to reuse and improve my KeyGenerator                      * _
  ***********************************************************************
Option Explicit

Dim mi_HashAlgorithm As Integer ' the "mi_" prefix stands for:
                                ' scope = Module, type = Integer

Private Sub Command1_Click()
    'Making some variables
    Dim strKey As String
    Dim strSoftwareCode As String
    Dim strSoftwareName As String
    Dim intLiberationKeyLength As Integer
    
    'Catching the Softwarename, make sure you type it exactly as you _
     set it in ActiveLock propertypages for your software
    strSoftwareName = Me.Text(0)
    
    'Catching the length of LiberationKey, make sure you _
     type it exactly as you set it in ActiveLock propertypages _
     for your software
    intLiberationKeyLength = Me.Text(1)
    
    'Here you type or paste your softwarecode under runtime _
     Remember "exactly" is very important.
    strSoftwareCode = Me.Text(2)
    
    'Here we call for Function "hash"
    strKey = hash(strSoftwareCode & strSoftwareName)
    
    'Here we get the proper part of the generated LiberationKey based _
     on what we type in for LiberationKeyLength under runtime.
    strKey = UCase(Left(strKey, intLiberationKeyLength))
    
    'Printing the key to the textbox.
    Text(3).Text = strKey
    
    'Selecting the LiberatingKey in the textbox, now you just copy it _
     and place it where you decide
    Text(3).SetFocus
    Text(3).SelStart = 0
    Text(3).SelLength = Len(Text(3).Text)
    
    'Save default settings for next session
    SaveSetting "KeyGenerator", "Default", "SoftwareName", strSoftwareName
    SaveSetting "KeyGenerator", "Default", "LiberationKeyLength", intLiberationKeyLength
    SaveSetting "KeyGenerator", "Default", "HashAlgorithm", mi_HashAlgorithm
End Sub
Private Function hash(ByVal strHashThis As String)

'Here we call for LiberatingKeys by several functions made by ActiveLock _
 author based on the algorithm we choose

Select Case mi_HashAlgorithm
    Case 0: hash = SHA1AA1Hash(strHashThis)
    Case 1: hash = SHA1AA2Hash(strHashThis)
    Case 2: hash = MD5AA1Hash(strHashThis)
    Case 3: hash = MD5AA2Hash(strHashThis)
    Case 4: hash = MD5AB1Hash(strHashThis)
    Case 5: hash = MD5AB2Hash(strHashThis)
    Case Else: hash = SHA1AA1Hash(strHashThis) 'Default
End Select

End Function

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Dim strSoftwareName As String
Dim intLiberationKeyLength As Integer, intHashAlgorithm As Integer

'Making some text in the stausbarpanel..
lblPanel = "KeyGenerator for ActiveLock 1.7.2, Made by " & App.Comments & " @ http://www.industri.no"

'Get default values from previous session
strSoftwareName = GetSetting("KeyGenerator", "Default", "SoftwareName", "My SoftWareName")
intLiberationKeyLength = GetSetting("KeyGenerator", "Default", "LiberationKeyLength", 16)
intHashAlgorithm = GetSetting("KeyGenerator", "Default", "HashAlgorithm", 0)

'Load some standard text
Me.Text(0).Text = strSoftwareName
Me.Text(1).Text = intLiberationKeyLength
optHashAlgorithm(intHashAlgorithm) = True

End Sub

Private Sub optHashAlgorithm_Click(Index As Integer)
'Setting conditions for the select event in Function "hash"
mi_HashAlgorithm = optHashAlgorithm(Index).Index
End Sub

Private Sub optHashAlgorithm_KeyPress(Index As Integer, KeyAscii As Integer)
'With this code EnterKey act like TAB key in the _
 optHashAlgorithm - optionbuttonarray
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optHashAlgorithm_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Calling proceduer under Command1 button when _
 a option in the optHashAlgorithm array is picked with spacebar

Command1_Click
End Sub

Private Sub optHashAlgorithm_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Calling proceduer under Command1 button when _
 a option in the optHashAlgorithm array is picked with mousepointer

Command1_Click
End Sub

Private Sub Text_GotFocus(Index As Integer)
Dim I
'Selecting the all the text in textboxarray while _
 the box is getting focus
For I = 0 To Text.Count - 1
  Text(I).SelStart = 0
  Text(I).SelLength = Len(Text(I).Text)
Next I
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
'With this code EnterKey act like TAB key in the Text - textboxarray
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
