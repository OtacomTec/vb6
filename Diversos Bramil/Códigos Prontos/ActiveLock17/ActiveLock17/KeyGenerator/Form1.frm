VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sample KeyGenerator"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLiberationKeyLength 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "txtLiberationKeyLength"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtLiberationKey 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "txtLiberationKey"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Key"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cmbHashType 
      Height          =   315
      ItemData        =   "Form1.frx":030A
      Left            =   1440
      List            =   "Form1.frx":0320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtSoftwareCode 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "txtSoftwareCode"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtSoftwareName 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "txtSoftwareName"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "LibKeyLength"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "LiberationKey"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "HashType"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "SoftwareCode"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "SoftwareName"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project: Sample KeyGenerator
' Author:  Nelson Ferraz
'
' TO-DO:   ActiveX KeyGenerator

Option Explicit

Public Enum THastType 'Added 19 Apr 2002 --- Francois de Wet
  htSHA1AA1 = 0
  htSHA1AA2 = 1
  htMD5AA1 = 2
  htMD5AA2 = 3
  htMD5AB1 = 4
  htMD5AB2 = 5
End Enum

Private Sub cmdGenerate_Click()
    txtLiberationKey = Left(Hash(txtSoftwareCode & txtSoftwareName), txtLiberationKeyLength)
End Sub

Private Sub Form_Load()
    txtSoftwareName = ""
    txtSoftwareCode = ""
    txtLiberationKeyLength = "6"
    txtLiberationKey = ""
    cmbHashType.ListIndex = 0
End Sub

Private Function Hash(strHashThis As String) As String
  ' Allow different hash types
  
  Dim ht_HashAlgorithm As THastType
  ht_HashAlgorithm = cmbHashType.ListIndex

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

'' THE END ''
