VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1305
      Left            =   3840
      TabIndex        =   0
      Top             =   3270
      Width           =   4785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

     Dim lRtn As Long
     
     lRtn = mdlTwain.TransferWithUI("C:\ui.bmp")
     
     MsgBox lRtn
    
End Sub
