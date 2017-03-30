VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1275
   ClientLeft      =   3555
   ClientTop       =   3750
   ClientWidth     =   3255
   LinkTopic       =   "Form2"
   ScaleHeight     =   1275
   ScaleWidth      =   3255
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   345
      Left            =   1710
      TabIndex        =   2
      Top             =   780
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   420
      TabIndex        =   1
      Top             =   780
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   210
      Width           =   2325
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    AlterarRegistroTXTBinário Form1.CommonDialog.FileName, RegAntigo, Form1.MSFlexGrid1.Row
    
End Sub
