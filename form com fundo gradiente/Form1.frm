VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.AutoRedraw = True
Me.DrawStyle = 6
Me.DrawMode = 13
Me.DrawWidth = 2
Me.ScaleMode = 3
Me.ScaleHeight = (256 * 2)
For i = 255 To 0 Step -1
Me.Line (0, Y)-(Me.Width, Y + 2), RGB(255, i, 0), BF
Y = Y + 2
Next i


End Sub
