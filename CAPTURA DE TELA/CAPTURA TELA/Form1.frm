VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   13335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCaptura 
      Caption         =   "Captura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   9420
      TabIndex        =   0
      Top             =   8070
      Width           =   3735
   End
   Begin VB.Image imgTitulo_Rec 
      Height          =   7575
      Left            =   210
      Top             =   300
      Width           =   12915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Const VK_SNAPSHOT = &H2C

Private Sub cmdCaptura_Click()
    
    'Assim vc copia a tela do Windows.
    Call keybd_event(vbKeySnapshot, 0, 0, 0)
    DoEvents
    
    'e assim vc mostra numa Image
    imgTitulo_Rec.Picture = Clipboard.GetData()
    
    
End Sub
