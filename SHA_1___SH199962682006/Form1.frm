VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHA1 TestBed"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "V     HASH FILE     V"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   7335
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "V     HASH TEXT     V"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7560
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "SHA1 - Philip Ciebiera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   7575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hash Digest:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' If you like this code give me a vote!
' -Phil

Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Command1_Click()
    Dim oSHA1 As New clsSHA1
    Dim LngStart, LngEnd As Long
    LngStart = GetTickCount
    Text2.Text = oSHA1.SHA1(Text1.Text)
    LngEnd = GetTickCount
    Set oSHA1 = Nothing
    MsgBox "Hashing took: " & LngEnd - LngStart & "ms"
End Sub

Private Function BinaryRead(ByRef sFileName As String) As String
    Dim fh As Integer
    fh = FreeFile
    
    Open sFileName For Binary As #fh
    BinaryRead = Input$(LOF(fh), fh)
    Close #fh
End Function

Private Sub Command2_Click()
    With CommonDialog1
        .Filter = "*.*"
        .DialogTitle = "Select file to hash, doesn't modify file."
        .ShowOpen
        If .FileName = "" Then Exit Sub
    End With
    Dim oSHA1 As New clsSHA1
    Dim LngStart, LngEnd As Long
    Dim sFile As String
    Me.MousePointer = 11
    sFile = BinaryRead(CommonDialog1.FileName)
    LngStart = GetTickCount
    Text2.Text = oSHA1.SHA1(sFile)
    LngEnd = GetTickCount
    Set oSHA1 = Nothing
    Me.MousePointer = 0
    MsgBox "Hashing took: " & LngEnd - LngStart & "ms" & vbCrLf & "on a " & Format(Len(sFile), "###,###,###,##0") & " byte file."
End Sub
