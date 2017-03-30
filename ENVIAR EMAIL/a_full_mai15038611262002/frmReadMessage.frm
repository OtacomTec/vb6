VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReadMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Read Message"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   4440
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "Save attachment"
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtAttachment 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtSender 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdRespond 
      Caption         =   "Respond"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtMessage 
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Attachment:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   390
      TabIndex        =   5
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sender:"
      Height          =   195
      Left            =   420
      TabIndex        =   3
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmReadMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AttachmentData  As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Unload Me
End Sub

Private Sub cmdRespond_Click()
    sAdressData.message = txtMessage
    sAdressData.responseAdress = "irnchen@web.de"
    sAdressData.subject = txtSubject
    CreateNewSendMailForm sAdressData
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dialog.Filter = txtAttachment.Text & "|" & txtAttachment.Text
    Dialog.ShowSave
    If Dialog.FileName <> "" Then
        modUUCode.UUDecodeToFile AttachmentData, Dialog.FileName
    End If
End Sub

Private Sub Form_Load()
    txtMessage.Text = MessageBuffer(frmMain.lstMessages.ListIndex)
    txtSender.Text = FromBuffer(frmMain.lstMessages.ListIndex)
    txtSubject.Text = SubjectBuffer(frmMain.lstMessages.ListIndex)
    If InStr(1, LCase(txtMessage.Text), "begin 664") <> 0 Then
        SplittedData = Split(txtMessage.Text, vbCrLf)
        For i = 0 To UBound(SplittedData)
            If InStr(1, SplittedData(i), "begin 664") <> 0 Then
            txtAttachment.Text = Mid(SplittedData(i), 11)
            AttachmentData = Mid(txtMessage.Text, InStr(1, txtMessage.Text, "begin 664"))
            txtMessage.Text = Left(txtMessage.Text, InStr(1, txtMessage.Text, "begin 664") - 1)
            End If
        Next i
    End If
End Sub
