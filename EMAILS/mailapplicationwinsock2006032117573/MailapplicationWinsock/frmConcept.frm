VERSION 5.00
Begin VB.Form frmConcept 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Draft"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnnuleren 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtNaam 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label lblNaam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Give a name for the draft:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1890
   End
End
Attribute VB_Name = "frmConcept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub cmdAnnuleren_Click()
    frmMain.strConcept = vbNullString
    Unload Me                                'Unload the me object
End Sub

Private Sub cmdOK_Click()
    Dim bestaatal As Boolean, lngIndex As Long 'Declare bestaatal for local use as boolean, lngindex as long
    For i = 0 To frmMain.cmbConcept.ListCount - 1
        If LCase$(frmMain.cmbConcept.List(i)) = LCase$(txtNaam.Text) Then
            bestaatal = True: lngIndex = i: Exit For
        End If
    Next i

    If bestaatal = True Then
        If MsgBox("Draft '" & txtNaam.Text & "' already exist. Overwrite?", vbQuestion + vbYesNo, "Draft") = vbNo Then 'Inform the user with a messagebox
            Exit Sub                         'Leave this sub
        Else: frmMain.cmbConcept.RemoveItem lngIndex
        End If
    End If

    frmMain.strConcept = txtNaam.Text
    Unload Me                                'Unload the me object
End Sub

Private Sub txtNaam_Change()
    If Trim$(txtNaam.Text) <> vbNullString Then
        cmdOK.Enabled = True                 'enable the cmdok object
    Else: cmdOK.Enabled = False              'disable the else: cmdok object
    End If
End Sub

Private Sub txtNaam_KeyPress(KeyAscii As Integer)
    Dim lijst As New Collection              'Declare lijst for local use as new collection
    lijst.Add 47: lijst.Add 63: lijst.Add 92: lijst.Add 42: lijst.Add 34
    lijst.Add 60: lijst.Add 62: lijst.Add 124

    For i = 1 To lijst.Count
        If KeyAscii = lijst(i) Then KeyAscii = 0
    Next i
End Sub

