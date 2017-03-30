VERSION 5.00
Begin VB.Form frmContactpersoon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contact"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContactpersoon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   136
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   345
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAnnuleren 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtAdres 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   795
      Width           =   4215
   End
   Begin VB.TextBox txtNaam 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   315
      Width           =   4215
   End
   Begin VB.Label lblAdres 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mailadress:"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   960
   End
   Begin VB.Label lblNaam 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   465
   End
End
Attribute VB_Name = "frmContactpersoon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnnuleren_Click()
    Unload Me                                'Unload the me object
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub cmdOK_Click()
    AddContact txtNaam.Text, txtAdres.Text
    
    Unload Me                                'Unload the me object
End Sub

Private Sub txtAdres_Change()
    txtNaam_Change
End Sub

Private Sub txtNaam_Change()
    Dim punt As Boolean                      'Declare punt for local use as boolean

    If Len(txtAdres.Text) > 5 Then
        If InStr(Len(txtAdres.Text) - 5, txtAdres.Text, ".") > 0 Then punt = True Else punt = False
    Else: punt = False
    End If

    If Trim$(txtNaam.Text) <> vbNullString And InStr(1, txtAdres.Text, "@") > 0 And punt = True Then
        cmdOK.Enabled = True                 'enable the cmdok object
    Else: cmdOK.Enabled = False              'disable the else: cmdok object
    End If
End Sub

