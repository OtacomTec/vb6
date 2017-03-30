VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormAguarde 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1545
   ClientLeft      =   3450
   ClientTop       =   2220
   ClientWidth     =   5565
   ControlBox      =   0   'False
   Icon            =   "FRM9002-01-F2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.Animation Animation 
      Height          =   795
      Left            =   4650
      TabIndex        =   5
      Top             =   30
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1402
      _Version        =   393216
      FullWidth       =   55
      FullHeight      =   53
   End
   Begin VB.CommandButton CommandCancela 
      Caption         =   "&Cancela"
      Enabled         =   0   'False
      Height          =   435
      Left            =   2250
      TabIndex        =   2
      Top             =   570
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   1215
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   5085
      TabIndex        =   4
      Top             =   975
      Width           =   405
   End
   Begin VB.Label LabelReg 
      AutoSize        =   -1  'True
      Caption         =   "Registros Lidos:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   975
      Width           =   1125
   End
   Begin VB.Label LabelTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Aguarde, selecionando registros..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4485
   End
End
Attribute VB_Name = "FormAguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandCancela_Click()
    pboProcesso = False
    DoEvents
End Sub

Private Sub Form_Load()
    pboProcesso = True
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    If Dir(Trim(pstrLocacaoAcessoriostLogin) & "FINDFILE.AVI") <> "" Then
        Animation.Open Trim(pstrLocacaoAcessoriostLogin) & "FINDFILE.AVI"
        Animation.Play
    End If
    DoEvents
End Sub

