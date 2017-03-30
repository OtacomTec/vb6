VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VALIDA Nº DE SÉRIE DE IMPRESSORA"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton bematech 
      Caption         =   "Bematech"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   3
      Top             =   150
      Width           =   2835
   End
   Begin VB.OptionButton daruma 
      Caption         =   "Daruma"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   210
      TabIndex        =   2
      Top             =   150
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VALIDA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4200
      TabIndex        =   1
      Top             =   1230
      Width           =   2745
   End
   Begin VB.TextBox txtNumero_Serie 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   0
      Top             =   630
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     Dim clsCriptografia As New clsCriptografia
     If Me.bematech.value = True Then
           MsgBox Mid(clsCriptografia.CriptSenha(txtNumero_Serie.Text), 1, 13), vbInformation
     ElseIf Me.daruma.value = True Then
           MsgBox Mid(clsCriptografia.CriptSenha(txtNumero_Serie.Text), 1, 15), vbInformation
     End If
End Sub
