VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   1740
   ClientTop       =   3270
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   5295
   Begin VB.CommandButton Command4 
      Caption         =   "Logoff"
      Height          =   315
      Left            =   390
      TabIndex        =   8
      Top             =   3810
      Width           =   1125
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reiniciar"
      Height          =   315
      Left            =   390
      TabIndex        =   7
      Top             =   3480
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Desligar"
      Height          =   315
      Left            =   390
      TabIndex        =   6
      Top             =   3150
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6x 
      Caption         =   "Pasta Windows System"
      Height          =   195
      Index           =   2
      Left            =   30
      TabIndex        =   14
      Top             =   1740
      Width           =   1755
   End
   Begin VB.Label Label6x 
      Caption         =   "Pasta Windows"
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   13
      Top             =   1500
      Width           =   1755
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   195
      Left            =   1830
      TabIndex        =   12
      Top             =   1740
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   195
      Left            =   1830
      TabIndex        =   11
      Top             =   1500
      Width           =   90
   End
   Begin VB.Label Label6x 
      Caption         =   "Nome do Computador"
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   10
      Top             =   1260
      Width           =   1755
   End
   Begin VB.Label label6 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   195
      Left            =   1830
      TabIndex        =   9
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   195
      Left            =   1830
      TabIndex        =   5
      Top             =   1032
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   195
      Left            =   1830
      TabIndex        =   4
      Top             =   804
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   195
      Left            =   1830
      TabIndex        =   3
      Top             =   576
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   195
      Left            =   1830
      TabIndex        =   2
      Top             =   348
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   195
      Left            =   1830
      TabIndex        =   0
      Top             =   120
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vWin As eWin

Private Sub Command1_Click()
    Label1.Caption = vWin.VerificarNomeSistOperacional
    Label2.Caption = vWin.SistemaOperacional
    Label3.Caption = vWin.Detalhe
    Label4.Caption = "Versão " & vWin.VersãoMaior & "." & vWin.VersãoMenor & "." & vWin.Compilação
    Label5.Caption = "Plataforma " & vWin.NomeDaPlataforma
    label6.Caption = vWin.NomeDoComputador
    vWin.VerificarPastaWin
    vWin.VerificarPastaWinSys
    Label7.Caption = vWin.PastaWindows
    Label8.Caption = vWin.PastaWinSys
End Sub

Private Sub Command2_Click()
    vWin.SairDoWindows IIf(vWin.Plataforma < 2, SDW_SHUTDOWN, SDW_POWEROFF)
End Sub

Private Sub Command3_Click()
    vWin.SairDoWindows SDW_REBOOT
End Sub

Private Sub Command4_Click()
    vWin.SairDoWindows SDW_LOGOFF
    End
End Sub

Private Sub Form_Load()
    Set vWin = New eWin
End Sub
