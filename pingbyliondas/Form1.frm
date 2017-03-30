VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Utilitario para Redes"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   4125
      TabIndex        =   8
      Text            =   "192.168.0.1"
      Top             =   2250
      Width           =   1740
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   1140
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2250
      Width           =   3765
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalhes da rede"
      Height          =   1740
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   5790
      Begin VB.CommandButton Command1 
         Caption         =   "Atualizar"
         Height          =   765
         Left            =   4200
         TabIndex        =   2
         Top             =   525
         Width           =   1290
      End
      Begin VB.Label Label8 
         Caption         =   "Endereço IP           :"
         Height          =   240
         Left            =   225
         TabIndex        =   12
         Top             =   975
         Width           =   1440
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "?????"
         Height          =   240
         Left            =   1725
         TabIndex        =   11
         Top             =   975
         Width           =   2265
      End
      Begin VB.Label Label6 
         Caption         =   "Endereço MAC      :"
         Height          =   240
         Left            =   225
         TabIndex        =   10
         Top             =   1275
         Width           =   1440
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "?????"
         Height          =   240
         Left            =   1725
         TabIndex        =   9
         Top             =   1275
         Width           =   2265
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "?????"
         Height          =   240
         Left            =   1725
         TabIndex        =   6
         Top             =   675
         Width           =   2265
      End
      Begin VB.Label Label3 
         Caption         =   "Nome Usuário        :"
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "?????"
         Height          =   240
         Left            =   1725
         TabIndex        =   4
         Top             =   375
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "Nome Computador :"
         Height          =   240
         Left            =   225
         TabIndex        =   3
         Top             =   375
         Width           =   1440
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ping"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4125
      TabIndex        =   0
      Top             =   2700
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "IP para efetuar Ping"
      Height          =   240
      Left            =   4125
      TabIndex        =   14
      Top             =   2025
      Width           =   1740
   End
   Begin VB.Label Label9 
      Caption         =   "Resposta do Ping"
      Height          =   240
      Left            =   150
      TabIndex        =   13
      Top             =   2025
      Width           =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Mostra algumas informacoes da rede
Private Sub Command1_Click()
   'Cria a variavel para a classe SystemView
   Dim ClasseSystem As SystemView
   'Abre a classe SystemView
   Set ClasseSystem = New SystemView
   'Pega as informacoes da rede
   Label2.Caption = ClasseSystem.ComputadorNome
   Label4.Caption = ClasseSystem.ComputadorUsuario
   Label7.Caption = ClasseSystem.RedeIP
   Label5.Caption = ClasseSystem.RedeMAC(0)
   'Fecha a classe SystemView
   Set ClasseSystem = Nothing
End Sub

'Efetua o ping
Private Sub Command2_Click()
   'Cria a variavel para a classe Rede
   Dim ClasseRede As Rede
   'Abre a classe Rede
   Set ClasseRede = New Rede
   'Mostra o status do ping em um textbox
   Text1.Text = "  Ping          " & vbTab & ": " & Text2.Text & vbCrLf
   DoEvents
   ClasseRede.Ping Text2.Text 'Inicia o processo de ping
   Text1.Text = Text1.Text & "  Status          " & vbTab & ": " & ClasseRede.GetStatusCode(ClasseRede.ECHO_Status) & vbCrLf
   Text1.Text = Text1.Text & "  Round Trip Time " & vbTab & ": " & ClasseRede.ECHO_RoundTripTime & " ms" & vbCrLf
   Text1.Text = Text1.Text & "  Data Size       " & vbTab & ": " & ClasseRede.ECHO_DataSize & " bytes" & vbCrLf
   'Fecha a classe Rede
   Set ClasseRede = Nothing
End Sub
