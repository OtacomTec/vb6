VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atualizador do bdGMS001"
   ClientHeight    =   2115
   ClientLeft      =   4725
   ClientTop       =   3855
   ClientWidth     =   3885
   Icon            =   "PAD0001-01-F1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Atualizar Agora"
      Height          =   555
      Left            =   1080
      TabIndex        =   0
      Top             =   1260
      Width           =   1725
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   270
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   525
      Left            =   270
      TabIndex        =   1
      Top             =   150
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim myshell As myshell
    Set myshell = New myshell
    
    myshell.SyncShell App.Path & "\bdgms001.exe", 1
    
    MsgBox "Base de Dados Atualizada!"
    
    Kill "bdgms001.*"
    End
    'Exit Sub
    
    
End Sub

Private Sub Form_Load()
    If Dir(App.Path & "\bdGMS001.exe") <> "" Then
        dtArquivo = FileDateTime(App.Path & "\bdGMS001.exe")
        Label1.Caption = "Este programa atualizará a sua base de dados!" & Chr(13) & _
                       "Esta base de dados foi gerada no dia:"
        Label2.Caption = dtArquivo
    Else
        Label1.Caption = "Não há base de dados para atualizar!"
        Me.Command1.Enabled = False
    End If
End Sub
