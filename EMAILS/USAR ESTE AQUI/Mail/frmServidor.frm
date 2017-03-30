VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Server 
   Caption         =   "::: Mail 1.1 ::: "
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServidor.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Mensagem 
      Height          =   4095
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7223
      _Version        =   393217
      TextRTF         =   $"frmServidor.frx":030A
   End
   Begin VB.CommandButton cmdVisit 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdTeste 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      Begin VB.TextBox Assunto 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   7335
      End
      Begin VB.TextBox CC 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox Para 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox De 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox PWR 
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox User 
         Height          =   315
         Left            =   4680
         TabIndex        =   6
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox SMTP 
         Height          =   315
         Left            =   4680
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cópia Para <email@email.com.br>"
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   2430
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Para <email@email.com>,<emai1@email.com>,..."
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   3480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De <email@email.com>"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha SMTP"
         Height          =   210
         Left            =   4680
         TabIndex        =   9
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário SMTP"
         Height          =   210
         Left            =   4680
         TabIndex        =   7
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP"
         Height          =   210
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   405
      End
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTeste_Click()
EnvioDeEmail

End Sub

Private Sub cmdVisit_Click()
    Unload Me
    
End Sub

Public Function EnvioDeEmail()

Dim Msg As CDO.Message
Dim Cof As CDO.Configuration
Dim Camp

Set Msg = New CDO.Message
Set Cof = New CDO.Configuration
Set Camp = Cof.Fields
  
 
    With Camp
    
      .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
      .Item(cdoSMTPServer) = SMTP.Text
      .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
      .Item(cdoSMTPAuthenticate) = 1
      .Item(cdoSendUserName) = User.Text
      .Item(cdoSendPassword) = LCase(PWR.Text)
      .Update
      
    End With
    
    With Msg
    
      Set .Configuration = Cof
      
          .To = Para.Text
          .From = De.Text
          .Subject = Assunto.Text
          .TextBody = Mensagem.Text
          .CC = CC.Text
          .AddAttachment ("C:\CLIENTESFF2482010182252100.TXT")
          .Send
          
    End With

End Function


