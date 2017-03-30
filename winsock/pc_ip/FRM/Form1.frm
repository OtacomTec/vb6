VERSION 5.00
Object = "{6056838B-3FD3-48EF-A66A-575DBDFF9DD1}#1.0#0"; "PC_IP.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "PC_IP"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7455
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
   ScaleHeight     =   5055
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PCIP.PC_IP PC_IP1 
      Left            =   1680
      Top             =   2835
      _ExtentX        =   1429
      _ExtentY        =   1111
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   75
      TabIndex        =   3
      Top             =   4275
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   390
      Left            =   2010
      TabIndex        =   2
      Top             =   4590
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Carregar"
      Height          =   390
      Left            =   60
      TabIndex        =   1
      Top             =   4590
      Width           =   1950
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   7245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    PC_IP1.Listar_Maquinas_IP List1, ProgressBar1
End Sub

Private Sub Command2_Click()
    PC_IP1.AbrirFormPadrão
End Sub
