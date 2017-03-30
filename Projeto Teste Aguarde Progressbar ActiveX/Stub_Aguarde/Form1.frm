VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOcultar_Progresso 
      Caption         =   "Ocultar Progresso"
      Height          =   255
      Left            =   510
      TabIndex        =   0
      Top             =   150
      Width           =   1785
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Testar"
      Height          =   285
      Left            =   3300
      TabIndex        =   1
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Path_Aguarde = "\\Onlytech-dados\Projetos\Sistemas Teste\Projeto Teste Aguarde Progressbar ActiveX - LEANDRO NOLASCO\Teste_Aguarde\ActiveX\xExe.exe"

Private Sub Command1_Click()

    Dim i As Long
    Dim XExe As XExe.XClass
    
    Shell Path_Aguarde, vbNormalNoFocus
    
    Set XExe = New XExe.XClass
    XExe.setMax_Progresso 10000
    ProgressBar1.Max = 10000
    XExe.setMin_Progresso 0
    ProgressBar1.Min = 0
    
    XExe.AbrirInterface
    
    For i = 0 To 10000
         XExe.Incrementar_Progresso
         ProgressBar1.Value = i
    Next i

    XExe.Destruir
    
    Set XExe = Nothing


End Sub

