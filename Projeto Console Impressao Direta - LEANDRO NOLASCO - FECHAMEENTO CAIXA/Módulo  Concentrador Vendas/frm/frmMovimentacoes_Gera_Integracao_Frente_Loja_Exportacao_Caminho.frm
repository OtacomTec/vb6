VERSION 5.00
Begin VB.Form frmMovimentacoes_Gera_Integracao_Frente_Loja_Exportacao_Caminho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caminho"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Exportacao_Caminho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6960
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4410
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informe o caminho do arquivo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   90
      TabIndex        =   5
      Top             =   30
      Width           =   6765
      Begin VB.DriveListBox Drive1 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   630
         Width           =   3255
      End
      Begin VB.DirListBox Dir1 
         Height          =   1980
         Left            =   120
         TabIndex        =   1
         Top             =   1110
         Width           =   3255
      End
      Begin VB.TextBox txtCaminho 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3750
         Width           =   6495
      End
      Begin VB.FileListBox File1 
         Height          =   2490
         Left            =   3420
         Pattern         =   "*.dat"
         TabIndex        =   6
         Top             =   630
         Width           =   3225
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5370
         TabIndex        =   3
         ToolTipText     =   "Limpar o caminho do Arquivo"
         Top             =   3210
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Diretório"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   390
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Caminho do arquivo"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   3510
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmMovimentacoes_Gera_Integracao_Frente_Loja_Exportacao_Caminho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Transportes                                                    '
' Objetivo...............: Movimentação Gera Integração Frente de Loja                    '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rafael de Oliveira Gomes                                       '
' Data de Criação........: 18/12/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim strCaminho As String
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim I As Integer
Dim strArquivo As String
Dim log As New DLLSystemManager.log
Option Explicit

Private Sub cmdLimpar_Click()
    txtCaminho.Text = Empty
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    txtCaminho.Text = Dir1.Path
    strCaminho = Dir1.Path
End Sub

Private Sub Drive1_Change()
    If Drive1.Drive <> "a:" And Drive1.Drive <> "A:" Then
       Dir1.Path = Drive1.Drive
       txtCaminho.Text = Dir1.Path
    Else
       MsgBox "Não é recomendado que se gere arquivos diretamente para o drive 'A:', salve-os em sua máquina antes.", vbInformation, "Only Tech"
       Drive1.SetFocus
    End If
End Sub

Private Sub File1_DblClick()
    txtCaminho.Text = Dir1.Path & "\" & File1.FileName
End Sub

Private Sub cmdOk_Click()
  frmMovimentacoes_Gera_Integracao_Frente_Loja_Exportacao.txtCaminho.Text = txtCaminho.Text
   
  txtCaminho.Text = Empty
  Unload Me
End Sub

Private Sub Form_Load()
    Dir1.Path = "C:\Arquivos de programas"
    
    Drive1.Drive = "C:\"
End Sub
