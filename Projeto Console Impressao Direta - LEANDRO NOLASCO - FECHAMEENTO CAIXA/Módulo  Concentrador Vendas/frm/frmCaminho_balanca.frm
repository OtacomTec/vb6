VERSION 5.00
Begin VB.Form frmCaminho_balanca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caminho"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaminho_balanca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6255
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
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
      Height          =   3075
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   6195
      Begin VB.DriveListBox Drive1 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   630
         Width           =   2775
      End
      Begin VB.DirListBox Dir1 
         Height          =   900
         Left            =   120
         TabIndex        =   4
         Top             =   1050
         Width           =   2775
      End
      Begin VB.TextBox txtCaminho 
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   2580
         Width           =   5985
      End
      Begin VB.FileListBox File1 
         Height          =   1290
         Left            =   2910
         Pattern         =   "*.txt"
         TabIndex        =   2
         Top             =   630
         Width           =   3165
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
         Left            =   4860
         TabIndex        =   1
         ToolTipText     =   "Limpar o Caminho do Arquivo"
         Top             =   2100
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Diretório"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   390
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Caminho do arquivo"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   2340
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmCaminho_balanca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Encontrar Caminho                                              '
' Data de Criação........: 17/12/2004                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSql As String
Dim strCaminho As String
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim I As Integer
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim strArquivo As String
'------------------------------------------------------------
'Declaração da variavel do intercomunicador de mensagens
Private Cliente_mensagem_exe As VetorDeMensagens.ClienteDeMensagens
Dim log As New DLLSystemManager.log

Private Sub cmdLimpar_Click()
    txtCaminho.Text = ""
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
       MsgBox "Não é recomendado que se importe arquivos diretamente do drive 'A:', salve-os em sua máquina antes.", vbInformation, "Only Tech"
       Drive1.SetFocus
    End If
End Sub

Private Sub File1_DblClick()
    txtCaminho.Text = Dir1.Path & "\" & File1.FileName
End Sub
Private Sub cmdOk_Click()

  frmMovimentacoes_exportacao_balancas.txtCaminho.Text = txtCaminho.Text
  txtCaminho.Text = ""
  Unload Me
  
End Sub

