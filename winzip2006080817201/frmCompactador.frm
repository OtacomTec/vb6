VERSION 5.00
Begin VB.Form frmCompactador 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compactador / Descompactador"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtZip 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Text            =   "Esta TextBox é usado internamente pelo AddZip, podendo fica escondida do usuário."
      Top             =   2835
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.CommandButton cmdDescompacta 
      Caption         =   "&Descompactar"
      Height          =   540
      Left            =   3255
      TabIndex        =   7
      Top             =   1575
      Width           =   2850
   End
   Begin VB.CommandButton cmdCompactar 
      Caption         =   "&Compactar"
      Height          =   540
      Left            =   105
      TabIndex        =   6
      ToolTipText     =   "Compacta os arquivos"
      Top             =   1575
      Width           =   2850
   End
   Begin VB.TextBox txtExtrairPara 
      Height          =   330
      Left            =   2520
      TabIndex        =   5
      Text            =   "C:\Textos\"
      Top             =   1050
      Width           =   3585
   End
   Begin VB.TextBox txtArqCompactado 
      Height          =   330
      Left            =   2520
      TabIndex        =   4
      Text            =   "Textos.zip"
      Top             =   525
      Width           =   3585
   End
   Begin VB.TextBox txtArqOrigem 
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Text            =   "*.txt"
      Top             =   0
      Width           =   3585
   End
   Begin VB.Label lblProgresso 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1050
      TabIndex        =   10
      Top             =   2415
      Width           =   5055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Progresso"
      Height          =   195
      Left            =   105
      TabIndex        =   9
      Top             =   2483
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      X1              =   0
      X2              =   6720
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Local para descompactação"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   1118
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Nome do arquivo compactado"
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   593
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Arquivo(s) para compactação"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   68
      Width           =   2085
   End
End
Attribute VB_Name = "frmCompactador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub cmdCompactar_Click()
    Dim strArquivo_comp As String
    Dim strArquivo_add_comp As String
    
    strArquivo_comp = Funcoes_Gerais.Abrir_path_fva_registro("Otica", Me) & "\" & 2 & "\bdfva.zip"
    strArquivo_add_comp = Funcoes_Gerais.Abrir_path_fva_registro("Otica", Me) & "\" & 2 & "\bdfva.*"
    
    Compacta strArquivo_comp, strArquivo_add_comp

    'Compacta txtArqCompactado, txtArqOrigem
    
    lblProgresso = ""

End Sub

Private Sub cmdDescompacta_Click()
DesCompacta txtArqCompactado, "*.*", txtExtrairPara, False
End Sub

Private Sub Form_Load()
'É necessário inicializar a biblioteca de compactação.
InicializaZip Me, txtZip
End Sub

Private Sub txtZip_Change()
'Indicação de progresso da compactação/descompactação por arquivo
'----------------------------------------------------------------
'Tipo de ação que esta sendo feita no momento
lblProgresso = TipoAção(Val(GetAction(txtZip))) & " "
'Nome do arquivo que esta sendo compactado
lblProgresso = lblProgresso & GetFileName(txtZip) & " -> "
'Porcentagem de compactação do arquivo
lblProgresso = lblProgresso & GetPercentComplete(txtZip) & "%"
'Força a atualização da tela
DoEvents
End Sub
