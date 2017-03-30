VERSION 5.00
Begin VB.Form frmTeste_impressora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teste de impress�o"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTeste_impressora.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4110
   Begin VB.Frame Frame1 
      Caption         =   "Bematech N�o Fiscal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   90
      TabIndex        =   2
      Top             =   30
      Width           =   3945
      Begin VB.CommandButton cmdTestar 
         Caption         =   "Testar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2670
         TabIndex        =   0
         ToolTipText     =   "Testar Impressora"
         Top             =   330
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bematech n�o Fiscal"
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmTeste_impressora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' M�dulo.................: Concentrador Vendas                                            '
' Objetivo...............: Teste Impressora                                               '
' Data de Cria��o........: 22/06/2005                                                     '
' Equipe Respons�vel.....: Only Tech Solutions                                            '
' �ltima Manuten��o......:                                                                '
' Desenvolvedor..........:                                                                '
' Data �ltima manuten��o.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTestar_Click()
    '-------------------------------------------------------------------------------------------------------
    'Abrindo Impressora n�o fiscal
    '-------------------------------------------------------------------------------------------------------
    Dim intPorta As Integer
    Dim strComunica As String
    
    ' Fecha a porta que est� aberta
    intPorta = FechaPorta()
    If intPorta <= 0 Then
       MsgBox "Problemas ao Fechar a Porta de Comunica��o com a imp. n�o fiscal.Reinicie a aplica��o", vbCritical, "Only Tech"
    End If

    ' Abre a porta de comunicacao com imp. n�o fiscal
    intPorta = IniciaPorta("LPT1")
    If intPorta <= 0 Then
       MsgBox "Problemas ao Abrir a Porta de Comunica��o com a imp. n�o fiscal.Reinicie a aplica��o", vbCritical, "Only Tech"
    End If
    
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
        
    strLinha_Impressao = "T E S T E   D E   I M P R E S S � O"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
    
    strLinha_Impressao = "T E S T E   D E   I M P R E S S � O"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
    
    strLinha_Impressao = "T E N H A    U M    B O M     T R A B A L H O"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 3, 0, 0, 0, 0)
    
    strLinha_Impressao = "-----------------------------------------------------------"
    sBuffer = strLinha_Impressao + Chr(13) + Chr(10)
    iretorno = FormataTX(sBuffer, 1, 0, 0, 0, 0)
    
    MsgBox "Teste efetuado com sucesso !!!", vbInformation, "Only Tech"
    
End Sub
