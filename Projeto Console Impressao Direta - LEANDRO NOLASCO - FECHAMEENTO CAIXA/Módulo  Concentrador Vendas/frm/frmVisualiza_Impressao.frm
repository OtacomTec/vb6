VERSION 5.00
Begin VB.Form frmVisualiza_Impressao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizador de Impressão"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmVisualiza_Impressao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   9435
      TabIndex        =   6
      Top             =   6660
      Width           =   2460
      Begin VB.CommandButton cmdNavegacao 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   435
         TabIndex        =   10
         ToolTipText     =   "Primeira Pág."
         Top             =   435
         Width           =   375
      End
      Begin VB.CommandButton cmdNavegacao 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   795
         TabIndex        =   9
         ToolTipText     =   "Pág. Anterior"
         Top             =   435
         Width           =   375
      End
      Begin VB.CommandButton cmdNavegacao 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1155
         TabIndex        =   8
         ToolTipText     =   "Próxima Pág."
         Top             =   435
         Width           =   375
      End
      Begin VB.CommandButton cmdNavegacao 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1515
         TabIndex        =   7
         ToolTipText     =   "Última Pág."
         Top             =   435
         Width           =   375
      End
      Begin VB.Label lblPagina 
         Alignment       =   2  'Center
         Caption         =   "Página 1 de 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   435
         TabIndex        =   11
         Top             =   120
         Width           =   1470
      End
   End
   Begin VB.CommandButton cmdConfirma_Impressao 
      Caption         =   "I&mprimir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   90
      Picture         =   "frmVisualiza_Impressao.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir"
      Top             =   6735
      Width           =   1000
   End
   Begin VB.CommandButton cmdFehcar_Resultados 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1095
      Picture         =   "frmVisualiza_Impressao.frx":1A8C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sair ou Fechar"
      Top             =   6735
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   6585
      Left            =   90
      TabIndex        =   2
      Top             =   30
      Width           =   11805
      Begin VB.HScrollBar vschVisualiza 
         Height          =   240
         LargeChange     =   1000
         Left            =   90
         Max             =   0
         SmallChange     =   200
         TabIndex        =   12
         Top             =   6210
         Width           =   11400
      End
      Begin VB.Frame fraImpressao 
         BorderStyle     =   0  'None
         Height          =   5955
         Left            =   90
         TabIndex        =   4
         Top             =   135
         Width           =   11400
         Begin VB.PictureBox picVisualiza 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   5865
            Left            =   0
            ScaleHeight     =   5835
            ScaleWidth      =   11370
            TabIndex        =   5
            Top             =   45
            Width           =   11400
            Begin VB.Image imgGrafico 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   915
               Index           =   0
               Left            =   4950
               Top             =   120
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Image imgLogo 
               Height          =   720
               Left            =   180
               Picture         =   "frmVisualiza_Impressao.frx":1D96
               Stretch         =   -1  'True
               Top             =   135
               Visible         =   0   'False
               Width           =   720
            End
         End
      End
      Begin VB.VScrollBar vscVisualiza 
         Height          =   6315
         LargeChange     =   1000
         Left            =   11520
         Max             =   0
         SmallChange     =   200
         TabIndex        =   3
         Top             =   135
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVisualiza_Impressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim int_Pag_Atual As Integer
Dim int_Ultima As Integer

Private Sub cmdConfirma_Impressao_Click()

    If (MsgBox("Será(ão) impressa(s) " & Impressao.Numero_Paginas & " página(s). Deseja realmente imprimir?", vbYesNo Or vbQuestion Or _
                 vbDefaultButton1, wNomeSistema) = vbYes) Then
        'seta impressao para impressora
        Impressao.Destino = True
        Impressao.Executa_Impressao 1
        Impressao.Destino = False
        
    End If
    
End Sub

Private Sub cmdFehcar_Resultados_Click()
    
    Unload Me
    
End Sub


Private Sub cmdNavegacao_Click(Index As Integer)
    
    Select Case Index
    
        Case 0
            int_Pag_Atual = 1
            Impressao.Executa_Impressao int_Pag_Atual
            
        Case 1
            int_Pag_Atual = IIf(int_Pag_Atual = 1, 1, int_Pag_Atual - 1)
            Impressao.Executa_Impressao int_Pag_Atual

        Case 2
            int_Pag_Atual = IIf(int_Pag_Atual = int_Ultima, int_Ultima, int_Pag_Atual + 1)
            Impressao.Executa_Impressao int_Pag_Atual

        Case 3
            int_Pag_Atual = Impressao.Numero_Paginas
            Impressao.Executa_Impressao int_Pag_Atual

    End Select
    
    Habilita_Botoes
    
End Sub

Private Sub Form_Activate()

    int_Pag_Atual = 1
    Impressao.Executa_Impressao (int_Pag_Atual)
    int_Ultima = Impressao.Numero_Paginas
    
    Habilita_Botoes

End Sub

Private Sub Form_Load()
    
'    CentralizaForm Me
'    Implementa_Acesso Me
    
End Sub

Private Sub picVisualiza_Resize()

'    shpBordaPagina.Height = picVisualiza.Height - 1675
'    shpBordaPagina.width = picVisualiza.width - 425

End Sub

Private Sub vschVisualiza_Change()

    picVisualiza.Left = -(vschVisualiza.Value)
    
End Sub

Private Sub vscVisualiza_Change()

    picVisualiza.Top = -(vscVisualiza.Value)
        
End Sub

Private Sub Habilita_Botoes()
    
    cmdNavegacao(0).Enabled = int_Pag_Atual <> 1
    cmdNavegacao(1).Enabled = int_Pag_Atual <> 1
    cmdNavegacao(2).Enabled = int_Pag_Atual <> int_Ultima
    cmdNavegacao(3).Enabled = int_Pag_Atual <> int_Ultima
    
End Sub
