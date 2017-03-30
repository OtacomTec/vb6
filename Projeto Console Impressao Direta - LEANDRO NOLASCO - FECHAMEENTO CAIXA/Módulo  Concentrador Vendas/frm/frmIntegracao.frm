VERSION 5.00
Begin VB.Form frmIntegracao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Integração"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2610
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIntegracao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   2610
   Begin VB.Frame Frame3 
      Height          =   525
      Left            =   90
      TabIndex        =   6
      Top             =   30
      Width           =   2415
      Begin VB.Label lblForm_solicitante 
         Alignment       =   2  'Center
         Caption         =   "Form Solicitante"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   7
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Legenda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   2
      Top             =   1530
      Width           =   2415
      Begin VB.Shape Shape3 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   225
         Left            =   120
         Shape           =   3  'Circle
         Top             =   300
         Width           =   165
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   225
         Left            =   120
         Shape           =   3  'Circle
         Top             =   630
         Width           =   165
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Integrado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   4
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Não Integrado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   3
         Top             =   630
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   2415
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Portal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   5
         Top             =   570
         Width           =   1695
      End
      Begin VB.Shape shpStatus_Portal 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   225
         Left            =   120
         Shape           =   3  'Circle
         Top             =   570
         Width           =   165
      End
      Begin VB.Shape shpStatus_Retaguarda 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   225
         Left            =   120
         Shape           =   3  'Circle
         Top             =   270
         Width           =   165
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Retaguarda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   1
         Top             =   270
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmIntegracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strForm_Ativo As String
Private objTemp As Object
Dim Form_Ativo As String

Private Sub tmrTemporizador_Timer()
    Unload Me
End Sub

Public Function Verifica_Integracao(ID_Campo As String, Valor_ID_Campo As String, Campo_Integracao_Retaguarda As String, Tabela As String, Aplicacao As String, Banco As String, Optional Campo_Integracao_Portal As String, Optional Top As Integer, Optional Left As Integer, Optional Width As Integer, Optional Height As Integer, Optional Caption_Form_Max23Carct As String)
        
    Dim strSql As String
    Dim rstIntegracao As New ADODB.Recordset
    Dim conexao_Integracao As New DLLConexao_Sistema.conexao
           
    'INDICANDO O BANCO A CONECTAR-SE
    conexao_Integracao.Initial_Catalog = Banco
    
    'ABRINDO CONEXAO COM BANCO
    conexao_Integracao.Abrir_conexao (Aplicacao)
    
    DoEvents
    
    rstIntegracao.CursorLocation = adUseClient
    
    'STRING QUE COLETA DADOS RELATIVOS A INTEGRAÇÃO DO REGISTRO
    If Campo_Integracao_Portal = Empty Then
       strSql = "SELECT " & Campo_Integracao_Retaguarda & " FROM " & Tabela & _
                " WHERE " & ID_Campo & " = " & Valor_ID_Campo & ""
    Else
       strSql = "SELECT  " & Campo_Integracao_Retaguarda & "," & Campo_Integracao_Portal & _
                " FROM " & Tabela & " WHERE " & ID_Campo & " = " & Valor_ID_Campo & ""
    End If
             
    rstIntegracao.Open strSql, conexao_Integracao.CNConexao, adOpenStatic, adLockReadOnly
    
    rstIntegracao.MoveFirst
    Me.Show
    
    'VERIFICANDO SE INTEGRADO
    If rstIntegracao.BOF = False Then
       
       'RETAGUARDA
       If rstIntegracao(Campo_Integracao_Retaguarda) = True Then
          shpStatus_Retaguarda.BackColor = &HFF0000   'Azul
       Else
          shpStatus_Retaguarda.BackColor = &HFF&      'Vermelho
       End If
       
       'PORTAL
       If Campo_Integracao_Portal <> Empty Then
          If rstIntegracao(Campo_Integracao_Portal) = True Then
             shpStatus_Portal.BackColor = &HFF0000    'Azul
          Else
             shpStatus_Portal.BackColor = &HFF&       'Vermelho
          End If
       End If
       
    End If
      
    Set rstIntegracao = Nothing
    
    conexao_Integracao.Fechar_conexao
    
    DoEvents
    
    'POSICIONA FORM
    If Width > 9500 Then
       'POSICIONA FORM INTEGRAÇÃO AO LADO DO FORM SOLICITANTE
       Me.Left = Left + (Width / 2) - 1350
       Me.Top = (Top + (Height / 2)) - 1477
    Else
       Me.Left = Left + Width + 60
       Me.Top = Top
       'O valor 2445 é a altura do Form Integração
    End If
    
    lblForm_solicitante.Caption = Caption_Form_Max23Carct
    
    Exit Function
    
End Function

