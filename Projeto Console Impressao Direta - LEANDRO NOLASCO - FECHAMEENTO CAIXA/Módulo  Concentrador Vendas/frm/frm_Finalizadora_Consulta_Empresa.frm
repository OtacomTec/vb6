VERSION 5.00
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmFinalizadora_Consulta_Empresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresa para Consulta Finalizadora"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   6015
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
      Height          =   360
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
   Begin AutoCompletar.CbCompleta cbbEmpresa 
      Height          =   360
      Left            =   90
      TabIndex        =   1
      Top             =   300
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   750
   End
End
Attribute VB_Name = "frmFinalizadora_Consulta_Empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' M�dulo.................: Compras                                                        '
' Objetivo...............: Empresa para Consulta Finalizadora                             '
' Equipe Respons�vel.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rodrigo Santos                                                 '
' Data de Cria��o........: /06/2006                                                     '
' Desenvolvedor..........:                                                                '
' Data �ltima manuten��o.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Option Explicit

Private Sub cmdOk_Click()
    Dim rstEmpresa_Codigo As New ADODB.Recordset
    
    strSql = "SELECT PKCodigo_TBEmpresa FROM TBEmpresa WHERE DFRazao_Social_TBEmpresa = '" & cbbEmpresa.Text & "' "
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstEmpresa_Codigo, "Otica", Me
    
    If rstEmpresa_Codigo.RecordCount <> 0 Then
       frmFinalizadora.strCodigo_Empresa_Consulta = rstEmpresa_Codigo.Fields!PKCodigo_TBEmpresa
    Else
       frmFinalizadora.strCodigo_Empresa_Consulta = 0
    End If
    
    Set rstEmpresa_Codigo = Nothing
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rstEmpresa_Razao_Social As New ADODB.Recordset
    
    strSql = "SELECT DFRazao_Social_TBEmpresa FROM TBEmpresa "
    
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstEmpresa_Razao_Social, "Otica", Me
    
    If rstEmpresa_Razao_Social.RecordCount <> 0 Then
       cbbEmpresa.Clear
       cbbEmpresa.AddItem ("TODOS")
       
       rstEmpresa_Razao_Social.MoveFirst
       
       Do While rstEmpresa_Razao_Social.EOF <> True
          cbbEmpresa.AddItem (rstEmpresa_Razao_Social.Fields!DFRazao_Social_TBEmpresa)
          
          rstEmpresa_Razao_Social.MoveNext
       Loop
    End If
    
    Set rstEmpresa_Razao_Social = Nothing
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub




