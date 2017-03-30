VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMarcha_Analitica_Acompanhamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acompanhamento de Marcha"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMarcha_Analitica_Acompanhamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7800
   Begin VB.TextBox txtNumero_Sequencial 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   "Peso"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtDescricao_Tipo_Marcha 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1845
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Nº da Nota Fiscal"
      Top             =   240
      Width           =   5865
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgAcompanhamento 
      Height          =   2655
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   4683
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "Marcha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   525
   End
End
Attribute VB_Name = "frmMarcha_Analitica_Acompanhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Transportes                                                    '
' Objetivo...............: Informções de Acompanhamento Marcha                            '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Sá                                                       '
' Data de Criação........: 30/12/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim strSql As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim intContador As Integer
    
    Movimentacoes.Monta_HFlex_Grid Me.hfgAcompanhamento, "2000,1000,1000,1000,1000,1400", "Estágio,Dt.Entrada,Hr.Entrada,Dt.Saída,Hr.Saída,Usuário", 6, "Otica", Me
    
    strSql = "SELECT DFEstagio_TBAcompanhamento_marcha," & _
             "DFData_inicio_TBAcompanhamento_marcha," & _
             "DFHora_inicio_TBAcompanhamento_marcha," & _
             "DFData_fim_TBAcompanhamento_marcha," & _
             "DFHora_fim_TBAcompanhamento_marcha," & _
             "DFUsuario_DFHora_inicio_TBAcompanhamento_marcha " & _
             "FROM TBAcompanhamento_marcha " & _
             "WHERE FKId_TBMarcha = '" & frmMarcha_Analitica.strID_Marcha & "'"
    
    Movimenta_HFlex_Grid strSql, hfgAcompanhamento, "2000,900,900,900,900,1500", "Estágio,Dt.Entrada,Hr.Entrada,Dt.Saída,Hr.Saída,Usuário", "BDRetaguarda", "Otica", Me, "N"
    hfgAcompanhamento.Col = 0
    hfgAcompanhamento.Row = 1
    If hfgAcompanhamento.Text = Empty Then
       hfgAcompanhamento.Rows = 2
       Movimentacoes.Monta_HFlex_Grid Me.hfgAcompanhamento, "2000,1000,1000,900,900,1580", "Estágio,Dt.Entrada,Hr.Entrada,Dt.Saída,Hr.Saída,Usuário", 6, "Otica", Me
    Else
       intContador = 1
       hfgAcompanhamento.Col = 1
       Do While intContador <= hfgAcompanhamento.Rows - 1
          hfgAcompanhamento.Row = intContador
          If hfgAcompanhamento.Text = 1 Then
             hfgAcompanhamento.Text = "Recebimento"
          ElseIf hfgAcompanhamento.Text = 2 Then
             hfgAcompanhamento.Text = "Triagem"
          ElseIf hfgAcompanhamento.Text = 3 Then
             hfgAcompanhamento.Text = "Amostragem"
          ElseIf hfgAcompanhamento.Text = 4 Then
             hfgAcompanhamento.Text = "Laboratório"
          ElseIf hfgAcompanhamento.Text = 5 Then
             hfgAcompanhamento.Text = "Micro"
          ElseIf hfgAcompanhamento.Text = 6 Then
             hfgAcompanhamento.Text = "Físico-Químico"
          ElseIf hfgAcompanhamento.Text = 7 Then
             hfgAcompanhamento.Text = "Digitação"
          End If
          intContador = intContador + 1
       Loop
    End If
    
    txtNumero_Sequencial.Text = frmMarcha_Analitica.txtNumero_Sequencial.Text
    txtDescricao_Tipo_Marcha.Text = frmMarcha_Analitica.dtcTipo_Marcha.Text
    hfgAcompanhamento.Col = 0
    hfgAcompanhamento.Row = 1
End Sub
