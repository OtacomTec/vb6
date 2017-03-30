VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmContrato_Servico_Informacoes_Adicionais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informações Adicionais"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContrato_Servico_Informacoes_Adicionais.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8550
   Begin VB.TextBox txtDescricao 
      Enabled         =   0   'False
      Height          =   360
      Left            =   1620
      MaxLength       =   5
      TabIndex        =   5
      ToolTipText     =   "Código Plano de Serviços"
      Top             =   270
      Width           =   6855
   End
   Begin VB.TextBox txtPlano_Servico 
      Enabled         =   0   'False
      Height          =   360
      Left            =   90
      MaxLength       =   5
      TabIndex        =   3
      ToolTipText     =   "Código Plano de Serviços"
      Top             =   270
      Width           =   1485
   End
   Begin VB.Frame Frame3 
      Caption         =   "Serviços"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   90
      TabIndex        =   0
      Top             =   660
      Width           =   8385
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgServico 
         Height          =   1635
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   2884
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Plano de Serviço"
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   30
      Width           =   1425
   End
   Begin VB.Label lblDescricao 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   3000
      Width           =   8385
   End
End
Attribute VB_Name = "frmContrato_Servico_Informacoes_Adicionais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Informações Adicionais do Plano de Serviços                    '
' Data de Criação........: 27/01/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........: Jones Sá Peixoto                                               '
' Data última manutenção.: 23/08/2005                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim conexao As New DLLConexao_Sistema.conexao
Dim intIndice As Integer
Dim rstAplicacao As New ADODB.Recordset
Option Explicit

Private Sub Form_Activate()
    Dim strTamanho_servico As String
    Dim strNomes_servico As String
    
    lblDescricao.Caption = frmContrato_Servico.dtcPlano_servico.Text
    
    strTamanho_servico = "0,850,4000,1200,1500,1200,1300,1300,1300"
    strNomes_servico = "ID,Código,Descrição,Limite,Controle,Período,Pr. Conv. 1,Pr. Conv. 2,Pr. Conv. 3"
    
    txtPlano_Servico.Text = frmContrato_Servico.txtPlano_Servico.Text
    txtDescricao.Text = frmContrato_Servico.dtcPlano_servico.Text
    
    'Abastecendo Grid
    strSql = "SELECT PkId_TBPlano_servico_servico_laboratorio," & _
             "FKCodigo_TBServico_laboratorio,DFDescricao_TBServico_laboratorio," & _
             "DFQuantidade_TBPlano_servico_servico_laboratorio," & _
             "DFControle_TBPlano_servico_servico_laboratorio,DFPeriodo_TBPlano_servico_servico_laboratorio," & _
             "DFPreco1_conveniado_TBServico_laboratorio," & _
             "DFPreco2_conveniado_TBServico_laboratorio,DFPreco3_conveniado_TBServico_laboratorio " & _
             "FROM TBPlano_servico_servico_laboratorio " & _
             "INNER JOIN TBServico_laboratorio ON TBPlano_servico_servico_laboratorio.FKCodigo_TBServico_laboratorio = TBServico_laboratorio.pKCodigo_TBServico_laboratorio " & _
             "WHERE FKCodigo_TBPlano_servico = " & frmContrato_Servico.txtPlano_Servico.Text & " " & _
             "ORDER BY PkId_TBPlano_servico_servico_laboratorio"
    
    Call Movimentacoes.Movimenta_HFlex_Grid(strSql, hfgServico, strTamanho_servico, strNomes_servico, "BDRetaguarda", "Otica", Me)
    
    Dim intContador As Integer
    hfgServico.Row = 1
    hfgServico.Col = 0
    If hfgServico.Text = Empty Then
       hfgServico.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgServico, strTamanho_servico, strNomes_servico, 9, "Otica", Me
    Else
       intContador = 1
       hfgServico.Col = 5
       Do While intContador <= hfgServico.Rows - 1
          hfgServico.Row = intContador
          If hfgServico.Text = "1" Then
             hfgServico.Text = "Valor Contrato"
          ElseIf hfgServico.Text = "2" Then
             hfgServico.Text = "Serviços"
          ElseIf hfgServico.Text = "3" Then
             hfgServico.Text = "Grupo Serviços"
          End If
          intContador = intContador + 1
       Loop
       hfgServico.ColAlignment(5) = 1
    End If
    
    Set rstAplicacao = Nothing
    
    hfgServico.Row = 1
    hfgServico.Col = 0
    hfgServico.SetFocus
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmContrato_Servico.txtBanco.SetFocus
End Sub
