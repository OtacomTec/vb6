VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogradouro_Solicitacao_Visita 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logradouro"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "frmLogradouro_Solicitacao_Visita.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10290
   Begin VB.Frame Frame1 
      Caption         =   "Consulta Detalhada"
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
      Height          =   1065
      Left            =   90
      TabIndex        =   32
      Top             =   1170
      Width           =   10125
      Begin VB.ComboBox cmbTipo_Logradouro 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmLogradouro_Solicitacao_Visita.frx":0FC2
         Left            =   3510
         List            =   "frmLogradouro_Solicitacao_Visita.frx":0FD8
         TabIndex        =   5
         Top             =   570
         Width           =   1335
      End
      Begin VB.ComboBox cmbEstado 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmLogradouro_Solicitacao_Visita.frx":1008
         Left            =   90
         List            =   "frmLogradouro_Solicitacao_Visita.frx":1060
         TabIndex        =   3
         Top             =   570
         Width           =   735
      End
      Begin VB.CommandButton cmdConsulta_especifica 
         Caption         =   "Consultar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   8970
         Picture         =   "frmLogradouro_Solicitacao_Visita.frx":10D3
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1065
      End
      Begin VB.ComboBox cmbFiltar_Especifica 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmLogradouro_Solicitacao_Visita.frx":145D
         Left            =   7830
         List            =   "frmLogradouro_Solicitacao_Visita.frx":1479
         TabIndex        =   7
         Top             =   570
         Width           =   1035
      End
      Begin MSDataListLib.DataCombo dtcLocalidade 
         Height          =   360
         Left            =   840
         TabIndex        =   4
         Top             =   570
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         ForeColor       =   8388608
         Text            =   ""
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
      Begin MSDataListLib.DataCombo dtcLogradouro 
         Height          =   360
         Left            =   4890
         TabIndex        =   6
         Top             =   570
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         ForeColor       =   8388608
         Text            =   ""
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
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   37
         Top             =   330
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tp. Logradouro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3510
         TabIndex        =   36
         Top             =   330
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Localidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   870
         TabIndex        =   35
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Logradouro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4950
         TabIndex        =   34
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7860
         TabIndex        =   33
         Top             =   330
         Width           =   510
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Resumo"
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
      Height          =   1785
      Left            =   90
      TabIndex        =   17
      Top             =   4980
      Width           =   10095
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
         Height          =   585
         Left            =   8910
         Picture         =   "frmLogradouro_Solicitacao_Visita.frx":14A2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   1065
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "Atualizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   8910
         Picture         =   "frmLogradouro_Solicitacao_Visita.frx":182C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label lblCep 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
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
         Height          =   240
         Left            =   4380
         TabIndex        =   31
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
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
         Height          =   240
         Left            =   750
         TabIndex        =   30
         Top             =   1410
         Width           =   660
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         Height          =   240
         Left            =   4620
         TabIndex        =   29
         Top             =   690
         Width           =   735
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tp. Logradouro:"
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
         Height          =   240
         Left            =   1560
         TabIndex        =   28
         Top             =   330
         Width           =   1545
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Height          =   240
         Left            =   750
         TabIndex        =   27
         Top             =   690
         Width           =   600
      End
      Begin VB.Label lblComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
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
         Height          =   240
         Left            =   1470
         TabIndex        =   26
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label lblLocalidade 
         AutoSize        =   -1  'True
         Caption         =   "Localidade:"
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
         Height          =   240
         Left            =   4980
         TabIndex        =   25
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Localidade:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3900
         TabIndex        =   24
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   1050
         Width           =   1260
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   690
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tp. Logradouro:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3900
         TabIndex        =   20
         Top             =   690
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3900
         TabIndex        =   18
         Top             =   1050
         Width           =   405
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Resultado"
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
      Height          =   2655
      Left            =   90
      TabIndex        =   16
      Top             =   2280
      Width           =   10125
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgResultado 
         Height          =   2325
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   4101
         _Version        =   393216
         FixedCols       =   0
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
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Opções de Consulta"
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
      Height          =   1095
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   6855
      Begin VB.OptionButton optConsulta_CEP 
         Caption         =   "Consulta por CEP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4260
         TabIndex        =   13
         Top             =   330
         Width           =   1815
      End
      Begin VB.OptionButton optConsulta_Especifica 
         Caption         =   "Consulta Detalhada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   1965
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Consulta por CEP"
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
      Height          =   1095
      Left            =   6930
      TabIndex        =   14
      Top             =   30
      Width           =   3285
      Begin VB.CommandButton cmdConsultar_cep 
         Caption         =   "Consultar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2100
         Picture         =   "frmLogradouro_Solicitacao_Visita.frx":286E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox txtcep 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   1
         Top             =   570
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   15
         Top             =   330
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmLogradouro_Solicitacao_Visita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Consultar Logradouros de Solicitação Visitas                   '
' Data de Criação........: 07/05/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strCampo_consulta As String
Dim strSql As String
Dim strCaptions As String
Dim strTamanho_colunas As String
Dim intCont As Integer
Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmbTipo_Logradouro_LostFocus()
     If cmbEstado.Text = "" Then Exit Sub
     If cmbTipo_Logradouro.Text = "" Then Exit Sub
     If dtcLocalidade.BoundColumn = "" Then Exit Sub
     
     frmAguarde.Show

     strSql = Empty
     strSql = "SELECT LOG_NU_SEQUENCIAL,LOG_NOME FROM LOG_LOGRADOURO " & _
              "WHERE UFE_SG = '" & Me.cmbEstado.Text & "'" & _
              "AND LOG_TIPO_LOGRADOURO = '" & cmbTipo_Logradouro.Text & "'" & _
              "AND LOC_NU_SEQUENCIAL = " & Me.dtcLocalidade.BoundColumn & " "
              
     Call Movimentacoes.Movimenta_DataCombo("LOG_NU_SEQUENCIAL", "LOG_NOME", dtcLogradouro, strSql, "BDGPB", "Otica", Me)
     
     Unload frmAguarde
     dtcLogradouro.SetFocus
End Sub

Private Sub cmdAtualizar_Click()
    frmSolicitacao_Visita.txtBairro.Text = lblBairro.Caption
    frmSolicitacao_Visita.txtcep.Text = lblCep.Caption
    frmSolicitacao_Visita.txtComplemento.Text = lblComplemento.Caption
    frmSolicitacao_Visita.txtEndereco.Text = lblNome.Caption
    frmSolicitacao_Visita.dtcCidade.Text = lblLocalidade.Caption
    frmSolicitacao_Visita.txtUf.Text = lblEstado.Caption
    frmSolicitacao_Visita.dtcCidade.SetFocus
    Unload Me
End Sub

Private Sub cmdConsulta_especifica_Click()
    'Removendo linhas do grid, evitando assim que fiquem linhas em branco.
    hfgResultado.ClearStructure
    Do While hfgResultado.Rows <= hfgResultado.Rows + 1
       hfgResultado.Col = 1
       If hfgResultado.Text = "" And hfgResultado.Rows = 2 Then
          Exit Do
       End If
       hfgResultado.Row = hfgResultado.Rows - 1
       hfgResultado.RemoveItem hfgResultado.Rows - 1
    Loop
    
    If cmbEstado.Text = "" Or cmbTipo_Logradouro.Text = "" Or dtcLocalidade.BoundColumn = "" Then Exit Sub
      
    frmAguarde.Show
    DoEvents
    Call Reposicao
    Unload frmAguarde
    cmbEstado.SetFocus
End Sub

Private Sub cmdConsultar_cep_Click()
    'Removendo linhas do grid, evitando assim que fiquem linhas em branco.
    hfgResultado.ClearStructure
    Do While hfgResultado.Rows <= hfgResultado.Rows + 1
       hfgResultado.Col = 1
       If hfgResultado.Text = "" And hfgResultado.Rows = 2 Then
          Exit Do
       End If
       hfgResultado.Row = hfgResultado.Rows - 1
       hfgResultado.RemoveItem hfgResultado.Rows - 1
    Loop
    
    If txtcep.Text = Empty Then Exit Sub
    
    frmAguarde.Show
    DoEvents
    Call Reposicao
    Unload frmAguarde
    txtcep.SetFocus
End Sub

Private Sub dtcLocalidade_GotFocus()
    strSql = Empty
    strSql = "SELECT LOC_NU_SEQUENCIAL,LOC_NO FROM LOG_LOCALIDADE WHERE UFE_SG_LOCALIDADE = '" & Me.cmbEstado.Text & "'"
    Call Movimentacoes.Movimenta_DataCombo("LOC_NU_SEQUENCIAL", "LOC_NO", dtcLocalidade, strSql, "BDGPB", "Otica", Me)
End Sub

Private Sub Form_Load()
    intCont = 1
    hfgResultado.Clear
    'Limpando campos
    dtcLocalidade.Text = Empty
    dtcLogradouro.Text = Empty
    txtcep.Text = Empty
    'Frame de consulta especifica
    Frame1.Enabled = False
    cmbEstado.Enabled = False
    cmbTipo_Logradouro.Enabled = False
    cmdConsulta_especifica.Enabled = False
    dtcLocalidade.Enabled = True
    dtcLogradouro.Enabled = True
    cmdConsulta_especifica.Enabled = False
    Label8.Enabled = False
    cmbFiltar_Especifica.Enabled = False
    'Frame de Consulta por CEP
    optConsulta_CEP.Value = True
    Frame3.Enabled = True
    txtcep.Enabled = True
    cmdConsultar_cep.Enabled = True
    Label5.Enabled = True
End Sub

Private Sub hfgResultado_Click()
    
   If hfgResultado.Col = 0 Then
        On Error Resume Next
        lblTipo.Caption = hfgResultado.TextArray((hfgResultado.Row * hfgResultado.Cols + hfgResultado.Col + 1))
        lblNome.Caption = hfgResultado.TextArray((hfgResultado.Row * hfgResultado.Cols + hfgResultado.Col + 3))
        lblComplemento.Caption = hfgResultado.TextArray((hfgResultado.Row * hfgResultado.Cols + hfgResultado.Col + 4))
        lblBairro.Caption = hfgResultado.TextArray((hfgResultado.Row * hfgResultado.Cols + hfgResultado.Col + 5))
        lblLocalidade.Caption = hfgResultado.TextArray((hfgResultado.Row * hfgResultado.Cols + hfgResultado.Col + 6))
        lblEstado.Caption = hfgResultado.TextArray((hfgResultado.Row * hfgResultado.Cols + hfgResultado.Col + 7))
        lblCep.Caption = Format(hfgResultado.TextArray((hfgResultado.Row * hfgResultado.Cols + hfgResultado.Col + 8)), "#####-###")
   End If
    
End Sub

Private Sub optConsulta_CEP_Click()
    'Removendo linhas do grid, evitando assim que fiquem linhas em branco.
    hfgResultado.ClearStructure
    Do While hfgResultado.Rows <= hfgResultado.Rows + 1
       hfgResultado.Col = 1
       If hfgResultado.Text = "" And hfgResultado.Rows = 2 Then
          Exit Do
       End If
       hfgResultado.Row = hfgResultado.Rows - 1
       hfgResultado.RemoveItem hfgResultado.Rows - 1
    Loop
    
    intCont = 1
    hfgResultado.Clear
    'Limpando campos
    dtcLocalidade.Text = Empty
    dtcLogradouro.Text = Empty
    txtcep.Text = Empty
    'Frame de consulta especifica
    Frame1.Enabled = False
    cmbEstado.Enabled = False
    cmbTipo_Logradouro.Enabled = False
    cmdConsulta_especifica.Enabled = False
    dtcLocalidade.Enabled = False
    dtcLogradouro.Enabled = False
    cmdConsulta_especifica.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Label8.Enabled = False
    cmbFiltar_Especifica.Enabled = False
    'Frame de Consulta por CEP
    Frame3.Enabled = True
    txtcep.Enabled = True
    cmdConsultar_cep.Enabled = True
    Label5.Enabled = True
End Sub

Private Sub optConsulta_Especifica_Click()
    'Removendo linhas do grid, evitando assim que fiquem linhas em branco.
    hfgResultado.ClearStructure
    Do While hfgResultado.Rows <= hfgResultado.Rows + 1
       hfgResultado.Col = 1
       If hfgResultado.Text = "" And hfgResultado.Rows = 2 Then
          Exit Do
       End If
       hfgResultado.Row = hfgResultado.Rows - 1
       hfgResultado.RemoveItem hfgResultado.Rows - 1
    Loop
    
    Me.hfgResultado.Clear
    intCont = 1
    'Limpando campos
    dtcLocalidade.Text = Empty
    dtcLogradouro.Text = Empty
    txtcep.Text = Empty
    'Frame de consulta especifica
    cmbFiltar_Especifica.ListIndex = 0
    Frame1.Enabled = True
    cmbEstado.Enabled = True
    cmbTipo_Logradouro.Enabled = True
    cmdConsulta_especifica.Enabled = True
    dtcLocalidade.Enabled = True
    dtcLogradouro.Enabled = True
    cmdConsulta_especifica.Enabled = True
    Label1.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label8.Enabled = True
    cmbFiltar_Especifica.Enabled = True
    'Frame de Consulta por CEP
    Frame3.Enabled = False
    txtcep.Enabled = False
    cmdConsultar_cep.Enabled = False
    Label5.Enabled = False
End Sub

Private Function Reposicao()
    'Removendo linhas do grid, evitando assim que fiquem linhas em branco.
    hfgResultado.ClearStructure
    Do While hfgResultado.Rows <= hfgResultado.Rows + 1
       hfgResultado.Col = 1
       If hfgResultado.Text = "" And hfgResultado.Rows = 2 Then
          Exit Do
       End If
       hfgResultado.Row = hfgResultado.Rows - 1
       hfgResultado.RemoveItem hfgResultado.Rows - 1
    Loop
    
    strSql = Empty
   
   If Me.cmbFiltar_Especifica <> "" Then
       If cmbFiltar_Especifica.Text = "Tudo" Then
          strSql = strSql + "SELECT "
       Else
          strSql = strSql + "SELECT TOP " & cmbFiltar_Especifica.Text & " "
       End If
    Else
       strSql = "SELECT "
    End If
    
    strSql = strSql + "LOG_LOGRADOURO.LOG_TIPO_LOGRADOURO," & _
                      "LOG_LOGRADOURO.LOG_NO," & _
                      "LOG_LOGRADOURO.LOG_NOME," & _
                      "LOG_LOGRADOURO.LOG_COMPLEMENTO," & _
                      "LOG_BAIRRO.BAI_NO," & _
                      "LOG_LOCALIDADE.LOC_NO," & _
                      "LOG_LOGRADOURO.UFE_SG," & _
                      "LOG_LOGRADOURO.Cep " & _
                      "FROM LOG_LOGRADOURO " & _
                      "INNER JOIN LOG_BAIRRO " & _
                      "ON LOG_LOGRADOURO.BAI_NU_SEQUENCIAL_INI = LOG_BAIRRO.BAI_NU_SEQUENCIAL " & _
                      "INNER JOIN LOG_LOCALIDADE " & _
                      "ON LOG_LOCALIDADE.LOC_NU_SEQUENCIAL = LOG_BAIRRO.LOC_NU_SEQUENCIAL "
             
    If txtcep.Text <> "" Then
       strSql = strSql + "WHERE LOG_LOGRADOURO.Cep = " & Me.txtcep.Text & " "
    End If
    
    'Montando o WHERE da consulta especifíca
    If Frame1.Enabled = True Then
       strSql = strSql + "WHERE "
       If cmbEstado.Text <> "" Then
          strSql = strSql + "LOG_LOGRADOURO.UFE_SG =  '" & cmbEstado.Text & "' "
       End If
       If dtcLocalidade.Text <> "" Then
          strSql = strSql + "AND LOG_LOCALIDADE.LOC_NO = '" & dtcLocalidade.Text & "'"
       End If
       If cmbTipo_Logradouro.Text <> "" Then
          strSql = strSql + " AND LOG_LOGRADOURO.LOG_TIPO_LOGRADOURO = '" & cmbTipo_Logradouro.Text & "'"
      End If
       If Me.dtcLogradouro.Text <> "" Then
          strSql = strSql + " AND LOG_LOGRADOURO.LOG_NOME = '" & dtcLogradouro.Text & "'"
       End If
    End If
    
    strTamanho_colunas = "1000,2000,3000,2000,1500,1500,500,1000"
    strCaptions = "Tipo Logradouro,Nome Logradouro(res),Nome Logradouro,Complemento,Bairro,Cidade,UF,CEP"
    
    strSql = strSql + " ORDER BY LOG_LOGRADOURO.LOG_NOME"
     
    Call Movimentacoes.Movimenta_HFlex_Grid(strSql, Me.hfgResultado, strTamanho_colunas, strCaptions, "BDGPB", "Otica", Me)
    
    Me.txtcep.Text = Empty
    Me.dtcLocalidade.Text = Empty
    Me.dtcLogradouro.Text = Empty
    Me.cmbEstado.Text = Empty
    Me.cmbFiltar_Especifica.Text = Empty
    Me.cmbTipo_Logradouro.Text = Empty
End Function

Private Sub txtcep_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub


