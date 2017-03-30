VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAbertura_Caixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura de Caixa"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbertura_Caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFundo_Caixa 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Código do Cliente(Informado Automaticamente)"
      Top             =   270
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5130
      TabIndex        =   3
      Top             =   1020
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
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
      Height          =   405
      Left            =   5130
      TabIndex        =   2
      Top             =   570
      Width           =   1245
   End
   Begin MSDataListLib.DataCombo dtcFinalizadora 
      Height          =   360
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Finalizadora"
      Top             =   1050
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
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
   Begin VB.Label lblPDV 
      Caption         =   "PDV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5130
      TabIndex        =   7
      Top             =   270
      Width           =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo de Caixa"
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   30
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Finalizadora"
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   810
      Width           =   1035
   End
   Begin VB.Label lblOperador 
      Caption         =   "Operador:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1920
      TabIndex        =   4
      Top             =   270
      Width           =   3120
   End
End
Attribute VB_Name = "frmAbertura_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strOperador As String
Public strCodigo_Operador As String
Public strEmpresa_Operador As String
Public strPDV As String
Public strNumero_PDV As String
Dim strSQL As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim strID_Finalizadora As String
    Dim strCampos As String
    Dim strValores As String
    
    On Error GoTo Erro
    
    If Me.dtcFinalizadora.BoundText = "" Then
        MsgBox "Favor insira uma finalizadora válida", vbInformation, "Only Tech"
        dtcFinalizadora.SetFocus
        Exit Sub
    End If
    
    'Id Finalizadora
    strID_Finalizadora = Funcoes_Gerais.Localiza_ID("PKId_TBFinalizadora", "IXCodigo_TBFinalizadora", dtcFinalizadora.BoundText, "TBFinalizadora", "Otica", Me, "BDRetaguarda")

    strCampos = "FKCodigo_TBPdv,FKId_TBFinalizadora,FKCodigo_TBOperadores_ecf,DFData_TBOperacao_caixa," & _
               "DFHora_TBOperacao_caixa,DFValor_TBOperacao_caixa,DFTipo_operacao_TBOperacao_caixa,DFStatus_aberto_fechado_TBOperacao_caixa," & _
               "DFObservacao_TBOperacao_caixa"
              
    strValores = "" & strNumero_PDV & "," & _
                 "" & strID_Finalizadora & "," & _
                 "" & strCodigo_Operador & "," & _
                 "'" & Format(Now, "YYYYMMDD") & "'," & _
                 "'" & Format(Now, "hh:mm:ss") & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(Me.txtFundo_Caixa.Text) & "," & _
                 "1," & _
                 "0, '" & dtcFinalizadora.Text & "'"
    
    funcoes_banco.Gravar "TBoperacao_caixa", strCampos, strValores, "Otica", Me, "BDRetaguarda"
    
    MsgBox "Abertura de caixa feito com sucesso!", vbInformation, "Only Tech"
    
    frmTela_Venda.strOperador = strOperador
    frmTela_Venda.strPDV = strPDV
    frmTela_Venda.strEmpresa_Operador = strEmpresa_Operador
    frmTela_Venda.strCodigo_Operador = strCodigo_Operador
    
    frmTela_Venda.Show
    Unload Me
    
    Exit Sub
    
Erro:
    Call Erro.Erro(Me, "Otica")
    
End Sub

Private Sub Form_Load()

    'Operador
    lblOperador.Caption = "Operador: " & strOperador
    Me.lblPDV.Caption = "N° PDV: " & strNumero_PDV
    
    'Carregando a combo de finalizadora
    strSQL = Empty
    strSQL = "SELECT IXCodigo_TBFinalizadora,DFDescricao_TBFinalizadora FROM TBFinalizadora WHERE DFControle_venda_TBFinalizadora = 0"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSQL, "BDRetaguarda", "Otica", Me

    
End Sub
Private Sub txtFundo_Caixa_LostFocus()
    txtFundo_Caixa.Text = Format(txtFundo_Caixa, "#,###0.00")
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
