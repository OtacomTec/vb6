VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFechamento_dia 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Fechamento de Dia"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFechamento_dia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   3120
      Width           =   15
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5820
      ScaleHeight     =   465
      ScaleWidth      =   285
      TabIndex        =   7
      Top             =   3150
      Width           =   285
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   240
      ScaleHeight     =   45
      ScaleWidth      =   5655
      TabIndex        =   6
      Top             =   3600
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   240
      ScaleHeight     =   45
      ScaleWidth      =   5655
      TabIndex        =   5
      Top             =   3150
      Width           =   5655
   End
   Begin MSDataListLib.DataCombo dtcFinalizadora 
      Height          =   465
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Finalizadora"
      Top             =   3150
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   820
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   8454143
      ForeColor       =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5910
      Picture         =   "frmFechamento_dia.frx":1782
      ScaleHeight     =   615
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   30
      Width           =   435
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   200
      ScaleHeight     =   525
      ScaleWidth      =   5925
      TabIndex        =   8
      Top             =   3120
      Width           =   5925
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtFundo_Caixa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   200
      MaxLength       =   14
      TabIndex        =   0
      Top             =   1950
      Width           =   3015
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   495
      Left            =   4410
      Shape           =   4  'Rounded Rectangle
      Top             =   2460
      Width           =   1485
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   495
      Left            =   4410
      Shape           =   4  'Rounded Rectangle
      Top             =   1740
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Valor da Gaveta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   200
      TabIndex        =   13
      Top             =   1470
      Width           =   2235
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   525
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1980
      Width           =   3045
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   5925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Finalizadora:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   200
      TabIndex        =   12
      Top             =   2640
      Width           =   1770
   End
   Begin VB.Label lblOperador 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Operador:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   200
      TabIndex        =   11
      Top             =   900
      Width           =   1440
   End
   Begin VB.Label lblPDV 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "PDV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   4590
      TabIndex        =   10
      Top             =   900
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   2280
      X2              =   60
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Finalizando Dia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   240
      TabIndex        =   9
      Top             =   210
      Width           =   2730
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   4740
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line5 
      X1              =   6360
      X2              =   0
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Line Line3 
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   4740
   End
End
Attribute VB_Name = "frmFechamento_dia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Dim booIntegracao_online As Boolean
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim strID_Finalizadora As String
    Dim strCampos As String
    Dim strValores As String
    Dim strValores_Set As String
    
    On Error GoTo Erro
    
    'Verifica se confirma a operação
    If MsgBox("Esta operação irá executar uma redução Z. A impressora não poderá mais ser usada até às 23:59:59 hs. Você realmente deseja continuar com essa operação?", vbYesNo + vbInformation + vbDefaultButton2, "Onlytech") = vbNo Then
       Exit Sub
    End If
       
    If frmTela_Venda.booCupom_fiscal = True And frmTela_Venda.intImpressoes_suportadas <> 2 Then
       Call Comandos_impressoras_fiscais.Leitura_z(frmTela_Venda.strImpresora)
    End If
    
    If Me.dtcFinalizadora.BoundText = "" Then
        MsgBox "Favor insira uma finalizadora válida", vbInformation, "Only Tech"
        dtcFinalizadora.SetFocus
        Exit Sub
    End If
    
    'Id Finalizadora
    If booIntegracao_online = True Then
       strID_Finalizadora = Funcoes_Gerais.Localiza_ID("PKId_TBFinalizadora", "IXCodigo_TBFinalizadora", dtcFinalizadora.BoundText, "TBFinalizadora", "Otica", Me, "BDRetaguarda")
    Else
       strID_Finalizadora = Funcoes_Gerais.Localiza_ID("PKId_TBFinalizadora", "IXCodigo_TBFinalizadora", dtcFinalizadora.BoundText, "TBFinalizadora", "PDV", Me, "BDPDV")
    End If

    strCampos = "FKCodigo_TBPdv,FKId_TBFinalizadora,FKCodigo_TBOperadores_ecf,DFData_TBOperacao_caixa," & _
                "DFHora_TBOperacao_caixa,DFValor_TBOperacao_caixa,DFTipo_operacao_TBOperacao_caixa,DFStatus_aberto_fechado_TBOperacao_caixa," & _
                "DFObservacao_TBOperacao_caixa"
              
    strValores = "" & frmTela_Venda.txtNumero_check_out & "," & _
                 "" & strID_Finalizadora & "," & _
                 "" & frmTela_Venda.strCodigo_Operador & "," & _
                 "'" & Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD") & "'," & _
                 "'" & Format(frmTela_Venda.dtpData_operacao, "hh:mm:ss") & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(Me.txtFundo_Caixa.Text) & "," & _
                 "1," & _
                 "1, '" & dtcFinalizadora.Text & "'"
                 
    If booIntegracao_online = True Then
       funcoes_banco.Gravar "TBoperacao_caixa", strCampos, strValores, "Otica", Me, "BDRetaguarda"
    End If
    
    funcoes_banco.Gravar "TBoperacao_caixa", strCampos, strValores, "PDV", Me, "BDPDV"
    
'    If frmTela_Venda.booCupom_fiscal = True Then
'       Retorno = Bematech_FI_LeituraX()
'       Call VerificaRetornoImpressora("", "", "Leitura X")
'    End If
    
    strValores_Set = "SET DFStatus_dia_TBDia_Operacao_pdv = 'F'"
       
    If frmTela_Venda.booIntegracao_Retaguarda = True Then
       funcoes_banco.Alterar "TBDia_Operacao_pdv", strValores_Set, "DFDia_TBDia_operacao_pdv", Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD"), "Otica", Me, "BDRetaguarda"
    Else
       funcoes_banco.Alterar "TBDia_Operacao_pdv", strValores_Set, "DFDia_TBDia_operacao_pdv", Format(frmTela_Venda.dtpData_operacao, "YYYYMMDD"), "Otica", Me, "BDPDV"
    End If
    
    MsgBox "Fechamento de dia feito com sucesso!", vbInformation, "Only Tech"
  
    End
    
    Exit Sub
    
Erro:

    If booIntegracao_online = True Then
       Call Erro.Erro(Me, "Otica")
    Else
       Call Erro.Erro(Me, "PDV")
    End If
    
End Sub

Private Sub Form_Load()

    'Operador
    lblOperador.Caption = "Operador: " & strOperador
    Me.lblPDV.Caption = "N° PDV: " & strNumero_PDV
    booIntegracao_online = frmTela_Venda.booIntegracao_Retaguarda
    
    'Carregando a combo de finalizadora
    strSql = Empty
    strSql = "SELECT IXCodigo_TBFinalizadora,DFDescricao_TBFinalizadora FROM TBFinalizadora WHERE DFControle_venda_TBFinalizadora = 0"
    
    If booIntegracao_online = True Then
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSql, "BDRetaguarda", "Otica", Me
    Else
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSql, "BDPDV", "PDV", Me
    End If
   
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
Private Sub dtcFinalizadora_GotFocus()
    Call Movimentacoes.Verifica_DataCombo(Me.dtcFinalizadora)
    Me.dtcFinalizadora.SetFocus
End Sub

Private Sub dtcFinalizadora_LostFocus()
    If Me.dtcFinalizadora.Text = "" Then Me.dtcFinalizadora.SetFocus
End Sub
