VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAbertura_Caixa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Abertura de Caixa"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
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
   ScaleHeight     =   3915
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5910
      Picture         =   "frmAbertura_Caixa.frx":1782
      ScaleHeight     =   615
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   30
      Width           =   435
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   45
      TabIndex        =   14
      Top             =   3120
      Width           =   45
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
      Left            =   4680
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
      Left            =   4680
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
      Left            =   255
      MaxLength       =   14
      TabIndex        =   0
      Top             =   1950
      Width           =   3015
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5850
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
      Top             =   3120
      Width           =   5655
   End
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
   Begin MSDataListLib.DataCombo dtcFinalizadora 
      Height          =   465
      Left            =   255
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
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   225
      ScaleHeight     =   525
      ScaleWidth      =   5925
      TabIndex        =   8
      Top             =   3120
      Width           =   5925
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   -30
      Y2              =   4710
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   495
      Left            =   4560
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
      Left            =   4560
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
      Left            =   255
      TabIndex        =   13
      Top             =   1470
      Width           =   2235
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   465
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   525
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   3210
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
      Left            =   255
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
      X1              =   2220
      X2              =   0
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Abertura de Caixa"
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
      Left            =   750
      TabIndex        =   9
      Top             =   210
      Width           =   3195
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
Public intImpressoes_suportadas As Integer
Dim strSql As String
Public dtpData_operacao As Date
Public booIntegracao_Retaguarda As Boolean
Public booLeitor_serial As Boolean
Public strCom_leitor_serial As String
Public intTipo_imp_orcamento As Integer
Public booGaveta_integrada As Boolean
'
Public intIP_Concentrador As String
Public booPreco_online As Boolean
Public booComissao_vendedor As Boolean
Public strNumero_check_out As String
Public strNumero_Nome_Operadora As String
Public strVersao_software As String
Public strNumero_loja As String
Public intFinalizadora_sangria  As Integer
'
Public strTipo_quantidade As String
Public strCasas_Decimais As String
Public strTipo_desconto As String
Public strDigito_Peso_Variavel As String
'
Public booPreco_peso_balanca_TBParametros_ecf As Boolean
Public strCaminho_impComum As String
Public booFinaliza_direto As Boolean
Public intPerfil_ECF As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim strID_Finalizadora As String
    Dim strCampos As String
    Dim strValores As String
    
    On Error GoTo Erro
    
    If Me.dtcFinalizadora.Text = "" Then
        MsgBox "Favor insira uma finalizadora v�lida", vbInformation, "Only Tech"
        dtcFinalizadora.SetFocus
        Exit Sub
    End If
    
    If Me.txtFundo_Caixa.Text = "" Or Me.txtFundo_Caixa.Text = 0 Then
        MsgBox "Favor insira um fundo de caixa v�lido", vbInformation, "Only Tech"
        Me.txtFundo_Caixa.SetFocus
        Exit Sub
    End If

    'Id Finalizadora
    If booIntegracao_Retaguarda = True Then
       strID_Finalizadora = Funcoes_Gerais.Localiza_ID("PKId_TBFinalizadora", "IXCodigo_TBFinalizadora", dtcFinalizadora.BoundText, "TBFinalizadora", "Otica", Me, "BDRetaguarda")
    Else
       strID_Finalizadora = Funcoes_Gerais.Localiza_ID("PKId_TBFinalizadora", "IXCodigo_TBFinalizadora", dtcFinalizadora.BoundText, "TBFinalizadora", "PDV", Me, "BDPDV")
    End If
    
    strCampos = "FKCodigo_TBPdv,FKId_TBFinalizadora,FKCodigo_TBOperadores_ecf,DFData_TBOperacao_caixa," & _
                "DFHora_TBOperacao_caixa,DFValor_TBOperacao_caixa,DFTipo_operacao_TBOperacao_caixa,DFStatus_aberto_fechado_TBOperacao_caixa," & _
                "DFObservacao_TBOperacao_caixa,FKCodigo_TBEmpresa,DFNumero_Cupom_TBOperacao_caixa"
              
    strValores = "" & strNumero_PDV & "," & _
                 "" & strID_Finalizadora & "," & _
                 "" & strCodigo_Operador & "," & _
                 "'" & Format(Now, "YYYYMMDD") & "'," & _
                 "'" & Format(Now, "hh:mm:ss") & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(Me.txtFundo_Caixa.Text) & "," & _
                 "1," & _
                 "0, '" & dtcFinalizadora.Text & "'," & strEmpresa_Operador & ",0"
    
    If booIntegracao_Retaguarda = True Then
       funcoes_banco.Gravar "TBoperacao_caixa", strCampos, strValores, "Otica", Me, "BDRetaguarda"
    Else
       funcoes_banco.Gravar "TBoperacao_caixa", strCampos, strValores, "PDV", Me, "BDPDV"
    End If
    
    MsgBox "Abertura de caixa feito com sucesso!", vbInformation, "Only Tech"
    
    Comandos_impressoras_fiscais.Abrir_gaveta (frmTela_Venda.strImpresora)
    
    frmTela_Venda.strOperador = strOperador
    frmTela_Venda.strPDV = strPDV
    frmTela_Venda.strEmpresa_Operador = strEmpresa_Operador
    frmTela_Venda.strCodigo_Operador = strCodigo_Operador
    frmTela_Venda.booIntegracao_Retaguarda = booIntegracao_Retaguarda
    frmTela_Venda.intImpressoes_suportadas = frmAbertura_Caixa.intImpressoes_suportadas
    frmTela_Venda.booLeitor_serial = booLeitor_serial
    frmTela_Venda.strCom_leitor_serial = strCom_leitor_serial
    frmTela_Venda.dtpData_operacao = dtpData_operacao
    frmTela_Venda.intTipo_imp_orcamento = intTipo_imp_orcamento
    frmTela_Venda.booGaveta_integrada = booGaveta_integrada
    frmTela_Venda.intIP_Concentrador = intIP_Concentrador
    frmTela_Venda.booPreco_online = booPreco_online
    frmTela_Venda.booComissao_vendedor = booComissao_vendedor
    frmTela_Venda.txtNumero_check_out.Text = strNumero_check_out
    frmTela_Venda.txtNumero_Nome_Operadora.Text = strNumero_Nome_Operadora
    frmTela_Venda.txtVersao_software.Text = strVersao_software
    frmTela_Venda.txtNumero_loja.Text = strNumero_loja
    frmTela_Venda.intFinalizadora_sangria = intFinalizadora_sangria
    frmTela_Venda.strTipo_quantidade = strTipo_quantidade
    frmTela_Venda.strCasas_Decimais = strCasas_Decimais
    frmTela_Venda.strTipo_desconto = strTipo_desconto
    frmTela_Venda.strTipo_quantidade = strTipo_quantidade
    frmTela_Venda.strDigito_Peso_Variavel = strDigito_Peso_Variavel
    frmTela_Venda.booPreco_peso_balanca_TBParametros_ecf = booPreco_peso_balanca_TBParametros_ecf
    frmTela_Venda.strCaminho_impComum = strCaminho_impComum
    frmTela_Venda.intPerfil_ECF = intPerfil_ECF
    frmTela_Venda.booFinaliza_direto = Me.booFinaliza_direto
    
    If frmTela_Venda.booCupom_fiscal = True Then
       Call Comandos_impressoras_fiscais.Abertura_Dia(Me.dtcFinalizadora.Text, Me.txtFundo_Caixa.Text)
    End If
    
    frmTela_Venda.Show
    
    Unload Me
    
    'Posto de Gasolina
    If intPerfil_ECF = 2 Then
       frmTela_Venda.booConsulta = True
       frmAbertura_Encerrantes.strOperador = strOperador
       frmAbertura_Encerrantes.strNumero_PDV = strNumero_PDV
       frmAbertura_Encerrantes.strCasas_Decimais = strCasas_Decimais
       Call frmAbertura_Encerrantes.Show(1)
    End If
    

    
    Exit Sub
    
Erro:
    If booIntegracao_Retaguarda = True Then
       Call Erro.Erro(Me, "Otica")
    Else
       Call Erro.Erro(Me, "PDV")
    End If
End Sub

Private Sub Form_Load()
   
    'Operador
    lblOperador.Caption = "Operador: " & strOperador
    Me.lblPDV.Caption = "N� PDV: " & strNumero_PDV
    
    'Carregando a combo de finalizadora
    strSql = Empty
    strSql = "SELECT IXCodigo_TBFinalizadora,DFDescricao_TBFinalizadora FROM TBFinalizadora WHERE DFControle_venda_TBFinalizadora = 0"
    
    If booIntegracao_Retaguarda = True Then
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSql, "BDRetaguarda", "Otica", Me
    Else
       Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSql, "BDPDV", "PDV", Me
    End If
    
End Sub
Private Sub dtcFinalizadora_GotFocus()
    Call Movimentacoes.Verifica_DataCombo(Me.dtcFinalizadora)
    Me.dtcFinalizadora.SetFocus
End Sub

Private Sub dtcFinalizadora_LostFocus()
    If Me.dtcFinalizadora.Text = "" Then Me.dtcFinalizadora.SetFocus
End Sub

Private Sub txtFundo_Caixa_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 44 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtFundo_Caixa_LostFocus()
    If Me.txtFundo_Caixa.Text <> "" Then
        If txtFundo_Caixa.Text = "," Then
           txtFundo_Caixa.Text = 0
        End If
        If txtFundo_Caixa.Text = 0 Then
           txtFundo_Caixa.Text = Empty
           txtFundo_Caixa.SetFocus
        Else
           txtFundo_Caixa.Text = Format(txtFundo_Caixa, "#,###0.00")
        End If
    Else
       txtFundo_Caixa.Text = 0
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    'Habilita a saida com ESC
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
