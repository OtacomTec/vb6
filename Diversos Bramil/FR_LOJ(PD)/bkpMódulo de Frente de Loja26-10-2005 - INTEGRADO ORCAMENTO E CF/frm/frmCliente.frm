VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Cliente"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   Icon            =   "frmCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6540
      Picture         =   "frmCliente.frx":1782
      ScaleHeight     =   615
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   30
      Width           =   435
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0080FFFF&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5220
      Width           =   1635
   End
   Begin VB.OptionButton optCodigo_interno 
      BackColor       =   &H0080FFFF&
      Caption         =   "código Interno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   150
      TabIndex        =   2
      Top             =   2100
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   8190
      ScaleHeight     =   1935
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   3390
      Width           =   15
   End
   Begin VB.OptionButton optCartao_afinidade 
      BackColor       =   &H0080FFFF&
      Caption         =   "cartão Afinidade"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   150
      TabIndex        =   0
      Top             =   1050
      Width           =   1995
   End
   Begin VB.OptionButton optDescricao 
      BackColor       =   &H0080FFFF&
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   150
      TabIndex        =   1
      Top             =   1560
      Width           =   1995
   End
   Begin VB.TextBox txtCliente 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   2490
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1710
      Width           =   4215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   2460
      ScaleHeight     =   75
      ScaleWidth      =   4215
      TabIndex        =   7
      Top             =   1650
      Width           =   4215
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   2460
      ScaleHeight     =   75
      ScaleWidth      =   4215
      TabIndex        =   6
      Top             =   2250
      Width           =   4215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgCliente 
      Height          =   1995
      Left            =   300
      TabIndex        =   4
      Top             =   2790
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   3519
      _Version        =   393216
      BackColor       =   8454143
      BackColorFixed  =   8454143
      BackColorBkg    =   8454143
      BackColorUnpopulated=   8454143
      GridColorFixed  =   8454143
      GridColorUnpopulated=   8454143
      AllowBigSelection=   0   'False
      FocusRect       =   2
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2460
      TabIndex        =   12
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   3600
      X2              =   60
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1260
      X2              =   1365
      Y1              =   1320
      Y2              =   1335
   End
   Begin VB.Line Line5 
      X1              =   1260
      X2              =   1320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6990
      Y1              =   6030
      Y2              =   6030
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   645
      Left            =   5070
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   1665
   End
   Begin VB.Label lblAguarde 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aguarde......."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      TabIndex        =   10
      Top             =   5370
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   -60
      Y2              =   6000
   End
   Begin VB.Line Line4 
      X1              =   6990
      X2              =   6990
      Y1              =   -30
      Y2              =   6030
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   345
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   2190
      Width           =   2055
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2295
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   2670
      Width           =   6675
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   2235
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   2850
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   345
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   1140
      Width           =   2055
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   345
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   1650
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6990
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   885
      Left            =   2370
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   4425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cliente Especial (Afinidade)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   750
      TabIndex        =   8
      Top             =   210
      Width           =   4920
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   915
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   1620
      Width           =   4395
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim rstFinalizadora As New ADODB.Recordset
    Dim rstAfinidade As New ADODB.Recordset
    Dim dblFinalizafora_paga As Double
    Dim dblLimite_Superado As Double
    
    Me.hfgCliente.Col = 1
    
    If Me.hfgCliente.Text = "" Then
       MsgBox "Cliente não selecionado!Verifique.", vbCritical, "Only Tech"
       Me.txtCliente.SetFocus
       Exit Sub
    End If
    
    'Verificando o limite do Cliente
    strSql = Empty
    strSql = "SELECT DFLimite_credito_TBCliente,DFDia_vencimento_TBCliente FROM TBCLIENTE WHERE IXCODIGO_TBCLIENTE = " & Me.hfgCliente.Text & " " & _
             "AND IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAfinidade, "Otica", Me
    
    rstAfinidade.MoveFirst
    
    If rstAfinidade.BOF = True And rstAfinidade.EOF = True Then
       MsgBox "Este Cliente não possui limite disponível para esta operação!Verifique.", vbCritical, "Only Tech"
       Set rstAfinidade = Nothing
       Unload Me
    End If
    
    If CDbl(rstAfinidade!DFLimite_credito_TBCliente) <= 0 Then
       MsgBox "Este cliente esta com o seu limite esgotado!Verifique.", vbCritical, "Only Tech"
       Set rstAfinidade = Nothing
       Unload Me
       frmFechamento_Cupom.dtcFinalizadora_cupom.SetFocus
       Exit Sub
    End If
    
    If CInt(rstAfinidade!DFDia_vencimento_TBCliente) = 0 Or IsNull(rstAfinidade!DFDia_vencimento_TBCliente) = True Then
       MsgBox "Este cliente esta com o dia de vencimento não definido!Verifique.", vbCritical, "Only Tech"
       Set rstAfinidade = Nothing
       Unload Me
       frmFechamento_Cupom.dtcFinalizadora_cupom.SetFocus
       Exit Sub
    End If
    
    'Selecionando a finalizadora do parametro
    strSql = Empty
    strSql = "SELECT DFFinalizadora_cartao_afinidade_TBParametros_ecf  FROM TBParametros_ecf " & _
             "WHERE FKCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstFinalizadora, "Otica", Me
    
    If rstFinalizadora.BOF And rstFinalizadora.EOF = True Then
       MsgBox "Finalizadora para este tipo de operação não cadastrada no parametros ECF!Verifique.", vbCritical, "Only Tech"
       Set rstFinalizadora = Nothing
       Set rstAfinidade = Nothing
       Exit Sub
    Else
       If Not IsNull(rstFinalizadora!DFFinalizadora_cartao_afinidade_TBParametros_ecf) Then
          frmFechamento_Cupom.dtcFinalizadora_cupom.BoundText = rstFinalizadora!DFFinalizadora_cartao_afinidade_TBParametros_ecf
          Me.hfgCliente.Col = 1
          frmFechamento_Cupom.Cod_Cliente = Me.hfgCliente.Text
          frmFechamento_Cupom.intDia_Vencimento_Cliente = rstAfinidade!DFDia_vencimento_TBCliente
          'Total por finalizadora
          frmFechamento_Cupom.HflexGrid.Col = 1
          If frmFechamento_Cupom.HflexGrid.Text <> "" Then
             frmFechamento_Cupom.HflexGrid.Col = 2
             frmFechamento_Cupom.HflexGrid.Row = 3
             If frmFechamento_Cupom.HflexGrid.Text <> "" Then
                frmFechamento_Cupom.HflexGrid.Rows = frmFechamento_Cupom.HflexGrid.Rows + 1
                Do While frmFechamento_Cupom.HflexGrid.Text <> ""
                   dblFinalizafora_paga = CDbl(frmFechamento_Cupom.HflexGrid.Text) + dblFinalizafora_paga
                   frmFechamento_Cupom.HflexGrid.Row = frmFechamento_Cupom.HflexGrid.Row + 1
                Loop
                'Somando o limite à compra e comparando se ainda resta limite
                If CDbl(rstAfinidade!DFLimite_credito_TBCliente) > CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text) - CDbl(dblFinalizafora_paga) Then
                    frmFechamento_Cupom.txtValor_pago.Text = CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text) - CDbl(dblFinalizafora_paga)
                    frmFechamento_Cupom.dblPago_Afinidade = CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text) - CDbl(dblFinalizafora_paga)
                    frmFechamento_Cupom.dblLimite = CDbl(rstAfinidade!DFLimite_credito_TBCliente) - frmFechamento_Cupom.txtValor_pago.Text
                    frmFechamento_Cupom.intDia_Vencimento_Cliente = rstAfinidade!DFDia_vencimento_TBCliente
                    frmFechamento_Cupom.lngCliente_Afinidade = Me.hfgCliente.Text
                    frmFechamento_Cupom.HflexGrid.Rows = frmFechamento_Cupom.HflexGrid.Rows - 1
                    Set rstFinalizadora = Nothing
                    Set rstAfinidade = Nothing
                    Unload Me
                    Call frmFechamento_Cupom.txtValor_pago_LostFocus
                Else
                    dblLimite_Superado = CDbl(rstAfinidade!DFLimite_credito_TBCliente) - (CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text) - CDbl(dblFinalizafora_paga))
                    dblLimite_Superado = dblLimite_Superado * (-1)
                    
                    MsgBox "Com o total deste cupom superará o limite permitido para este Cliente.Verifique! LIMITE PERMITIDO --> " & Format(CDbl(rstAfinidade!DFLimite_credito_TBCliente), "#,###0.00") & " & TOTAL EXCEDIDO DO LIMITE: " & Format(dblLimite_Superado, "#,###0.00") & "", vbCritical, "Only Tech"
                    
                    Set rstFinalizadora = Nothing
                    Set rstAfinidade = Nothing
                    Unload Me
                    frmFechamento_Cupom.HflexGrid.Rows = frmFechamento_Cupom.HflexGrid.Rows - 1
                    frmFechamento_Cupom.dtcFinalizadora_cupom.SetFocus
                End If
             Else
                'Somando o limite à compra e comparando se ainda resta limite
                If CDbl(rstAfinidade!DFLimite_credito_TBCliente) > CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text) Then
                    frmFechamento_Cupom.txtValor_pago.Text = CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text)
                    frmFechamento_Cupom.dblPago_Afinidade = CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text)
                    frmFechamento_Cupom.intDia_Vencimento_Cliente = rstAfinidade!DFDia_vencimento_TBCliente
                    frmFechamento_Cupom.dblLimite = CDbl(rstAfinidade!DFLimite_credito_TBCliente) - frmFechamento_Cupom.txtTotal_Cupom.Text
                    frmFechamento_Cupom.lngCliente_Afinidade = Me.hfgCliente.Text
                    Set rstFinalizadora = Nothing
                    Set rstAfinidade = Nothing
                    Unload Me
                    Call frmFechamento_Cupom.txtValor_pago_LostFocus
                Else
                    dblLimite_Superado = CDbl(rstAfinidade!DFLimite_credito_TBCliente) - CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text)
                    dblLimite_Superado = dblLimite_Superado * (-1)
                    
                    MsgBox "Com o total deste cupom superará o limite permitido para este Cliente.Verifique! LIMITE PERMITIDO --> " & Format(CDbl(rstAfinidade!DFLimite_credito_TBCliente), "#,###0.00") & " & TOTAL EXCEDIDO DO LIMITE: " & Format(dblLimite_Superado, "#,###0.00") & "", vbCritical, "Only Tech"
                    
                    Set rstFinalizadora = Nothing
                    Set rstAfinidade = Nothing
                    frmFechamento_Cupom.HflexGrid.Rows = frmFechamento_Cupom.HflexGrid.Rows - 1
                    frmFechamento_Cupom.dtcFinalizadora_cupom.SetFocus
                    Unload Me
                End If
             End If
          Else
             'Somando o limite à compra e comparando se ainda resta limite
             If CDbl(rstAfinidade!DFLimite_credito_TBCliente) > CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text) Then
                frmFechamento_Cupom.txtValor_pago.Text = CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text)
                frmFechamento_Cupom.dblLimite = CDbl(rstAfinidade!DFLimite_credito_TBCliente) - frmFechamento_Cupom.txtTotal_Cupom.Text
                frmFechamento_Cupom.dblPago_Afinidade = CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text)
                frmFechamento_Cupom.intDia_Vencimento_Cliente = rstAfinidade!DFDia_vencimento_TBCliente
                frmFechamento_Cupom.lngCliente_Afinidade = Me.hfgCliente.Text
                Set rstFinalizadora = Nothing
                Set rstAfinidade = Nothing
                Unload Me
                Call frmFechamento_Cupom.txtValor_pago_LostFocus
             Else
                dblLimite_Superado = CDbl(frmFechamento_Cupom.txtTotal_Cupom.Text) - CDbl(rstAfinidade!DFLimite_credito_TBCliente)
                dblLimite_Superado = dblLimite_Superado * (-1)
                
                MsgBox "Com o total deste cupom superará o limite permitido para este Cliente.Verifique! LIMITE PERMITIDO --> " & Format(CDbl(rstAfinidade!DFLimite_credito_TBCliente), "#,###0.00") & " & TOTAL EXCEDIDO DO LIMITE: " & Format(dblLimite_Superado, "#,###0.00") & "", vbCritical, "Only Tech"
                Set rstFinalizadora = Nothing
                Set rstAfinidade = Nothing
                Unload Me
                frmFechamento_Cupom.HflexGrid.Rows = frmFechamento_Cupom.HflexGrid.Rows - 1
                frmFechamento_Cupom.dtcFinalizadora_cupom.SetFocus
             End If
          End If
       End If
    End If
    
    Unload Me
    
    Set rstFinalizadora = Nothing
    Set rstAfinidade = Nothing
    
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
    
    'Verifica se foi prescionado o CTRL + A
    If KeyAscii = 1 Then
        Me.optCartao_afinidade.SetFocus
    End If
    
    'Verifica se foi prescionado o CTRL + I
    If KeyAscii = 9 Then
        Me.optCodigo_interno.SetFocus
    End If
    
    'Verifica se foi prescionado o CTRL + D
    If KeyAscii = 4 Then
        Me.optDescricao.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
    frmFechamento_Cupom.dblLimite = 0
    frmFechamento_Cupom.dblPago_Afinidade = 0
    If frmTela_Venda.booComissao_vendedor = True Then
       frmVendedor.Show (1)
       frmVendedor.Visible = False
    End If
End Sub

Private Sub optCartao_afinidade_Click()
    Me.txtCliente.SetFocus
    Me.txtCliente.Text = Empty
End Sub

Private Sub optCodigo_interno_Click()
    Me.txtCliente.SetFocus
    Me.txtCliente.Text = Empty
End Sub

Private Sub optDescricao_Click()
    Me.txtCliente.SetFocus
    Me.txtCliente.Text = Empty
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    If Me.optCartao_afinidade.Value = True Or Me.optCodigo_interno.Value = True Then
       If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
       End If
    End If
End Sub

Private Sub txtCliente_LostFocus()

    If Me.txtCliente.Text <> "" Then
    
        Dim rstCliente As New ADODB.Recordset
        
        Me.lblAguarde.Visible = True
        
        If Me.optCartao_afinidade.Value = True Then
            strSql = Empty
            strSql = "SELECT IXCodigo_TBCliente,PKId_TBCliente,DFNome_TBCliente,DFLimite_credito_TBCliente FROM TBCliente WHERE DFNumero_contrato_TBCliente = '" & Me.txtCliente.Text & "' " & _
                     "AND IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
            Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCliente, "Otica", Me
            If rstCliente.BOF = True And rstCliente.EOF = True Then
               MsgBox "Cliente não cadastrado!Verifique.", vbCritical, "Only Tech"
               Me.txtCliente.Text = Empty
               Me.txtCliente.SetFocus
               Set rstCliente = Nothing
               Me.lblAguarde.Visible = False
              Exit Sub
            Else
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgCliente, "800,0,5000", "Código,ID,Nome", "BDRetaguarda", "Otica", Me, "S"
            End If
        End If
        
        If Me.optDescricao.Value = True Then
            strSql = Empty
            strSql = "SELECT IXCodigo_TBCliente,PKId_TBCliente,DFNome_TBCliente FROM TBCliente WHERE convert(nvarchar,TBCliente.DFNome_TBCliente) LIKE '%" & Me.txtCliente.Text & "%' " & _
                     "AND IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
            Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCliente, "Otica", Me
            If rstCliente.BOF = True And rstCliente.EOF = True Then
               MsgBox "Cliente não cadastrado!Verifique.", vbCritical, "Only Tech"
               Me.txtCliente.Text = Empty
               Me.txtCliente.SetFocus
               Set rstCliente = Nothing
               Me.lblAguarde.Visible = False
              Exit Sub
            Else
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgCliente, "800,0,5000", "Código,ID,Nome", "BDRetaguarda", "Otica", Me, "S"
            End If
        End If
        
        If Me.optCodigo_interno.Value = True And IsNumeric(Me.txtCliente.Text) Then
            strSql = Empty
            strSql = "SELECT IXCodigo_TBCliente,PKId_TBCliente,DFNome_TBCliente FROM TBCliente WHERE IXCodigo_TBCliente = " & Me.txtCliente.Text & " " & _
                     "AND IXCodigo_TBEmpresa = " & frmTela_Venda.strEmpresa_Operador & ""
            Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCliente, "Otica", Me
            If rstCliente.BOF = True And rstCliente.EOF = True Then
               MsgBox "Cliente não cadastrado!Verifique.", vbCritical, "Only Tech"
               Me.txtCliente.Text = Empty
               Me.txtCliente.SetFocus
               Set rstCliente = Nothing
               Me.lblAguarde.Visible = False
              Exit Sub
            Else
              Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgCliente, "800,0,5000", "Código,ID,Nome", "BDRetaguarda", "Otica", Me, "S"
            End If
        End If
        Me.lblAguarde.Visible = False
    End If
    
End Sub
