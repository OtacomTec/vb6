VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmMarcha_Consulta_Cliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Detalhada Cliente"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMarcha_Consulta_Cliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7470
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar"
      Enabled         =   0   'False
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
      Left            =   6300
      Picture         =   "frmMarcha_Consulta_Cliente.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3990
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros"
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
      Height          =   1725
      Left            =   3480
      TabIndex        =   6
      Top             =   30
      Width           =   3885
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
         Height          =   585
         Left            =   2700
         Picture         =   "frmMarcha_Consulta_Cliente.frx":1B0C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1020
         Width           =   1065
      End
      Begin VB.CommandButton cmdConsultar 
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
         Left            =   1560
         Picture         =   "frmMarcha_Consulta_Cliente.frx":1E96
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   120
         MaxLength       =   40
         TabIndex        =   7
         Top             =   570
         Width           =   2925
      End
      Begin AutoCompletar.CbCompleta cbbUf 
         Height          =   360
         Left            =   3090
         TabIndex        =   10
         Top             =   570
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
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
      Begin VB.Label lblVariavel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Width           =   60
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opções de Filtro"
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
      Height          =   1725
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   3375
      Begin VB.OptionButton optContrato 
         Caption         =   "Nº Contrato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1980
         TabIndex        =   16
         Top             =   660
         Width           =   1305
      End
      Begin VB.OptionButton optCidade 
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1980
         TabIndex        =   5
         Top             =   330
         Width           =   885
      End
      Begin VB.OptionButton optCnpj 
         Caption         =   "CNPJ/CPF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1125
      End
      Begin VB.OptionButton optNome_fantasia 
         Caption         =   "Nome Fantasia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1575
      End
      Begin VB.OptionButton optInscricao_estadual 
         Caption         =   "Inscrição Estadual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   990
         Width           =   1755
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   885
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgConsulta_cliente 
      Height          =   2145
      Left            =   90
      TabIndex        =   13
      Top             =   1770
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   3784
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   150
      TabIndex        =   15
      Top             =   4140
      Width           =   60
   End
   Begin VB.Label lblDescricao 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   14
      Top             =   4140
      Width           =   4995
   End
End
Attribute VB_Name = "frmMarcha_Consulta_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Transportes                                                    '
' Objetivo...............: Consulta Detalhada do Cliente para Marcha Analítica            '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Sá                                                       '
' Data de Criação........: 30/12/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim strSql As String
Dim intContador As Integer

Private Sub cmdAplicar_Click()
    frmMarcha_Analitica.txtCliente.Text = hfgConsulta_cliente.TextArray((hfgConsulta_cliente.Row * hfgConsulta_cliente.Cols + hfgConsulta_cliente.Col + 1))
    frmMarcha_Analitica.txtCliente.SetFocus
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
   'Limpando o grid de itens

    hfgConsulta_cliente.ClearStructure

    hfgConsulta_cliente.Rows = 2
    Movimentacoes.Monta_HFlex_Grid hfgConsulta_cliente, "800,3700,1900,2500", "Código,Destinatário,CNPJ/CPF,Cidade", 4, "OTICA", Me

    lblCodigo.Caption = Empty
    lblDescricao.Caption = Empty
    cmdAplicar.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdConsultar_Click()
    Dim booRetorno As Boolean
    lblCodigo.Caption = Empty
    lblDescricao.Caption = Empty

    If txtConsulta.Text = Empty And optTodos.Value = False Then
       MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
       txtConsulta.SetFocus
       Exit Sub
    End If

    ''Parametrização
     If Me.optCnpj = True Then
        txtConsulta.Text = Replace(txtConsulta.Text, ".", "")
        txtConsulta.Text = Replace(txtConsulta.Text, "-", "")
        txtConsulta.Text = Replace(txtConsulta.Text, "/", "")
        If Len(txtConsulta.Text) = 11 Then
           Call CGC_CPF.FormatarCPF(txtConsulta.Text, txtConsulta)
        ElseIf Len(txtConsulta.Text) = 14 Then
           Call CGC_CPF.FormatarCNPJ(txtConsulta.Text, txtConsulta)
        Else
           MsgBox "Número de digitos invalido. Redigite", vbInformation, "Only Tech"
           txtConsulta.Text = Empty
           txtConsulta.SetFocus
           Exit Sub
        End If
    Else
        If Me.optInscricao_estadual = True Then
           If txtConsulta.Text <> "ISENTO" And cbbUf.Text = Empty Then
              MsgBox "UF Inválida.", vbInformation, "Only Tech"
              Me.txtConsulta.SetFocus
              cbbUf.SetFocus
              Exit Sub
           End If
           txtConsulta.Text = Replace(txtConsulta.Text, ".", "")
           txtConsulta.Text = Replace(txtConsulta.Text, "-", "")
           txtConsulta.Text = Replace(txtConsulta.Text, "/", "")

           txtConsulta.Text = UCase(txtConsulta.Text)

           booRetorno = Inscricao_Estadual.ChecaInscrE(cbbUf.Text, txtConsulta.Text, txtConsulta)
           If booRetorno = False Then
                MsgBox "Inscrição Inválida.", vbInformation, "Only Tech"
                Me.txtConsulta.SetFocus
                Exit Sub
           End If
        End If
     End If
    ''''''''''''''''''''''''''''''''

    strSql = Empty
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente,DFCpf_TBCliente," & _
             "TBCidade_Otica.DFNome_TBCidade_Otica " & _
             "FROM TBCliente " & _
             "INNER JOIN TBCidade_Otica " & _
             "ON TBCliente.FKId_TBCidade_otica = TBCidade_Otica.PKId_TBCidade_Otica " & _
             "INNER JOIN TBContrato_cliente " & _
             "ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente "
    
    If optNome_fantasia.Value = True Then
        strSql = strSql & "WHERE DFNome_Fantasia_TBCliente LIKE '%" & Me.txtConsulta.Text & "%' "
    ElseIf optInscricao_estadual.Value = True Then
        strSql = strSql & "WHERE DFInscricao_estadual_TBCliente LIKE '%" & Me.txtConsulta.Text & "%' "
    ElseIf Me.optCnpj = True Then
       strSql = strSql & "WHERE DFCpf_TBCliente LIKE '%" & Me.txtConsulta.Text & "%' "
    ElseIf optCidade.Value = True Then
        strSql = strSql & "WHERE DFNome_TBCidade_otica LIKE '%" & Me.txtConsulta.Text & "%' "
    ElseIf optContrato.Value = True Then
        strSql = strSql & "WHERE PKCodigo_TBContrato_cliente = '" & Me.txtConsulta.Text & "' "
    End If

    frmAguarde.Show
    Call Movimentacoes.Movimenta_HFlex_Grid(strSql, Me.hfgConsulta_cliente, "800,3700,1900,2500", "Código,Cliente,CNPJ/CPF,Cidade", "BDRetaguarda", "Otica", Me)

    hfgConsulta_cliente.Row = 1
    hfgConsulta_cliente.Col = 0
    If hfgConsulta_cliente.Text = Empty Then
       hfgConsulta_cliente.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgConsulta_cliente, "800,3700,1900,2500", "Código,Cliente,CNPJ/CPF,Cidade", 4, "Otica", Me
    End If
    cmdAplicar.Enabled = False
    Unload frmAguarde
End Sub

Private Sub Form_Load()
    Movimentacoes.Monta_HFlex_Grid Me.hfgConsulta_cliente, "800,3700,1900,2500", "Código,Cliente,CNPJ/CPF,Cidade", 4, "Otica", Me
    Call Monta_Combo
End Sub

Private Sub optCidade_Click()
    txtConsulta.Width = 3645
    txtConsulta.Text = Empty
    Me.cbbUf.Visible = False
    lblVariavel.Caption = "Cidade"
    txtConsulta.Visible = True
    txtConsulta.SetFocus
End Sub

Private Sub optNome_fantasia_Click()
    txtConsulta.Width = 3645
    txtConsulta.Text = Empty
    Me.cbbUf.Visible = False
    lblVariavel.Caption = "Nome Fantasia"
    txtConsulta.Visible = True
    txtConsulta.SetFocus
End Sub

Private Sub optContrato_Click()
    txtConsulta.Width = 3645
    txtConsulta.Text = Empty
    Me.cbbUf.Visible = False
    lblVariavel.Caption = "Nome Fantasia"
    txtConsulta.Visible = True
    txtConsulta.SetFocus
End Sub
Private Sub optTodos_Click()
    txtConsulta.Visible = False
    Me.cbbUf.Visible = False
    lblVariavel.Caption = ""
    cmdConsultar.SetFocus
End Sub

Private Sub optCnpj_Click()
    cbbUf.Visible = False
    txtConsulta.Width = 3645
    lblVariavel.Caption = "CNPJ/CPF"
    txtConsulta.Visible = True
    txtConsulta.Text = Empty
    txtConsulta.SetFocus
End Sub

Private Sub optInscricao_estadual_Click()
    cbbUf.Visible = True
    txtConsulta.Width = 2925
    lblVariavel.Caption = "Inscrição Estadual"
    txtConsulta.Text = Empty
    txtConsulta.Visible = True
    cbbUf.Text = Empty
    txtConsulta.SetFocus
End Sub

Private Sub hfgConsulta_cliente_Click()
    If hfgConsulta_cliente.Text <> Empty Then
       hfgConsulta_cliente.Col = 0
       lblCodigo.Caption = hfgConsulta_cliente.TextArray((hfgConsulta_cliente.Row * hfgConsulta_cliente.Cols + hfgConsulta_cliente.Col + 1))
       lblDescricao.Caption = hfgConsulta_cliente.TextArray((hfgConsulta_cliente.Row * hfgConsulta_cliente.Cols + hfgConsulta_cliente.Col + 2))
       cmdAplicar.Enabled = True
    End If
End Sub

Private Function Monta_Combo()

   cbbUf.Clear
   cbbUf.AddItem ("AC")
   cbbUf.AddItem ("AL")
   cbbUf.AddItem ("AM")
   cbbUf.AddItem ("AP")
   cbbUf.AddItem ("BA")
   cbbUf.AddItem ("CE")
   cbbUf.AddItem ("DF")
   cbbUf.AddItem ("ES")
   cbbUf.AddItem ("GO")
   cbbUf.AddItem ("MA")
   cbbUf.AddItem ("MT")
   cbbUf.AddItem ("MS")
   cbbUf.AddItem ("MG")
   cbbUf.AddItem ("PA")
   cbbUf.AddItem ("PB")
   cbbUf.AddItem ("PE")
   cbbUf.AddItem ("PI")
   cbbUf.AddItem ("PR")
   cbbUf.AddItem ("RJ")
   cbbUf.AddItem ("RN")
   cbbUf.AddItem ("RO")
   cbbUf.AddItem ("RR")
   cbbUf.AddItem ("RS")
   cbbUf.AddItem ("SC")
   cbbUf.AddItem ("SE")
   cbbUf.AddItem ("SP")
   cbbUf.AddItem ("TO")

End Function

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConsulta_LostFocus()
    If optInscricao_estadual.Value = True Then
       If txtConsulta.Text = "ISENTO" Or txtConsulta.Text = "isento" Then
          txtConsulta.Width = 3645
          cbbUf.Visible = False
          cmdConsultar.SetFocus
       Else
          txtConsulta.Width = 2925
          cbbUf.Visible = True
          cbbUf.SetFocus
       End If
    End If

   txtConsulta.Text = UCase(txtConsulta.Text)
End Sub
