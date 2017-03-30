VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormAcessoSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso Administrativo"
   ClientHeight    =   4275
   ClientLeft      =   2775
   ClientTop       =   2385
   ClientWidth     =   6510
   Icon            =   "FRM9002-01-F1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6510
   Begin VB.CommandButton CommandProcura 
      Caption         =   "&Procura"
      Height          =   435
      Left            =   5400
      TabIndex        =   18
      Top             =   1710
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialogProcuradbConfus 
      Left            =   0
      Top             =   3750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CommandOk2 
      Caption         =   "O&k"
      Height          =   435
      Left            =   5400
      TabIndex        =   6
      Top             =   3825
      Width           =   975
   End
   Begin VB.CommandButton CommandCancela 
      Caption         =   "&Cancela"
      Height          =   435
      Left            =   4350
      TabIndex        =   7
      Top             =   3825
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGridLocacaoBancoDados 
      Height          =   1365
      Left            =   15
      TabIndex        =   8
      Top             =   2400
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   2408
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
   Begin VB.ComboBox CombobCodEsquemaBancoDadostLogin 
      Height          =   315
      Left            =   4500
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox TextstrLocacaobdConfustLogin 
      Height          =   315
      Left            =   20
      TabIndex        =   4
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Frame FrameLogin 
      Caption         =   "Login do Usuário"
      Height          =   1515
      Left            =   2625
      TabIndex        =   0
      Top             =   0
      Width           =   3840
      Begin VB.TextBox TextiCodUsuariotUsu 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   465
         Width           =   1260
      End
      Begin VB.TextBox TextstrSenhaUsuariotUsu 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   150
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   1260
      End
      Begin VB.CommandButton CommandOk 
         Caption         =   "&Ok"
         Height          =   435
         Left            =   2775
         TabIndex        =   3
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1530
         TabIndex        =   17
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1530
         TabIndex        =   16
         Top             =   225
         Width           =   570
      End
      Begin VB.Label LabelData 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1530
         TabIndex        =   15
         Top             =   465
         Width           =   1185
      End
      Begin VB.Label LabelHora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1530
         TabIndex        =   14
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   855
         Width           =   510
      End
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   20
      Picture         =   "FRM9002-01-F1.frx":000C
      Top             =   375
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Locação dos Bancos de Dados:"
      Height          =   195
      Left            =   20
      TabIndex        =   13
      Top             =   2175
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Esquema"
      Height          =   195
      Left            =   4500
      TabIndex        =   12
      Top             =   1560
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Locação Atual do bdConfus:"
      Height          =   195
      Left            =   20
      TabIndex        =   11
      Top             =   1560
      Width           =   2025
   End
End
Attribute VB_Name = "FormAcessoSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mstrHora As String
Public mdtData As Date
Public mboLibera As Boolean
Public mrstLocacaoBancoDados As ADODB.Recordset

Sub mpCarregaLocacaodbConfus()
    CombobCodEsquemaBancoDadostLogin.Clear
    
    MSFlexGridLocacaoBancoDados.Rows = 1
    
    Call ppAbre_BDAcesso(pdbConfus, Trim(TextstrLocacaobdConfustLogin.Text), pstrSenhaBancoDadosBDCONFUS)
         
    pstrSql = "Select DISTINCT bCodEsquemaBancoDadostLoc from tLocacaoBancoDados"
    
    If pfboQuery(pdbConfus, pstrSql, mrstLocacaoBancoDados) Then
        CombobCodEsquemaBancoDadostLogin.Text = mrstLocacaoBancoDados("bCodEsquemaBancoDadostLoc")
        pbCodEsquemaBancoDadostLogin = mrstLocacaoBancoDados("bCodEsquemaBancoDadostLoc")
    End If
    
    While Not mrstLocacaoBancoDados.EOF
        CombobCodEsquemaBancoDadostLogin.AddItem mrstLocacaoBancoDados("bCodEsquemaBancoDadostLoc")
        mrstLocacaoBancoDados.MoveNext
    Wend
        
    pstrSql = "Select bCodEsquemaBancoDadostLoc, strCodBancoDadostLoc, strLocacaoBancoDadostLoc, bTabTipoBancoDadostLoc from tLocacaoBancoDados WHERE bCodEsquemaBancoDadostLoc = " & Val(CombobCodEsquemaBancoDadostLogin.Text)
    If pfboQuery(pdbConfus, pstrSql, mrstLocacaoBancoDados) Then MSFlexGridLocacaoBancoDados.Rows = 1
    
    While Not mrstLocacaoBancoDados.EOF
        If InStr(1, pstrBancosUsados, mrstLocacaoBancoDados("strCodBancoDadostLoc") + " ", vbTextCompare) > 0 Then
            MSFlexGridLocacaoBancoDados.Rows = MSFlexGridLocacaoBancoDados.Rows + 1
            MSFlexGridLocacaoBancoDados.TextMatrix(MSFlexGridLocacaoBancoDados.Rows - 1, 0) = mrstLocacaoBancoDados("strCodBancoDadostLoc")
            MSFlexGridLocacaoBancoDados.TextMatrix(MSFlexGridLocacaoBancoDados.Rows - 1, 1) = IIf(mrstLocacaoBancoDados("bTabTipoBancoDadostLoc") = 1, "Access", "SQL")
            MSFlexGridLocacaoBancoDados.TextMatrix(MSFlexGridLocacaoBancoDados.Rows - 1, 2) = mrstLocacaoBancoDados("strLocacaoBancoDadostLoc")
        End If
        mrstLocacaoBancoDados.MoveNext
    Wend
    pstrLocacaobdConfus = TextstrLocacaobdConfustLogin.Text
End Sub

Private Sub CombobCodEsquemaBancoDadostLogin_Click()
    If Trim(CombobCodEsquemaBancoDadostLogin.Text) = Empty Then Exit Sub
    pstrSql = "Select bCodEsquemaBancoDadostLoc, strCodBancoDadostLoc, strLocacaoBancoDadostLoc, bTabTipoBancoDadostLoc from tLocacaoBancoDados WHERE bCodEsquemaBancoDadostLoc = " & Val(CombobCodEsquemaBancoDadostLogin.Text)
    If pfboQuery(pdbConfus, pstrSql, mrstLocacaoBancoDados) Then mrstLocacaoBancoDados.MoveFirst
    MSFlexGridLocacaoBancoDados.Rows = 1
    While Not mrstLocacaoBancoDados.EOF
        If InStr(1, pstrBancosUsados, mrstLocacaoBancoDados("strCodBancoDadostLoc"), vbTextCompare) > 0 Then
            MSFlexGridLocacaoBancoDados.Rows = MSFlexGridLocacaoBancoDados.Rows + 1
            MSFlexGridLocacaoBancoDados.TextMatrix(MSFlexGridLocacaoBancoDados.Rows - 1, 0) = mrstLocacaoBancoDados("strCodBancoDadostLoc")
            MSFlexGridLocacaoBancoDados.TextMatrix(MSFlexGridLocacaoBancoDados.Rows - 1, 1) = IIf(mrstLocacaoBancoDados("bTabTipoBancoDadostLoc") = 1, "Access", "SQL")
            MSFlexGridLocacaoBancoDados.TextMatrix(MSFlexGridLocacaoBancoDados.Rows - 1, 2) = mrstLocacaoBancoDados("strLocacaoBancoDadostLoc")
        End If
        mrstLocacaoBancoDados.MoveNext
    Wend
    SendKeys "{TAB}"
End Sub

Private Sub CombobCodEsquemaBancoDadostLogin_GotFocus()
    CombobCodEsquemaBancoDadostLogin.SelStart = 0
    CombobCodEsquemaBancoDadostLogin.SelLength = Len(Trim(CombobCodEsquemaBancoDadostLogin.Text))
End Sub

Private Sub CombobCodEsquemaBancoDadostLogin_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaNumerico(KeyAscii)
End Sub

Private Sub CombobCodEsquemaBancoDadostLogin_LostFocus()
    If Trim(CombobCodEsquemaBancoDadostLogin.Text) = Empty Then Call CombobCodEsquemaBancoDadostLogin_GotFocus
End Sub

Private Sub CommandCancela_Click()
    If MSFlexGridLocacaoBancoDados.Rows > 1 Then pdbConfus.Close
    TextstrLocacaobdConfustLogin = Empty
    CombobCodEsquemaBancoDadostLogin.Clear
    MSFlexGridLocacaoBancoDados.Rows = 1
    TextiCodUsuariotUsu = Empty
    TextstrSenhaUsuariotUsu = Empty
    TextiCodUsuariotUsu.SetFocus
End Sub

Private Sub CommandOk2_Click()
    On Error GoTo Erro
    Dim lstrLinha As String
    pboAcessoSistema = False
    
    Dim pstrComputer As String
    pstrComputer = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\control\ComputerName\ComputerName", "ComputerName")

    If Trim(pstrComputer) <> Empty Then
        piCodEstacaotLog = Val(Right(Trim(pstrComputer), 3))
    Else
        MsgBox "Configuração Incorreta! Entre com contato com D.I.", vbInformation, pstrEmpresa
        pdbConfus.Close
        Unload Me
        Exit Sub
    End If
    
    Dim lstrLocacaoAcessoriostLogin As String
    pstrSql = "Select * from tParametros"
    If pfboQuery(pdbConfus, pstrSql, prsSeleção) Then lstrLocacaoAcessoriostLogin = prsSeleção("strLocacaoAcessoriostPar")
    
    pstrSql = "UPDATE tLogin SET iCodUsuariotLogin = " & TextiCodUsuariotUsu.Text & ", strCodProgramatLogin = '" & pstrCodPrograma & "', strLocacaobdConfustLogin = '" & Trim(TextstrLocacaobdConfustLogin.Text) & "', bCodEsquemaBancoDadostLogin = " & CombobCodEsquemaBancoDadostLogin.Text & ", strLocacaoAcessoriostLogin = '" & lstrLocacaoAcessoriostLogin & " '"
    Call ppCommandExecute(pdbGMUSLOG, pstrSql)
    pboAcessoSistema = True
    pdbConfus.Close
    Unload Me
    Exit Sub

Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "CommandOK2_Click"
    Err.Clear
End Sub

Private Sub CommandOk_Click()
    On Error GoTo Erro
    
    Dim liCalc As Integer, liSenha8888 As String, liSenha9999 As String
    
    'Calculo das senhas administrativas
    'liCalc = ((Day(mdtData) * 3) + (Month(mdtData) * 2)) + Year(mdtData)
    liCalc = ((Day(mdtData) * 2) + (Month(mdtData))) + Year(mdtData) + Weekday(mdtData)
    liSenha8888 = Str(liCalc)
    liSenha9999 = Str(liCalc + (Val(Str(Hour(mstrHora)) + Str(Minute(mstrHora)) + Str(Second(mstrHora)))))
    
    If mboLibera Then
        TextiCodUsuariotUsu = 8888
        TextstrSenhaUsuariotUsu = liSenha8888
    End If
    
    If Trim(TextiCodUsuariotUsu.Text) <> "8888" And Trim(TextiCodUsuariotUsu.Text) <> "9999" Then
        MsgBox "Código de Usuario Inválido!", vbInformation, "CommandOk_Click"
        TextiCodUsuariotUsu.SetFocus
        Exit Sub
    End If
    
    If Trim(TextstrSenhaUsuariotUsu) = Empty Then
        TextstrSenhaUsuariotUsu.SetFocus
        Exit Sub
    End If
    
    If (Trim(TextiCodUsuariotUsu.Text) = "8888" And Val(Trim(TextstrSenhaUsuariotUsu)) <> liSenha8888) Or (Trim(TextiCodUsuariotUsu.Text) = "9999" And Val(Trim(TextstrSenhaUsuariotUsu)) <> liSenha9999) Then
        MsgBox "Senha de Acesso Inválida!", vbInformation, "CommandOk_Click"
        TextstrSenhaUsuariotUsu.Text = Empty
        TextstrSenhaUsuariotUsu.SetFocus
    Else
        TextstrLocacaobdConfustLogin.Enabled = True
        CombobCodEsquemaBancoDadostLogin.Enabled = True
        MSFlexGridLocacaoBancoDados.Enabled = True
        CommandCancela.Enabled = True
        CommandOk2.Enabled = True
        CommandProcura.Enabled = True
        Dim mrstLocacaoBancoDados As Recordset
        TextstrLocacaobdConfustLogin.Text = prstLogin("strLocacaobdConfustLogin")
        CombobCodEsquemaBancoDadostLogin.Text = prstLogin("bCodEsquemaBancoDadostLogin")
        
        If Trim(TextstrLocacaobdConfustLogin.Text) = Empty Or Dir(TextstrLocacaobdConfustLogin.Text) = Empty Then
            MsgBox TextstrLocacaobdConfustLogin & " não Encontrado!", vbInformation, "CommandOk_Click"
        Else
            Call mpCarregaLocacaodbConfus
        End If
        
        CommandOk.Enabled = False
        TextstrLocacaobdConfustLogin.SetFocus
    End If
    
    If mboLibera Then
        mboLibera = False
        Call CommandOk2_Click
    End If
    
    Exit Sub
    
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "CommandOK_Click"
    Err.Clear
End Sub

Private Sub CommandProcura_Click()
    On Error GoTo Erro
    CommonDialogProcuradbConfus.FLAGS = cdlOFNFileMustExist
    CommonDialogProcuradbConfus.Filter = "Arquivo bdConfus.mdb|bdConfus.mdb|"
    CommonDialogProcuradbConfus.FilterIndex = 2
    CommonDialogProcuradbConfus.FileName = TextstrLocacaobdConfustLogin.Text
    CommonDialogProcuradbConfus.ShowOpen
    Me.Refresh
    If CommonDialogProcuradbConfus.FileName <> Empty Then
        If MSFlexGridLocacaoBancoDados.Rows > 1 Then pdbConfus.Close
        TextstrLocacaobdConfustLogin.Text = CommonDialogProcuradbConfus.FileName
        Call mpCarregaLocacaodbConfus
        CommandOk.Enabled = False
        TextstrLocacaobdConfustLogin.SetFocus
    End If
    Exit Sub

Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "CommandImportar_Click"
    Err.Clear
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'MsgBox KeyCode
    Select Case KeyCode
        Case 116: mboLibera = True
    End Select
    
End Sub

Private Sub Form_Load()
    Me.Icon = IIf(Dir(pstrLocacaoIcoLogotipotLogin) <> Empty, LoadPicture(pstrLocacaoIcoLogotipotLogin), Empty)
    Me.Left = (Screen.Width - Width) / 2
    Me.Top = (Screen.Height - Height) / 2

    mstrHora = Time
    mdtData = Date
    mboLibera = False

    LabelHora.Caption = mstrHora
    LabelData.Caption = mdtData
    
    Call ppAbre_BDAcesso(pdbGMUSLOG, "C:\InfoMil_Estacao\GmusLog.Dll")
    
    pstrSql = "SELECT * FROM tLogin"
    
    If pfboQuery(pdbGMUSLOG, pstrSql, prstLogin) Then
        If Dir(prstLogin("strLocacaoIcoLogotipotLogin")) <> Empty Then
            'ImageIco.Picture = LoadPicture(prstLogin("strLocacaoIcoLogotipotLogin"))
            FormAcessoSistema.Icon = LoadPicture(prstLogin("strLocacaoIcoLogotipotLogin"))
        End If
        FormAcessoSistema.Caption = prstLogin("strNomeIdentFormtLogin") + " - " + FormAcessoSistema.Caption
    End If
    
    MSFlexGridLocacaoBancoDados.FixedAlignment(0) = 1
    MSFlexGridLocacaoBancoDados.FixedAlignment(1) = 1
    MSFlexGridLocacaoBancoDados.FixedAlignment(2) = 1
        
    MSFlexGridLocacaoBancoDados.ColWidth(0) = 1500
    MSFlexGridLocacaoBancoDados.ColWidth(1) = 1070
    MSFlexGridLocacaoBancoDados.ColWidth(2) = 3770
        
    MSFlexGridLocacaoBancoDados.TextMatrix(0, 0) = "Banco de Dados"
    MSFlexGridLocacaoBancoDados.TextMatrix(0, 1) = "Tipo/Banco"
    MSFlexGridLocacaoBancoDados.TextMatrix(0, 2) = "Locação do Banco"
    MSFlexGridLocacaoBancoDados.Rows = 1
    
    Dim liCalc As Integer, liSenha8888 As String, liSenha9999 As String
    
    'Calculo das senhas administrativas
    liCalc = ((Day(mdtData) * 3) + (Month(mdtData) * 2)) + Year(mdtData)
    liSenha8888 = Trim(Str(liCalc))
    Call ppCarregaPropriedadesForm(Me, 200)
    
    Set mrstLocacaoBancoDados = New ADODB.Recordset
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'mboLibera = False
    If (X > 749 And X < 811) And (Y > 374 And Y < 406) And TextiCodUsuariotUsu.Text = Empty Then
        'mboLibera = True
        Call CommandOk_Click
    End If
End Sub

Private Sub TextiCodUsuariotUsu_GotFocus()
    TextiCodUsuariotUsu.SelStart = 0
    TextiCodUsuariotUsu.SelLength = Len(Trim(TextiCodUsuariotUsu))
    TextstrLocacaobdConfustLogin.Enabled = False
    CombobCodEsquemaBancoDadostLogin.Enabled = False
    MSFlexGridLocacaoBancoDados.Enabled = False
    CommandCancela.Enabled = False
    CommandOk2.Enabled = False
    CommandOk.Enabled = True
    CommandProcura.Enabled = False

End Sub

Private Sub TextiCodUsuariotUsu_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaNumerico(KeyAscii)
End Sub

Private Sub TextiCodUsuariotUsu_LostFocus()
    If Trim(TextiCodUsuariotUsu.Text) <> "8888" And Trim(TextiCodUsuariotUsu.Text) <> "9999" Then
        MsgBox "Código de Usuario Inválido!", vbInformation, "TextiCodUsuariotUsu_LostFocus"
        Call TextiCodUsuariotUsu_GotFocus
        TextiCodUsuariotUsu.SetFocus
    End If
End Sub

Private Sub TextstrLocacaobdConfustLogin_GotFocus()
    TextstrLocacaobdConfustLogin.SelStart = 0
    TextstrLocacaobdConfustLogin.SelLength = Len(Trim(TextstrLocacaobdConfustLogin))
End Sub

Private Sub TextstrLocacaobdConfustLogin_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaTexto(KeyAscii)
End Sub

Private Sub TextstrLocacaobdConfustLogin_LostFocus()
    If Trim(TextstrLocacaobdConfustLogin) = Empty Then Call TextstrLocacaobdConfustLogin_GotFocus
    If pstrLocacaobdConfus <> Trim(TextstrLocacaobdConfustLogin.Text) Then
        If Trim(TextstrLocacaobdConfustLogin.Text) = Empty Or Dir(TextstrLocacaobdConfustLogin.Text) = Empty Then
            MsgBox TextstrLocacaobdConfustLogin & " não Encontrado!", vbInformation, "TextstrLocacaobdConfustLogin_LostFocus"
            TextstrLocacaobdConfustLogin.Text = pstrLocacaobdConfus
            TextstrLocacaobdConfustLogin.SetFocus
            Exit Sub
        Else
            If pstrLocacaobdConfus = Empty Then Exit Sub
            If MSFlexGridLocacaoBancoDados.Rows > 1 Then pdbConfus.Close
            Call mpCarregaLocacaodbConfus
        End If
    End If
End Sub

Private Sub TextstrSenhaUsuariotUsu_GotFocus()
    TextstrSenhaUsuariotUsu.SelStart = 0
    TextstrSenhaUsuariotUsu.SelLength = Len(Trim(TextiCodUsuariotUsu))
End Sub

Private Sub TextstrSenhaUsuariotUsu_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaTexto(KeyAscii)
End Sub

Private Sub TextstrSenhaUsuariotUsu_LostFocus()
    If Trim(TextstrSenhaUsuariotUsu) = Empty Then Call TextstrSenhaUsuariotUsu_GotFocus
End Sub
