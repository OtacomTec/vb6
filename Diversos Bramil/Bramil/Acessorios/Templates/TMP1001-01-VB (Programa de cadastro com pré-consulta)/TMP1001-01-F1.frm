VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FormCadastro1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro1"
   ClientHeight    =   1755
   ClientLeft      =   690
   ClientTop       =   3135
   ClientWidth     =   8190
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   8190
   Begin VB.CommandButton CommandOk 
      Caption         =   "&Ok"
      Height          =   435
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CommandCancela 
      Caption         =   "&Cancela"
      Height          =   435
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame FrameCadastro 
      Caption         =   " Dados Cadastrais "
      Height          =   1755
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7095
      Begin VB.Frame FrameDados 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   75
         TabIndex        =   7
         Top             =   960
         Width           =   6975
         Begin VB.TextBox TextstrCampo 
            DataField       =   "strSenhaUsuariotUsu"
            DataSource      =   "DatatUsuarios"
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1275
            MaxLength       =   8
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   240
            Width           =   5535
         End
         Begin MSMask.MaskEdBox MaskEdBoxdtData 
            DataField       =   "dtDataValidadeUsuariotUsu"
            DataSource      =   "DatatUsuarios"
            Height          =   315
            Left            =   30
            TabIndex        =   2
            Top             =   240
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   327681
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   45
            TabIndex        =   9
            Top             =   0
            Width           =   345
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Outro Campo"
            Height          =   195
            Left            =   1275
            TabIndex        =   8
            Top             =   0
            Width           =   930
         End
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2040
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   2  'Snapshot
         RecordSource    =   "Nome da Tabela"
         Top             =   240
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.TextBox TextiCodigo 
         Height          =   315
         Left            =   120
         MaxLength       =   4
         TabIndex        =   0
         Top             =   600
         Width           =   735
      End
      Begin MSDBCtls.DBCombo DBCombostrNome 
         Bindings        =   "TMP1001-01-F1.frx":0000
         DataField       =   "Nome do Campo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   327681
         ListField       =   "Nome do Campo "
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "FormCadastro1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mpCarregaDados()
    On Error GoTo Erro
    
    If pstrFuncaoToolbar = "Consulta" Then piQuantidadeCon = piQuantidadeCon + 1
    
    TextiCodigo.Text = Data1.Recordset.Fields("iCodigo")
    DBCombostrNome.Text = Data1.Recordset.Fields("strNome")
    MaskEdBoxdtData.Mask = ""
    MaskEdBoxdtData.Text = Data1.Recordset.Fields("dtData")
    MaskEdBoxdtData.Mask = "##/##/####"
    TextstrCampo.Text = Data1.Recordset.Fields("strCampo")
    
    DBCombostrNome.SelStart = 0
    DBCombostrNome.SelLength = 0
                
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "mpCarregaDados"
    Err.Clear
End Sub

Public Sub mpInicializaForm()
    TextiCodigo.Text = ""
    DBCombostrNome.Text = ""
    
    MaskEdBoxdtData.Mask = ""
    MaskEdBoxdtData.Text = ""
    MaskEdBoxdtData.Mask = "##/##/####"
    
    TextstrCampo.Text = ""
    
    piCodigo1 = 0
    FrameCadastro.Enabled = True
    FrameDados.Enabled = False
    TextiCodigo.Enabled = True
    
    TextiCodigo.SetFocus
End Sub

Private Sub CommandCancela_Click()
    Call mpInicializaForm
End Sub

Private Sub DBCombostrNome_Click(Area As Integer)
    If Area <> 2 Then Exit Sub
    Area = 0
    TextiCodigo.Enabled = True
    TextiCodigo.Text = DBCombostrNome.BoundText
    TextiCodigo.SetFocus
End Sub

Private Sub DBCombostrNome_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaTexto(KeyAscii)
End Sub

Private Sub DBCombostrNome_LostFocus()
    If TextiCodigo.Text = "" And DBCombostrNome.BoundText = "" Then DBCombostrNome.Text = ""
End Sub

Private Sub TextiCodigo_GotFocus()
    If TextiCodigo.Text <> "" Then SendKeys ("{TAB}"):  Exit Sub
    Call mpInicializaForm
End Sub

Private Sub TextiCodigo_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaNumerico(KeyAscii)
End Sub

Private Sub TextiCodigo_LostFocus()
    On Error GoTo Erro
        
    piCodigo1 = 0

    If TextiCodigo.Text = "" Or piCodigo1 <> 0 Then Exit Sub
    
    pstrSql = "SELECT DISTINCTROW tUsuarios.iCodUsuariotUsu, tUsuarios.strNomeUsuariotUsu, tUsuarios.strNomeUsuarioResumidotUsu, tUsuarios.bCodGrupoUsuariotUsu, tUsuariosGrupo.strDescrGrupoUsuariotGrpUsu, tUsuarios.bTabNivelUsuariotUsu, tUsuarios.strSenhaUsuariotUsu, tUsuarios.bDiasValidadeSenhatUsu, tUsuarios.dtDataSenhatUsu, tUsuarios.dtDataValidadeUsuariotUsu FROM tUsuariosGrupo INNER JOIN tUsuarios ON tUsuariosGrupo.bCodGrupoUsuariotGrpUsu = tUsuarios.bCodGrupoUsuariotUsu WHERE tUsuarios.iCodUsuariotUsu= " & TextiCodigo.Text
    
    Data1.RecordSource = pstrSql
    Data1.Refresh
    
    If Data1.Recordset.EOF And pstrFuncaoToolbar <> "Inclusão" Then
        MsgBox "Usuário não encontrado", vbCritical, "Nome do Evento"
        TextiCodigo = ""
        Call TextiCodigo_GotFocus
        Exit Sub
    End If

    If Not Data1.Recordset.EOF And pstrFuncaoToolbar = "Inclusão" Then
        MsgBox "Usuário Já Cadastrado", vbCritical, "Nome do Evento"
        TextiCodigo = ""
        Call TextiCodigo_GotFocus
        Exit Sub
    End If
    
    piCodigo1 = TextiCodigo.Text
    TextiCodigo.Enabled = False
    FrameCadastro.Enabled = True
    FrameDados.Enabled = True
    
    If pstrFuncaoToolbar <> "Inclusão" Then Call mpCarregaDados
    If pstrFuncaoToolbar = "Inclusão" Or pstrFuncaoToolbar = "Alteração" Then FrameCadastro.Enabled = True Else FrameCadastro.Enabled = False
    If pstrFuncaoToolbar = "Chave" Then FormAlteraChave2.Show 1
                
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "TextiCodigo_LostFocus"
    Err.Clear
End Sub

Private Sub MaskEdBoxdtData_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaNumerico(KeyAscii)
End Sub

Private Sub Form_Load()
    On Error GoTo Erro
    
    Call ppCarregaPropriedadesForm(Me, 1001)
    
    Data1.DatabaseName = pstrLocacaobdConfus
    
    Data1.Refresh
    
    If pstrFuncaoToolbar <> "Inclusão" And piCodigo1 <> 0 Then TextiCodigo.Text = piCodigo1: piCodigo1 = 0
                
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "Form_Load"
    Err.Clear
End Sub

Private Sub CommandOk_Click()
    On Error GoTo Erro
    
    If TextiCodigo.Text = "" Then Unload Me: Exit Sub
    If pstrFuncaoToolbar = "Consulta" Then Call mpInicializaForm: Exit Sub
    If pstrFuncaoToolbar = "Exclusão" Then If MsgBox("Confirma Exclusão", vbInformation + vbYesNo + vbDefaultButton2, "CommandOk_Click") = vbNo Then Exit Sub
    
    If Not IsDate(MaskEdBoxdtData) Then
        MsgBox "Data inválida", vbCritical, "CommandOk_Click"
        MaskEdBoxdtData.SetFocus
        Exit Sub
    End If
    
    If TextiCodigo.Text = "" Then
        MsgBox "Campo <Código> requerido na tabela <tTabela1>", vbCritical, "CommandOk_Click"
        TextiCodigo.SetFocus
        Exit Sub
    End If
    
    If DBCombostrNome.Text = "" Then
        MsgBox "Campo <Nome> requerido na tabela <tTabela1>", vbCritical, "CommandOk_Click"
        DBCombostrNome.SetFocus
        Exit Sub
    End If
    If TextstrCampo.Text = "" Then
        MsgBox "Campo <Campo> requerido na tabela <tTabela1>", vbCritical, "CommandOk_Click"
        TextstrCampo.SetFocus
        Exit Sub
    End If

    If MaskEdBoxdtData.ClipText = "" Then
        MsgBox "Campo <Data> requerido na tabela <tTabela1>", vbCritical, "CommandOk_Click"
        MaskEdBoxdtData.SetFocus
        Exit Sub
    End If
    
    If pstrFuncaoToolbar = "Inclusão" Then
        pstrSql = "INSERT INTO ..."
        piQuantidadeInc = piQuantidadeInc + 1
    End If
    
    If pstrFuncaoToolbar = "Alteração" Then
        pstrSql = "UPDATE ... Where iCodigo = " & piCodigo1
        piQuantidadeAlt = piQuantidadeAlt + 1
    End If
    
    If pstrFuncaoToolbar = "Chave" Then
        pstrSql = "UPDATE ... Where iCodigo = " & piCodigo1
        piQuantidadeChv = piQuantidadeChv + 1
    End If
    
    If pstrFuncaoToolbar = "Exclusão" Then
        pstrSql = "DELETE * FROM ... WHERE iCodigo = " & TextiCodigo.Text
        piQuantidadeExc = piQuantidadeExc + 1
    End If
    
    pdbConfus.Execute pstrSql, dbFailOnError
        
    Data1.Refresh
    Data1Geral.Refresh
    Call mpInicializaForm
                
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "CommandOk_Click"
    Err.Clear
    
End Sub

Private Sub TextstrCampo_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaTexto(KeyAscii)
End Sub
