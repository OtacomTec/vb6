VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#1.3#0"; "CRYSTL32.OCX"
Begin VB.Form FormPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Extrato"
   ClientHeight    =   2415
   ClientLeft      =   2340
   ClientTop       =   2535
   ClientWidth     =   5535
   Icon            =   "TMP3001-01-F0.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5535
   Begin Crystal.CrystalReport CrystalReport 
      Left            =   3000
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data DataEmpresa 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "H:\APLICVB\BD\Bdsuport.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "tEmpresa"
      Top             =   840
      Visible         =   0   'False
      Width           =   912
   End
   Begin MSDBCtls.DBCombo DBComboEmpresa 
      Bindings        =   "TMP3001-01-F0.frx":0442
      DataSource      =   "DataEmpresa"
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   327681
      ListField       =   "strNomeResumidotEmp"
      BoundColumn     =   "iCodEmpresatEmp"
      Text            =   ""
   End
   Begin VB.Data DataSolicitação 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "H:\APLICVB\BD\Bdsuport.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "SELECT  DISTINCT  tSolicitacao.strTipoSolicitacaotSol from tSolicitacao"
      Top             =   1440
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.CommandButton CommandImprimir 
      Caption         =   "Imprimir"
      Height          =   435
      Left            =   4440
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton CommandLimpar 
      Caption         =   "Limpar"
      Height          =   435
      Left            =   4440
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TextEmpresa 
      Height          =   288
      Left            =   120
      MaxLength       =   3
      TabIndex        =   2
      Top             =   840
      Width           =   612
   End
   Begin VB.OptionButton OptionAnalítico 
      Caption         =   "Analítico"
      Height          =   435
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton OptionSintético 
      Caption         =   "Sintético"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSMask.MaskEdBox MaskEdBoxInicial 
      Height          =   285
      Left            =   855
      TabIndex        =   5
      Top             =   2010
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      _Version        =   327681
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBoxFinal 
      Height          =   285
      Left            =   2265
      TabIndex        =   6
      Top             =   2010
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      _Version        =   327681
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSDBCtls.DBCombo DBComboSolicitação 
      Bindings        =   "TMP3001-01-F0.frx":0458
      DataSource      =   "DataSolicitação"
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   327681
      ListField       =   "strTipoSolicitacaotSol"
      BoundColumn     =   "strTipoSolicitacaotSol"
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Index           =   3
      Left            =   2070
      TabIndex        =   12
      Top             =   2025
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período"
      Height          =   195
      Index           =   2
      Left            =   855
      TabIndex        =   11
      Top             =   1770
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Solicitação"
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   10
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   660
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------'
'Codigo Programa: TMP3001-01-VB                                                              '
'Descr.Programa.: Modelo de Programa de Relatorio                                            '
'Analista.......: Geraldo Coimbra                                                            '
'Programador....: Jorge A M de Carvalho / Victor Augusto Faria Dias                          '
'Data Criação...: 24/04/1998                                                                 '
'Data Alteração.:                                                                            '
'--------------------------------------------------------------------------------------------'
Option Explicit
Public mboFlagRotinasIniciais As Boolean

Private Sub CommandImprimir_Click()
    If Me.MaskEdBoxInicial.Text = "" Or Not IsDate(Me.MaskEdBoxInicial.Text) Then
        MsgBox "Data Inicial inválida", vbCritical + vbOKOnly, "CommandImprimir_Click"
        Me.MaskEdBoxInicial.SetFocus
        Exit Sub
    End If
    
    If Me.MaskEdBoxFinal.Text = "" Or Not IsDate(Me.MaskEdBoxFinal.Text) Then
        MsgBox "Data Final inválida", vbCritical + vbOKOnly, "CommandImprimir_Click"
        Me.MaskEdBoxFinal.SetFocus
        Exit Sub
    End If
    
    pstrSql = "{tSolicitacao.dtDataExecGeraltSol} >= date(" & Mid(Me.MaskEdBoxInicial.ClipText, 5) & "," & Mid(Me.MaskEdBoxInicial.ClipText, 3, 2) & "," & Mid(Me.MaskEdBoxInicial.ClipText, 1, 2) & ") AND {tSolicitacao.dtDataExecGeraltSol} <= date(" & Mid(Me.MaskEdBoxFinal.ClipText, 5) & "," & Mid(Me.MaskEdBoxInicial.ClipText, 3, 2) & "," & Mid(Me.MaskEdBoxInicial.ClipText, 1, 2) & ")"
                
    If Me.TextEmpresa.Text <> "" Then pstrSql = pstrSql & " AND {tSolicitacao.iCodEmpresatSol} = " & Me.TextEmpresa.Text
    If Me.DBComboSolicitação.BoundText <> "" Then pstrSql = pstrSql & " AND {tSolicitacao.strTipoSolicitacaotSol} = '" & Me.DBComboSolicitação.Text & "'"
    
    Me.CrystalReport.Formulas(0) = "TipoRelatorio = '" & IIf(Me.OptionSintético.Value = True, "S", "A")
    Me.CrystalReport.SelectionFormula = pstrSql
    Me.CrystalReport.ReportFileName = App.Path & IIf(Right(App.Path, 1) <> "/", "/", "") & "SGH3001-01-R1.rpt"
    Me.CrystalReport.Action = 1
End Sub

Private Sub CommandLimpar_Click()
    Me.TextEmpresa.Text = ""
    Me.DBComboSolicitação.BoundText = ""
    Me.DBComboEmpresa.BoundText = ""
    Me.MaskEdBoxInicial.Mask = ""
    Me.MaskEdBoxInicial.Text = ""
    Me.MaskEdBoxInicial.Mask = "##/##/####"
    Me.MaskEdBoxFinal.Mask = ""
    Me.MaskEdBoxFinal.Text = ""
    Me.MaskEdBoxFinal.Mask = "##/##/####"
    Me.TextEmpresa.SetFocus
End Sub

Private Sub DBComboEmpresa_Click(Area As Integer)
    If Area <> 2 Then Exit Sub
    TextEmpresa.Text = DBComboEmpresa.BoundText
End Sub

Private Sub DBComboEmpresa_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaTexto(KeyAscii)
End Sub

Private Sub DBComboEmpresa_LostFocus()
    If Me.DBComboEmpresa.Text = "" Then Me.TextEmpresa.Text = ""
End Sub

Private Sub DBComboSolicitação_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaTexto(KeyAscii)
End Sub

Private Sub Form_Load()
    On Error GoTo Erro
    
    mboFlagRotinasIniciais = False
    
    If Not pfboRotinasIniciais("SGH3001-01-VB") Then End
    
    Call ppCarregaPropriedadesForm(Me)
    Call ppAbre_BDAcesso(pstrCodPrograma, pdbConfus, pstrLocacaobdConfus)
    
    pstrSql = "Select * FROM tLocacaoBancoDados where bCodEsquemaBancoDadostLoc = " & pbCodEsquemaBancoDadostLogin & " AND strCodBancoDadostLoc = 'BDSuport'"
    
    Set prsSeleção = pfrsSelecao(pdbConfus, pstrSql)
    
    If prsSeleção.EOF Then
        MsgBox "Locação do Banco de Dados BDSuport nao encontrado na Tabela tLocacaoBancoDados", vbCritical + vbOKOnly, "Form_Load"
        End
    End If
    
    pstrLocacaobdBanco = prsSeleção.Fields("strLocacaoBancoDadostLoc")
    
    Call ppAbre_BDAcesso(pstrCodPrograma, pdbBanco, pstrLocacaobdBanco)
                    
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "Form_Load"
    Err.Clear

End Sub

Private Sub Form_Terminate()
    Call ppRotinasFinais: pdbBanco.Close: End
End Sub

Private Sub MaskEdBoxFinal_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaNumerico(KeyAscii)
End Sub

Private Sub MaskEdBoxInicial_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaNumerico(KeyAscii)
End Sub

Private Sub OptionAnalítico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub OptionSintético_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TextEmpresa_KeyPress(KeyAscii As Integer)
    Call ppValidaTeclaNumerico(KeyAscii)
End Sub

Private Sub TextEmpresa_LostFocus()
    On Error GoTo Erro
    
    If Me.TextEmpresa.Text = "" Then Me.DBComboEmpresa.BoundText = "": Exit Sub
    
    pstrSql = "SELECT * From tEmpresa Where iCodEmpresatEmp = " & Me.TextEmpresa.Text

    Set prsSeleção = pfrsSelecao(pdbBanco, pstrSql)

    If prsSeleção.EOF Then
        MsgBox "Empresa não encontrada!", vbCritical, "TextEmpresa_LostFocus"
        Me.TextEmpresa.Text = ""
        Me.TextEmpresa.SetFocus
        Exit Sub
    End If

    prsSeleção.Close
                    
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "TextEmpresa_LostFocus"
    Err.Clear

End Sub
