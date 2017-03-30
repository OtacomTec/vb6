VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de Manutenção"
   ClientHeight    =   4605
   ClientLeft      =   6075
   ClientTop       =   2040
   ClientWidth     =   6465
   ClipControls    =   0   'False
   Icon            =   "FormPrincipal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6465
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   4260
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Exclusão de Registros"
      TabPicture(0)   =   "FormPrincipal.frx":0A8A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameTabelaDeGiro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CheckTabelaDeGiro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CheckMovimento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Base de Dados"
      TabPicture(1)   =   "FormPrincipal.frx":0AA6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "Reparar"
         Height          =   1635
         Left            =   -71790
         TabIndex        =   21
         Top             =   390
         Width           =   1485
         Begin VB.CheckBox CheckManutencao 
            Caption         =   "bdGMS001"
            Height          =   195
            Index           =   20
            Left            =   150
            TabIndex        =   24
            ToolTipText     =   "Cadastro de Produtos da Matriz"
            Top             =   330
            Width           =   1155
         End
         Begin VB.CheckBox CheckManutencao 
            Caption         =   "bdGMS002"
            Height          =   195
            Index           =   21
            Left            =   150
            TabIndex        =   23
            ToolTipText     =   "Base Geral do InfoMil"
            Top             =   570
            Width           =   1155
         End
         Begin VB.CheckBox CheckManutencao 
            Caption         =   "bdGMS005"
            Height          =   195
            Index           =   22
            Left            =   150
            TabIndex        =   22
            ToolTipText     =   "Base de Dados de Movimento de Entrada e Saída"
            Top             =   810
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Compactar"
         Height          =   1635
         Left            =   -73350
         TabIndex        =   17
         Top             =   390
         Width           =   1485
         Begin VB.CheckBox CheckManutencao 
            Caption         =   "bdGMS005"
            Height          =   195
            Index           =   12
            Left            =   150
            TabIndex        =   20
            ToolTipText     =   "Base de Dados de Movimento de Entrada e Saída"
            Top             =   810
            Width           =   1155
         End
         Begin VB.CheckBox CheckManutencao 
            Caption         =   "bdGMS002"
            Height          =   195
            Index           =   11
            Left            =   150
            TabIndex        =   19
            ToolTipText     =   "Base Geral do InfoMil"
            Top             =   570
            Width           =   1155
         End
         Begin VB.CheckBox CheckManutencao 
            Caption         =   "bdGMS001"
            Height          =   195
            Index           =   10
            Left            =   150
            TabIndex        =   18
            ToolTipText     =   "Cadastro de Produtos da Matriz"
            Top             =   330
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Gerar Nova Base"
         Height          =   1635
         Left            =   -74910
         TabIndex        =   13
         Top             =   390
         Width           =   1485
         Begin VB.CheckBox CheckManutencao 
            Caption         =   "bdGMS001"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   16
            ToolTipText     =   "Cadastro de Produtos da Matriz"
            Top             =   330
            Width           =   1155
         End
         Begin VB.CheckBox CheckManutencao 
            Caption         =   "bdGMS002"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   15
            ToolTipText     =   "Base Geral do InfoMil"
            Top             =   570
            Width           =   1155
         End
         Begin VB.CheckBox CheckManutencao 
            Caption         =   "bdGMS005"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   14
            ToolTipText     =   "Base de Dados de Movimento de Entrada e Saída"
            Top             =   810
            Width           =   1155
         End
      End
      Begin VB.CheckBox CheckMovimento 
         Caption         =   "Movimento de Entrada e Saída"
         Height          =   225
         Left            =   3390
         TabIndex        =   12
         Top             =   720
         Width           =   2625
      End
      Begin VB.Frame Frame7 
         Height          =   1095
         Left            =   3240
         TabIndex        =   8
         Top             =   720
         Width           =   2955
         Begin VB.OptionButton OptionMovimento 
            Caption         =   "Manter apenas os últimos 60 dias"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   10
            Top             =   330
            Width           =   2685
         End
         Begin VB.OptionButton OptionMovimento 
            Caption         =   "Excluir até o dia"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   9
            Top             =   630
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPickerMov 
            Height          =   285
            Index           =   1
            Left            =   1590
            TabIndex        =   11
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            Format          =   24444929
            CurrentDate     =   37189
         End
      End
      Begin VB.CheckBox CheckTabelaDeGiro 
         Caption         =   "Giro Diário"
         Height          =   225
         Left            =   270
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Frame FrameTabelaDeGiro 
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2985
         Begin VB.OptionButton OptionGiro 
            Caption         =   "Excluir até o dia"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   6
            Top             =   630
            Width           =   1455
         End
         Begin VB.OptionButton OptionGiro 
            Caption         =   "Manter apenas os últimos 60 dias"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   5
            Top             =   330
            Width           =   2685
         End
         Begin MSComCtl2.DTPicker DTPickerGiro 
            Height          =   285
            Index           =   1
            Left            =   1620
            TabIndex        =   4
            Top             =   600
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            Format          =   24444929
            CurrentDate     =   37189
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Executar"
      Height          =   375
      Left            =   5130
      TabIndex        =   0
      Top             =   2550
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2970
      Width           =   6345
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lboGiroDiário As Boolean
Dim lboNotas As Boolean

Enum GM_enGiroDiário
    GM_GiroNenhum = 0
    GM_GiroManter60Dias = 1
    GM_GiroExcluirAté = 2
    GM_GiroExcluirPeríodo = 3
End Enum

Enum GM_enMovimento
    GM_MovNenhum = 0
    GM_MovManter60Dias = 1
    GM_MovExcluirAté = 2
    GM_MovExcluirPeríodo = 3
End Enum

Private Type GM_tpManuteção
    GM_GiroDiário As GM_enGiroDiário
    GM_Movimento As GM_enMovimento
    
    GM_GerarNovoGMS001 As Boolean
    GM_GerarNovoGMS002 As Boolean
    GM_GerarNovoGMS005 As Boolean
    GM_GerarNovoBDLOG As Boolean
    
    GM_CompactarGMS001 As Boolean
    GM_CompactarGMS002 As Boolean
    GM_CompactarGMS005 As Boolean
    
    GM_RepararGMS001 As Boolean
    GM_RepararGMS002 As Boolean
    GM_RepararGMS005 As Boolean
End Type

Dim mtpManutenção As GM_tpManuteção


Private Sub CheckManutencao_Click(Index As Integer)

    Select Case Index
        Case 0: mtpManutenção.GM_GerarNovoGMS001 = CheckManutencao(Index).Value
        Case 1: mtpManutenção.GM_GerarNovoGMS002 = CheckManutencao(Index).Value
        Case 2: mtpManutenção.GM_GerarNovoGMS005 = CheckManutencao(Index).Value
        Case 3: mtpManutenção.GM_GerarNovoBDLOG = CheckManutencao(Index).Value
        
        Case 10: mtpManutenção.GM_CompactarGMS001 = CheckManutencao(Index).Value
        Case 11: mtpManutenção.GM_CompactarGMS002 = CheckManutencao(Index).Value
        Case 12: mtpManutenção.GM_CompactarGMS005 = CheckManutencao(Index).Value
        'Case 10: mtpManutenção.GM_CompactarGMS001 = CheckManutencao(Index).Value
        
        Case 20: mtpManutenção.GM_RepararGMS001 = CheckManutencao(Index).Value
        Case 21: mtpManutenção.GM_RepararGMS002 = CheckManutencao(Index).Value
        Case 22: mtpManutenção.GM_RepararGMS005 = CheckManutencao(Index).Value
        
    End Select
    
End Sub

Private Sub CheckMovimento_Click()
    If CheckMovimento.Value = 0 Then
        mtpManutenção.GM_Movimento = GM_MovNenhum
        OptionMovimento(1).Enabled = False
        OptionMovimento(2).Enabled = False
        OptionMovimento(1).Value = False
        OptionMovimento(2).Value = False
        DTPickerMov(1).Enabled = False
    Else
        OptionMovimento(1).Enabled = True
        OptionMovimento(2).Enabled = True
    
    End If
End Sub

Private Sub CheckTabelaDeGiro_Click()
    If CheckTabelaDeGiro.Value = 0 Then
        mtpManutenção.GM_GiroDiário = GM_GiroNenhum
        OptionGiro(1).Enabled = False
        OptionGiro(2).Enabled = False
        OptionGiro(1).Value = False
        OptionGiro(2).Value = False
        DTPickerGiro(1).Enabled = False

    Else
        OptionGiro(1).Enabled = True
        OptionGiro(2).Enabled = True
    End If
    
End Sub

Private Sub Command1_Click()
    Dim lstrCaminhoPastaBD  As String
    Dim lstrCaminhobdGMS001 As String
    Dim lstrSql             As String
    Dim liResp              As Integer
    Dim lstrmsg             As String
    
              lstrmsg = "Esta operação requer uso EXCLUSIVO!" & vbCrLf
    lstrmsg = lstrmsg & "Cerifique-se de que em todas as estações os programas" & vbCrLf
    lstrmsg = lstrmsg & "do InfoMil estejam fechados, incluido o Navegador." & vbCrLf & vbCrLf
    lstrmsg = lstrmsg & "Continuar a operação?"
    
    liResp = MsgBox(lstrmsg, vbYesNo + vbQuestion, "GMS8003-01-VB")
    If liResp <> 6 Then Exit Sub
    Command1.Enabled = False
    Text1.Text = Text1.Text & ".......... Iniciando Processo de Manutenção: " & Date & " " & Time & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Dim dtInicial As Date
    Dim dtFinal   As Date
    Dim wrk As Workspace
    Dim bd As Database
    
    'Me.Refresh
    DoEvents
    lstrCaminhoPastaBD = "h:\aplicvb\BD"
    lstrCaminhobdGMS001 = pstrLocacaobdGMS001 ' lstrCaminhoPastaBD & "\bdGMS001.mdb"
    lstrCaminhobdGMS002 = pstrLocacaobdGMS002 ' lstrCaminhoPastaBD & "\bdGMS002.mdb"
    lstrCaminhobdGMS005 = pstrLocacaobdGMS005 ' lstrCaminhoPastaBD & "\bdGMS005.mdb"
    
    'Inicia Exclusão no bdGMS002 tabela de Giro
    If mtpManutenção.GM_GiroDiário <> GM_GiroNenhum Then
        Text1.Text = Text1.Text & ".......... INICIANDO EXCLUSÃO NA TABELA DE GIRO" & vbCrLf
        Text1.SelStart = Len(Text1.Text)
        'Me.Refresh
        DoEvents
        
        Set wrk = Workspaces(0)
        Set bd = wrk.OpenDatabase(lstrCaminhobdGMS002, True, False)
        Select Case mtpManutenção.GM_GiroDiário
            Case GM_GiroManter60Dias
                dtFinal = Date - 60
                lstrFragSQL = " WHERE dtSaidatGirDia <#" & Format(dtFinal, "mm/dd/yyyy") & "#"
            Case GM_GiroExcluirAté
                dtFinal = DTPickerGiro(1).Value
                lstrFragSQL = " WHERE dtSaidatGirDia <=#" & Format(dtFinal, "mm/dd/yyyy") & "#"
        End Select
        'Me.Refresh
        DoEvents
        
        Text1.Text = Text1.Text & "Iniciando Instrução" & vbCrLf
        Text1.SelStart = Len(Text1.Text)
        
        lstrSql = "DELETE FROM tProdutosBarraGiroDiario " & lstrFragSQL
        bd.Execute lstrSql
        
        
        Text1.Text = Text1.Text & "Tabela de Giro: " & bd.RecordsAffected & " Registros excluídos" & vbCrLf
        Text1.SelStart = Len(Text1.Text)
        'Me.Refresh
        DoEvents
    
        bd.Close
        Set wrk = Nothing
    End If
    'Me.Refresh
    DoEvents
    
    'Inicia Exclusão no bdGMS005
     If mtpManutenção.GM_Movimento <> GM_MovNenhum Then
        Text1.Text = vbCrLf & Text1.Text & ".......... INICIANDO EXCLUSÃO NA TABELA DE MOVIMENTO DE ENTRADA E SAÍDA" & vbCrLf
        Text1.SelStart = Len(Text1.Text)
        'Me.Refresh
        DoEvents
        
        Set wrk = Workspaces(0)
        Set bd = wrk.OpenDatabase(lstrCaminhobdGMS005, True, False)
        Select Case mtpManutenção.GM_Movimento
            Case GM_MovManter60Dias
                dtFinal = Date - 60
                lstrFragSQL = " WHERE dtEmissaoDoctotMov <#" & Format(dtFinal, "mm/dd/yyyy") & "#"
            Case GM_MovExcluirAté
                dtFinal = DTPickerGiro(1).Value
                lstrFragSQL = " WHERE dtEmissaoDoctotMov <=#" & Format(dtFinal, "mm/dd/yyyy") & "#"
        End Select
        
        lstrSql = "DELETE DISTINCTROW tMovto.lSeqMovtotMov, " & _
                                     "tMovtoItens.lSeqMovtotMovItem, " & _
                                     "tMovto.dtEmissaoDoctotMov, " & _
                                     "tMovtoItens.* " & _
                                "FROM tMovto " & _
                          "INNER JOIN tMovtoItens " & _
                                  "ON tMovto.lSeqMovtotMov = tMovtoItens.lSeqMovtotMovItem " & _
                                      lstrFragSQL
        Text1.Text = Text1.Text & "Iniciando Instrução" & vbCrLf
        Text1.SelStart = Len(Text1.Text)
        'Me.Refresh
        DoEvents
        
        bd.Execute lstrSql
        'Me.Refresh
        DoEvents
        
        llTotalItens = bd.RecordsAffected
        lstrSql = "DELETE tMovto.lSeqMovtotMov, " & _
                         "tMovtoItens.lSeqMovtotMovItem, " & _
                         "tMovto.* " & _
                    "FROM tMovto " & _
               "LEFT JOIN tMovtoItens " & _
                      "ON tMovto.lSeqMovtotMov = tMovtoItens.lSeqMovtotMovItem " & _
                   "WHERE tMovtoItens.lSeqMovtotMovItem Is Null"
        bd.Execute lstrSql
        'Me.Refresh
        DoEvents
        
        Text1.Text = Text1.Text & "Tabela Cabeçalho da Nota: " & bd.RecordsAffected & " Registros excluídos" & vbCrLf
        Text1.Text = Text1.Text & "Tabela Itens da Nota: " & llTotalItens & " Registros excluídos" & vbCrLf
        Text1.SelStart = Len(Text1.Text)
        'Me.Refresh
        DoEvents
        
        bd.Close
        Set wrk = Nothing
    End If
   
    If mtpManutenção.GM_GerarNovoGMS001 = True Then CriarNovoGMS001
    If mtpManutenção.GM_GerarNovoGMS002 = True Then CriarNovoGMS002
    If mtpManutenção.GM_GerarNovoGMS005 = True Then CriarNovoGMS005
    If mtpManutenção.GM_GerarNovoBDLOG = True Then CriarNovoBDLOG
    If mtpManutenção.GM_CompactarGMS001 = True Then CompactarBaseDeDados (pstrLocacaobdGMS001)
    If mtpManutenção.GM_CompactarGMS002 = True Then CompactarBaseDeDados (pstrLocacaobdGMS002)
    If mtpManutenção.GM_CompactarGMS005 = True Then CompactarBaseDeDados (pstrLocacaobdGMS005)
    If mtpManutenção.GM_RepararGMS001 = True Then RepararBaseDeDados (pstrLocacaobdGMS001)
    If mtpManutenção.GM_RepararGMS002 = True Then RepararBaseDeDados (pstrLocacaobdGMS002)
    If mtpManutenção.GM_RepararGMS005 = True Then RepararBaseDeDados (pstrLocacaobdGMS005)
    
    Text1.Text = Text1.Text & ".......... Processo de Manutenção Finalizado: " & Date & " " & Time & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    Command1.Enabled = True
    
End Sub

Private Sub DTPickerGiro_Change(Index As Integer)
    Select Case Index
        Case 1
            If DTPickerGiro(1).Value > Date - 60 Then DTPickerGiro(1).Value = Date - 60
    End Select
End Sub

Private Sub DTPickerMov_Change(Index As Integer)
    Select Case Index
        Case 1
            If DTPickerMov(1).Value > Date - 60 Then DTPickerMov(1).Value = Date - 60
    End Select
End Sub

Private Sub Form_Load()

    Call ppCarregaPropriedadesForm(Me, 190)
    Call mpVerificaPermissao
    
    
    DTPickerGiro(1).Value = Date - 61
    DTPickerMov(1).Value = Date - 61
    
    
    For Each Control In FormPrincipal.Controls
        If TypeOf Control Is DTPicker Then Control.Enabled = False
        If TypeOf Control Is OptionButton Then Control.Enabled = False
     Next Control
    
    'Frame3.Enabled = False
        
End Sub
Sub CriarNovoGMS002()
    On Error GoTo Erro
    Dim lboBancoFechado As Boolean
    
    ''Me.Refresh
    DoEvents
    
    Text1.Text = vbCrLf & Text1.Text & ".......... Iniciando Novo BDGMS002" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Screen.MousePointer = 11
    pstrLocacaobdGMS002Novo = Replace(pstrLocacaobdGMS002, ".mdb", "Novo.mdb")
    
    lstrLocacaobdGMS002Limpo = Replace(pstrLocacaobdGMS002Novo, "Novo", "Limpo")
    FileCopy lstrLocacaobdGMS002Limpo, pstrLocacaobdGMS002Novo
    Dim Dbs As Database
    Set Dbs = OpenDatabase(pstrLocacaobdGMS002)
    'Me.Refresh
    DoEvents
        
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosRegistro: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosRegistro IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
                     "SELECT tProdutosRegistro.* FROM tProdutosRegistro"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tSecao: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tSecao IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
                "SELECT tSecao.* FROM tSecao"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tLocalizacao: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tLocalizacao IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tLocalizacao.* FROM tLocalizacao"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    Text1.Text = Text1.Text & "Exportando Tabela tLocalizacaoNumero: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tLocalizacaoNumero IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tLocalizacaoNumero.* FROM tLocalizacaoNumero"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tClientes: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tClientes IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tClientes.* FROM tClientes"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tTabelaPreco: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tTabelaPreco IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tTabelaPreco.* FROM tTabelaPreco"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tSubTabelaPreco: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tSubTabelaPreco IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tSubTabelaPreco.* FROM tSubTabelaPreco"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosBarra: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosBarra IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tProdutosBarra.* FROM tProdutosBarra"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosBarraAlternativa: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosBarraAlternativa IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tProdutosBarraAlternativa.* FROM tProdutosBarraAlternativa"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosBarratTabelaPreco: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosBarratTabelaPreco IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tProdutosBarratTabelaPreco.* FROM tProdutosBarratTabelaPreco"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosBarraComposicao: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosBarraComposicao IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tProdutosBarraComposicao.* FROM tProdutosBarraComposicao"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosBarraGiroDiario: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosBarraGiroDiario IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tProdutosBarraGiroDiario.* FROM tProdutosBarraGiroDiario"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tParametros: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tParametros IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tParametros.* FROM tParametros"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tVendasProcessadas: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tVendasProcessadas IN '" & Trim(pstrLocacaobdGMS002Novo) & "' " & _
            "SELECT tVendasProcessadas.* FROM tVendasProcessadas"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Dbs.Close: lboBancoFechado = True
    'Me.Refresh
    DoEvents
    
    Text1.Text = Text1.Text & "Base de Dados bdGMS002.mdb Exportada." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    
    Text1.Text = Text1.Text & "Excluindo Base Antiga: "
    Text1.SelStart = Len(Text1.Text)
        Kill pstrLocacaobdGMS002
        'Me.Refresh
        DoEvents
    Text1.Text = Text1.Text & "OK " & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    
    Text1.Text = Text1.Text & "Renomeando Base Nova: "
    Text1.SelStart = Len(Text1.Text)
        Name pstrLocacaobdGMS002Novo As pstrLocacaobdGMS002
        'Me.Refresh
        DoEvents
    Text1.Text = Text1.Text & "OK " & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    Screen.MousePointer = 0
    MsgBox "Erro: " & Err.Number & ". " & Err.Description, vbCritical, "CriarNovoGMS002"
    Text1.Text = Text1.Text & "FALHOU em CriarNovoGMS002" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    Text1.Text = Text1.Text & "Mensagem de Erro: " & Err.Number & " - " & Err.Description & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    If lboBancoFechado = False Then Dbs.Close
    'Me.Refresh
    DoEvents
End Sub
Sub CriarNovoGMS001()
    On Error GoTo Erro
    Dim lboBancoFechado As Boolean
    Screen.MousePointer = 11
    
    Text1.Text = Text1.Text & ".......... Iniciando Novo BDGMS001" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    pstrLocacaobdGMS001Novo = Replace(pstrLocacaobdGMS001, ".mdb", "Novo.mdb")
    lstrLocacaobdGMS001Limpo = Replace(pstrLocacaobdGMS001Novo, "Novo", "Limpo")
    FileCopy lstrLocacaobdGMS001Limpo, pstrLocacaobdGMS001Novo
    'Me.Refresh
    DoEvents
    
    
    Dim Dbs As Database
    Set Dbs = OpenDatabase(pstrLocacaobdGMS001)
    'Me.Refresh
    DoEvents
    
    Text1.Text = Text1.Text & "Exportando Tabela tComprador: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tComprador IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
                     "SELECT tComprador.* FROM tComprador"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tFornecedores: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tFornecedores IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
                     "SELECT tFornecedores.* FROM tFornecedores"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tFornecedoresLojas: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tFornecedoresLojas IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
                     "SELECT tFornecedoresLojas.* FROM tFornecedoresLojas"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
        
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosGrupoDef: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosGrupoDef IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
                     "SELECT tProdutosGrupoDef.* FROM tProdutosGrupoDef"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
                            
                            
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosGrupo1: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosGrupo1 IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
                     "SELECT tProdutosGrupo1.* FROM tProdutosGrupo1"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
            
            
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosGrupo2: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosGrupo2 IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
                     "SELECT tProdutosGrupo2.* FROM tProdutosGrupo2"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosGrupo3: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosGrupo3 IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
                     "SELECT tProdutosGrupo3.* FROM tProdutosGrupo3"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosGrupo4: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosGrupo4 IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
            "SELECT tProdutosGrupo4.* FROM tProdutosGrupo4"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
        
        
    Text1.Text = Text1.Text & "Exportando Tabela tTributacao: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tTributacao IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
            "SELECT tTributacao.* FROM tTributacao"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
        
    Text1.Text = Text1.Text & "Exportando Tabela tSincronizacao: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tSincronizacao IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
            "SELECT tSincronizacao.* FROM tSincronizacao"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tOperacao: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tOperacao IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
            "SELECT tOperacao.* FROM tOperacao"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tCondicaoPagto: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tCondicaoPagto IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
            "SELECT tCondicaoPagto.* FROM tCondicaoPagto"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutos: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutos IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
                     "SELECT tProdutos.* FROM tProdutos"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosGrupos: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutostGrupos IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
            "SELECT tProdutostGrupos.* FROM tProdutostGrupos"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tTributacaoGrupos: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tTributacaoGrupos IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
            "SELECT tTributacaoGrupos.* FROM tTributacaoGrupos"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tProdutosFilial: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tProdutosFilial IN '" & Trim(pstrLocacaobdGMS001Novo) & "' " & _
            "SELECT tProdutosFilial.* FROM tProdutosFilial"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Dbs.Close: lboBancoFechado = True
    'Me.Refresh
    DoEvents
        
    Text1.Text = Text1.Text & "Base de Dados BDGMS001.mdb Exportada." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Text1.Text = Text1.Text & "Excluindo Base Antiga: "
    Text1.SelStart = Len(Text1.Text)
        Kill pstrLocacaobdGMS001
        'Me.Refresh
        DoEvents
    Text1.Text = Text1.Text & "OK " & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    
    Text1.Text = Text1.Text & "Renomeando Base Nova: "
    Text1.SelStart = Len(Text1.Text)
        Name pstrLocacaobdGMS001Novo As pstrLocacaobdGMS001
        'Me.Refresh
        DoEvents
    Text1.Text = Text1.Text & "OK " & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    Screen.MousePointer = 0
    MsgBox "Erro: " & Err.Number & ". " & Err.Description, vbCritical, "CriarNovoGMS001"
    
    Text1.Text = Text1.Text & "FALHOU em CriarNovoGMS001" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Text1.Text = Text1.Text & "Mensagem de Erro: " & Err.Number & " - " & Err.Description & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    If lboBancoFechado = False Then Dbs.Close
    'Me.Refresh
    DoEvents
End Sub


Sub CriarNovoBDLOG()
    On Error GoTo Erro
    Dim lboBancoFechado As Boolean
    Screen.MousePointer = 11
    
    Text1.Text = vbCrLf & Text1.Text & ".......... INICIANDO EXPORTAÇÃO BDLOG" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    pstrLocacaobdLogNovo = Replace(pstrLocacaobdLog, ".mdb", "Novo.mdb")
    lstrLocacaobdLogLimpo = Replace(pstrLocacaobdLogNovo, "Novo", "Limpo")
    FileCopy lstrLocacaobdLogLimpo, pstrLocacaobdLogNovo
    'Me.Refresh
    DoEvents
    
    
    Dim Dbs As Database
    Set Dbs = OpenDatabase(pstrLocacaobdLog, False, False, ";pwd=" & pstrSenhabdLog)
        
    Text1.Text = Text1.Text & "Exportando Tabela tLogAcesso: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tLogAcesso IN '" & Trim(DbX) & "' " & _
                     "SELECT tLogAcesso.* FROM tLogAcesso"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tLogAcessoFuncao: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tLogAcessoFuncao IN '" & Trim(pstrLocacaobdLogNovo) & "' " & _
                     "SELECT tLogAcessoFuncao.* FROM tMovtoItens"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
        
    Dbs.Close: lboBancoFechado = True
    'Me.Refresh
    DoEvents
    
    Text1.Text = Text1.Text & "Base de Dados BDLOG.mdb Exportada com SUCESSO!" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Text1.Text = Text1.Text & "Excluindo Base Antiga: "
    Text1.SelStart = Len(Text1.Text)
        Kill pstrLocacaobdLog
        'Me.Refresh
        DoEvents
    Text1.Text = Text1.Text & "OK " & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    
    Text1.Text = Text1.Text & "Renomeando Base Nova: "
    Text1.SelStart = Len(Text1.Text)
        Name pstrLocacaobdLogNovo As pstrLocacaobdLog
        'Me.Refresh
        DoEvents
    Text1.Text = Text1.Text & "OK " & vbCrLf
    Text1.SelStart = Len(Text1.Text)
  
    
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    Screen.MousePointer = 0
    MsgBox "Erro: " & Err.Number & ". " & Err.Description, vbCritical, "CriarNovoGMS005"
    
    Text1.Text = Text1.Text & "FALHOU em CriarNovoGMS005" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Text1.Text = Text1.Text & "Mensagem de Erro: " & Err.Number & " - " & Err.Description & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    If lboBancoFechado = False Then Dbs.Close
    'Me.Refresh
    DoEvents
End Sub

Sub CriarNovoGMS005()
    On Error GoTo Erro
    Dim lboBancoFechado As Boolean
    Screen.MousePointer = 11
    
    Text1.Text = vbCrLf & Text1.Text & ".......... INICIANDO EXPORTAÇÃO BDGMS005" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    'pstrLocacaobdGMS001Novo = pstrLocacaobdGMS001 & "Novo"
    pstrLocacaobdGMS005Novo = Replace(pstrLocacaobdGMS005, ".mdb", "Novo.mdb")
    
    lstrLocacaobdGMS005Limpo = Replace(pstrLocacaobdGMS005Novo, "Novo", "Limpo")
    FileCopy lstrLocacaobdGMS005Limpo, pstrLocacaobdGMS005Novo
    
    
    Dim Dbs As Database
    Set Dbs = OpenDatabase(pstrLocacaobdGMS005)
        
    Text1.Text = Text1.Text & "Exportando Tabela tMovto: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tMovto IN '" & Trim(pstrLocacaobdGMS005Novo) & "' " & _
                     "SELECT tMovto.* FROM tMovto"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tMovtoItens: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tMovtoItens IN '" & Trim(pstrLocacaobdGMS005Novo) & "' " & _
                     "SELECT tMovtoItens.* FROM tMovtoItens"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    
    Text1.Text = Text1.Text & "Exportando Tabela tMovtoTitulos: "
    Text1.SelStart = Len(Text1.Text)
    Dbs.Execute "INSERT INTO tMovtoTitulos IN '" & Trim(pstrLocacaobdGMS005Novo) & "' " & _
                     "SELECT tMovtoTitulos.* FROM tMovtoTitulos"
    Text1.Text = Text1.Text & "OK - " & Dbs.RecordsAffected & " Registros." & vbCrLf
    Text1.SelStart = Len(Text1.Text)
        
    Dbs.Close: lboBancoFechado = True
       
    
    Text1.Text = Text1.Text & "Base de Dados BDGMS005.mdb Exportada com SUCESSO!" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Text1.Text = Text1.Text & "Excluindo Base Antiga: "
    Text1.SelStart = Len(Text1.Text)
        Kill pstrLocacaobdGMS005
    Text1.Text = Text1.Text & "OK " & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    
    Text1.Text = Text1.Text & "Renomeando Base Nova: "
    Text1.SelStart = Len(Text1.Text)
        Name pstrLocacaobdGMS005Novo As pstrLocacaobdGMS005
    Text1.Text = Text1.Text & "OK " & vbCrLf
    Text1.SelStart = Len(Text1.Text)
       
    
    
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    Screen.MousePointer = 0
    MsgBox "Erro: " & Err.Number & ". " & Err.Description, vbCritical, "CriarNovoGMS005"
    
    Text1.Text = Text1.Text & "FALHOU em CriarNovoGMS005" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Text1.Text = Text1.Text & "Mensagem de Erro: " & Err.Number & " - " & Err.Description & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    If lboBancoFechado = False Then Dbs.Close
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ppRotinasFinais
    pWrkArea.Close
End Sub

Private Sub OptionGiro_Click(Index As Integer)
    mtpManutenção.GM_GiroDiário = Index
    Select Case Index
        Case 1
            DTPickerGiro(1).Enabled = False
        Case 2
            DTPickerGiro(1).Enabled = True
    End Select
End Sub

Private Sub OptionMovimento_Click(Index As Integer)
        mtpManutenção.GM_Movimento = Index
    Select Case Index
        Case 1
            DTPickerMov(1).Enabled = False
        Case 2
            DTPickerMov(1).Enabled = True
        Case 3
            DTPickerMov(1).Enabled = False
    End Select
End Sub
Public Sub mpVerificaPermissao()
    Me.Enabled = pboPermissaoExc
    
    For Each Control In FormPrincipal.Controls
        'If TypeOf Control Is DTPicker Then Control.Enabled = pboPermissaoExc
        'If TypeOf Control Is OptionButton Then Control.Enabled = pboPermissaoExc
        If TypeOf Control Is CheckBox Then Control.Enabled = pboPermissaoExc
    Next Control
    
        
    'pboPermissaoInc
    'pboPermissaoAlt
    'pboPermissaoExc
    'pboPermissaoCon
    'pboPermissaoChv
    'pboPermissaoPrt
    'pboPermissaoAtu
    'pboPermissaoExe
    'pboPermissaoImp
    'pboPermissaoExp

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Sub CompactarBaseDeDados(BancoDeDados As String)
    On Error GoTo Erro
    Dim lboBancoFechado As Boolean
    Dim lstrBaseDeDados As String
    lstrBaseDeDados = Mid(BancoDeDados, InStrRev(BancoDeDados, "\") + 1, Len(BancoDeDados) - InStrRev(BancoDeDados, "\") + 1)
    
    Screen.MousePointer = 11
    
    Text1.Text = vbCrLf & Text1.Text & ".......... INICIANDO COMPACTAÇÃO DO " & lstrBaseDeDados & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    pstrLocacaobdNovo = Replace(BancoDeDados, ".mdb", "Novo.mdb")
    lstrLocacaobdLimpo = Replace(pstrLocacaobdNovo, "Novo", "Limpo")
    DBEngine.CompactDatabase BancoDeDados, pstrLocacaobdNovo
    'Me.Refresh
    DoEvents
    
    Dim Dbs As Database
    Set Dbs = OpenDatabase(pstrLocacaobdNovo)
    If Err.Number = 0 Then
        Dbs.Close
        Text1.Text = Text1.Text & "Excluindo Base Antiga: "
        Text1.SelStart = Len(Text1.Text)
            Kill BancoDeDados
            'Me.Refresh
            DoEvents
        Text1.Text = Text1.Text & "OK " & vbCrLf
        Text1.SelStart = Len(Text1.Text)
        
    
        Text1.Text = Text1.Text & "Renomeando Base Nova: "
        Text1.SelStart = Len(Text1.Text)
            Name pstrLocacaobdNovo As BancoDeDados
            'Me.Refresh
            DoEvents
        Text1.Text = Text1.Text & "OK " & vbCrLf
        Text1.SelStart = Len(Text1.Text)
    Else
        Text1.Text = Text1.Text & "Excluindo Base Incompleta: "
        Text1.SelStart = Len(Text1.Text)
            Kill pstrLocacaobdNovo
            'Me.Refresh
            DoEvents
        Text1.Text = Text1.Text & "OK " & vbCrLf
        Text1.SelStart = Len(Text1.Text)
    End If
  
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    Screen.MousePointer = 0
    MsgBox "Erro: " & Err.Number & ". " & Err.Description, vbCritical, "CriarNovoGMS005"
    
    Text1.Text = Text1.Text & "FALHOU em CriarNovoGMS005" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Text1.Text = Text1.Text & "Mensagem de Erro: " & Err.Number & " - " & Err.Description & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    Resume Next
    
End Sub

Sub RepararBaseDeDados(BancoDeDados As String)
    On Error GoTo Erro
    Screen.MousePointer = 11
    Dim lstrBaseDeDados As String
    lstrBaseDeDados = Mid(BancoDeDados, InStrRev(BancoDeDados, "\") + 1, Len(BancoDeDados) - InStrRev(BancoDeDados, "\") + 1)
        
    Text1.Text = vbCrLf & Text1.Text & ".......... INICIANDO REPARAÇÃO " & lstrBaseDeDados & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    
    
    DBEngine.RepairDatabase BancoDeDados
    'Me.Refresh
    DoEvents
    
    If Err.Number = 0 Then
        Text1.Text = Text1.Text & "Reparação do " & lstrBaseDeDados & " Bem Sucedida!" & vbCrLf
        Text1.SelStart = Len(Text1.Text)
        'Me.Refresh
        DoEvents
    End If
  
    Screen.MousePointer = 0
    Exit Sub
    
Erro:
    Screen.MousePointer = 0
    MsgBox "Erro: " & Err.Number & ". " & Err.Description, vbCritical, "RepararBaseDeDados"
    
    Text1.Text = Text1.Text & "FALHOU em CriarNovoGMS005" & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    
    Text1.Text = Text1.Text & "Mensagem de Erro: " & Err.Number & " - " & Err.Description & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    'Me.Refresh
    DoEvents
    Resume Next
    
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub
