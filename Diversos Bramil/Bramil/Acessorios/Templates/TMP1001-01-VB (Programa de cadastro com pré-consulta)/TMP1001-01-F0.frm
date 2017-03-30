VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form FormPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro_Padrao"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   1470
   ClientWidth     =   9360
   Icon            =   "TMP1001-01-F0.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9360
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   1535
      ButtonWidth     =   1640
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "    Incluir    "
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "    Alterar    "
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "    Excluir    "
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "    Chave    "
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   600
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8916
      _Version        =   327681
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro 1"
      TabPicture(0)   =   "TMP1001-01-F0.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFlexGrid(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cadastro 2"
      TabPicture(1)   =   "TMP1001-01-F0.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexGrid(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
         Height          =   4455
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7858
         _Version        =   65541
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
         Height          =   4455
         Index           =   1
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7858
         _Version        =   65541
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TMP1001-01-F0.frx":047A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TMP1001-01-F0.frx":058C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TMP1001-01-F0.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TMP1001-01-F0.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TMP1001-01-F0.frx":0EDA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------'
'Codigo Programa: TMP1001-01-VB                                                              '
'Descr.Programa.: Modelo de Manutencao de Cadastro Simples com Tela de Pré-Consulta          '
'Analista.......: Geraldo Coimbra                                                            '
'Programador....: Jorge A M de Carvalho / Victor Augusto Faria Dias                          '
'Data Criação...: 24/04/1998                                                                 '
'Data Alteração.:                                                                            '
'--------------------------------------------------------------------------------------------'
Option Explicit

Private mboCarregaGrid As Boolean

Public Sub mpVerificaPermissao()
    Toolbar.Buttons(1).Enabled = pboPermissaoInc
    Toolbar.Buttons(2).Enabled = pboPermissaoAlt
    Toolbar.Buttons(3).Enabled = pboPermissaoExc
    Toolbar.Buttons(4).Enabled = pboPermissaoChv

    Toolbar.Buttons(1).Value = tbrUnpressed
    Toolbar.Buttons(2).Value = tbrUnpressed
    Toolbar.Buttons(3).Value = tbrUnpressed
    Toolbar.Buttons(4).Value = tbrUnpressed
End Sub

Public Sub mpInicializaGrid(liTipoGrid As Integer)
    Dim lbIndex As Byte

    With MSFlexGrid(liTipoGrid)
    
        .Clear
        .Rows = 1
        .Cols = 4
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000
        
        .TextMatrix(0, 1) = "Código"
        .TextMatrix(0, 2) = "Descrição"
        
        Select Case liTipoGrid
            Case 0 'Grid do Cadastro 1
                .ColWidth(2) = 1000
                .ColWidth(3) = 0
                .TextMatrix(0, 2) = "Nome"
            Case 1 'Grid do Cadastro 2
                .ColWidth(2) = 1000
                .ColWidth(3) = 1000
                .TextMatrix(0, 2) = "Valor"
                .TextMatrix(0, 3) = "Data"
        End Select
        'Alinhamento das Colunas
        For lbIndex = 1 To Cols - 1
            .FixedAlignment(lbIndex) = 4
            .ColAlignment(lbIndex) = 1
        Next lbIndex
        
        .SelectionMode = flexSelectionByRow
    End With
End Sub

Private Sub Form_Activate()
    
    Call mpCarregaGrid(SSTab.Tab)  'Atualizando o Grid dos Cadastros
    
    pstrFuncaoToolbar = "Consulta"
    mboCarregaGrid = False
    
    'Inicialização das Variaveis auxiliares das chaves
    piCodigo2 = 0: piCodigo1 = 0
    
    MSFlexGrid.Item(0).Enabled = pboPermissaoCon
    MSFlexGrid.Item(1).Enabled = pboPermissaoCon
End Sub

Private Sub Form_Load()
    'Verificação da permissao de execução do programa
    If Not pfboRotinasIniciais("Codigo-do-Programa") Then End

    '1001 = Código do arquivo de help
    Call ppCarregaPropriedadesForm(Me, 1001)
    
    'Abre o Banco de Dados
    Call ppAbre_BDAcesso(pstrCodPrograma, pdbConfus, pstrLocacaobdConfus)
    
    'Verifica Permissões
    Call mpVerificaPermissao
    
    SSTab.Tab = 0
End Sub

Public Sub mpCarregaGrid(liTipoGrid As Integer)
    On Error GoTo Erro

    mboCarregaGrid = True
    
    'Escolher uma das duas formas de carregamento do Grid (1-Manual / 2-Data Control)
    
    'Forma 1 (Manual)
    '---------------------------------------------------------------------------------------
    Call mpInicializaGrid(liTipoGrid)
    
    If pboPermissaoCon = False Then Exit Sub
    Screen.MousePointer = 11
    
    Err.Clear
    'Select de preenchimento do Grid
    pstrSql = "SELECT DISTINCTROW tUsuariosGrupo.bCodGrupoUsuariotGrpUsu, tUsuariosGrupo.strDescrGrupoUsuariotGrpUsu, tUsuarios.iCodUsuariotUsu, tUsuarios.strNomeUsuariotUsu, tUsuarios.bTabNivelUsuariotUsu, tUsuarios.strSenhaUsuariotUsu, tUsuarios.bDiasValidadeSenhatUsu, tUsuarios.dtDataSenhatUsu, tUsuarios.dtDataValidadeUsuariotUsu, tUsuarios.strNomeUsuarioResumidotUsu FROM tUsuariosGrupo LEFT JOIN tUsuarios ON tUsuariosGrupo.bCodGrupoUsuariotGrpUsu = tUsuarios.bCodGrupoUsuariotUsu " & IIf(liTipoGrid = 1, " ORDER BY tUsuarios.strNomeUsuariotUsu, tUsuarios.iCodUsuariotUsu;", " ORDER BY tUsuariosGrupo.strDescrGrupoUsuariotGrpUsu, tUsuarios.strNomeUsuariotUsu, tUsuarios.iCodUsuariotUsu;")
       
    Set prsSeleção = pdbConfus.OpenRecordset(pstrSql)
                      
    '-- Variaveis de Preencimento do Grid --'
    Dim lbCodGrupoUsuariotUsu As Byte
    Dim liCodUsuariotUsu As Integer
    Dim lstrDescrGrupoUsuario0 As String
    Dim lstrDescrGrupoUsuario1 As String
    Dim lbTabNivelUsuariotUsu As String
    Dim lstrDescrNivelUsuario As String
    Dim lstrSenhaUsuariotUsu As String
    Dim lbDiasValidadeSenhatUsu As String
    Dim ldtDataSenhatUsu As String
    Dim ldtDataValidadeUsuariotUsu As String
    Dim lstrNomeUsuarioResumidotUsu As String
    '----------------------------------------'
    While Not prsSeleção.EOF
        lstrDescrGrupoUsuario0 = ""
        lstrDescrGrupoUsuario1 = ""
        lbTabNivelUsuariotUsu = ""
        lstrDescrNivelUsuario = ""
        lstrSenhaUsuariotUsu = ""
        lbDiasValidadeSenhatUsu = ""
        ldtDataSenhatUsu = ""
        ldtDataValidadeUsuariotUsu = ""
        lstrNomeUsuarioResumidotUsu = ""
        
        If prsSeleção.Fields("bCodGrupoUsuariotGrpUsu") <> lbCodGrupoUsuariotUsu Then
            '-- Linha de Grupo --'
            
            lbCodGrupoUsuariotUsu = Format(prsSeleção.Fields("bCodGrupoUsuariotGrpUsu"), "00")
            lstrDescrGrupoUsuario0 = Format(prsSeleção.Fields("bCodGrupoUsuariotGrpUsu"), "00")
            If liTipoGrid = 2 Then lstrDescrGrupoUsuario1 = Format(prsSeleção.Fields("bCodGrupoUsuariotGrpUsu"), "00") & " - " & prsSeleção.Fields("strDescrGrupoUsuariotGrpUsu") Else lstrDescrGrupoUsuario1 = prsSeleção.Fields("strDescrGrupoUsuariotGrpUsu")
        Else
            '-- Linha de Usuário --'
            If prsSeleção.Fields("bTabNivelUsuariotUsu") > 0 And liTipoGrid <> 0 Then
                Select Case prsSeleção.Fields("bTabNivelUsuariotUsu")
                    Case 1
                        lstrDescrNivelUsuario = " - Operador"
                    Case 2
                        lstrDescrNivelUsuario = " - Coordenador"
                    Case 3
                        lstrDescrNivelUsuario = " - Supervisor"
                    Case 4
                        lstrDescrNivelUsuario = " - Administrador"
                    Case Else
                        lstrDescrNivelUsuario = " - Desconhecido"
                End Select
                lstrNomeUsuarioResumidotUsu = IIf(IsNull(prsSeleção.Fields("strNomeUsuarioResumidotUsu")), "", prsSeleção.Fields("strNomeUsuarioResumidotUsu"))
                lbTabNivelUsuariotUsu = IIf(IsNull(prsSeleção.Fields("bTabNivelUsuariotUsu")), "", prsSeleção.Fields("bTabNivelUsuariotUsu"))
                lstrSenhaUsuariotUsu = pfCriptaSenha(IIf(IsNull(prsSeleção.Fields("strSenhaUsuariotUsu")), "", prsSeleção.Fields("strSenhaUsuariotUsu")))
                lbDiasValidadeSenhatUsu = IIf(IsNull(prsSeleção.Fields("bDiasValidadeSenhatUsu")), "", prsSeleção.Fields("bDiasValidadeSenhatUsu"))
                ldtDataSenhatUsu = Format(IIf(IsNull(prsSeleção.Fields("dtDataSenhatUsu")), "", prsSeleção.Fields("dtDataSenhatUsu")), "dd/mm/yyyy")
                ldtDataValidadeUsuariotUsu = Format(IIf(IsNull(prsSeleção.Fields("dtDataValidadeUsuariotUsu")), "", prsSeleção.Fields("dtDataValidadeUsuariotUsu")), "dd/mm/yyyy")
                lstrDescrGrupoUsuario0 = Format(IIf(IsNull(prsSeleção.Fields("iCodUsuariotUsu")), "", prsSeleção.Fields("iCodUsuariotUsu")), "000")
                If liTipoGrid = 2 Then lstrDescrGrupoUsuario1 = "        " & Format(IIf(IsNull(prsSeleção.Fields("iCodUsuariotUsu")), "", prsSeleção.Fields("iCodUsuariotUsu")), "000") & " - " & prsSeleção.Fields("strNomeUsuariotUsu") Else lstrDescrGrupoUsuario1 = prsSeleção.Fields("strNomeUsuariotUsu")
            End If
            prsSeleção.MoveNext
        End If
        
        If liTipoGrid = 0 And lbTabNivelUsuariotUsu = "" Or _
            liTipoGrid = 1 And lbTabNivelUsuariotUsu <> "" Or _
            liTipoGrid = 2 Then
        
            If Len(Trim(lstrDescrGrupoUsuario1)) > 0 Then
                With MSFlexGrid(liTipoGrid)
                    .Rows = .Rows + 1: .Row = .Rows - 1
                    
                    If liTipoGrid = 2 Then
                        .Col = 2
                        .CellFontBold = IIf(lbTabNivelUsuariotUsu = "", True, False)
                    End If
                    
                    .TextMatrix(.Row, 1) = lstrDescrGrupoUsuario0
                    .TextMatrix(.Row, 2) = lstrDescrGrupoUsuario1
                    .TextMatrix(.Row, 3) = IIf(lbTabNivelUsuariotUsu = "", "", lbTabNivelUsuariotUsu & lstrDescrNivelUsuario)
                    .TextMatrix(.Row, 4) = lstrNomeUsuarioResumidotUsu
                    .TextMatrix(.Row, 5) = ldtDataValidadeUsuariotUsu
                    .TextMatrix(.Row, 6) = lstrSenhaUsuariotUsu
                    .TextMatrix(.Row, 7) = lbDiasValidadeSenhatUsu
                    .TextMatrix(.Row, 8) = ldtDataSenhatUsu
                End With
            End If
        End If
    Wend
    MSFlexGrid(liTipoGrid).Row = 0
    prsSeleção.Close
    mboCarregaGrid = False
    Screen.MousePointer = 0
    '----------------------------------------------------------------------------------------
    
    'Forma 2 (Data Control)
    '----------------------------------------------------------------------------------------
    If pboPermissaoCon = False Then Exit Sub
    
    Screen.MousePointer = 11
    
    Err.Clear
    
    If liTipoGrid = 0 Then
        pstrSql = "SELECT ..."
    End If
    
    If liTipoGrid = 1 Then
        pstrSql = "SELECT ..."

    End If
    
    DataMsFlexGrid.RecordSource = pstrSql
    DataMsFlexGrid.Refresh
    MSFlexGrid(liTipoGrid).Refresh
    
    Call mpInicializaGrid(liTipoGrid)
       
    mboCarregaGrid = False
    
    Screen.MousePointer = 0
    '----------------------------------------------------------------------------------------
    Exit Sub
    
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "mpCarregaGrid"
    Err.Clear
End Sub

Private Sub Form_Terminate()
    Call ppRotinasFinais
    pdbConfus.Close
    End
End Sub

Private Sub MSFlexGrid_DblClick(Index As Integer)
    If Index = 0 Then piCodigo1 = MSFlexGrid(Index).TextMatrix(MSFlexGrid(Index).Row, 1)
    If Index = 1 Then piCodigo2 = MSFlexGrid(Index).TextMatrix(MSFlexGrid(Index).Row, 1)
    
    If piCodigo1 <> 0 Then FormCadastro1.Show 1
    If piCodigo2 <> 0 Then FormCadastro2.Show 1
End Sub

Private Sub MSFlexGrid_EnterCell(Index As Integer)
    On Error GoTo Erro
    
    If mboCarregaGrid = True Then Exit Sub
    
    piCodigo2 = 0: piCodigo1 = 0
    
    If Index = 0 Then piCodigo1 = MSFlexGrid(Index).TextMatrix(MSFlexGrid(Index).Row, 1)
    If Index = 1 Then piCodigo2 = MSFlexGrid(Index).TextMatrix(MSFlexGrid(Index).Row, 1)
        
    Exit Sub
    
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "MSFlexGrid_EnterCell"
    Err.Clear
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    Call mpCarregaGrid(SSTab.Tab)
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Toolbar.Buttons.Item(Button.Index).Value = tbrPressed
    
    Select Case Button.Index
        Case 1
            pstrFuncaoToolbar = "Inclusão"
        Case 2
            pstrFuncaoToolbar = "Alteração"
        Case 3
            pstrFuncaoToolbar = "Exclusão"
        Case 4
            pstrFuncaoToolbar = "Chave"
    End Select
    
    If SSTab.Tab = 0 Or piCodigo1 <> 0 Then FormCadastro1.Show 1
    If SSTab.Tab = 1 Or piCodigo2 <> 0 Then FormCadastro2.Show 1
    
    Toolbar.Buttons.Item(Button.Index).Value = tbrUnpressed
End Sub
