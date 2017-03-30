VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmPlano_Servicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plano de Serviços"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlano_Servicos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7845
   Begin TabDlg.SSTab sstPlano 
      Height          =   5055
      Left            =   0
      TabIndex        =   14
      Top             =   330
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Geral"
      TabPicture(0)   =   "frmPlano_Servicos.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "hfgServico"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCodigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDescricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmPlano_Servicos.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "cbbCampos"
      Tab(1).Control(2)=   "hfgPlano"
      Tab(1).Control(3)=   "txtConsulta"
      Tab(1).Control(4)=   "cmdRefresh"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdConsulta"
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame1 
         Caption         =   "Serviços"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   120
         TabIndex        =   20
         Top             =   1230
         Width           =   7575
         Begin VB.TextBox txtServico 
            Height          =   360
            Left            =   120
            TabIndex        =   2
            Top             =   570
            Width           =   1365
         End
         Begin VB.TextBox txtLimite 
            Height          =   360
            Left            =   2100
            TabIndex        =   5
            Top             =   1230
            Width           =   1425
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "Incluir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5100
            TabIndex        =   7
            ToolTipText     =   "Incluir"
            Top             =   1230
            Width           =   1125
         End
         Begin VB.CommandButton cmdRemover 
            Caption         =   "Remover"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6300
            TabIndex        =   8
            ToolTipText     =   "Remover"
            Top             =   1230
            Width           =   1155
         End
         Begin MSDataListLib.DataCombo dtcServico 
            Height          =   360
            Left            =   1530
            TabIndex        =   3
            Top             =   570
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   635
            _Version        =   393216
            ForeColor       =   8388608
            Text            =   ""
         End
         Begin AutoCompletar.CbCompleta cbbControle 
            Height          =   360
            Left            =   120
            TabIndex        =   4
            Top             =   1230
            Width           =   1935
            _ExtentX        =   3413
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
            ForeColor       =   8388608
         End
         Begin AutoCompletar.CbCompleta cbbPeriodo 
            Height          =   360
            Left            =   3570
            TabIndex        =   6
            Top             =   1230
            Width           =   1455
            _ExtentX        =   2566
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
            ForeColor       =   8388608
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Período"
            Height          =   240
            Left            =   3570
            TabIndex        =   24
            Top             =   990
            Width           =   645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   330
            Width           =   585
         End
         Begin VB.Label lblControle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   240
            Left            =   2100
            TabIndex        =   22
            Top             =   990
            Width           =   510
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Controle"
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   990
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdConsulta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68040
         Picture         =   "frmPlano_Servicos.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -67650
         Picture         =   "frmPlano_Servicos.frx":34B4
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72900
         TabIndex        =   11
         Top             =   780
         Width           =   4785
      End
      Begin VB.TextBox txtDescricao 
         Height          =   360
         Left            =   1410
         MaxLength       =   40
         TabIndex        =   1
         Top             =   780
         Width           =   6285
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   780
         Width           =   1245
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgPlano 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   13
         Top             =   1200
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   6588
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
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
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   -74880
         TabIndex        =   10
         Top             =   780
         Width           =   1935
         _ExtentX        =   3413
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
         ForeColor       =   8388608
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgServico 
         Height          =   1875
         Left            =   120
         TabIndex        =   9
         Top             =   3030
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3307
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   18
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   1410
         TabIndex        =   17
         Top             =   540
         Width           =   825
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   540
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alt + N"
            Description     =   "Novo"
            Object.ToolTipText     =   "Novo registro - CTRL+N"
            ImageIndex      =   4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Confirmar"
            Object.ToolTipText     =   "Gravar registro - CTRL+G"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Excluir"
            Object.ToolTipText     =   "Excluir registro - CTRL+E"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Imprimir"
            Object.ToolTipText     =   "Imprimir - CTRL+I"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sair"
            Object.ToolTipText     =   "Sair - CTRL+S"
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Integração"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9060
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlano_Servicos.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlano_Servicos.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlano_Servicos.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlano_Servicos.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlano_Servicos.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlano_Servicos.frx":5578
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPlano_Servicos.frx":5892
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPlano_Servicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro de Planos de Serviços                                 '
' Data de Criação........: 25/08/2003                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.: 07/11/2005                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSet As String
Dim strTamanho As String
Dim strNomes As String
Dim strTamanho_servico As String
Dim strNomes_servico As String
Dim strIDServiço As String
Public strCombo As String
Public strConsulta As String
Dim strId_Estados As String
Dim strClique_servico As String
Dim strId_remover As String
Public strSql As String
Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim conexao As New DLLConexao_Sistema.conexao
Dim I As Integer
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log

Function Imprimir()
    On Error GoTo Erro
    'Tratamento de Erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Logicx"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Plano_Servico.Show
    
    Unload frmAguarde
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty

    If cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       cmdConsulta.SetFocus
    Else
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdIncluir_Click()
    Dim strIndice As String
    Dim intContador As Integer
    Dim strLimite As String
    Dim intResp As Integer
    Dim intVerifica_Grupo As Integer
    Dim booAcrescenta_Grupo As Integer
    
    If txtServico.Text = Empty Then
       MsgBox "Serviço não definido. Verifique.", vbInformation, "Only Tech"
       txtServico.SetFocus
       Exit Sub
    End If

    'Verificar se o item está no grid de itens do pedido
    intContador = 1
    Do While intContador <= hfgServico.Rows - 1
        hfgServico.Row = intContador
        hfgServico.Col = 2
        If cmdIncluir.Caption = "Alterar" Then
           If hfgServico.Text = txtServico.Text And hfgServico.Row <> strClique_servico Then
              MsgBox "O Serviço alterado pertence a outro item neste cadastro. Verifique.", vbInformation, "Only Tech"
              txtServico.SetFocus
              Exit Sub
           End If
        Else
           If hfgServico.Text = txtServico.Text Then
              MsgBox "Serviço já definido para este cadastro. Verifique.", vbInformation, "Only Tech"
              'Limpando os campos dos Itens
              txtServico.Text = Empty
              txtLimite.Text = Empty
              cbbControle.Text = Empty
              cbbPeriodo.Text = Empty
              txtServico.SetFocus
              Exit Sub
           End If
        End If
        intContador = intContador + 1
    Loop
    
    'Verificando se o serviço é por grupo, caso seja, o limite e o período devem ser o mesmo
    If cbbControle.Text = "Grupo Serviços" Then
       intVerifica_Grupo = 1
       booAcrescenta_Grupo = False
       Do While intVerifica_Grupo <= hfgServico.Rows - 1
          hfgServico.Row = intVerifica_Grupo
          hfgServico.Col = 4
          If hfgServico.Text = "Grupo Serviços" Then
             hfgServico.Col = 5
             strLimite = hfgServico.Text
             hfgServico.Col = 6
             If strLimite <> txtLimite.Text Or cbbPeriodo.Text <> hfgServico.Text Then
                booAcrescenta_Grupo = True
             ElseIf strLimite = txtLimite.Text And cbbPeriodo.Text = hfgServico.Text Then
                booAcrescenta_Grupo = False
                Exit Do
             End If
          End If
          intVerifica_Grupo = intVerifica_Grupo + 1
       Loop
       'a boolean verifica se existem grupos com esse limite e período
       'caso exista, é incluído, caso nao exista, há opção de se criar um novo grupo
       If booAcrescenta_Grupo = True Then
          intResp = MsgBox("Já existe um ou mais grupos com limites e períodos diferentes deste determinado. Deseja incluir um novo grupo com esses parâmetros?", vbYesNo, "Only Tech")
          If intResp = 7 Then
             Exit Sub
          End If
       End If
    End If
    
    hfgServico.Row = hfgServico.TopRow
    If cmdIncluir.Caption = "Incluir" Then
       If hfgServico.Text <> Empty Then
          strIndice = intContador
          hfgServico.Rows = hfgServico.Rows + 1
       Else
          strIndice = intContador - 1
       End If
    Else
       strIndice = strClique_servico
    End If
    
    hfgServico.Row = strIndice
    
    hfgServico.Col = 0
    hfgServico.ColWidth(0) = 500
    hfgServico.Font.Name = "Tahoma"
    hfgServico.CellFontSize = 7
    hfgServico.CellFontBold = False
    hfgServico.CellBackColor = &H80FFFF
    hfgServico.Text = strIndice
    
    hfgServico.Col = 2
    hfgServico.Text = txtServico.Text
    
    hfgServico.Col = 3
    hfgServico.Text = dtcServico.Text
    
    hfgServico.Col = 4
    hfgServico.Text = cbbControle.Text
    
    hfgServico.Col = 5
    hfgServico.Text = txtLimite.Text
    
    hfgServico.Col = 6
    hfgServico.Text = cbbPeriodo.Text
    
    hfgServico.Refresh
    
    txtServico.Text = Empty
    txtLimite.Text = Empty
    cbbControle.Text = Empty
    cbbPeriodo.Text = Empty
    
    cmdIncluir.Caption = "Incluir"
    hfgServico.Col = 0: hfgServico.Row = 1
    txtServico.SetFocus
    
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub cmdRemover_Click()
    Dim intContador As Integer
    
    hfgServico.Col = 0
    If hfgServico.Text = Empty Then
       MsgBox "Não há Serviço selecionado.", vbInformation, "Only Tech"
       txtServico.SetFocus
       Exit Sub
    End If
    
    'Guardando os Ids removidos para serem deletados no evento gravar
    hfgServico.Col = 1
    If hfgServico.Text <> Empty Then
       If strId_remover = Empty Then
          strId_remover = hfgServico.Text
       Else
          strId_remover = strId_remover + "," + hfgServico.Text
       End If
    End If

    If hfgServico.Rows <= 2 Then
       txtServico.Text = Empty
       txtLimite.Text = Empty
       hfgServico.Clear
       Movimentacoes.Monta_HFlex_Grid hfgServico, strTamanho_servico, strNomes_servico, 6, "Otica", Me
    Else
       hfgServico.RemoveItem (hfgServico.Row)
       intContador = 1
       hfgServico.Col = 0
       Do While intContador <= hfgServico.Rows - 1
          hfgServico.Row = intContador
          hfgServico.Text = intContador
          intContador = intContador + 1
       Loop
    End If
    
    hfgServico.Refresh
    
    txtServico.Text = Empty
    txtLimite.Text = Empty
    cbbControle.Text = Empty
    cbbPeriodo.Text = Empty

    cmdIncluir.Caption = "Incluir"

    hfgServico.Col = 0
    hfgServico.Row = 1
End Sub

Private Sub dtcServico_GotFocus()
   If txtServico.Text = Empty Then
      Call Movimentacoes.Verifica_DataCombo(dtcServico.Text)
   End If
End Sub

Private Sub dtcServico_LostFocus()
   txtServico.Text = dtcServico.BoundText
   If IsNumeric(txtServico.Text) = False Or dtcServico.Text = Empty Then txtServico.Text = Empty: Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 78: If booPrivilegio_Incluir = True Then Call Novo     'CTRL+N
                       Case 71: If booPrivilegio_Incluir = True Then Call Gravar   'CTRL+G
                       Case 67: If booPrivilegio_Incluir = True Then Call Cancelar 'CTRL+C
                       Case 69: If booPrivilegio_Excluir = True Then Call Excluir  'CTRL+E
                       Case 83: Unload Me  'CTRL+S
                End Select
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Erro
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Planos de Serviço "
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstPlano, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando Cadastro Planos de Serviços"
    'Gravando o log
    log.Gravar_log "Otica", Me
        
    sstPlano.TabEnabled(0) = False
    sstPlano.Tab = 1
    strClique_servico = 0
    Call Reposicao
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    strEvento_log = "Unload"
    
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    
    strCombo = Empty
    strConsulta = Empty
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
        
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Unload")
    Exit Sub
End Sub

Private Sub hfgPlano_Click()
    If hfgPlano.Col = 0 And hfgPlano.Text <> Empty Then
    
       On Error Resume Next
       cbbControle.Text = Empty
       cbbPeriodo.Text = Empty
    
       'Novo
       tlbBotoes.Buttons.Item(1).Enabled = False
       'Gravar
       tlbBotoes.Buttons.Item(2).Enabled = booPrivilegio_Alterar
       'Cancelar
       tlbBotoes.Buttons.Item(3).Enabled = booPrivilegio_Alterar
       'Excluir
       tlbBotoes.Buttons.Item(4).Enabled = booPrivilegio_Excluir
       'Imprimir
       tlbBotoes.Buttons.Item(5).Enabled = False
       'Integração
       If booIntegra_Portal = True Then
          tlbBotoes.Buttons.Item(9).Enabled = True
       End If
       
       frmAguarde.Show
       DoEvents
    
       txtCodigo.Text = hfgPlano.TextArray((hfgPlano.Row * hfgPlano.Cols + hfgPlano.Col + 1))
       txtDescricao.Text = hfgPlano.TextArray((hfgPlano.Row * hfgPlano.Cols + hfgPlano.Col + 2))

       'ABASTECENDO OS SERVICOS
       strSql = "SELECT PkId_TBPlano_servico_servico_laboratorio," & _
                "FKCodigo_TBServico_laboratorio,DFDescricao_TBServico_laboratorio," & _
                "DFControle_TBPlano_servico_servico_laboratorio," & _
                "DFQuantidade_TBPlano_servico_servico_laboratorio," & _
                "DFPeriodo_TBPlano_servico_servico_laboratorio " & _
                "FROM TBPlano_servico_servico_laboratorio " & _
                "INNER JOIN TBServico_laboratorio ON TBPlano_servico_servico_laboratorio.FKCodigo_TBServico_laboratorio = TBServico_laboratorio.pKCodigo_TBServico_laboratorio " & _
                "WHERE FKCodigo_TBPlano_servico = " & txtCodigo.Text & " " & _
                "ORDER BY PkId_TBPlano_servico_servico_laboratorio"
        
       Call Movimentacoes.Movimenta_HFlex_Grid(strSql, hfgServico, strTamanho_servico, strNomes_servico, "BDRetaguarda", "Otica", Me)
        
       hfgServico.Row = 1
       hfgServico.Col = 0
       If hfgServico.Text = Empty Then
          hfgServico.Rows = 2
          Movimentacoes.Monta_HFlex_Grid hfgServico, strTamanho_servico, strNomes_servico, 6, "Otica", Me
       Else
          intContador = 1
          hfgServico.Col = 4
          Do While intContador <= hfgServico.Rows - 1
             hfgServico.Row = intContador
             If hfgServico.Text = "1" Then
                hfgServico.Text = "Valor Contrato"
             ElseIf hfgServico.Text = "2" Then
                hfgServico.Text = "Serviços"
             ElseIf hfgServico.Text = "3" Then
                hfgServico.Text = "Grupo Serviços"
             End If
             intContador = intContador + 1
          Loop
          hfgServico.ColAlignment(4) = 1
       End If
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstPlano.TabEnabled(0) = True
       sstPlano.Tab = 0
       txtDescricao.SetFocus
    End If
    Unload frmAguarde
End Sub

Private Sub hfgPlano_DblClick()
    hfgPlano.Sort = 1
End Sub

Private Sub hfgPlano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgPlano_Click
    End If
End Sub

Private Sub hfgServico_Click()
    If hfgServico.Col = 0 And hfgServico.Text <> Empty And hfgServico.Row <> strClique_servico Then
        txtServico.Text = Empty
        txtLimite.Text = Empty
        cbbControle.Text = Empty
        cbbPeriodo.Text = Empty
        cmdIncluir.Caption = "Incluir"
    End If
End Sub

Private Sub hfgServico_DblClick()
   If hfgServico.Col = 0 And hfgServico.Text <> Empty Then
       strClique_servico = hfgServico.Row
       cmdIncluir.Caption = "Alterar"
       txtServico.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 2))
       cbbControle.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 4))
       txtLimite.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 5))
       cbbPeriodo.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 6))
    End If
    hfgServico.SetFocus
End Sub

Private Sub sstPlano_Click(PreviousTab As Integer)
    If sstPlano.Tab = 0 Then
       txtDescricao.SetFocus
    ElseIf sstPlano.Tab = 1 Then
        If frmIntegracao.Visible = True Then
           Unload frmIntegracao
        End If
        If strCombo <> Empty And strCombo <> "Todos" And strCombo <> "Controle" Then
           cbbCampos.Text = strCombo
           txtConsulta.SetFocus
        ElseIf strCombo = "Todos" Then
           hfgPlano.Row = 1
           hfgPlano.Col = 0
           hfgPlano.SetFocus
        End If
    End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstPlano.Tab = 0: Call Gravar
           Case 3: Call Cancelar
           Case 4 And sstPlano.Tab = 0: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
    End Select
End Sub

Function Gravar()
    On Error GoTo Erro
    
    'Verifica se os campos necessarios para gravar não estão nulos
    If txtDescricao.Text = Empty Then
       MsgBox "O campo descrição do Serviço não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtDescricao.SetFocus
       Exit Function
    End If
    
    hfgServico.Col = 0
    hfgServico.Row = 1
    If hfgServico.Text = Empty Then
       MsgBox "Não há Serviços a serem cadastrados. Verifique.", vbInformation, "Only Tech"
       Exit Function
    End If

    Dim strCampo As String
    Dim strValores As String
    Dim intContador As Integer
    Dim intControle As Integer
    Dim intPeriodo As Integer
    Dim strQuantidade As String
    Dim strCodigo As String
    
    Dim intEmpresa As Integer
    Dim strCodigo_Plano_Servico As String
    Dim strProx_Cod_Plano As String
    Dim rstVerifica As New ADODB.Recordset
    
    intEmpresa = MDIPrincipal.OCXUsuario.Empresa
    
    Call Objetos.Maiusculo_TXT(Me)
    
    If booAlterar = False Then
       strProx_Cod_Plano = Funcoes_Gerais.Localiza_Proximo_Codigo("DFProximo_plano_servico_TBParametros_servicos", "FKCodigo_TBEmpresa", intEmpresa, "TBParametros_Servicos", "Otica", Me, "BDRetaguarda")
       txtCodigo.Text = strProx_Cod_Plano
    End If
    
    strCampo = "PKCodigo_TBPlano_servico,DFDescricao_TBPlano_servico," & _
               "DFData_alteracao_TBPlano_servico,DFIntegrado_filiais_TBPlano_servico "
               
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBPlano_servico "
    End If

    strValores = "" & txtCodigo.Text & ",'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
        strValores = strValores & ",0"
    End If
    
    'Abrindo conexao
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    On Error GoTo Erro_transacao
    
    If booAlterar = True Then
       
       log.Evento = "Alterar"
       
       strSet = "UPDATE TBPlano_servico " & _
                "SET DFDescricao_TBPlano_servico = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                "DFData_alteracao_TBPlano_servico = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBPlano_servico = 0 "
                
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBPlano_servico = 0 "
       End If
       
       strSet = strSet & "WHERE PKCodigo_TBPlano_servico = " & txtCodigo.Text & ""
       
       conexao.CNConexao.Execute strSet

       If hfgServico.Text <> Empty Then
          intContador = 1
          Do While intContador <= hfgServico.Rows - 1
          
              hfgServico.Row = intContador
              
              hfgServico.Col = 2
              strCodigo = hfgServico.Text
              
              hfgServico.Col = 4
              
              If hfgServico.Text = "Valor Contrato" Then
                 intControle = 1
              ElseIf hfgServico.Text = "Serviços" Then
                 intControle = 2
              ElseIf hfgServico.Text = "Grupo Serviços" Then
                 intControle = 3
              Else
                 intControle = 0
              End If
              
              hfgServico.Col = 5
              strQuantidade = hfgServico.Text
              
              hfgServico.Col = 6
              intPeriodo = hfgServico.Text
              
              hfgServico.Col = 1
              If hfgServico.Text <> Empty Then
              
                 strSql = Empty
                 strSql = "UPDATE TBPlano_servico_servico_laboratorio SET FKCodigo_TBServico_laboratorio = " & strCodigo & "," & _
                          "FKCodigo_TBPlano_servico = '" & txtCodigo.Text & "'," & _
                          "DFQuantidade_TBPlano_servico_servico_laboratorio = '" & strQuantidade & "'," & _
                          "DFControle_TBPlano_servico_servico_laboratorio = '" & intControle & "'," & _
                          "DFPeriodo_TBPlano_servico_servico_laboratorio = '" & intPeriodo & "'," & _
                          "DFData_alteracao_TBPlano_servico_servico_laboratorio = '" & Format(Date, "YYYYMMDD") & "'," & _
                          "DFIntegrado_filiais_TBPlano_servico_servico_laboratorio = 0 "
                          
                 If booIntegra_Portal = True Then
                    strSql = strSql & ",DFIntegrado_portal_TBPlano_servico_servico_laboratorio = 0 "
                 End If
                          
                 strSql = strSql & "WHERE PkId_TBPlano_servico_servico_laboratorio = " & hfgServico.Text & ""

                 conexao.CNConexao.Execute strSql
  
              ElseIf hfgServico.Text = Empty Then
              
                 strSql = Empty
                 strSql = "INSERT INTO TBPlano_servico_servico_laboratorio (FKCodigo_TBServico_laboratorio," & _
                          "FKCodigo_TBPlano_servico,DFQuantidade_TBPlano_servico_servico_laboratorio," & _
                          "DFControle_TBPlano_servico_servico_laboratorio," & _
                          "DFPeriodo_TBPlano_servico_servico_laboratorio," & _
                          "DFData_alteracao_TBPlano_servico_servico_laboratorio," & _
                          "DFIntegrado_filiais_TBPlano_servico_servico_laboratorio"
                          
                 If booIntegra_Portal = True Then
                    strSql = strSql & ",DFIntegrado_portal_TBPlano_servico_servico_laboratorio) "
                 Else
                    strSql = strSql & ") "
                 End If
                          
                 strSql = strSql & "SELECT '" & strCodigo & "','" & txtCodigo.Text & "'," & _
                                   "'" & strQuantidade & "','" & intControle & "','" & intPeriodo & "'," & _
                                   "'" & Format(Date, "YYYYMMDD") & "',0"
                 
                 If booIntegra_Portal = True Then
                    strSql = strSql & ",0"
                 End If
                                              
                 conexao.CNConexao.Execute strSql
              End If
              
              intContador = intContador + 1
           Loop
       End If
       
       'Deletando registros antes da nova gravacao
        If strId_remover <> Empty Then
           strSql = "DELETE FROM TBPlano_servico_servico_laboratorio WHERE PkId_TBPlano_servico_servico_laboratorio IN (" & strId_remover & ")"
           conexao.CNConexao.Execute strSql
           strId_remover = Empty
        End If

       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "Otica", Me
       
    Else
    
       log.Evento = "Incluir Novo"

       strSql = "INSERT INTO TBPlano_servico (" & strCampo & ") " & _
                "SELECT " & strValores & " "
       
       conexao.CNConexao.Execute strSql
       
       If hfgServico.Text <> Empty Then
          intContador = 1
          Do While intContador <= hfgServico.Rows - 1
          
              hfgServico.Row = intContador
              
              hfgServico.Col = 2
              strCodigo = hfgServico.Text
              
              hfgServico.Col = 4
              
              If hfgServico.Text = "Valor Contrato" Then
                 intControle = 1
              ElseIf hfgServico.Text = "Serviços" Then
                 intControle = 2
              ElseIf hfgServico.Text = "Grupo Serviços" Then
                 intControle = 3
              Else
                 intControle = 0
              End If
              
              hfgServico.Col = 5
              strQuantidade = hfgServico.Text
              
              hfgServico.Col = 6
              intPeriodo = hfgServico.Text
              
              strSql = "INSERT INTO TBPlano_servico_servico_laboratorio (FKCodigo_TBServico_laboratorio," & _
                       "FKCodigo_TBPlano_servico,DFQuantidade_TBPlano_servico_servico_laboratorio," & _
                       "DFControle_TBPlano_servico_servico_laboratorio," & _
                       "DFPeriodo_TBPlano_servico_servico_laboratorio," & _
                       "DFData_alteracao_TBPlano_servico_servico_laboratorio," & _
                       "DFIntegrado_filiais_TBPlano_servico_servico_laboratorio"

              If booIntegra_Portal = True Then
                 strSql = strSql & ",DFIntegrado_portal_TBPlano_servico_servico_laboratorio) "
              Else
                 strSql = strSql & ") "
              End If
              
              strSql = strSql & "SELECT '" & strCodigo & "','" & txtCodigo.Text & "'," & _
                       "'" & strQuantidade & "','" & intControle & "','" & intPeriodo & "'," & _
                       "'" & Format(Date, "YYYYMMDD") & "',0"
                       
              If booIntegra_Portal = True Then
                 strSql = strSql & ",0"
              End If
                                   
              conexao.CNConexao.Execute strSql
              
              intContador = intContador + 1
           Loop
       End If
       log.Descricao = "Incluindo o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "Otica", Me
       
       'ATUALIZAÇÃO DA TABELA TBParametros_Servicos
        
       'Somente para mostrar ao usuario o código que o plano foi incluido
       strCodigo_Plano_Servico = strProx_Cod_Plano
       
       If strCodigo_Plano_Servico <> Empty Then
          MsgBox "** O código desse Plano é: " & strCodigo_Plano_Servico & "", vbOKOnly, "Only Tech"
       End If
         
       strProx_Cod_Plano = strProx_Cod_Plano + 1
       
       strSet = "SET DFProximo_plano_servico_TBParametros_servicos = " & strProx_Cod_Plano & ""
       
       Call funcoes_banco.Alterar("TBParametros_Servicos", strSet, "FKCodigo_TBEmpresa", MDIPrincipal.OCXUsuario.Empresa, "Otica", Me, "BDRetaguarda")
       
    End If
    
    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
       
    Call Objetos.Limpa_TXT(Me)

    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       Me.hfgPlano.Visible = False
    End If
    
    sstPlano.TabEnabled(0) = False
    sstPlano.Tab = 1
        
    Exit Function
    
Erro_transacao:
    'cancelando as alteracoes
    conexao.CNConexao.RollbackTrans
    'fechando conexao
    conexao.Fechar_conexao
Erro:
    Call Erro.Erro(Me, "Otica", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    
    strSql = "SELECT PKCodigo_TBContrato_cliente FROM TBContrato_cliente " & _
             "WHERE FKCodigo_TBPlano_servico = " & txtCodigo.Text & ""
    
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       MsgBox "Este Plano de Serviço está vinculado ao Contrato de Código " & rstAplicacao.Fields("PKCodigo_TBContrato_cliente") & " e não pode ser excluído. Verifique.", vbInformation, "Only Tech"
       Set rstAplicacao = Nothing
       Exit Function
    End If
    Set rstAplicacao = Nothing
    
    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "Otica", Me
    
    'abrindo conexao
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    'Excluindo Registro
    strSql = "DELETE FROM TBPlano_servico_servico_laboratorio WHERE FKCodigo_TBPlano_servico = '" & txtCodigo.Text & "'"
    
    conexao.CNConexao.Execute strSql
    
    'Excluindo Registro Principal
    strSql = "DELETE FROM TBPlano_servico WHERE PKCodigo_TBPlano_servico = '" & txtCodigo.Text & "'"
    
    conexao.CNConexao.Execute strSql
    
    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
       
    Call Objetos.Limpa_TXT(Me)
    cbbControle.Text = Empty
    cbbPeriodo.Text = Empty
    
    'Novo
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = False
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = False
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = False
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       Me.hfgPlano.Visible = False
    End If
    
    sstPlano.TabEnabled(0) = False
    sstPlano.Tab = 1
        
    Exit Function
Erro:
    conexao.CNConexao.RollbackTrans
    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
    
    Call Erro.Erro(Me, "Otica", "Excluir")
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)

    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = booPrivilegio_Excluir
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgPlano.Visible = False
    End If
    
    sstPlano.TabEnabled(0) = False
    sstPlano.Tab = 1
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    cbbControle.Text = Empty
    cbbPeriodo.Text = Empty
    
    sstPlano.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Cancelar")
    Exit Function
End Function

Private Function Novo()

    Dim rstBusca_Parametro As New ADODB.Recordset
    Dim strCodigo_Plano_Servico As String
    
    On Error GoTo Erro
    Call Objetos.Limpa_TXT(Me)

    'verificacao de código
    strSql = Empty
    strSql = "SELECT * FROM TBParametros_Servicos " & _
             "WHERE TBParametros_Servicos.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ""
             
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Parametro, "Otica", Me)
        
    strCodigo_Plano_Servico = rstBusca_Parametro.Fields("DFProximo_plano_servico_TBParametros_servicos")
    Set rstBusca_Parametro = Nothing
        
    strSql = Empty
    strSql = "SELECT * FROM TBPlano_servico WHERE TBPlano_servico.PKCodigo_TBPlano_servico = " & strCodigo_Plano_Servico & ""
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Parametro, "Otica", Me)
    
    If rstBusca_Parametro.RecordCount <> 0 Then
       MsgBox "O Código " & strCodigo_Plano_Servico & " já existe, por favor, verifique o cadastro Parâmetros de Serviços e atualize o código do próximo Plano.", vbInformation, "Only Tech"
       Set rstBusca_Parametro = Nothing
       Call Objetos.Limpa_TXT(Me)
       sstPlano.TabEnabled(1) = True
       sstPlano.Tab = 1
       Exit Function
    End If
    Set rstBusca_Parametro = Nothing
 
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
        
    'Gravando Log
    log.Gravar_log "Otica", Me
            
    tlbBotoes.Buttons.Item(1).Enabled = False
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = booPrivilegio_Incluir
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = booPrivilegio_Incluir
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = False
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = False
               
    sstPlano.TabEnabled(0) = True
    sstPlano.Tab = 0
    
    hfgServico.Rows = 2
    Call Movimentacoes.Monta_HFlex_Grid(hfgServico, strTamanho_servico, strNomes_servico, 6, "Otica", Me)
    
    booAlterar = False
    
    cbbControle.Text = Empty
    cbbPeriodo.Text = Empty
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    On Error GoTo Erro
    
    strTamanho = "1250,7000"
    strNomes = "Código Plano,Descrição Plano"
    
    Movimentacoes.Monta_HFlex_Grid hfgPlano, strTamanho, strNomes, 2, "Otica", Me

    strTamanho_servico = "0,900,3500,1400,1200,1100"
    strNomes_servico = "ID,Código,Descrição,Controle,Limite,Período"
    
    Movimentacoes.Monta_HFlex_Grid hfgServico, strTamanho_servico, strNomes_servico, 6, "Otica", Me
    
    strSql = "SELECT PKCodigo_TBServico_laboratorio,DFDescricao_TBServico_laboratorio FROM TBServico_laboratorio"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBServico_laboratorio", "DFDescricao_TBServico_laboratorio", dtcServico, strSql, "BDRetaguarda", "Otica", Me

    Call Monta_Combo
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Reposicao")
    Resume Next
End Function

Private Sub txtCodigo_Change()
    dtcServico.BoundText = txtServico.Text
    If IsNumeric(txtServico.Text) = False Then txtServico.Text = Empty: Exit Sub
End Sub

Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> Empty And booAlterar = False Then
       Movimentacoes.Verifica_Numero "PKCodigo_TBPlano_servico", "TBPlano_servico", txtCodigo, "OTICA", Me
    End If
End Sub

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDescricao_LostFocus()
    txtDescricao.Text = UCase(txtDescricao.Text)
End Sub

Private Sub txtLimite_LostFocus()
    If Left(txtLimite.Text, 1) = "0" Then
       txtLimite.Text = Right(txtLimite.Text, 1)
    End If
End Sub

Private Sub txtServico_Change()
    dtcServico.BoundText = txtServico.Text
    If IsNumeric(txtServico.Text) = False Then txtServico.Text = Empty: Exit Sub
End Sub

Private Sub txtServico_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtLimite_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Public Function Consulta()
    Dim intControle As Integer
    Dim intContador As Integer
    
    If cbbCampos.Text = Empty Then
       MsgBox "Selecione um campo para consulta.", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    ElseIf cbbCampos.Text <> "Todos" And txtConsulta.Text = Empty Then
       MsgBox "Selecione um parâmetro para consulta.", vbInformation, "Only Tech"
       txtConsulta.SetFocus
       Exit Function
    End If
    
    strSql = "SELECT TBPlano_servico.PKCodigo_TBPlano_servico," & _
             "TBPlano_servico.DFDescricao_TBPlano_servico " & _
             "FROM TBPlano_servico "
             
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    If cbbCampos.Text <> "Todos" Then
        If cbbCampos.Text = "Código" Then
           strSql = strSql & " WHERE convert(nvarchar,TBPlano_servico.PKCodigo_TBPlano_servico) = '" & txtConsulta.Text & "'"
        ElseIf cbbCampos.Text = "Descrição" Then
            strSql = strSql & " WHERE TBPlano_servico.DFDescricao_TBPlano_servico LIKE '%" & txtConsulta.Text & "%'"
        End If
    End If
    
    frmAguarde.Show
    DoEvents
    
    strSql = strSql & " ORDER BY TBPlano_servico.PKCodigo_TBPlano_servico"
         
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgPlano, strTamanho, strNomes, "BDRetaguarda", "Otica", Me

    hfgPlano.Row = 1
    hfgPlano.Col = 0
    If hfgPlano.Text = Empty Then
       hfgPlano.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgPlano, strTamanho, strNomes, 2, "Otica", Me
    End If
    
    Unload frmAguarde
    hfgPlano.Row = 1
    hfgPlano.Col = 0
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código")
    cbbCampos.AddItem ("Descrição")

    cbbControle.Clear
    cbbControle.AddItem ("Valor Contrato")
    cbbControle.AddItem ("Serviços")
    cbbControle.AddItem ("Grupo Serviços")
    
    cbbPeriodo.Clear
    cbbPeriodo.AddItem ("1")
    cbbPeriodo.AddItem ("2")
    cbbPeriodo.AddItem ("3")
    cbbPeriodo.AddItem ("4")
    cbbPeriodo.AddItem ("5")
    cbbPeriodo.AddItem ("6")
    cbbPeriodo.AddItem ("7")
    cbbPeriodo.AddItem ("8")
    cbbPeriodo.AddItem ("9")
    cbbPeriodo.AddItem ("10")
    cbbPeriodo.AddItem ("11")
    cbbPeriodo.AddItem ("12")
    
End Function

Private Sub txtCodigo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDescricao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtServico_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtLimite_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo_TBPlano_servico", txtCodigo.Text, "DFIntegrado_filiais_TBPlano_servico", "TBPlano_servico", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBPlano_servico", Me.Top, Me.Left, Me.Width, Me.Height, "Plano de Serviços")
    
End Function
