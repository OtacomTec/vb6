VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEquipamento_Laboratorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipamento Laboratório"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEquipamento_Laboratorio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6045
   Begin TabDlg.SSTab sstEquipamento 
      Height          =   2685
      Left            =   0
      TabIndex        =   8
      Top             =   330
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   4736
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmEquipamento_Laboratorio.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(2)=   "Label16"
      Tab(0).Control(3)=   "dtpAfericao"
      Tab(0).Control(4)=   "txtDescricao"
      Tab(0).Control(5)=   "txtCodigo"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmEquipamento_Laboratorio.frx":179E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblA"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dtpFinal"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "dtpInicial"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cbbCampos"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "hfgEquipamento"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdConsulta"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdRefresh"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtConsulta"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   1830
         TabIndex        =   6
         Top             =   780
         Width           =   3255
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   5550
         Picture         =   "frmEquipamento_Laboratorio.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   5160
         Picture         =   "frmEquipamento_Laboratorio.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   360
         Left            =   -74880
         TabIndex        =   0
         ToolTipText     =   "Código do Tipo Sinistro"
         Top             =   780
         Width           =   1110
      End
      Begin VB.TextBox txtDescricao 
         Height          =   360
         Left            =   -74880
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Descrição do Tipo Sinistro"
         Top             =   1440
         Width           =   4185
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgEquipamento 
         Height          =   1365
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   2408
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   1665
         _ExtentX        =   2937
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
      Begin MSComCtl2.DTPicker dtpAfericao 
         Height          =   360
         Left            =   -70650
         TabIndex        =   2
         ToolTipText     =   "Ultima Aferição"
         Top             =   1440
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20381697
         CurrentDate     =   37858
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   360
         Left            =   1830
         TabIndex        =   4
         ToolTipText     =   "Ultima Aferição"
         Top             =   780
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20381697
         CurrentDate     =   37858
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   360
         Left            =   3690
         TabIndex        =   5
         ToolTipText     =   "Ultima Aferição"
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20381697
         CurrentDate     =   37858
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   240
         Left            =   3420
         TabIndex        =   16
         Top             =   900
         Width           =   105
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Última Aferição"
         Height          =   240
         Left            =   -70650
         TabIndex        =   15
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   -74880
         TabIndex        =   13
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   -74880
         TabIndex        =   11
         Top             =   1200
         Width           =   825
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6045
      _ExtentX        =   10663
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
      Left            =   6270
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
            Picture         =   "frmEquipamento_Laboratorio.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipamento_Laboratorio.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipamento_Laboratorio.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipamento_Laboratorio.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipamento_Laboratorio.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipamento_Laboratorio.frx":5578
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipamento_Laboratorio.frx":5892
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEquipamento_Laboratorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro de Equipamentos                                       '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Peixoto                                                  '
' Data de Criação........: 24/12/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Public strSql As String
Dim strId As String
Dim booAlterar As Boolean
Dim conexao As New DLLConexao_Sistema.conexao
Dim log As New DLLSystemManager.log
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean

Function Imprimir()
    On Error GoTo Erro
    'Tratamento de Erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique.", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Equipamento_Laboratorio.Show
    
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
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       lblA.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       dtpInicial.Visible = True
       dtpFinal.Visible = True
       lblA.Visible = True
       dtpInicial.SetFocus
    Else
       txtConsulta.Visible = True
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       lblA.Visible = False
       txtConsulta.SetFocus
    End If
    
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
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
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub
  
Private Sub Form_Load()
    On Error GoTo Erro
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Equipamento"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstEquipamento, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o cadastro de Equipamento"
    
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    dtpAfericao.Value = Date
    dtpInicial.Value = Date - 7
    dtpFinal.Value = Date
    
    sstEquipamento.Tab = 1
    sstEquipamento.TabEnabled(0) = False
    
    Call Reposicao
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
        
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando o cadastro de Equipamento"
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Set log = Nothing
    
    strCombo = Empty
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgEquipamento_Click()

   If hfgEquipamento.Col = 0 And hfgEquipamento.Text <> Empty Then
     
        On Error Resume Next
        
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
       
       txtCodigo.Text = hfgEquipamento.TextArray((hfgEquipamento.Row * hfgEquipamento.Cols + hfgEquipamento.Col + 1))
       txtDescricao.Text = hfgEquipamento.TextArray((hfgEquipamento.Row * hfgEquipamento.Cols + hfgEquipamento.Col + 2))
       If hfgEquipamento.TextArray((hfgEquipamento.Row * hfgEquipamento.Cols + hfgEquipamento.Col + 3)) <> Empty And _
          hfgEquipamento.TextArray((hfgEquipamento.Row * hfgEquipamento.Cols + hfgEquipamento.Col + 3)) <> " " Then
          dtpAfericao.Value = hfgEquipamento.TextArray((hfgEquipamento.Row * hfgEquipamento.Cols + hfgEquipamento.Col + 3))
       Else
          dtpAfericao.Value = "1/1/1900"
       End If

       booAlterar = True
       txtConsulta.Text = Empty
       sstEquipamento.TabEnabled(0) = True
       sstEquipamento.Tab = 0
       
       txtCodigo.Enabled = False
                
   End If
   
   Unload frmAguarde
   
End Sub

Private Sub hfgEquipamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgEquipamento_Click
    End If
End Sub

Private Sub sstEquipamento_Click(PreviousTab As Integer)
   If sstEquipamento.Tab = 0 Then
      txtDescricao.SetFocus
   ElseIf sstEquipamento.Tab = 1 Then
      If frmIntegracao.Visible = True Then
         Unload frmIntegracao
      End If
      If strCombo <> Empty And strCombo <> "Todos" And strCombo <> "Última Aferição" Then
          cbbCampos.Text = strCombo
          txtConsulta.SetFocus
      ElseIf strCombo = "Última Aferição" Then
         cbbCampos.Text = strCombo
         dtpInicial.SetFocus
      ElseIf strCombo = "Todos" Then
         hfgEquipamento.Row = 1
         hfgEquipamento.Col = 0
         hfgEquipamento.SetFocus
      End If
   End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstEquipamento.Tab <> 1: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
    End Select
End Sub

Function Gravar()
    On Error GoTo Erro
    
    'Verifica se os campos necessarios para gravar não estão nulos
    If txtDescricao.Text = Empty Then
       MsgBox "O campo descrição não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtDescricao.SetFocus
       Exit Function
    End If
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim intEmpresa As Integer
    Dim strCodigo_Equipamento As String
    Dim strProx_Cod_Equipamento As String
    Dim rstVerifica_Titulo As New ADODB.Recordset
    
    intEmpresa = MDIPrincipal.OCXUsuario.Empresa
    
    Call Objetos.Maiusculo_TXT(Me)
    
    If booAlterar = False Then
       strProx_Cod_Equipamento = Funcoes_Gerais.Localiza_Proximo_Codigo("DFProximo_equipamento_laboratorio", "FKCodigo_TBEmpresa", intEmpresa, "TBParametros_servicos", "Otica", Me, "BDRetaguarda")
       txtCodigo.Text = strProx_Cod_Equipamento
    End If
    
    strCampo = "PKCodigo_TBEquipamento_laboratorio,DFDescricao_TBEquipamento_laboratorio," & _
               "DFData_ultima_afericao_TBEquipamento_laboratorio," & _
               "DFData_alteracao_TBEquipamento_laboratorio," & _
               "DFIntegrado_filiais_TBEquipamento_laboratorio "
               
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBEquipamento_laboratorio "
    End If
    
    strValores = "" & txtCodigo.Text & ",'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                 "'" & Format(dtpAfericao.Value, "YYYYMMDD") & "'," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
        strValores = strValores & ",0"
    End If
                              
    If booAlterar = True Then
       log.Evento = "Alterar"
       
       strSet = "SET DFDescricao_TBEquipamento_laboratorio = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                "DFData_ultima_afericao_TBEquipamento_laboratorio = '" & Format(dtpAfericao.Value, "YYYYMMDD") & "'," & _
                "DFData_alteracao_TBEquipamento_laboratorio = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBEquipamento_laboratorio = 0"
                
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBEquipamento_laboratorio = 0"
       End If
       
       Call funcoes_banco.Alterar("TBEquipamento_laboratorio", strSet, "PKCodigo_TBEquipamento_laboratorio", txtCodigo.Text, "Otica", Me, "BDRetaguarda")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       
       Call funcoes_banco.Gravar("TBEquipamento_laboratorio", strCampo, strValores, "Otica", Me, "BDRetaguarda")
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
       
       ''''' aqui começa a ATUALIZAÇÃO DA TABELA TBParametros_servicos '''''
        
       'Somente para mostrar ao usuario o código que o cliente foi incluido
       strCodigo_Equipamento = strProx_Cod_Equipamento
       
    If strCodigo_Equipamento <> Empty Then
       MsgBox "** O código desse Equipamento é: " & strCodigo_Equipamento & "", vbOKOnly, "Only Tech"
    End If
    
       strProx_Cod_Equipamento = strProx_Cod_Equipamento + 1
       strSet = "SET DFProximo_equipamento_laboratorio = " & strProx_Cod_Equipamento & ""
       Call funcoes_banco.Alterar("TBParametros_servicos", strSet, "FKCodigo_TBEmpresa", MDIPrincipal.OCXUsuario.Empresa, "Otica", Me, "BDRetaguarda")
    End If
    
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
       Me.hfgEquipamento.Visible = False
    End If
    
    sstEquipamento.TabEnabled(0) = False
    sstEquipamento.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    'VERIFICACAO PARA EXCLUSAO
    strSql = "SELECT FKCodigo_TBEquipamento_laboratorio,FKCodigo_TBTipo_marcha FROM TBEquipamento_tipo_marcha " & _
             "WHERE FKCodigo_TBEquipamento_laboratorio = " & txtCodigo.Text & ""
    
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
    If rstAplicacao.RecordCount <> 0 Then
       MsgBox "Este Equipamento está vinculada ao Tipo de Marcha de código " & rstAplicacao.Fields("FKCodigo_TBTipo_marcha") & " e não poderá ser excluído. Verifique.", vbInformation, "Only Tech"
       Set rstAplicacao = Nothing
       Exit Function
    End If
    Set rstAplicacao = Nothing
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBEquipamento_laboratorio", "PKCodigo_TBEquipamento_laboratorio", txtCodigo.Text, "Otica", Me, "BDRetaguarda")
        
    Call Objetos.Limpa_TXT(Me)
    
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
       hfgEquipamento.Visible = False
    End If
           
    sstEquipamento.TabEnabled(0) = False
    sstEquipamento.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Excluir")
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
       
    'Novo
     tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = False
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = False
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
       hfgEquipamento.Visible = False
    End If
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstEquipamento.TabEnabled(0) = False
    sstEquipamento.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
    Dim rstBusca_Parametro As New ADODB.Recordset
    Dim strCodigo_Equipamento As String
               
    Call Objetos.Limpa_TXT(Me)
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
        
    strSql = Empty
    strSql = "SELECT * FROM TBParametros_servicos " & _
             "WHERE TBParametros_servicos.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ""
             
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Parametro, "Otica", Me)
        
    strCodigo_Equipamento = rstBusca_Parametro.Fields("DFProximo_equipamento_laboratorio")
    Set rstBusca_Parametro = Nothing
        
    strSql = Empty
    strSql = "SELECT * FROM TBEquipamento_laboratorio WHERE TBEquipamento_laboratorio.PKCodigo_TBEquipamento_laboratorio = " & strCodigo_Equipamento & ""
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Parametro, "Otica", Me)
    
    If rstBusca_Parametro.RecordCount <> 0 Then
       MsgBox "O Código " & strCodigo_Equipamento & " já existe, por favor, verifique o cadastro parâmetros de serviços e atualize o código da próxima Equipamento.", vbInformation, "Only Tech"
       Set rstBusca_Parametro = Nothing
       Call Objetos.Limpa_TXT(Me)
       sstEquipamento.TabEnabled(0) = False
       sstEquipamento.Tab = 1
       Exit Function
    End If
    Set rstBusca_Parametro = Nothing
    
    dtpAfericao.Value = Date
    
    sstEquipamento.TabEnabled(0) = True
    sstEquipamento.Tab = 0
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    'Novo
    tlbBotoes.Buttons.Item(1).Enabled = False
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = booPrivilegio_Incluir
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = booPrivilegio_Incluir
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = False
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = False

    txtCodigo.Enabled = False
    txtDescricao.SetFocus
       
    booAlterar = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Sub txtCodigo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> Empty And booAlterar = False Then
        Movimentacoes.Verifica_Numero "PKCodigo_TBEquipamento_laboratorio", "TBEquipamento_laboratorio", txtCodigo, "Otica", Me
    End If
End Sub

Private Function Reposicao()
    On Error GoTo Erro
    
    Movimentacoes.Monta_HFlex_Grid hfgEquipamento, "1000,3500,1500", "Código,Descrição,Ult. Aferição", 3, "Otica", Me
        
   ' Call Monta_DataCombo
    Call Monta_Combo
    
    strSql = Empty
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Reposicao")
    Resume Next
End Function

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Sub txtDescricao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDescricao_LostFocus()
    txtDescricao.Text = UCase(txtDescricao.Text)
End Sub

Private Function Consulta()
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
    
    strSql = "SELECT PKCodigo_TBEquipamento_laboratorio,DFDescricao_TBEquipamento_laboratorio," & _
             "DFData_ultima_afericao_TBEquipamento_laboratorio " & _
             "FROM TBEquipamento_laboratorio "
            
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    Funcoes_Gerais.Grava_String (txtConsulta.Text)
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE TBEquipamento_laboratorio.PKCodigo_TBEquipamento_laboratorio = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Descrição" Then
          strSql = strSql & " WHERE TBEquipamento_laboratorio.DFDescricao_TBEquipamento_laboratorio LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Última Aferição" Then
          strSql = strSql & " WHERE TBEquipamento_laboratorio.DFData_ultima_afericao_TBEquipamento_laboratorio >= '" & Format(dtpInicial.Value) & "' " & _
                            " AND TBEquipamento_laboratorio.DFData_ultima_afericao_TBEquipamento_laboratorio <= '" & Format(dtpFinal.Value) & "' "
       End If
    End If
    
    frmAguarde.Show
    DoEvents
                   
    strSql = strSql & " ORDER BY TBEquipamento_laboratorio.PKCodigo_TBEquipamento_laboratorio"
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgEquipamento, "1000,3500,1500", "Código,Descrição,Ult. Aferição", "BDRetaguarda", "Otica", Me
    
    hfgEquipamento.Col = 0
    hfgEquipamento.Row = 1
    If hfgEquipamento.Text = Empty Then
       hfgEquipamento.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgEquipamento, "1000,3500,1500", "Código,Descrição,Ult. Aferição", 3, "Otica", Me
    End If
    
    Unload frmAguarde
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código")
    cbbCampos.AddItem ("Descrição")
    cbbCampos.AddItem ("Última Aferição")
End Function

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo_TBEquipamento_laboratorio", txtCodigo.Text, "DFIntegrado_filiais_TBEquipamento_laboratorio", "TBEquipamento_laboratorio", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBEquipamento_laboratorio", Me.Top, Me.Left, Me.Width, Me.Height, "Equipamento Laboratório")
    
End Function
