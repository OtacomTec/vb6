VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmPrioridade_Pendencia 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prioridade Pendência"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrioridade_Pendencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5205
   Begin TabDlg.SSTab sstPrioridade_Pendencia 
      Height          =   2685
      Left            =   0
      TabIndex        =   3
      Top             =   330
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   4736
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
      TabPicture(0)   =   "frmPrioridade_Pendencia.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPrioridade"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmPrioridade_Pendencia.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "cbbCampos"
      Tab(1).Control(2)=   "hfgPrioridade_Pendencia"
      Tab(1).Control(3)=   "cmdConsulta"
      Tab(1).Control(4)=   "cmdRefresh"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtConsulta"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtPrioridade 
         Height          =   375
         Left            =   4140
         MaxLength       =   9
         TabIndex        =   2
         ToolTipText     =   "Prioridade da Pendência"
         Top             =   1440
         Width           =   915
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -73170
         TabIndex        =   6
         Top             =   780
         Width           =   2415
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   -70290
         Picture         =   "frmPrioridade_Pendencia.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   -70680
         Picture         =   "frmPrioridade_Pendencia.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         MaxLength       =   9
         TabIndex        =   0
         ToolTipText     =   "Código da Prioridade Pendência"
         Top             =   780
         Width           =   1110
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         ToolTipText     =   "Descrição da Prioridade Pendência"
         Top             =   1440
         Width           =   3975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgPrioridade_Pendencia 
         Height          =   1365
         Left            =   -74880
         TabIndex        =   7
         Top             =   1200
         Width           =   4965
         _ExtentX        =   8758
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
         Left            =   -74880
         TabIndex        =   8
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prioridade"
         Height          =   240
         Left            =   4140
         TabIndex        =   14
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   10
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   120
         TabIndex        =   9
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
      TabIndex        =   12
      Top             =   0
      Width           =   5205
      _ExtentX        =   9181
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
      Left            =   5280
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
            Picture         =   "frmPrioridade_Pendencia.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrioridade_Pendencia.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrioridade_Pendencia.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrioridade_Pendencia.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrioridade_Pendencia.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrioridade_Pendencia.frx":5578
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrioridade_Pendencia.frx":5892
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   240
      Left            =   0
      TabIndex        =   13
      Top             =   4890
      Width           =   825
   End
End
Attribute VB_Name = "frmPrioridade_Pendencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro Prioridade Pendência                                  '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rodrigo Santos                                                 '
' Data de Criação........: 17/01/2006                                                     '
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
    
    Call frmConsole_Prioridade_Pendencia.Show
    
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
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    Else
       txtConsulta.Visible = True
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
    log.Programa = "Cadastro de Prioridade Pendência"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstPrioridade_Pendencia, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o cadastro de Prioridade Pendência"
    
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    sstPrioridade_Pendencia.Tab = 1
    sstPrioridade_Pendencia.TabEnabled(0) = False
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
    log.Descricao = "Finalizando o cadastro de Prioridade Pendência"
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

Private Sub hfgPrioridade_Pendencia_click()

   If hfgPrioridade_Pendencia.Col = 0 And hfgPrioridade_Pendencia.Text <> Empty Then
     
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
       
      txtCodigo.Text = hfgPrioridade_Pendencia.TextArray((hfgPrioridade_Pendencia.Row * hfgPrioridade_Pendencia.Cols + hfgPrioridade_Pendencia.Col + 1))
      txtDescricao.Text = hfgPrioridade_Pendencia.TextArray((hfgPrioridade_Pendencia.Row * hfgPrioridade_Pendencia.Cols + hfgPrioridade_Pendencia.Col + 2))
      txtPrioridade.Text = hfgPrioridade_Pendencia.TextArray((hfgPrioridade_Pendencia.Row * hfgPrioridade_Pendencia.Cols + hfgPrioridade_Pendencia.Col + 3))
       
      booAlterar = True
      txtConsulta.Text = Empty
      sstPrioridade_Pendencia.TabEnabled(0) = True
      sstPrioridade_Pendencia.Tab = 0
      txtCodigo.Enabled = False
                
   End If
   
   Unload frmAguarde
   
End Sub

Private Sub hfgPrioridade_Pendencia_DblClick()
    hfgPrioridade_Pendencia.Sort = 1
End Sub

Private Sub hfgPrioridade_Pendencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgPrioridade_Pendencia_click
    End If
End Sub

Private Sub sstPrioridade_Pendencia_Click(PreviousTab As Integer)
    If sstPrioridade_Pendencia.Tab = 0 Then
       txtDescricao.SetFocus
    ElseIf sstPrioridade_Pendencia.Tab = 1 Then
       If frmIntegracao.Visible = True Then
          Unload frmIntegracao
       End If
       If strCombo <> Empty And strCombo <> "Todos" Then
           cbbCampos.Text = strCombo
           txtConsulta.SetFocus
       ElseIf strCombo = "Todos" Then
           hfgPrioridade_Pendencia.Row = 1
           hfgPrioridade_Pendencia.Col = 0
           hfgPrioridade_Pendencia.SetFocus
       End If
    End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstPrioridade_Pendencia.Tab <> 1: Call Gravar
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
    
    If txtCodigo.Text <> Empty And booAlterar = False Then
        Movimentacoes.Verifica_Numero "PKCodigo__TBPrioridade_pendencia_servico", "TBPrioridade_Pendencia_servico", txtCodigo, "Otica", Me
    End If

    Call Objetos.Maiusculo_TXT(Me)
    
    strCampo = "PKCodigo__TBPrioridade_pendencia_servico,DFDescricao__TBPrioridade_pendencia_servico," & _
               "DFScala_TBPrioridade_pendencia_servico,DFData_alteracao_TBPrioridade_Pendencia_servico," & _
               "DFIntegrado_filiais_TBPrioridade_Pendencia_servico"
               
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBPrioridade_Pendencia_servico"
    End If
                           
    strValores = "'" & txtCodigo.Text & "','" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                 "'" & txtPrioridade.Text & "','" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
        strValores = strValores & ",0"
    End If
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       
       strSet = "SET DFDescricao__TBPrioridade_pendencia_servico = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "', " & _
                "DFScala_TBPrioridade_pendencia_servico = '" & txtPrioridade.Text & "', " & _
                "DFData_alteracao_TBPrioridade_Pendencia_servico = '" & Format(Date, "YYYYMMDD") & "', " & _
                "DFIntegrado_filiais_TBPrioridade_Pendencia_servico = 0"
                
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBPrioridade_Pendencia_servico = 0"
       End If
                          
       Call funcoes_banco.Alterar("TBPrioridade_Pendencia_servico", strSet, "PKCodigo__TBPrioridade_pendencia_servico", txtCodigo.Text, "Otica", Me, "BDRetaguarda")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       
       Call funcoes_banco.Gravar("TBPrioridade_Pendencia_servico", strCampo, strValores, "Otica", Me, "BDRetaguarda")
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    End If
    
    Call Objetos.Limpa_TXT(Me)
           
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       Me.hfgPrioridade_Pendencia.Visible = False
    End If
    
    sstPrioridade_Pendencia.TabEnabled(0) = False
    sstPrioridade_Pendencia.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    'VERIFICACAO PARA EXCLUSAO
    strSql = "SELECT PKID_TBPendencia_servico,FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico " & _
             "FROM TBPendencia_Servicos where FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico = '" & txtCodigo.Text & "'"
    
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
    If rstAplicacao.RecordCount <> 0 Then
       MsgBox "Esta Prioridade está vinculado a tabela Pendência no registro " & rstAplicacao.Fields("PKID_TBPendencia_servico") & " e não poderá ser excluída. Verifique.", vbInformation, "Only Tech"
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
    Call funcoes_banco.Excluir("TBPrioridade_Pendencia_servico", "PKCodigo__TBPrioridade_pendencia_servico", txtCodigo.Text, "Otica", Me, "BDRetaguarda")
        
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
       hfgPrioridade_Pendencia.Visible = False
    End If
           
    sstPrioridade_Pendencia.TabEnabled(0) = False
    sstPrioridade_Pendencia.Tab = 1
    
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
       hfgPrioridade_Pendencia.Visible = False
    End If
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstPrioridade_Pendencia.TabEnabled(0) = False
    sstPrioridade_Pendencia.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
        
    sstPrioridade_Pendencia.TabEnabled(0) = True
    sstPrioridade_Pendencia.Tab = 0
               
    Call Objetos.Limpa_TXT(Me)
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
                   
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
             
    booAlterar = False
    
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    
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
        Movimentacoes.Verifica_Numero "PKCodigo__TBPrioridade_pendencia_servico", "TBPrioridade_Pendencia_servico", txtCodigo, "Otica", Me
    End If
End Sub

Private Function Reposicao()
    On Error GoTo Erro
    
    Movimentacoes.Monta_HFlex_Grid hfgPrioridade_Pendencia, "1000,3500,1000", "Código,Descrição,Prioridade", 3, "Otica", Me
        
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
    cmdConsulta.SetFocus
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
          MsgBox "Selecione um campo e digite os dados para consulta.", vbInformation, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
    
    strSql = "SELECT TBPrioridade_Pendencia_servico.PKCodigo__TBPrioridade_pendencia_servico As " & _
             "Codigo,TBPrioridade_Pendencia_servico.DFDescricao__TBPrioridade_pendencia_servico As Descricao, " & _
             "DFScala_TBPrioridade_pendencia_servico as Prioridade " & _
             "FROM TBPrioridade_Pendencia_servico "
            
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    Funcoes_Gerais.Grava_String (txtConsulta.Text)
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código" Then
          strSql = strSql & " WHERE convert(nvarchar,TBPrioridade_Pendencia_servico.PKCodigo__TBPrioridade_pendencia_servico) LIKE '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Descrição" Then
          strSql = strSql & " WHERE TBPrioridade_Pendencia_servico.DFDescricao__TBPrioridade_pendencia_servico LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Prioridade" Then
          strSql = strSql & " WHERE TBPrioridade_Pendencia_servico.DFScala_TBPrioridade_pendencia_servico = '" & txtConsulta.Text & "' "
       End If
    End If
    
    frmAguarde.Show
    DoEvents
            
    strSql = strSql & " ORDER BY TBPrioridade_Pendencia_servico.PKCodigo__TBPrioridade_pendencia_servico"
       
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgPrioridade_Pendencia, "1000,3500,1000", "Código,Descrição,Prioridade", "BDRetaguarda", "Otica", Me
    
    hfgPrioridade_Pendencia.Col = 0
    hfgPrioridade_Pendencia.Row = 1
    If hfgPrioridade_Pendencia.Text = Empty Then
       hfgPrioridade_Pendencia.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgPrioridade_Pendencia, "1000,3500,1000", "Código,Descrição,Prioridade", 3, "Otica", Me
    End If
    
    Unload frmAguarde
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código")
    cbbCampos.AddItem ("Descrição")
    cbbCampos.AddItem ("Prioridade")
End Function

Private Sub txtPrioridade_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPrioridade_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo__TBPrioridade_pendencia_servico", txtCodigo.Text, "DFIntegrado_filiais_TBPrioridade_Pendencia_servico", "TBPrioridade_Pendencia_servico", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBPrioridade_Pendencia_servico", Me.Top, Me.Left, Me.Width, Me.Height, "Prioridade Serviço")
    
End Function
