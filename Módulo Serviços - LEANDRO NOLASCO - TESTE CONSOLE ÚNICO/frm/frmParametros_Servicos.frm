VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmParametros_Servicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parâmetros de Serviços"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmParametros_Servicos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6180
   Begin TabDlg.SSTab sstParametros_Servicos 
      Height          =   3315
      Left            =   0
      TabIndex        =   10
      Top             =   330
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   5847
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
      TabPicture(0)   =   "frmParametros_Servicos.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(7)=   "dtcEmpresa"
      Tab(0).Control(8)=   "txtFuncao_Insumo"
      Tab(0).Control(9)=   "txtInsumo"
      Tab(0).Control(10)=   "txtServico"
      Tab(0).Control(11)=   "txtPlano_Servico"
      Tab(0).Control(12)=   "txtTipo_Marcha"
      Tab(0).Control(13)=   "txtEquipamento"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmParametros_Servicos.frx":179E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cbbCampos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgParametros_Servicos"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdParametros_Consulta_Empresa"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtConsulta"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdRefresh"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdConsulta"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtEquipamento 
         Height          =   375
         Left            =   -71910
         TabIndex        =   5
         ToolTipText     =   "Próximo Equipamento Laboratório"
         Top             =   2760
         Width           =   2925
      End
      Begin VB.TextBox txtTipo_Marcha 
         Height          =   375
         Left            =   -74880
         TabIndex        =   4
         ToolTipText     =   "Próximo Tipo de Marcha"
         Top             =   2760
         Width           =   2925
      End
      Begin VB.TextBox txtPlano_Servico 
         Height          =   375
         Left            =   -71910
         TabIndex        =   1
         ToolTipText     =   "Próximo Plano de Serviços"
         Top             =   1410
         Width           =   2925
      End
      Begin VB.TextBox txtServico 
         Height          =   375
         Left            =   -74880
         TabIndex        =   0
         ToolTipText     =   "Próximo Serviço"
         Top             =   1410
         Width           =   2925
      End
      Begin VB.TextBox txtInsumo 
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         ToolTipText     =   "Próximo Insumo"
         Top             =   2070
         Width           =   2925
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
         Left            =   5250
         Picture         =   "frmParametros_Servicos.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   5640
         Picture         =   "frmParametros_Servicos.frx":34B4
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   1920
         TabIndex        =   7
         Top             =   780
         Width           =   2865
      End
      Begin VB.TextBox txtFuncao_Insumo 
         Height          =   375
         Left            =   -71910
         TabIndex        =   3
         ToolTipText     =   "Próxima Função Insumo"
         Top             =   2070
         Width           =   2925
      End
      Begin VB.CommandButton cmdParametros_Consulta_Empresa 
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
         Left            =   4860
         Picture         =   "frmParametros_Servicos.frx":44F6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   780
         Width           =   375
      End
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   360
         Left            =   -74880
         TabIndex        =   13
         Top             =   780
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         BackColor       =   -2147483639
         ForeColor       =   8388608
         Text            =   ""
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgParametros_Servicos 
         Height          =   1935
         Left            =   120
         TabIndex        =   9
         Top             =   1230
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3413
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
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   1755
         _ExtentX        =   3096
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Prox. Equipamento Laboratório"
         Height          =   240
         Left            =   -71910
         TabIndex        =   22
         Top             =   2520
         Width           =   2640
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Próximo Tipo de Marcha"
         Height          =   240
         Left            =   -74880
         TabIndex        =   21
         Top             =   2520
         Width           =   2085
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Próximo Plano de Serviços"
         Height          =   240
         Left            =   -71910
         TabIndex        =   20
         Top             =   1170
         Width           =   2265
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa [ F2 ]"
         Height          =   375
         Left            =   -74880
         TabIndex        =   18
         Top             =   540
         Width           =   1290
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Próximo Serviço"
         Height          =   240
         Left            =   -74880
         TabIndex        =   17
         Top             =   1170
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Próximo Insumo"
         Height          =   240
         Left            =   -74880
         TabIndex        =   15
         Top             =   1830
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Próxima Função Insumo"
         Height          =   240
         Left            =   -71910
         TabIndex        =   14
         Top             =   1830
         Width           =   2055
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6180
      _ExtentX        =   10901
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
      Left            =   7860
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
            Picture         =   "frmParametros_Servicos.frx":5538
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Servicos.frx":5852
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Servicos.frx":5B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Servicos.frx":5F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Servicos.frx":62A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Servicos.frx":65BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParametros_Servicos.frx":68D4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParametros_Servicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro de Parâmetros Gerais de Serviços                      '
' Data de Criação........: 30/11/2004                                                     '
' Equipe Responsável.....: Jones Sá Peixoto,Marcos Baião,Alex Baião,Rafael Gomes          '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strTamanho As String
Dim strNomes As String
Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim strID_Parametros As String
Public strSql As String
Public strCodigo_Empresa_Consulta As String
Dim conexao As New DLLConexao_Sistema.conexao
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
Dim log As New DLLSystemManager.log

Function Imprimir()
    'Tratamento de erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Parametro_Servico.Show
    
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

Private Sub cmdParametros_Consulta_Empresa_Click()
    'STRING QUE COLETA DADOS RELATIVOS A ACESSIBILIDADE DO USUARIO
    Dim rstAcesso_Consulta_Empresa As New ADODB.Recordset

    strSql = Empty
    strSql = "SELECT  DFNivel_TBUsuario FROM TBUsuario " & _
             "WHERE DFNome_TBUsuario = '" & MDIPrincipal.OCXUsuario.Nome & "'"

    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAcesso_Consulta_Empresa, "Otica", Me

    If rstAcesso_Consulta_Empresa!DFNivel_TBUsuario < 5 Then
       Exit Sub
    End If

    Set rstAcesso_Consulta_Empresa = Nothing

    Unload frmParametro_Servico_Consulta_Empresa
    frmAguarde.Show
    DoEvents
    frmParametro_Servico_Consulta_Empresa.Show
    Unload frmAguarde
End Sub



Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub dtcEmpresa_LostFocus()
    dtcEmpresa.Enabled = False
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
    If KeyCode = "113" And booAlterar = False Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
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
    log.Programa = "Cadastro de Parâmetros Gerais de Serviços"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstParametros_Servicos, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o cadastro de Parâmetros Gerais de Serviços"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    Call Reposicao
    strCodigo_Empresa_Consulta = Empty
    
    sstParametros_Servicos.TabEnabled(0) = False
    sstParametros_Servicos.Tab = 1
    
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
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Set log = Nothing
    
    strCombo = Empty
    strCodigo_Empresa_Consulta = Empty
    
    If frmIntegracao.Visible = True Then
        Unload frmIntegracao
    End If

    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgParametros_servicos_Click()
    If hfgParametros_Servicos.Col = 0 And hfgParametros_Servicos.Text <> Empty Then
        
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
        
       strID_Parametros = hfgParametros_Servicos.TextArray((hfgParametros_Servicos.Row * hfgParametros_Servicos.Cols + hfgParametros_Servicos.Col + 1))
       txtServico.Text = hfgParametros_Servicos.TextArray((hfgParametros_Servicos.Row * hfgParametros_Servicos.Cols + hfgParametros_Servicos.Col + 2))
       txtPlano_Servico.Text = hfgParametros_Servicos.TextArray((hfgParametros_Servicos.Row * hfgParametros_Servicos.Cols + hfgParametros_Servicos.Col + 3))
       txtInsumo.Text = hfgParametros_Servicos.TextArray((hfgParametros_Servicos.Row * hfgParametros_Servicos.Cols + hfgParametros_Servicos.Col + 4))
       txtFuncao_Insumo.Text = hfgParametros_Servicos.TextArray((hfgParametros_Servicos.Row * hfgParametros_Servicos.Cols + hfgParametros_Servicos.Col + 5))
       txtTipo_Marcha.Text = hfgParametros_Servicos.TextArray((hfgParametros_Servicos.Row * hfgParametros_Servicos.Cols + hfgParametros_Servicos.Col + 6))
       txtEquipamento.Text = hfgParametros_Servicos.TextArray((hfgParametros_Servicos.Row * hfgParametros_Servicos.Cols + hfgParametros_Servicos.Col + 7))
       dtcEmpresa.BoundText = hfgParametros_Servicos.TextArray((hfgParametros_Servicos.Row * hfgParametros_Servicos.Cols + hfgParametros_Servicos.Col + 8))
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstParametros_Servicos.TabEnabled(0) = True
       sstParametros_Servicos.Tab = 0
       Me.txtServico.SetFocus
   End If
   
   Unload frmAguarde
   
End Sub

Private Sub hfgParametros_Servicos_DblClick()
    hfgParametros_Servicos.Sort = 1
End Sub

Private Sub hfgParametros_servicos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgParametros_servicos_Click
    End If
End Sub


Private Sub sstParametros_Servicos_Click(PreviousTab As Integer)
    If sstParametros_Servicos.Tab = 0 Then
       Me.txtServico.SetFocus
    ElseIf sstParametros_Servicos.Tab = 1 Then
        If frmIntegracao.Visible = True Then
            Unload frmIntegracao
        End If
        If strCombo <> Empty And strCombo <> "Todos" Then
           cbbCampos.Text = strCombo
           txtConsulta.SetFocus
        ElseIf strCombo = "Todos" Then
           hfgParametros_Servicos.Row = 1
           hfgParametros_Servicos.Col = 0
           hfgParametros_Servicos.SetFocus
        End If
    End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
    End Select
End Sub

Function Gravar()
    On Error GoTo Erro
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    
    strCampo = "DFProximo_tipo_marcha_TBParametros_servicos,DFProximo_insumo_TBParametros_servicos," & _
               "DFProximo_funcao_insumo_TBParametros_servicos," & _
               "DFProximo_servico_TBParametros_servicos,DFProximo_plano_servico_TBParametros_servicos," & _
               "FKCodigo_TBEmpresa,DFProximo_equipamento_laboratorio,DFData_alteracao_TBParametros_servicos," & _
               "DFIntegrado_filiais_TBParametros_servicos"
    
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBParametros_servicos"
    End If
               
    strValores = "'" & txtTipo_Marcha.Text & "','" & txtInsumo.Text & "'," & _
                 "'" & txtFuncao_Insumo.Text & "'," & _
                 "'" & txtServico.Text & "','" & txtPlano_Servico.Text & "'," & _
                 "" & dtcEmpresa.BoundText & ",'" & txtEquipamento.Text & "'," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
       strValores = strValores & ",0"
    End If
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       
       strSet = "SET DFProximo_tipo_marcha_TBParametros_servicos = '" & txtTipo_Marcha.Text & "'," & _
                "    DFProximo_insumo_TBParametros_servicos = '" & txtInsumo.Text & "'," & _
                "    DFProximo_funcao_insumo_TBParametros_servicos = '" & txtFuncao_Insumo.Text & "'," & _
                "    DFProximo_servico_TBParametros_servicos = '" & txtServico.Text & "'," & _
                "    DFProximo_plano_servico_TBParametros_servicos = '" & txtPlano_Servico.Text & "'," & _
                "    DFProximo_equipamento_laboratorio = '" & txtEquipamento.Text & "'," & _
                "    DFData_alteracao_TBParametros_servicos = '" & Format(Date, "YYYYMMDD") & "'," & _
                "    DFIntegrado_filiais_TBParametros_servicos = 0"
     
    If booIntegra_Portal = True Then
       strSet = strSet & ",DFIntegrado_portal_TBParametros_servicos = 0"
    End If

       Call funcoes_banco.Alterar("TBParametros_servicos", strSet, "PKId_TBParametros_servicos", strID_Parametros, "Otica", Me, "BDRetaguarda")
       
       log.Descricao = "Alterando o registro: " + strID_Parametros
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       
       Call funcoes_banco.Gravar("TBParametros_servicos", strCampo, strValores, "Otica", Me, "BDRetaguarda")
       
       log.Descricao = "Gravando o registro: " + strID_Parametros
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    End If
    
    Call Objetos.Limpa_TXT(Me)
    
    dtcEmpresa.Text = Empty
        
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
       hfgParametros_Servicos.Visible = False
    End If
    
    sstParametros_Servicos.TabEnabled(0) = False
    sstParametros_Servicos.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + strID_Parametros
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBParametros_servicos", "PKId_TBParametros_servicos", strID_Parametros, "Otica", Me, "BDRetaguarda")
    
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
       hfgParametros_Servicos.Visible = False
    End If
        
    sstParametros_Servicos.TabEnabled(0) = False
    sstParametros_Servicos.Tab = 1
    
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
       hfgParametros_Servicos.Visible = False
    End If
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstParametros_Servicos.TabEnabled(0) = False
    sstParametros_Servicos.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
          
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
    
    'dtcEmpresa.boundtext = ---- Inserir aqui informações da DLLIntercomunicador de EXE's
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
       
    sstParametros_Servicos.TabEnabled(0) = True
    sstParametros_Servicos.Tab = 0
    
    txtServico.SetFocus
    booAlterar = False
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    On Error GoTo Erro
    
    strNomes = "ID,Serviço,Plano Serviço,Insumo," & _
               "Função Insumo,Tipo Marcha,Equipamento,Cod.Empresa,Empresa"
    strTamanho = "0,1000,1500,1000," & _
                 "1500,1200,1200,0,0"
    
    Movimentacoes.Monta_HFlex_Grid hfgParametros_Servicos, strTamanho, strNomes, 9, "Otica", Me
    
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
        
    Call Monta_Combo
          
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Reposicao")
    Resume Next
End Function

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Sub txtServico_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtServico_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtPlano_servico_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPlano_servico_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtInsumo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtInsumo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtFuncao_Insumo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFuncao_Insumo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtTipo_Marcha_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtTipo_Marcha_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub
Private Sub txtEquipamento_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtEquipamento_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Serviço")
    cbbCampos.AddItem ("Plano Serviço")
    cbbCampos.AddItem ("Insumo")
    cbbCampos.AddItem ("Função Insumo")
    cbbCampos.AddItem ("Tipo Marcha")
    cbbCampos.AddItem ("Equipamento")
End Function

Private Function Consulta()
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If

    strSql = "SELECT TBParametros_servicos.PKId_TBParametros_servicos," & _
             "DFProximo_servico_TBParametros_servicos,DFProximo_plano_servico_TBParametros_servicos," & _
             "TBParametros_servicos.DFProximo_insumo_TBParametros_servicos," & _
             "TBParametros_servicos.DFProximo_funcao_insumo_TBParametros_servicos," & _
             "TBParametros_servicos.DFProximo_tipo_marcha_TBParametros_servicos," & _
             "DFProximo_equipamento_laboratorio," & _
             "FKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa " & _
             "FROM TBParametros_servicos " & _
             "INNER JOIN TBEmpresa ON TBParametros_servicos.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa "
             
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Serviço" Then
          strSql = strSql & " WHERE convert(nvarchar,DFProximo_servico_TBParametros_servicos) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Plano Serviço" Then
          strSql = strSql & " WHERE convert(nvarchar,DFProximo_plano_servico_TBParametros_servicos) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Insumo" Then
          strSql = strSql & " WHERE convert(nvarchar,DFProximo_insumo_TBParametros_servicos) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Função Insumo" Then
          strSql = strSql & " WHERE convert(nvarchar,DFProximo_funcao_insumo_TBParametros_servicos) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Tipo Marcha" Then
          strSql = strSql & " WHERE convert(nvarchar,DFProximo_tipo_marcha_TBParametros_servicos) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Equipamento" Then
          strSql = strSql & " WHERE convert(nvarchar,DFProximo_equipamento_laboratorio) = '" & txtConsulta.Text & "'"
       End If
       If IsNumeric(strCodigo_Empresa_Consulta) = False Then
           strSql = strSql & " AND TBParametros_servicos.FKCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' "
       ElseIf IsNumeric(strCodigo_Empresa_Consulta) = True And CDbl(strCodigo_Empresa_Consulta) <> 0 Then
           strSql = strSql & " AND TBParametros_servicos.FKCodigo_TBEmpresa = '" & strCodigo_Empresa_Consulta & "' "
       End If
    Else
       If IsNumeric(strCodigo_Empresa_Consulta) = False Then
          strSql = strSql & " WHERE TBParametros_servicos.FKCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' "
       ElseIf IsNumeric(strCodigo_Empresa_Consulta) = True And CDbl(strCodigo_Empresa_Consulta) <> 0 Then
          strSql = strSql & " WHERE TBParametros_servicos.FKCodigo_TBEmpresa = '" & strCodigo_Empresa_Consulta & "' "
       End If
    End If
    
    frmAguarde.Show
    DoEvents

    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgParametros_Servicos, strTamanho, strNomes, "BDRetaguarda", "Otica", Me

    hfgParametros_Servicos.Row = 1
    hfgParametros_Servicos.Col = 0
    If hfgParametros_Servicos.Text = Empty Then
       hfgParametros_Servicos.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgParametros_Servicos, strTamanho, strNomes, 9, "Otica", Me
    End If
    
    Unload frmAguarde
    hfgParametros_Servicos.Refresh
    
End Function

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKId_TBParametros_servicos", strID_Parametros, "DFIntegrado_filiais_TBParametros_servicos", "TBParametros_servicos", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBParametros_servicos", Me.Top, Me.Left, Me.Width, Me.Height, "Parâmetros de Serviços")
    
End Function

