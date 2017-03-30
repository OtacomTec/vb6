VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmServicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serviços"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServicos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5970
   Begin TabDlg.SSTab sstServico 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   5900
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
      TabPicture(0)   =   "frmServicos.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(8)=   "txtDescricao"
      Tab(0).Control(9)=   "txtCodigo"
      Tab(0).Control(10)=   "txtPreco_Conveniado_1"
      Tab(0).Control(11)=   "txtPreco_Isolado_1"
      Tab(0).Control(12)=   "txtPreco_Conveniado_2"
      Tab(0).Control(13)=   "txtPreco_Isolado_2"
      Tab(0).Control(14)=   "txtPreco_Isolado_3"
      Tab(0).Control(15)=   "txtPreco_Conveniado_3"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmServicos.frx":179E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cbbCampos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgServico"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdConsulta"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdRefresh"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtConsulta"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtPreco_Conveniado_3 
         Height          =   375
         Left            =   -72210
         MaxLength       =   40
         TabIndex        =   8
         Top             =   2760
         Width           =   2985
      End
      Begin VB.TextBox txtPreco_Isolado_3 
         Height          =   375
         Left            =   -74880
         TabIndex        =   7
         Top             =   2760
         Width           =   2625
      End
      Begin VB.TextBox txtPreco_Isolado_2 
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         Top             =   2100
         Width           =   2625
      End
      Begin VB.TextBox txtPreco_Conveniado_2 
         Height          =   375
         Left            =   -72210
         MaxLength       =   40
         TabIndex        =   6
         Top             =   2100
         Width           =   2985
      End
      Begin VB.TextBox txtPreco_Isolado_1 
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   1440
         Width           =   2625
      End
      Begin VB.TextBox txtPreco_Conveniado_1 
         Height          =   375
         Left            =   -72210
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1440
         Width           =   2985
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   1
         Top             =   780
         Width           =   1485
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   -73350
         MaxLength       =   40
         TabIndex        =   2
         Top             =   780
         Width           =   4125
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   2070
         TabIndex        =   10
         Top             =   780
         Width           =   2895
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
         Left            =   5430
         Picture         =   "frmServicos.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
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
         Left            =   5040
         Picture         =   "frmServicos.frx":27FC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgServico 
         Height          =   1965
         Left            =   120
         TabIndex        =   12
         Top             =   1230
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   3466
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
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   1905
         _ExtentX        =   3360
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Preço Conveniado Tab. 3"
         Height          =   240
         Left            =   -72210
         TabIndex        =   23
         Top             =   2520
         Width           =   2145
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Isolado Tab. 3"
         Height          =   240
         Left            =   -74880
         TabIndex        =   22
         Top             =   2520
         Width           =   1770
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Isolado Tab. 2"
         Height          =   240
         Left            =   -74880
         TabIndex        =   21
         Top             =   1860
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Preço Conveniado Tab. 2"
         Height          =   240
         Left            =   -72210
         TabIndex        =   20
         Top             =   1860
         Width           =   2145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Isolado Tab. 1"
         Height          =   240
         Left            =   -74880
         TabIndex        =   19
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Preço Conveniado Tab. 1"
         Height          =   240
         Left            =   -72210
         TabIndex        =   18
         Top             =   1200
         Width           =   2145
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   -74880
         TabIndex        =   16
         Top             =   540
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   -73350
         TabIndex        =   15
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   540
         Width           =   435
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   5970
      _ExtentX        =   10530
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
      Left            =   6210
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
            Picture         =   "frmServicos.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServicos.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServicos.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServicos.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServicos.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServicos.frx":5578
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServicos.frx":5892
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro de Servicos                                           '
' Data de Criação........: 30/04/2006                                                     '
' Equipe Responsável.....: Jones Sá Peixoto                                               '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCombo As String
Dim strConsulta As String
Dim strTamanho As String
Dim strNomes As String
Dim intCodigo_empresa As Integer
Dim booAlterar As Boolean
Public strSql As String
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
Option Explicit

Function Imprimir()
    On Error GoTo Erro
    'Tratamento de erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório. Verifique.", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    frmConsole_Servicos.Show
        
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
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
  
Private Sub Form_Load()
    On Error GoTo Erro
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Serviços"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstServico, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando Cadastro de Serviços"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    intCodigo_empresa = MDIPrincipal.OCXUsuario.Empresa
    
    sstServico.TabEnabled(0) = False
    sstServico.Tab = 1
        
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
    log.Descricao = "Finalizando cadastro de Seção"
        
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    strCombo = Empty
    
    If frmIntegracao.Visible = True Then
        Unload frmIntegracao
    End If
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgServico_Click()
    
    If hfgServico.Col = 0 And hfgServico.Text <> Empty Then
        
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
       
       txtCodigo.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 1))
       txtDescricao.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 2))
       txtPreco_Isolado_1.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 3))
       txtPreco_Conveniado_1.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 4))
       txtPreco_Isolado_2.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 5))
       txtPreco_Conveniado_2.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 6))
       txtPreco_Isolado_3.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 7))
       txtPreco_Conveniado_3.Text = hfgServico.TextArray((hfgServico.Row * hfgServico.Cols + hfgServico.Col + 8))
       
       booAlterar = True
       txtCodigo.Enabled = False
       
       sstServico.TabEnabled(0) = True
       sstServico.Tab = 0
       Me.txtDescricao.SetFocus
   End If
   Unload frmAguarde
End Sub

Private Sub hfgServico_DblClick()
    hfgServico.Sort = 1
End Sub

Private Sub hfgServico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgServico_Click
    End If
End Sub

Private Sub sstServico_Click(PreviousTab As Integer)
    If sstServico.Tab = 0 Then
       txtDescricao.SetFocus
    ElseIf sstServico.Tab = 1 Then
       If frmIntegracao.Visible = True Then
          Unload frmIntegracao
       End If
       If strCombo <> Empty And strCombo <> "Todos" Then
          cbbCampos.Text = strCombo
          txtConsulta.SetFocus
       ElseIf strCombo = "Todos" Then
          hfgServico.Row = 1
          hfgServico.Col = 0
          hfgServico.SetFocus
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
    
    'Verifica se os campos necessarios para gravar não estão nulos
    If txtDescricao.Text = Empty Then
       MsgBox "O campo descrição não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtDescricao.SetFocus
       Exit Function
    End If
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
           
    Dim intEmpresa As Integer
    Dim strCodigo_Servico As String
    Dim strProx_Cod_Servico As String
    Dim rstVerifica As New ADODB.Recordset
    
    intEmpresa = MDIPrincipal.OCXUsuario.Empresa
    
    Call Objetos.Maiusculo_TXT(Me)
    
    If booAlterar = False Then
       strProx_Cod_Servico = Funcoes_Gerais.Localiza_Proximo_Codigo("DFProximo_servico_TBParametros_servicos", "FKCodigo_TBEmpresa", intEmpresa, "TBParametros_Servicos", "Otica", Me, "BDRetaguarda")
       txtCodigo.Text = strProx_Cod_Servico
    End If
    
    strCampo = "PKCodigo_TBServico_laboratorio,DFDescricao_TBServico_laboratorio," & _
               "DFPreco1_TBServico_laboratorio,DFPreco1_conveniado_TBServico_laboratorio," & _
               "DFPreco2_TBServico_laboratorio,DFPreco2_conveniado_TBServico_laboratorio," & _
               "DFPreco3_TBServico_laboratorio,DFPreco3_conveniado_TBServico_laboratorio," & _
               "DFData_alteracao_TBServico_laboratorio,DFIntegrado_filiais_TBServico_laboratorio"
    
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBServico_laboratorio"
    End If
    
    strValores = "" & txtCodigo.Text & ",'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtPreco_Isolado_1.Text) & "," & Funcoes_Gerais.Grava_Moeda(txtPreco_Conveniado_1.Text) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtPreco_Isolado_2.Text) & "," & Funcoes_Gerais.Grava_Moeda(txtPreco_Conveniado_2.Text) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtPreco_Isolado_3.Text) & "," & Funcoes_Gerais.Grava_Moeda(txtPreco_Conveniado_3.Text) & "," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
        strValores = strValores & ",0"
    End If

    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFDescricao_TBServico_laboratorio = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                "DFPreco1_TBServico_laboratorio = " & Funcoes_Gerais.Grava_Moeda(txtPreco_Isolado_1.Text) & ",DFPreco1_conveniado_TBServico_laboratorio = " & Funcoes_Gerais.Grava_Moeda(txtPreco_Conveniado_1.Text) & "," & _
                "DFPreco2_TBServico_laboratorio = " & Funcoes_Gerais.Grava_Moeda(txtPreco_Isolado_2.Text) & ",DFPreco2_conveniado_TBServico_laboratorio = " & Funcoes_Gerais.Grava_Moeda(txtPreco_Conveniado_2.Text) & "," & _
                "DFPreco3_TBServico_laboratorio = " & Funcoes_Gerais.Grava_Moeda(txtPreco_Isolado_3.Text) & ",DFPreco3_conveniado_TBServico_laboratorio = " & Funcoes_Gerais.Grava_Moeda(txtPreco_Conveniado_3.Text) & "," & _
                "DFData_alteracao_TBServico_laboratorio = '" & Format(Date, "YYYYMMDD") & "'," & "DFIntegrado_filiais_TBServico_laboratorio = 0"
                
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBServico_laboratorio = 0"
       End If
       
       Call funcoes_banco.Alterar("TBServico_laboratorio", strSet, "PKCodigo_TBServico_laboratorio", txtCodigo.Text, "OTICA", Me, "BDRetaguarda")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBServico_laboratorio", strCampo, strValores, "OTICA", Me, "BDRetaguarda")
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
       
       'ATUALIZAÇÃO DA TABELA TBParametros_Servicos
        
       'Somente para mostrar ao usuario o código que o serviço foi incluido
       strCodigo_Servico = strProx_Cod_Servico
       
       If strCodigo_Servico <> Empty Then
          MsgBox "** O código desse Serviço é: " & strCodigo_Servico & "", vbOKOnly, "Only Tech"
       End If
        
       strProx_Cod_Servico = strProx_Cod_Servico + 1
       
       strSet = "SET DFProximo_servico_TBParametros_servicos = " & strProx_Cod_Servico & ""
       
       Call funcoes_banco.Alterar("TBParametros_Servicos", strSet, "FKCodigo_TBEmpresa", MDIPrincipal.OCXUsuario.Empresa, "Otica", Me, "BDRetaguarda")

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
       hfgServico.Visible = False
    End If
    
    sstServico.TabEnabled(0) = False
    
    sstServico.Tab = 1
        
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    strSql = "SELECT FKCodigo_TBPlano_servico FROM TBPlano_servico_servico_laboratorio " & _
             "WHERE FKCodigo_TBServico_laboratorio = " & txtCodigo.Text & ""
   
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       MsgBox "Este Serviço está vinculado ao Plano de Serviços de Código " & rstAplicacao.Fields("FKCodigo_TBPlano_servico") & " e não pode ser excluído. Verifique.", vbInformation, "Only Tech"
       Set rstAplicacao = Nothing
       Exit Function
    End If
    Set rstAplicacao = Nothing
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBServico_laboratorio", "PKCodigo_TBServico_laboratorio", Me.txtCodigo.Text, "OTICA", Me, "BDRetaguarda")
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
        
    'Gravando log
     log.Gravar_log "OTICA", Me
           
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
       hfgServico.Visible = False
    End If
            
    sstServico.TabEnabled(0) = False
    sstServico.Tab = 1
    
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
       hfgServico.Visible = False
    End If
    
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Operação com Registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    txtCodigo.Enabled = False
    
    sstServico.TabEnabled(0) = False
    sstServico.Tab = 1
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    
    Dim rstBusca_Parametro As New ADODB.Recordset
    Dim strCodigo_Servico As String
    
    On Error GoTo Erro
    Call Objetos.Limpa_TXT(Me)

    'verificacao de código
    strSql = Empty
    strSql = "SELECT * FROM TBParametros_Servicos " & _
             "WHERE TBParametros_Servicos.FKCodigo_TBEmpresa = " & MDIPrincipal.OCXUsuario.Empresa & ""
             
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Parametro, "Otica", Me)
        
    strCodigo_Servico = rstBusca_Parametro.Fields("DFProximo_servico_TBParametros_servicos")
    Set rstBusca_Parametro = Nothing
        
    strSql = Empty
    strSql = "SELECT * FROM TBServico_laboratorio WHERE TBServico_laboratorio.PKCodigo_TBServico_laboratorio = " & strCodigo_Servico & ""
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Parametro, "Otica", Me)
    
    If rstBusca_Parametro.RecordCount <> 0 Then
       MsgBox "O Código " & strCodigo_Servico & " já existe, por favor, verifique o cadastro Parâmetros de Serviços e atualize o código do próximo Serviço.", vbInformation, "Only Tech"
       Set rstBusca_Parametro = Nothing
       Call Objetos.Limpa_TXT(Me)
       sstServico.TabEnabled(1) = True
       sstServico.Tab = 1
       Exit Function
    End If
    Set rstBusca_Parametro = Nothing
    
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
    
    sstServico.TabEnabled(0) = True
    sstServico.Tab = 0
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
       Movimentacoes.Verifica_Numero "PKCodigo_TBServico_laboratorio", "TBServico_laboratorio", txtCodigo, "OTICA", Me
    End If
End Sub

Private Function Reposicao()
    On Error GoTo Erro
    
    strTamanho = "1000,2200,1700,1850,1700," & _
                 "1850,1700,1850"
    
    strNomes = "Código,Descrição,P. Isolado Tab. 1,P. Conveniado Tab.1,P. Isolado Tab. 2," & _
               "P. Conveniado Tab.2,P. Isolado Tab. 3,P. Conveniado Tab.3"
    
    Movimentacoes.Monta_HFlex_Grid hfgServico, strTamanho, strNomes, 8, "OTICA", Me
    
    Call Monta_Combo
              
    hfgServico.Refresh
    Exit Function
Erro:
   Call Erro.Erro(Me, "OTICA", "Reposicao")
   Resume Next
End Function

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Sub txtDescricao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDescricao_LostFocus()
    txtDescricao.Text = UCase(txtDescricao.Text)
End Sub

Private Sub txtPreco_Conveniado_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_Conveniado_1_LostFocus()
    txtPreco_Conveniado_1.Text = Format(txtPreco_Conveniado_1.Text, "#,###0.00")
End Sub

Private Sub txtPreco_Conveniado_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_Conveniado_2_LostFocus()
    txtPreco_Conveniado_2.Text = Format(txtPreco_Conveniado_2.Text, "#,###0.00")
End Sub

Private Sub txtPreco_Conveniado_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_Conveniado_3_LostFocus()
    txtPreco_Conveniado_3.Text = Format(txtPreco_Conveniado_3.Text, "#,###0.00")
End Sub

Private Sub txtPreco_Isolado_1_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPreco_Isolado_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_Isolado_1_LostFocus()
    txtPreco_Isolado_1.Text = Format(txtPreco_Isolado_1.Text, "#,###0.00")
End Sub

Private Sub txtPreco_Isolado_2_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPreco_Isolado_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_Isolado_2_LostFocus()
    txtPreco_Isolado_2.Text = Format(txtPreco_Isolado_2.Text, "#,###0.00")
End Sub

Private Sub txtPreco_Isolado_3_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPreco_Conveniado_1_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPreco_Conveniado_2_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPreco_Conveniado_3_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Consulta()
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
            
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
           
    txtConsulta.Text = Funcoes_Gerais.Grava_String(txtConsulta.Text)
    
    strSql = "SELECT PKCodigo_TBServico_laboratorio," & _
             "DFDescricao_TBServico_laboratorio," & _
             "DFPreco1_TBServico_laboratorio," & _
             "DFPreco1_conveniado_TBServico_laboratorio," & _
             "DFPreco2_TBServico_laboratorio," & _
             "DFPreco2_conveniado_TBServico_laboratorio," & _
             "DFPreco3_TBServico_laboratorio," & _
             "DFPreco3_conveniado_TBServico_laboratorio FROM TBServico_laboratorio"
        
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código do Serviço" Then
          strSql = strSql & " WHERE convert(nvarchar,PKCodigo_TBServico_laboratorio) LIKE '" & txtConsulta.Text & "' "
       Else
          strSql = strSql & " WHERE convert(nvarchar,DFDescricao_TBServico_laboratorio) LIKE '%" & txtConsulta.Text & "%' "
       End If
    End If
    
    strSql = strSql & " ORDER BY TBServico_laboratorio.PKCodigo_TBServico_laboratorio"
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgServico, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
           
    frmAguarde.Show
    DoEvents
    
    hfgServico.Refresh
    
    hfgServico.Row = 1
    hfgServico.Col = 0
    If hfgServico.Text = Empty Then
       hfgServico.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgServico, strTamanho, strNomes, 8, "Otica", Me
    End If
    
    hfgServico.Row = 1
    hfgServico.Col = 0
    hfgServico.SetFocus
    
    Unload frmAguarde
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código do Serviço")
    cbbCampos.AddItem ("Descrição do Serviço")
    cbbCampos.AddItem ("Preço Isolado Tab. 1")
    cbbCampos.AddItem ("Preço Conveniado Tab. 1")
    cbbCampos.AddItem ("Preço Isolado Tab. 2")
    cbbCampos.AddItem ("Preço Conveniado Tab. 2")
    cbbCampos.AddItem ("Preço Isolado Tab. 3")
    cbbCampos.AddItem ("Preço Conveniado Tab. 3")
End Function

Private Sub txtPreco_Isolado_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtPreco_Isolado_3_LostFocus()
    txtPreco_Isolado_3.Text = Format(txtPreco_Isolado_3.Text, "#,###0.00")
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo_TBServico_laboratorio", txtCodigo.Text, "DFIntegrado_filiais_TBServico_laboratorio", "TBServico_laboratorio", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBServico_laboratorio", Me.Top, Me.Left, Me.Width, Me.Height, "Serviço")
    
End Function

