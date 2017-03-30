VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIntegracao_retaguarda_portal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Integração Retaguarda X Portal"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIntegracao_retaguarda_portal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   12270
   Begin VB.Frame Frame1 
      Caption         =   "Informações do Portal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6195
      Left            =   0
      TabIndex        =   11
      Top             =   390
      Width           =   12225
      Begin VB.Frame Frame4 
         Caption         =   "Resumos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6180
         TabIndex        =   22
         Top             =   5370
         Width           =   5925
         Begin VB.Label lblNum_registros_atualizar 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   240
            Left            =   150
            TabIndex        =   24
            Top             =   330
            Width           =   870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registros à Integrar no Retaguarda"
            Height          =   240
            Left            =   1110
            TabIndex        =   23
            Top             =   330
            Width           =   3030
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Legenda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   5370
         Width           =   5985
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Integrado no Retaguarda"
            Height          =   240
            Left            =   3630
            TabIndex        =   21
            Top             =   330
            Width           =   2145
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   6  'Inside Solid
            Height          =   255
            Left            =   3300
            Top             =   330
            Width           =   195
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Não Integrado no Retaguarda"
            Height          =   240
            Left            =   510
            TabIndex        =   20
            Top             =   330
            Width           =   2535
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   6  'Inside Solid
            Height          =   255
            Left            =   180
            Top             =   300
            Width           =   195
         End
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   6030
         TabIndex        =   6
         Top             =   1380
         Visible         =   0   'False
         Width           =   5190
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   11700
         Picture         =   "frmIntegracao_retaguarda_portal.frx":1782
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   1380
         Width           =   375
      End
      Begin VB.CommandButton cmdConsultar 
         Height          =   360
         Left            =   11280
         Picture         =   "frmIntegracao_retaguarda_portal.frx":27C4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Consultar"
         Top             =   1380
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opção de integração"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9120
         TabIndex        =   13
         Top             =   390
         Width           =   2955
         Begin VB.OptionButton optFabricante 
            Caption         =   "Fabricante"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   2
            Top             =   330
            Width           =   1365
         End
         Begin VB.OptionButton optTriagem 
            Caption         =   "Triagem"
            Height          =   285
            Left            =   210
            TabIndex        =   1
            Top             =   330
            Width           =   1815
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgIntegracao 
         Height          =   3465
         Left            =   120
         TabIndex        =   9
         Top             =   1830
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   6112
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
         TabIndex        =   3
         Top             =   1380
         Width           =   2325
         _ExtentX        =   4101
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
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   660
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   360
         Left            =   8700
         TabIndex        =   15
         Top             =   1380
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   55771137
         CurrentDate     =   37923
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   360
         Left            =   6030
         TabIndex        =   16
         Top             =   1380
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   55771137
         CurrentDate     =   37923
      End
      Begin AutoCompletar.CbCompleta cbbMes_competencia 
         Height          =   360
         Left            =   2490
         TabIndex        =   4
         Top             =   1380
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
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
      Begin AutoCompletar.CbCompleta cbbAno_competencia 
         Height          =   360
         Left            =   4260
         TabIndex        =   5
         Top             =   1380
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Período"
         Height          =   240
         Left            =   6060
         TabIndex        =   25
         Top             =   1140
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano Competência"
         Height          =   240
         Left            =   4290
         TabIndex        =   18
         Top             =   1140
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês Competência"
         Height          =   240
         Left            =   2520
         TabIndex        =   17
         Top             =   1140
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Empresa [ F2 ]"
         Height          =   240
         Left            =   150
         TabIndex        =   14
         Top             =   390
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   150
         TabIndex        =   12
         Top             =   1140
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   10470
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":44BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":47D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":4AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":4E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":5226
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":5540
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   10140
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":585A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":5B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":5E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":6228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":65C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIntegracao_retaguarda_portal.frx":68DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1(0)"
      HotImageList    =   "ImageList1(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Confirmar"
            Object.ToolTipText     =   "Gravar registro - CTRL+G"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sair"
            Object.ToolTipText     =   "Sair - CTRL+S"
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIntegracao_retaguarda_portal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Movimentações de Integração Portal X Retaguarda                '
' Data de Criação........: 16/05/2006                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:                                                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim conexao As New DLLConexao_Sistema.conexao
Dim booPrivilegio_Incluir As Boolean
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log

Private Sub cbbCampos_LostFocus()
    If Me.cbbCampos.Text = "Todos" Then
        txtConsulta.Visible = False
        dtpFinal.Visible = False
        Me.Label5.Visible = False
        dtpInicial.Visible = False
        Me.cbbAno_competencia.Visible = False
        Me.cbbMes_competencia.Visible = False
        Me.Label1.Visible = False
        Me.Label2.Visible = False
        Me.cmdConsultar.SetFocus
    Else
        If Me.cbbCampos.Text = "Data de Lançamento" Or Me.cbbCampos.Text = "Data de Validade" Then
            txtConsulta.Visible = False
            dtpFinal.Visible = True
            dtpInicial.Visible = True
            Me.Label5.Visible = True
            Me.cbbAno_competencia.Visible = True
            Me.cbbMes_competencia.Visible = True
            Me.Label1.Visible = True
            Me.Label2.Visible = True
            Me.dtpInicial.SetFocus
        Else
            txtConsulta.Visible = True
            dtpFinal.Visible = False
            Me.Label5.Visible = False
            dtpInicial.Visible = False
            Me.cbbAno_competencia.Visible = True
            Me.cbbMes_competencia.Visible = True
            Me.Label1.Visible = True
            Me.Label2.Visible = True
            Me.cbbMes_competencia.SetFocus
        End If
    End If
End Sub

Private Sub cmdConsultar_Click()
    frmAguarde.Show
    Call Reposicao
    Unload frmAguarde
End Sub

Private Sub cmdRefresh_Click()
    frmAguarde.Show
    Call Reposicao
    Unload frmAguarde
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 71: Call Gravar 'CTRL+G
                      'Case 67: Cancelar 'CTRL+C
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
    log.Programa = "Gera Plano Completo"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    strSql = Empty
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    optTriagem.Value = True

    log.Descricao = "Inicializando Movimentação de integração Retaguarda X Portal"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    Call Monta_cbbs
    
    Exit Sub
    
Erro:
    Call Erro.Erro(Me, "Otica", "Load")
    Exit Sub
End Sub

Private Function Reposicao()

    Dim rstTriagem As New ADODB.Recordset
    Dim contador_colunas As Long
    Dim Linhas As Long
    
    If Me.optTriagem.Value = True Then
        If Me.cbbCampos.Text = "" Then
           MsgBox "Favor selecione um filtro!", vbInformation, "Only Tech"
           Me.cbbAno_competencia.SetFocus
           Exit Function
        End If
        If Me.cbbCampos.Text <> "Todos" Then
            If Me.cbbCampos.Text <> "Data de Lançamento" And Me.cbbCampos.Text <> "Data de Validade" Then
                If Me.txtConsulta.Text = "" Then
                   MsgBox "Favor digite uma condição!", vbInformation, "Only Tech"
                   Me.txtConsulta.SetFocus
                   Exit Function
                End If
                If Me.cbbAno_competencia.Text = "" And Me.cbbMes_competencia.Text = "" Then
                   MsgBox "Favor selecione um ano/mês de competência!", vbInformation, "Only Tech"
                   Me.txtConsulta.SetFocus
                   Exit Function
                End If
            End If
        End If
        
        strSql = Empty
        strSql = "SELECT " & _
                 "TBCliente_portal.DFCodigo_TBCliente_portal," & _
                 "TBCliente_portal.DFRazao_social_TBCliente_portal," & _
                 "TBInsumo_portal.DFDescricao_TBInsumo_portal," & _
                 "TBFabricante_portal.PKCodigo_TBFabricante_portal," & _
                 "TBFabricante_portal.DFDescricao_TBFabricante_portal," & _
                 "TBTriagem_portal.DFData_fabricacao_TBTriagem_portal," & _
                 "TBTriagem_portal.DFData_lancamento_TBTriagem_portal," & _
                 "TBTriagem_portal.DFLote_TBTriagem_portal," & _
                 "TBTriagem_portal.DFData_validade_TBTriagem_portal," & _
                 "TBTriagem_portal.DFMes_ano_competencia_TBTriagem_portal," & _
                 "TBTriagem_portal.DFAno_competencia_TBTriagem_portal,TBTriagem_portal.DFAtualizado_no_retaguarda,TBTriagem_portal.PKId_TBTriagem_portal " & _
                 "FROM TBTRIAGEM_PORTAL " & _
                 "LEFT JOIN TBCliente_portal " & _
                 "ON TBTRIAGEM_PORTAL.FKId_TBCliente_portal = TBCliente_portal.DFCodigo_TBCliente_portal " & _
                 "LEFT JOIN TBRamo_atividade_portal " & _
                 "ON TBCliente_portal.FKCodigo_TBRamo_atividade_portal = TBRamo_atividade_portal.PKCodigo_TBRamo_atividade_portal " & _
                 "LEFT JOIN TBInsumo_portal " & _
                 "ON TBTRIAGEM_PORTAL.FKCodigo_TBInsumo_portal = TBInsumo_portal.PKCodigo_TBInsumo_portal " & _
                 "LEFT JOIN TBFabricante_portal " & _
                 "ON TBTRIAGEM_PORTAL.FKCodigo_TBFabricante_portal = TBFabricante_portal.PKCodigo_TBFabricante_portal "
                 
        If Me.cbbCampos.Text <> "Todos" Then
            strSql = strSql & "WHERE TBTriagem_portal.DFMes_ano_competencia_TBTriagem_portal = '" & Me.cbbMes_competencia.Text & "' " & _
                              "AND TBTriagem_portal.DFAno_competencia_TBTriagem_portal = '" & Me.cbbAno_competencia.Text & "' "
            If Me.cbbCampos.Text = "Cód.Cliente" Then
               strSql = strSql & "AND TBCliente_portal.DFCodigo_TBCliente_portal = " & Me.txtConsulta.Text & ""
            End If
            If Me.cbbCampos.Text = "Nome Cliente" Then
               strSql = strSql & "AND TBCliente_portal.DFRazao_social_TBCliente_portal = '" & Me.txtConsulta.Text & "'"
            End If
            If Me.cbbCampos.Text = "Insumo" Then
               strSql = strSql & "AND TBCliente_portal.DFDescricao_TBInsumo_portal = '" & Me.txtConsulta.Text & "'"
            End If
            If Me.cbbCampos.Text = "Cód Fabricante" Then
               strSql = strSql & "AND TBCliente_portal.PKCodigo_TBFabricante_portal = '" & Me.txtConsulta.Text & "'"
            End If
            If Me.cbbCampos.Text = "Nome Fabricante" Then
               strSql = strSql & "AND TBCliente_portal.DFDescricao_TBFabricante_portal = '" & Me.txtConsulta.Text & "'"
            End If
            If Me.cbbCampos.Text = "Data de Lançamento" Then
               strSql = strSql & "AND TBTriagem_portal.DFData_lancamento_TBTriagem_portal BETWEEN = '" & Me.dtpInicial.Value & "' AND '" & Me.dtpFinal.Value & "'"
            End If
            If Me.cbbCampos.Text = "Lote" Then
               strSql = strSql & "AND TBCliente_portal.DFLote_TBTriagem_portal = '" & Me.txtConsulta.Text & "'"
            End If
            If Me.cbbCampos.Text = "Data de Validade" Then
               strSql = strSql & "AND TBTriagem_portal.DFData_validade_TBTriagem_portal BETWEEN = '" & Me.dtpInicial.Value & "' AND '" & Me.dtpFinal.Value & "'"
            End If
        End If
        
        Movimentacoes.Select_geral strSql, "ortofarma1", rstTriagem, "Portal", Me
        
        If rstTriagem.BOF = True And rstTriagem.EOF = True Then hfgIntegracao.Clear: Exit Function
        
        If rstTriagem.RecordCount = 0 Then
           Set rstTriagem = Nothing
           Exit Function
        End If
        
        hfgIntegracao.Clear
        
        Call Monta_Cabecalhos
        
        contador_colunas = 2
        Linhas = 1
        
        hfgIntegracao.Cols = 15
        hfgIntegracao.Rows = rstTriagem.RecordCount + 1
        
        rstTriagem.MoveFirst
        
        Do While Linhas <= rstTriagem.RecordCount
           DoEvents
           hfgIntegracao.Row = Linhas
           hfgIntegracao.Col = 0
           Me.hfgIntegracao.ColWidth(0) = 300
           hfgIntegracao.CellBackColor = &H80FFFF
           hfgIntegracao.CellFontBold = False
           hfgIntegracao.CellFontSize = 7
           hfgIntegracao.Text = Linhas
           contador_colunas = 1
           Do While contador_colunas <= rstTriagem.Fields.Count
              hfgIntegracao.Col = 0
              hfgIntegracao.Text = rstTriagem.AbsolutePosition
              hfgIntegracao.Col = 1
              If IsNull(rstTriagem!DFAtualizado_no_retaguarda) = True Then
                 hfgIntegracao.CellBackColor = vbRed
              Else
                 If rstTriagem!DFAtualizado_no_retaguarda = True Then
                    hfgIntegracao.CellBackColor = vbGreen
                 Else
                    hfgIntegracao.CellBackColor = vbRed
                 End If
              End If
              
              hfgIntegracao.Col = 2
              hfgIntegracao.Text = ""
              
              hfgIntegracao.Col = 3
              If IsNull(rstTriagem!DFCodigo_TBCliente_portal) = True Then
                 hfgIntegracao.Text = ""
              Else
                 hfgIntegracao.Text = rstTriagem!DFCodigo_TBCliente_portal
              End If
              hfgIntegracao.Col = 4
              If IsNull(rstTriagem!DFRazao_social_TBCliente_portal) = True Then
                 hfgIntegracao.Text = ""
              Else
                 hfgIntegracao.Text = rstTriagem!DFRazao_social_TBCliente_portal
              End If
              hfgIntegracao.Col = 5
              If IsNull(rstTriagem!DFDescricao_TBInsumo_portal) = True Then
                 hfgIntegracao.Text = ""
              Else
                 hfgIntegracao.Text = rstTriagem!DFDescricao_TBInsumo_portal
              End If
              hfgIntegracao.Col = 6
              If IsNull(rstTriagem!PKCodigo_TBFabricante_portal) = True Then
                 hfgIntegracao.Text = ""
              Else
                 hfgIntegracao.Text = rstTriagem!PKCodigo_TBFabricante_portal
              End If
              hfgIntegracao.Col = 7
              If IsNull(rstTriagem!DFDescricao_TBFabricante_portal) = True Then
                 hfgIntegracao.Text = ""
              Else
                 hfgIntegracao.Text = rstTriagem!DFDescricao_TBFabricante_portal
              End If
              hfgIntegracao.Col = 8
              hfgIntegracao.Text = rstTriagem!DFData_fabricacao_TBTriagem_portal
              hfgIntegracao.Col = 9
              hfgIntegracao.Text = rstTriagem!DFData_lancamento_TBTriagem_portal
              hfgIntegracao.Col = 10
              hfgIntegracao.Text = rstTriagem!DFLote_TBTriagem_portal
              hfgIntegracao.Col = 11
              hfgIntegracao.Text = rstTriagem!DFData_validade_TBTriagem_portal
              hfgIntegracao.Col = 12
              hfgIntegracao.Text = rstTriagem!DFMes_ano_competencia_TBTriagem_portal
              hfgIntegracao.Col = 13
              hfgIntegracao.Text = rstTriagem!DFAno_competencia_TBTriagem_portal
              hfgIntegracao.Col = 14
              hfgIntegracao.Text = rstTriagem!PKId_TBTriagem_portal
              contador_colunas = contador_colunas + 1
           Loop
           rstTriagem.MoveNext
           contador_colunas = 1
           Linhas = Linhas + 1
        Loop
        
        Set rstTriagem = Nothing
        hfgIntegracao.Row = 1
        hfgIntegracao.Col = 0
        
        lblNum_registros_atualizar.Caption = 0
        
    End If
    
End Function

Private Function Monta_Cabecalhos()

    If Me.optTriagem.Value = True Then
        'Montando o cabeçalho do grid do item pedido
        Me.hfgIntegracao.ColWidth(0) = 100
        Me.hfgIntegracao.Cols = 15
        Me.hfgIntegracao.Font.Name = "Tahoma"
        Me.hfgIntegracao.Font.Size = 8
        Me.hfgIntegracao.Row = 0
        '------------------------------------
        Me.hfgIntegracao.Col = 1
        Me.hfgIntegracao.Text = "Int."
        Me.hfgIntegracao.ColWidth(1) = 200
        '------------------------------------
        Me.hfgIntegracao.Col = 2
        Me.hfgIntegracao.Text = "?"
        Me.hfgIntegracao.ColWidth(2) = 200
        '------------------------------------
        Me.hfgIntegracao.Col = 3
        Me.hfgIntegracao.Text = "Cliente"
        Me.hfgIntegracao.ColWidth(3) = 800
        '------------------------------------
        Me.hfgIntegracao.Col = 4
        Me.hfgIntegracao.Text = "Nome Cliente"
        Me.hfgIntegracao.ColWidth(4) = 3500
        '------------------------------------
        Me.hfgIntegracao.Col = 5
        Me.hfgIntegracao.Text = "Insumo"
        Me.hfgIntegracao.ColWidth(5) = 3500
        '------------------------------------
        Me.hfgIntegracao.Col = 6
        Me.hfgIntegracao.Text = "Fabr."
        Me.hfgIntegracao.ColWidth(6) = 800
        '------------------------------------
        Me.hfgIntegracao.Col = 7
        Me.hfgIntegracao.Text = "Nome Fabricante"
        Me.hfgIntegracao.ColWidth(7) = 3500
        '------------------------------------
        Me.hfgIntegracao.Col = 8
        Me.hfgIntegracao.Text = "Dt.Fabr."
        Me.hfgIntegracao.ColWidth(8) = 1000
        '------------------------------------
        Me.hfgIntegracao.Col = 9
        Me.hfgIntegracao.Text = "Dt.Lanc."
        Me.hfgIntegracao.ColWidth(8) = 1000
        '------------------------------------
        Me.hfgIntegracao.Col = 10
        Me.hfgIntegracao.Text = "Lote"
        Me.hfgIntegracao.ColWidth(9) = 1500
        '------------------------------------
        Me.hfgIntegracao.Col = 11
        Me.hfgIntegracao.Text = "Validade"
        Me.hfgIntegracao.ColWidth(10) = 1000
        '------------------------------------
        Me.hfgIntegracao.Col = 12
        Me.hfgIntegracao.Text = "Mês Compt."
        Me.hfgIntegracao.ColWidth(11) = 1000
        '------------------------------------
        Me.hfgIntegracao.Col = 13
        Me.hfgIntegracao.Text = "Ano Compt."
        Me.hfgIntegracao.ColWidth(12) = 1000
        '------------------------------------
        Me.hfgIntegracao.Col = 14
        Me.hfgIntegracao.Text = "ID."
        Me.hfgIntegracao.ColWidth(13) = 800
    End If
End Function

Private Function Monta_cbbs()
    
        '-----------------------------------------------------------------------------------------------------
        'Mês
        '-----------------------------------------------------------------------------------------------------
        Me.cbbMes_competencia.AddItem ("       ")
        Me.cbbMes_competencia.AddItem ("Janeiro")
        Me.cbbMes_competencia.AddItem ("Fevereiro")
        Me.cbbMes_competencia.AddItem ("Março")
        Me.cbbMes_competencia.AddItem ("Abril")
        Me.cbbMes_competencia.AddItem ("Maio")
        Me.cbbMes_competencia.AddItem ("Junho")
        Me.cbbMes_competencia.AddItem ("Julho")
        Me.cbbMes_competencia.AddItem ("Agosto")
        Me.cbbMes_competencia.AddItem ("Setembro")
        Me.cbbMes_competencia.AddItem ("Outubro")
        Me.cbbMes_competencia.AddItem ("Novembro")
        Me.cbbMes_competencia.AddItem ("Dezembro")
        '-----------------------------------------------------------------------------------------------------
        'ANO
        '-----------------------------------------------------------------------------------------------------
        Me.cbbAno_competencia.AddItem ("    ")
        Me.cbbAno_competencia.AddItem ("1990")
        Me.cbbAno_competencia.AddItem ("1991")
        Me.cbbAno_competencia.AddItem ("1992")
        Me.cbbAno_competencia.AddItem ("1993")
        Me.cbbAno_competencia.AddItem ("1994")
        Me.cbbAno_competencia.AddItem ("1995")
        Me.cbbAno_competencia.AddItem ("1996")
        Me.cbbAno_competencia.AddItem ("1997")
        Me.cbbAno_competencia.AddItem ("1998")
        Me.cbbAno_competencia.AddItem ("1999")
        Me.cbbAno_competencia.AddItem ("2000")
        Me.cbbAno_competencia.AddItem ("2001")
        Me.cbbAno_competencia.AddItem ("2002")
        Me.cbbAno_competencia.AddItem ("2003")
        Me.cbbAno_competencia.AddItem ("2004")
        Me.cbbAno_competencia.AddItem ("2005")
        Me.cbbAno_competencia.AddItem ("2006")
        Me.cbbAno_competencia.AddItem ("2007")
        Me.cbbAno_competencia.AddItem ("2008")
        Me.cbbAno_competencia.AddItem ("2009")
        Me.cbbAno_competencia.AddItem ("2010")
        Me.cbbAno_competencia.AddItem ("2011")
        Me.cbbAno_competencia.AddItem ("2012")
        Me.cbbAno_competencia.AddItem ("2013")
        Me.cbbAno_competencia.AddItem ("2014")
        Me.cbbAno_competencia.AddItem ("2015")
        Me.cbbAno_competencia.AddItem ("2016")
        Me.cbbAno_competencia.AddItem ("2017")
        Me.cbbAno_competencia.AddItem ("2018")
        Me.cbbAno_competencia.AddItem ("2019")
        Me.cbbAno_competencia.AddItem ("2020")
        Me.cbbAno_competencia.AddItem ("2021")
        Me.cbbAno_competencia.AddItem ("2022")
        Me.cbbAno_competencia.AddItem ("2023")
        Me.cbbAno_competencia.AddItem ("2024")
        Me.cbbAno_competencia.AddItem ("2025")
        Me.cbbAno_competencia.AddItem ("2026")
        Me.cbbAno_competencia.AddItem ("2027")
        Me.cbbAno_competencia.AddItem ("2028")
        Me.cbbAno_competencia.AddItem ("2029")
        Me.cbbAno_competencia.AddItem ("2030")
        Me.cbbAno_competencia.AddItem ("2031")
        Me.cbbAno_competencia.AddItem ("2032")
        Me.cbbAno_competencia.AddItem ("2033")
        Me.cbbAno_competencia.AddItem ("2034")
        Me.cbbAno_competencia.AddItem ("2035")
        Me.cbbAno_competencia.AddItem ("2036")
        Me.cbbAno_competencia.AddItem ("2037")
        Me.cbbAno_competencia.AddItem ("2038")
        Me.cbbAno_competencia.AddItem ("2039")
        Me.cbbAno_competencia.AddItem ("2040")
        Me.cbbAno_competencia.AddItem ("2041")
        Me.cbbAno_competencia.AddItem ("2042")
        Me.cbbAno_competencia.AddItem ("2043")
        Me.cbbAno_competencia.AddItem ("2044")
        Me.cbbAno_competencia.AddItem ("2045")
        Me.cbbAno_competencia.AddItem ("2046")
        Me.cbbAno_competencia.AddItem ("2047")
        Me.cbbAno_competencia.AddItem ("2048")
        Me.cbbAno_competencia.AddItem ("2049")
        Me.cbbAno_competencia.AddItem ("2050")
        'Filtros
        Me.cbbCampos.AddItem ("Todos")
        Me.cbbCampos.AddItem ("Cód.Cliente")
        Me.cbbCampos.AddItem ("Nome Cliente")
        Me.cbbCampos.AddItem ("Insumo")
        Me.cbbCampos.AddItem ("Cód Fabricante")
        Me.cbbCampos.AddItem ("Nome Fabricante")
        Me.cbbCampos.AddItem ("Data de Lançamento")
        Me.cbbCampos.AddItem ("Lote")
        Me.cbbCampos.AddItem ("Data de Validade")
        
        Me.cbbCampos.Text = "Todos"
        
End Function

Private Sub hfgIntegracao_Click()
    'VERIFICANDO SE O USUARIO CLICOU EM LINHA NÃO PERMITIDA
    If Me.hfgIntegracao.Row = 0 Then Exit Sub
    If Me.hfgIntegracao.Col = 0 Then Exit Sub
    
    'MARCAÇÃO DE SIM / NÃO - CONFORME O CLICK DO USUARIO
    If hfgIntegracao.Col = 2 Then
       If hfgIntegracao.Text = Empty Then
          hfgIntegracao.Text = Empty
          hfgIntegracao.CellFontBold = True
          hfgIntegracao.CellForeColor = &HC00000
          hfgIntegracao.Text = "X"
          lblNum_registros_atualizar.Caption = CDbl(lblNum_registros_atualizar.Caption) + 1
       ElseIf hfgIntegracao.Text = "X" Then
          hfgIntegracao.Text = Empty
          lblNum_registros_atualizar.Caption = CDbl(lblNum_registros_atualizar.Caption) - 1
       End If
    End If
End Sub

Private Sub hfgIntegracao_DblClick()
    
    If Me.hfgIntegracao.Row = 1 Then
       hfgIntegracao.Sort = 1
       Exit Sub
    End If
    
    If Me.hfgIntegracao.Col = 0 Then
       Dim intResult As Integer
       intResult = MsgBox("Deseja realmente excluir esta informação do Portal??", vbYesNo, "Only Tech")
       If intResult = 7 Then
          Exit Sub
       End If
       If intResult = 6 Then
          Me.hfgIntegracao.Col = 13
          Call Excluir_Triagem(hfgIntegracao.Text)
       End If
    End If
    
End Sub
Private Function Excluir_Triagem(lngID_Portal As String)

    On Error GoTo Erro_exclusao
    
    Call funcoes_banco.Excluir("TBTriagem_portal", "PKId_TBTriagem_portal", lngID_Portal, "Portal", Me, "ortofarma1")
    
    Exit Function
    
Erro_exclusao:
    
    Call Erro.Erro(Me, "Otica")
    
    Exit Function

End Function

Private Function Gravar()

    Dim intContador_gravar As Integer
    Dim strDFCodigo_TBCliente_portal As String
    Dim strDFRazao_social_TBCliente_portal As String
    Dim strDFDescricao_TBInsumo_portal As String
    Dim strPKCodigo_TBFabricante_portal As String
    Dim strDFDescricao_TBFabricante_portal As String
    Dim strDFData_lancamento_TBTriagem_portal As String
    Dim strDFLote_TBTriagem_portal As String
    Dim strDFData_validade_TBTriagem_portal As String
    Dim strDFMes_ano_competencia_TBTriagem_portal As String
    Dim strDFAno_competencia_TBTriagem_portal As String
    Dim strPKId_TBTriagem_portal As String
    Dim CNConexao As New DLLConexao_Sistema.conexao
    Dim CNConexao_PORTAL As New DLLConexao_Sistema.conexao
    Dim rstInsumo As New ADODB.Recordset
    
    strSql = Empty
    strSql = "SELECT * FROM TBInsumo"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstInsumo, "Otica", Me
        
    CNConexao.Abrir_conexao ("Otica")
    CNConexao.CNConexao.BeginTrans
    
    CNConexao_PORTAL.Abrir_conexao ("Portal")
    CNConexao_PORTAL.CNConexao.BeginTrans
    
    intContador_gravar = 1
    
    Do While hfgIntegracao.Rows > intContador_gravar
    
       hfgIntegracao.Row = intContador_gravar
       hfgIntegracao.Col = 2
       
       If hfgIntegracao.Text = "X" Then
          
          hfgIntegracao.Col = 3
          strDFCodigo_TBCliente_portal = hfgIntegracao.Text
          
          hfgIntegracao.Col = 4
          strDFRazao_social_TBCliente_portal = hfgIntegracao.Text
          
          hfgIntegracao.Col = 5
          strDFDescricao_TBInsumo_portal = hfgIntegracao.Text
          
          hfgIntegracao.Col = 6
          strPKCodigo_TBFabricante_portal = hfgIntegracao.Text
          
          hfgIntegracao.Col = 7
          strDFDescricao_TBFabricante_portal = hfgIntegracao.Text
          
          hfgIntegracao.Col = 8
          strDFData_fabricacao_TBTriagem_portal = Format(hfgIntegracao.Text, "YYYYMMDD")
          
          hfgIntegracao.Col = 9
          strDFData_lancamento_TBTriagem_portal = Format(hfgIntegracao.Text, "YYYYMMDD")
          
          hfgIntegracao.Col = 10
          strDFLote_TBTriagem_portal = hfgIntegracao.Text
          
          hfgIntegracao.Col = 11
          strDFData_validade_TBTriagem_portal = Format(hfgIntegracao.Text, "YYYYMMDD")
          
          hfgIntegracao.Col = 12
          strDFMes_ano_competencia_TBTriagem_portal = hfgIntegracao.Text
          
          hfgIntegracao.Col = 13
          strDFAno_competencia_TBTriagem_portal = hfgIntegracao.Text
          
          hfgIntegracao.Col = 14
          strPKId_TBTriagem_portal = hfgIntegracao.Text

          IDCLIENTE = Funcoes_Gerais.Localiza_ID("PKId_TBCliente", "IXCodigo_TBCliente", strDFCodigo_TBCliente_portal, "TBCliente", "Otica", Me, "BDRetaguarda")
          
          rstInsumo.MoveFirst
          rstInsumo.Find ("DFDescricao_TBInsumo = '" & strDFDescricao_TBInsumo_portal & "'")

          INSUMO = rstInsumo!PKCodigo_TBInsumo
          
          strSql = Empty
          strSql = "INSERT INTO TBTriagem(FKCodigo_TBFabricante,FKCodigo_TBInsumo,FKId_TBCliente,DFData_lancamento_TBTriagem,DFData_fabricacao_TBTriagem,DFData_validade_TBTriagem,DFLote_TBTriagem,DFMes_ano_competencia_TBTriagem,DFIntegrado_TBTriagem,DFAno_competencia_TBTriagem,DFCodigo_Identificador_TBTriagem) " & _
                   "VALUES (" & strPKCodigo_TBFabricante_portal & "," & INSUMO & "," & IDCLIENTE & ",'" & strDFData_lancamento_TBTriagem_portal & "','" & strDFData_fabricacao_TBTriagem_portal & "','" & strDFData_validade_TBTriagem_portal & "','" & strDFLote_TBTriagem_portal & "','" & strDFMes_ano_competencia_TBTriagem_portal & "',1,'" & strDFAno_competencia_TBTriagem_portal & "'," & strPKId_TBTriagem_portal & ")"
          CNConexao.CNConexao.Execute strSql
          
          'Atualizando o Portal
          strSql = Empty
          strSql = "UPDATE TBTriagem_portal set DFAtualizado_no_retaguarda = 1 WHERE PKId_TBTriagem_portal = " & strPKId_TBTriagem_portal & "                   "
          CNConexao_PORTAL.CNConexao.Execute strSql
          
       End If
       
       intContador_gravar = intContador_gravar + 1
       
    Loop
    
    CNConexao.CNConexao.CommitTrans
    CNConexao_PORTAL.CNConexao.CommitTrans
    
    Set rstInsumo = Nothing
    
    MsgBox "Registros atualizados com sucesso no Retaguarda", vbInformation, "Only Tech"
    
End Function

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Gravar
           'Case 2: Call Cancelar
           Case 4: Unload Me
    End Select
End Sub
