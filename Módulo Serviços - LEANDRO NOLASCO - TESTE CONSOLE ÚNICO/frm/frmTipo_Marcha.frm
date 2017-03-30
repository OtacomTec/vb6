VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmTipo_Marcha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo Marcha"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTipo_Marcha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6405
   Begin TabDlg.SSTab sstTipo_Marcha 
      Height          =   2715
      Left            =   0
      TabIndex        =   13
      Top             =   330
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   4789
      _Version        =   393216
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
      TabPicture(0)   =   "frmTipo_Marcha.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDescricao_resumida"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNumero_Sequencial"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCodigo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDescricao"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&Equipamentos"
      TabPicture(1)   =   "frmTipo_Marcha.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRemover"
      Tab(1).Control(1)=   "cmdIncluir"
      Tab(1).Control(2)=   "txtCodigo_Equipamento"
      Tab(1).Control(3)=   "hfgEquipamento"
      Tab(1).Control(4)=   "dtcEquipamento"
      Tab(1).Control(5)=   "Label5"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&Listagem"
      TabPicture(2)   =   "frmTipo_Marcha.frx":17BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtConsulta"
      Tab(2).Control(1)=   "cmdRefresh"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdConsulta"
      Tab(2).Control(3)=   "hfgTipo_Marcha"
      Tab(2).Control(4)=   "cbbCampos"
      Tab(2).Control(5)=   "Label6"
      Tab(2).ControlCount=   6
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
         Left            =   -69780
         TabIndex        =   7
         ToolTipText     =   "Remover"
         Top             =   780
         Width           =   1065
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
         Left            =   -70860
         TabIndex        =   6
         ToolTipText     =   "Incluir"
         Top             =   780
         Width           =   1035
      End
      Begin VB.TextBox txtCodigo_Equipamento 
         Height          =   360
         Left            =   -74880
         TabIndex        =   4
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72990
         TabIndex        =   10
         Top             =   780
         Width           =   3435
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
         Left            =   -69090
         Picture         =   "frmTipo_Marcha.frx":17D6
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   -69480
         Picture         =   "frmTipo_Marcha.frx":2818
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   780
         Width           =   1230
      End
      Begin VB.TextBox txtNumero_Sequencial 
         Height          =   375
         Left            =   4440
         MaxLength       =   12
         TabIndex        =   3
         Top             =   2130
         Width           =   1815
      End
      Begin VB.TextBox txtDescricao_resumida 
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2130
         Width           =   4275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgTipo_Marcha 
         Height          =   1365
         Left            =   -74880
         TabIndex        =   12
         Top             =   1230
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   2408
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
         TabIndex        =   9
         Top             =   780
         Width           =   1845
         _ExtentX        =   3254
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgEquipamento 
         Height          =   1365
         Left            =   -74880
         TabIndex        =   8
         Top             =   1230
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   2408
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
      Begin MSDataListLib.DataCombo dtcEquipamento 
         Height          =   360
         Left            =   -73920
         TabIndex        =   5
         Top             =   780
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   8388608
         Text            =   ""
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   -74880
         TabIndex        =   21
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   20
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1200
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número Sequencial"
         Height          =   240
         Left            =   4440
         TabIndex        =   15
         Top             =   1890
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição Resumida"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   1890
         Width           =   1725
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
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
      Left            =   7020
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
            Picture         =   "frmTipo_Marcha.frx":4512
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipo_Marcha.frx":482C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipo_Marcha.frx":4B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipo_Marcha.frx":4EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipo_Marcha.frx":527A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipo_Marcha.frx":5594
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipo_Marcha.frx":58AE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTipo_Marcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro Tipo Marcha                                           '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Peixoto                                                  '
' Data de Criação........: 04/03/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strCombo As String
Dim strConsulta As String
Dim strNomes As String
Dim strTamanho As String
Dim strCampo_consulta As String
Dim intContador As Integer
Public strSql As String
Dim strId As String
Dim intClique_Equipamento As Integer
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
    
    Call frmConsole_Tipo_Marcha.Show
    
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

Private Sub cmdIncluir_Click()
    Dim strIndice As Integer

    If txtCodigo_Equipamento.Text = Empty Or _
       dtcEquipamento.Text = Empty Then
       MsgBox "Dados de Equipamento inválidos. Verifique.", vbInformation, "Only Tech"
       txtCodigo_Equipamento.SetFocus
       Exit Sub
    End If
    
    intContador = 1
    Do While intContador <= hfgEquipamento.Rows - 1
        hfgEquipamento.Row = intContador
        hfgEquipamento.Col = 1
        If hfgEquipamento.Text = txtCodigo_Equipamento.Text And cmdIncluir.Caption = "Incluir" Then
           MsgBox "Equipamento já vinculado a este Cadastro.", vbInformation, "Only Tech"
           'Limpando os campos dos Itens
           txtCodigo_Equipamento.Text = Empty
           txtCodigo_Equipamento.SetFocus
           Exit Sub
        ElseIf hfgEquipamento.Text = txtCodigo_Equipamento.Text And cmdIncluir.Caption = "Alterar" Then
           If hfgEquipamento.Row <> intClique_Equipamento Then
              MsgBox "O Código de Equipamento alterado pertence a outro item neste Cadastro. Verifique.", vbInformation, "Only Tech"
              dtpValidade_qualificacao.SetFocus
              Exit Sub
           End If
        End If
        intContador = intContador + 1
    Loop
    
    hfgEquipamento.Row = hfgEquipamento.TopRow
    If cmdIncluir.Caption = "Incluir" Then
       If hfgEquipamento.Text <> Empty Then
          strIndice = intContador
          hfgEquipamento.Rows = hfgEquipamento.Rows + 1
       Else
          strIndice = intContador - 1
       End If
    Else
       strIndice = intClique_Equipamento
    End If
    
    hfgEquipamento.Row = strIndice
    
    hfgEquipamento.Col = 0
    hfgEquipamento.ColWidth(0) = 500
    hfgEquipamento.Font.Name = "Tahoma"
    hfgEquipamento.CellFontSize = 7
    hfgEquipamento.CellFontBold = False
    hfgEquipamento.CellBackColor = &H80FFFF
    hfgEquipamento.Text = strIndice
    
    hfgEquipamento.Col = 1
    hfgEquipamento.Text = txtCodigo_Equipamento.Text
    
    hfgEquipamento.Col = 2
    hfgEquipamento.Text = dtcEquipamento.Text
    
    cmdIncluir.Caption = "Incluir"
       
    txtCodigo_Equipamento.Text = Empty
    txtCodigo_Equipamento.SetFocus
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub cmdRemover_Click()
    If hfgEquipamento.Text = Empty Then
       MsgBox "Não há Equipamento selecionado.", vbInformation, "Only Tech"
       txtCodigo_Equipamento.SetFocus
       Exit Sub
    End If
    
    cmdIncluir.Caption = "Incluir"
    
    If hfgEquipamento.Rows <= 2 Then
       hfgEquipamento.Clear
       Movimentacoes.Monta_HFlex_Grid hfgEquipamento, "1000,5000", "Código,Descrição", 2, "Otica", Me
    Else
       hfgEquipamento.RemoveItem (hfgEquipamento.Row)
       hfgEquipamento.Col = 0
       intContador = 1
       Do While intContador <= hfgEquipamento.Rows - 1
          hfgEquipamento.Row = intContador
          hfgEquipamento.Text = intContador
          intContador = intContador + 1
       Loop
    End If

    hfgEquipamento.Refresh
    txtCodigo_Equipamento.Text = Empty
    txtCodigo_Equipamento.SetFocus
    hfgEquipamento.Col = 0
    hfgEquipamento.Row = 0
End Sub

Private Sub dtcEquipamento_GotFocus()
   If txtCodigo_Equipamento.Text = Empty Then
      Call Movimentacoes.Verifica_DataCombo(dtcEquipamento.Text)
   End If
End Sub

Private Sub dtcEquipamento_LostFocus()
   txtCodigo_Equipamento.Text = dtcEquipamento.BoundText
   If IsNumeric(txtCodigo_Equipamento.Text) = False Or dtcEquipamento.Text = Empty Then txtCodigo_Equipamento.Text = Empty: Exit Sub
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
    log.Programa = "Cadastro de Tipo Marcha"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstTipo_Marcha, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o cadastro de Tipo Marcha"
    
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    sstTipo_Marcha.TabEnabled(0) = False
    sstTipo_Marcha.TabEnabled(1) = False
    sstTipo_Marcha.Tab = 2
    
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
    log.Descricao = "Finalizando o cadastro de Tipo Marcha"
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
    If hfgEquipamento.Col = 0 And hfgEquipamento.Text <> Empty And hfgEquipamento.Row <> intClique_Equipamento Then
        txtCodigo_Equipamento.Text = Empty
        dtcEquipamento.Text = Empty
        cmdIncluir.Caption = "Incluir"
    End If
End Sub

Private Sub hfgEquipamento_DblClick()
    If hfgEquipamento.Col = 0 And hfgEquipamento.Text <> Empty Then
       intClique_Equipamento = hfgEquipamento.Row
       cmdIncluir.Caption = "Alterar"
       txtCodigo_Equipamento.Text = hfgEquipamento.TextArray((hfgEquipamento.Row * hfgEquipamento.Cols + hfgEquipamento.Col + 1))
       dtcEquipamento.Text = hfgEquipamento.TextArray((hfgEquipamento.Row * hfgEquipamento.Cols + hfgEquipamento.Col + 2))
    End If
    hfgEquipamento.SetFocus
End Sub

Private Sub hfgTipo_Marcha_Click()

   If hfgTipo_Marcha.Col = 0 And hfgTipo_Marcha.Text <> Empty Then
     
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
       
       txtCodigo.Text = hfgTipo_Marcha.TextArray((hfgTipo_Marcha.Row * hfgTipo_Marcha.Cols + hfgTipo_Marcha.Col + 1))
       txtDescricao.Text = hfgTipo_Marcha.TextArray((hfgTipo_Marcha.Row * hfgTipo_Marcha.Cols + hfgTipo_Marcha.Col + 2))
       txtDescricao_resumida.Text = hfgTipo_Marcha.TextArray((hfgTipo_Marcha.Row * hfgTipo_Marcha.Cols + hfgTipo_Marcha.Col + 3))
       txtNumero_Sequencial.Text = hfgTipo_Marcha.TextArray((hfgTipo_Marcha.Row * hfgTipo_Marcha.Cols + hfgTipo_Marcha.Col + 4))
       
       'EQUIPAMENTOS
       strSql = "SELECT FKCodigo_TBEquipamento_laboratorio," & _
                "DFDescricao_TBEquipamento_laboratorio " & _
                "FROM TBEquipamento_tipo_marcha " & _
                "INNER JOIN TBEquipamento_laboratorio " & _
                "ON TBEquipamento_tipo_marcha.FKCodigo_TBEquipamento_laboratorio = TBEquipamento_laboratorio.PKCodigo_TBEquipamento_laboratorio " & _
                "WHERE FKCodigo_TBTipo_marcha = " & txtCodigo.Text & ""
                
       Movimentacoes.Movimenta_HFlex_Grid strSql, hfgEquipamento, "1000,5000", "Código,Descrição", "BDRetaguarda", "Otica", Me
       
       hfgEquipamento.Col = 0
       hfgEquipamento.Row = 1
       If hfgEquipamento.Text = Empty Then
          hfgEquipamento.Rows = 2
          Movimentacoes.Monta_HFlex_Grid hfgEquipamento, "1000,5000", "Código,Descrição", 2, "Otica", Me
       End If
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstTipo_Marcha.TabEnabled(0) = True
       sstTipo_Marcha.TabEnabled(1) = True
       sstTipo_Marcha.Tab = 0
       
       txtCodigo.Enabled = False
                
   End If
   
   Unload frmAguarde
   
End Sub

Private Sub hfgTipo_Marcha_DblClick()
    hfgTipo_Marcha.Sort = 1
End Sub

Private Sub hfgTipo_Marcha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgTipo_Marcha_Click
    End If
End Sub

Private Sub sstTipo_Marcha_Click(PreviousTab As Integer)
   If sstTipo_Marcha.Tab = 0 Then
      txtDescricao.SetFocus
   ElseIf sstTipo_Marcha.Tab = 1 Then
      txtCodigo_Equipamento.SetFocus
   ElseIf sstTipo_Marcha.Tab = 2 Then
      If frmIntegracao.Visible = True Then
         Unload frmIntegracao
      End If
      If strCombo <> Empty And strCombo <> "Todos" Then
         cbbCampos.Text = strCombo
         txtConsulta.SetFocus
      ElseIf strCombo = "Todos" Then
         hfgTipo_Marcha.Row = 1
         hfgTipo_Marcha.Col = 0
         hfgTipo_Marcha.SetFocus
     End If
   End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstTipo_Marcha.Tab <> 2: Call Gravar
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
       MsgBox "O campo Descrição não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtDescricao.SetFocus
       Exit Function
    ElseIf txtDescricao_resumida.Text = Empty Then
       MsgBox "O campo Descrição Resumida não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtDescricao_resumida.SetFocus
       Exit Function
    ElseIf txtNumero_Sequencial.Text = Empty Then
       MsgBox "O campo Número Sequencial não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtNumero_Sequencial.SetFocus
       Exit Function
    End If
    
    Dim strCodigo_Equipamento As String
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim intEmpresa As Integer
    Dim strCodigo_Tipo_Marcha As String
    Dim strProx_Cod_Tipo_Marcha As String
    Dim rstVerifica_Titulo As New ADODB.Recordset
    
    Call Objetos.Maiusculo_TXT(Me)

    'buscando as informacoes do parametro
    If booAlterar = False Then
       intEmpresa = MDIPrincipal.OCXUsuario.Empresa
       strProx_Cod_Tipo_Marcha = Funcoes_Gerais.Localiza_Proximo_Codigo("DFProximo_tipo_marcha_TBParametros_servicos", "FKCodigo_TBEmpresa", intEmpresa, "TBParametros_servicos", "Otica", Me, "BDRetaguarda")
       txtCodigo.Text = strProx_Cod_Tipo_Marcha
    End If
    
    strCampo = "PKCodigo_TBTipo_marcha,DFDescricao_TBTipo_marcha," & _
               "DFDescricao_resumida_TBTipo_marcha,DFNumero_sequencial_TBTipo_marcha," & _
               "DFData_alteracao_TBTipo_marcha,DFIntegrado_filiais_TBTipo_marcha "
                
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBTipo_marcha "
    End If
                                  
    strValores = "" & txtCodigo.Text & ",'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtDescricao_resumida.Text) & "'," & txtNumero_Sequencial.Text & "," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0"
                 
    If booIntegra_Portal = True Then
        strValores = strValores & ",0"
    End If
    
    'INDICANDO O BANCO A CONECTAR-SE
    conexao.Initial_Catalog = "BDRetaguarda"

    'ESTABELECENDO CONEXÃO COM O BANCO
    conexao.Abrir_conexao ("Otica")

    'INDICA O INICIO DA TRANSAÇÃO JUNTO O BANCO
    conexao.CNConexao.BeginTrans
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       
       strSet = "UPDATE TBTipo_marcha " & _
                "SET DFDescricao_TBTipo_marcha = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'," & _
                "DFDescricao_resumida_TBTipo_marcha = '" & Funcoes_Gerais.Grava_String(txtDescricao_resumida.Text) & "'," & _
                "DFNumero_sequencial_TBTipo_marcha = " & txtNumero_Sequencial.Text & "," & _
                "DFData_alteracao_TBTipo_marcha = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBTipo_marcha = 0 "
                
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBTipo_marcha = 0 "
       End If
               
       strSet = strSet & "WHERE PKCodigo_TBTipo_marcha = " & txtCodigo.Text & ""
                
       conexao.CNConexao.Execute strSet
       
       strSql = "DELETE FROM TBEquipamento_tipo_marcha WHERE FKCodigo_TBTipo_marcha = " & txtCodigo.Text & ""
       conexao.CNConexao.Execute strSql
       
       hfgEquipamento.Col = 0
       hfgEquipamento.Row = 1
       If hfgEquipamento.Text <> Empty Then
          intContador = 1
          hfgEquipamento.Col = 1
          Do While intContador <= hfgEquipamento.Rows - 1
             hfgEquipamento.Row = intContador

             strSql = "INSERT INTO TBEquipamento_tipo_marcha (FKCodigo_TBEquipamento_laboratorio," & _
                      "FKCodigo_TBTipo_marcha,DFData_alteracao_TBEquipamento_tipo_marcha," & _
                      "DFIntegrado_filiais_TBEquipamento_tipo_marcha"
                      
             If booIntegra_Portal = True Then
                strSql = strSql & ",DFIntegrado_portal_TBEquipamento_tipo_marcha) "
             Else
                strSql = ") "
             End If
            
             strSql = strSql & "VALUES (" & hfgEquipamento.Text & "," & _
                      "" & txtCodigo.Text & ",'" & Format(Date, "YYYYMMDD") & "',0 "
            
             If booIntegra_Portal = True Then
                strSql = strSql & ",0)"
             Else
                strSql = strSql & ")"
             End If
             
             conexao.CNConexao.Execute strSql
             
             intContador = intContador + 1
          Loop
       End If
       
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"

       strSql = "INSERT INTO TBTipo_marcha (" & strCampo & ") " & _
                "VALUES (" & strValores & ")"
       
       conexao.CNConexao.Execute strSql

       hfgEquipamento.Col = 0
       hfgEquipamento.Row = 1
       If hfgEquipamento.Text <> Empty Then
          intContador = 1
          hfgEquipamento.Col = 1
          Do While intContador <= hfgEquipamento.Rows - 1
             hfgEquipamento.Row = intContador
             
             strSql = "INSERT INTO TBEquipamento_tipo_marcha (FKCodigo_TBEquipamento_laboratorio," & _
                      "FKCodigo_TBTipo_marcha,DFData_alteracao_TBEquipamento_tipo_marcha," & _
                      "DFIntegrado_filiais_TBEquipamento_tipo_marcha"
            
             If booIntegra_Portal = True Then
                strSql = strSql & ",DFIntegrado_portal_TBEquipamento_tipo_marcha) "
             Else
                strSql = strSql & ") "
             End If
            
             strSql = strSql & "VALUES (" & hfgEquipamento.Text & "," & _
                      "" & txtCodigo.Text & ",'" & Format(Date, "YYYYMMDD") & "',0"
                      
             If booIntegra_Portal = True Then
                strSql = strSql & ",0 ) "
             Else
                strSql = strSql & " ) "
             End If
             
             conexao.CNConexao.Execute strSql
             
             intContador = intContador + 1
          Loop
       End If
       
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
       
       ''''' aqui começa a ATUALIZAÇÃO DA TABELA TBParametros_servicos '''''
       'Somente para mostrar ao usuario o código que o registro foi incluido
       strCodigo_Tipo_Marcha = strProx_Cod_Tipo_Marcha
       
       If strCodigo_Tipo_Marcha <> Empty Then
          MsgBox "** O código dessa Tipo Marcha é: " & strCodigo_Tipo_Marcha & "", vbOKOnly, "Only Tech"
       End If
       
       strProx_Cod_Tipo_Marcha = strProx_Cod_Tipo_Marcha + 1
       
       strSet = "UPDATE TBParametros_servicos " & _
                "SET DFProximo_tipo_marcha_TBParametros_servicos = " & strProx_Cod_Tipo_Marcha & "," & _
                "DFData_alteracao_TBParametros_servicos = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBParametros_servicos = 0 "
                
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBParametros_servicos = 0 "
       End If
       
       strSet = strSet & "WHERE FKCodigo_TBEmpresa  = " & intEmpresa & ""
                
       conexao.CNConexao.Execute strSet
    End If
    
    'COMITANDO A TRANSACAO
    conexao.CNConexao.CommitTrans

    'FECHANDO A CONEXÃO
    conexao.CNConexao.Close
    
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
       Me.hfgTipo_Marcha.Visible = False
    End If
    
    sstTipo_Marcha.TabEnabled(0) = False
    sstTipo_Marcha.TabEnabled(1) = False
    sstTipo_Marcha.Tab = 2
    
    Exit Function
    
Erro:
    If conexao.CNConexao.State <> adStateClosed Then
       conexao.CNConexao.RollbackTrans
       conexao.Fechar_conexao
    End If
    
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    'VERIFICACAO PARA EXCLUSAO
'    strSql = "SELECT PKCodigo_TBInsumo FROM TBInsumo WHERE FKCodigo_TBTipo_marcha = " & txtCodigo.Text & ""
'
'    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'    If rstAplicacao.RecordCount <> 0 Then
'       MsgBox "Esta função está vinculada ao Insumo de código " & rstAplicacao.Fields("PKCodigo_TBInsumo") & " e não poderá ser excluída. Verifique.", vbInformation, "Only Tech"
'       Set rstAplicacao = Nothing
'       Exit Function
'    End If
'    Set rstAplicacao = Nothing
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'abrindo conexao
    conexao.Initial_Catalog = "BDRetaguarda"
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    'Excluindo Registro filho
    strSql = "DELETE FROM TBEquipamento_tipo_marcha WHERE FKCodigo_TBTipo_marcha = '" & txtCodigo.Text & "'"
    
    conexao.CNConexao.Execute strSql
    
    'Excluindo Registro Principal
    strSql = "DELETE FROM TBTipo_marcha WHERE PKCodigo_TBTipo_marcha = '" & txtCodigo.Text & "'"
    
    conexao.CNConexao.Execute strSql
    
    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
    
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
       hfgTipo_Marcha.Visible = False
    End If
           
    sstTipo_Marcha.TabEnabled(0) = False
    sstTipo_Marcha.TabEnabled(1) = False
    sstTipo_Marcha.Tab = 2
    
    Exit Function
Erro:
    If conexao.CNConexao.State <> adStateClosed Then
       conexao.CNConexao.RollbackTrans
       conexao.Fechar_conexao
    End If
    
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
       hfgTipo_Marcha.Visible = False
    End If
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstTipo_Marcha.TabEnabled(0) = False
    sstTipo_Marcha.TabEnabled(1) = False
    sstTipo_Marcha.Tab = 2
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
    Dim rstBusca_Parametro As New ADODB.Recordset
    Dim strCodigo_Tipo_Marcha As String

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
        
    strCodigo_Tipo_Marcha = rstBusca_Parametro.Fields("DFProximo_tipo_marcha_TBParametros_servicos")
    Set rstBusca_Parametro = Nothing
        
    strSql = Empty
    strSql = "SELECT * FROM TBTipo_marcha WHERE TBTipo_marcha.PKCodigo_TBTipo_marcha = " & strCodigo_Tipo_Marcha & ""
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstBusca_Parametro, "Otica", Me)
    
    If rstBusca_Parametro.RecordCount <> 0 Then
       MsgBox "O Código " & strCodigo_Tipo_Marcha & " já existe, por favor, verifique o cadastro parâmetros de serviços e atualize o código da próxima Tipo Marcha.", vbInformation, "Only Tech"
       Set rstBusca_Parametro = Nothing
       Call Objetos.Limpa_TXT(Me)
       sstTipo_Marcha.TabEnabled(0) = False
       sstTipo_Marcha.TabEnabled(1) = False
       sstTipo_Marcha.Tab = 2
       Exit Function
    End If
    Set rstBusca_Parametro = Nothing
    
    'limpando o grid
    hfgEquipamento.Rows = 2
    Movimentacoes.Monta_HFlex_Grid hfgEquipamento, "1000,5000", "Código,Descrição", 2, "Otica", Me

    sstTipo_Marcha.TabEnabled(0) = True
    sstTipo_Marcha.TabEnabled(1) = True
    sstTipo_Marcha.Tab = 0
                    
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

Private Sub txtCodigo_Equipamento_Change()
    dtcEquipamento.BoundText = txtCodigo_Equipamento.Text
    If IsNumeric(txtCodigo_Equipamento.Text) = False Then txtCodigo_Equipamento.Text = Empty: Exit Sub
End Sub

Private Sub txtCodigo_Equipamento_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_Equipamento_LostFocus()
    If dtcEquipamento.Text = Empty Then txtCodigo_Equipamento.Text = Empty
End Sub

Private Sub txtCodigo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_LostFocus()
    If txtCodigo.Text <> Empty And booAlterar = False Then
        Movimentacoes.Verifica_Numero "PKCodigo_TBTipo_marcha", "TBTipo_marcha", txtCodigo, "Otica", Me
    End If
End Sub

Private Function Reposicao()
    On Error GoTo Erro
    
    strNomes = "Código,Descrição,Desc. Resumida,Nº Sequencial"
    strTamanho = "1000,3500,1800,1500"
    
    Movimentacoes.Monta_HFlex_Grid hfgTipo_Marcha, strTamanho, strNomes, 4, "Otica", Me
        
    Movimentacoes.Monta_HFlex_Grid hfgEquipamento, "1000,5000", "Código,Descrição", 2, "Otica", Me
    
    strSql = "SELECT PKCodigo_TBEquipamento_laboratorio,DFDescricao_TBEquipamento_laboratorio FROM TBEquipamento_laboratorio"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEquipamento_laboratorio", "DFDescricao_TBEquipamento_laboratorio", dtcEquipamento, strSql, "BDRetaguarda", "Otica", Me

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

Private Sub txtDescricao_resumida_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDescricao_resumida_LostFocus()
    txtDescricao_resumida.Text = UCase(txtDescricao_resumida.Text)
End Sub

Private Sub txtNumero_sequencial_GotFocus()
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
    
    strSql = "SELECT PKCodigo_TBTipo_marcha,DFDescricao_TBTipo_marcha," & _
             "DFDescricao_resumida_TBTipo_marcha,DFNumero_sequencial_TBTipo_marcha " & _
             "FROM TBTipo_marcha "
            
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    Funcoes_Gerais.Grava_String (txtConsulta.Text)
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE TBTipo_marcha.PKCodigo_TBTipo_marcha = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Descrição" Then
          strSql = strSql & " WHERE DFDescricao_TBTipo_marcha LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Desc. Resumida" Then
          strSql = strSql & " WHERE DFDescricao_resumida_TBTipo_marcha LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Nº Sequencial" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE DFNumero_sequencial_TBTipo_marcha = '" & txtConsulta.Text & "' "
       End If
    End If

    frmAguarde.Show
    DoEvents
            
    strSql = strSql & " ORDER BY TBTipo_marcha.PKCodigo_TBTipo_marcha"
        
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgTipo_Marcha, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    hfgTipo_Marcha.Col = 0
    hfgTipo_Marcha.Row = 1
    If hfgTipo_Marcha.Text = Empty Then
       hfgTipo_Marcha.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgTipo_Marcha, strTamanho, strNomes, 4, "Otica", Me
    End If
    
    Unload frmAguarde
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código")
    cbbCampos.AddItem ("Descrição")
    cbbCampos.AddItem ("Desc. Resumida")
    cbbCampos.AddItem ("Nº Sequencial")
End Function

Private Sub txtNumero_Sequencial_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo_TBTipo_marcha", txtCodigo.Text, "DFIntegrado_filiais_TBTipo_marcha", "TBTipo_marcha", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBTipo_marcha", Me.Top, Me.Left, Me.Width, Me.Height, "Tipo Marcha")
    
End Function
