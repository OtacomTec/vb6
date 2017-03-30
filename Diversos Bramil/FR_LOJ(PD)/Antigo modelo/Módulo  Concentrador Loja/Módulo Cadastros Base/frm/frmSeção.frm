VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmSecao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seção"
   ClientHeight    =   2625
   ClientLeft      =   1830
   ClientTop       =   2040
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSeção.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4905
   Begin TabDlg.SSTab sstSecao 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   4048
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
      TabPicture(0)   =   "frmSeção.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtcImpressora"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDescricao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtImpressora"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmSeção.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cbbCampos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgSecao"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtConsulta"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdRefresh"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdConsulta"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdOrdenar"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtImpressora 
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CommandButton cmdOrdenar 
         Caption         =   "A"
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
         Left            =   -71370
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ordenar: (A) Alfabética/ (C) Código "
         Top             =   750
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
         Left            =   -70980
         Picture         =   "frmSeção.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Consultar"
         Top             =   750
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
         Left            =   -70590
         Picture         =   "frmSeção.frx":34B4
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   750
         Width           =   375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -73080
         TabIndex        =   8
         Top             =   750
         Width           =   1635
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSecao 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   7
         Top             =   1200
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   2355
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
         TabIndex        =   11
         Top             =   750
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
      Begin MSDataListLib.DataCombo dtcImpressora 
         Height          =   360
         Left            =   1260
         TabIndex        =   14
         Top             =   1500
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impressora"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número ECF"
         Height          =   240
         Left            =   1260
         TabIndex        =   4
         Top             =   600
         Width           =   1065
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
         TabIndex        =   3
         Top             =   600
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6060
      Top             =   330
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
            Picture         =   "frmSeção.frx":44F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeção.frx":4810
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeção.frx":4B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeção.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeção.frx":525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeção.frx":5578
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                         '
' Módulo.................: Compras                                                        '
' Objetivo...............: Cadastro de Seção                                              '
' Data de Criação........: 30/04/2003                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCombo As String
Dim strConsulta As String
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
Option Explicit

Function Imprimir()
    'Tratamento de erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       'cbbCampos.SetFocus
       Me.txtConsulta.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    frmConsole_Relatorio_secao.Show
        
    Unload frmAguarde
        
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Gravar")
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

Private Sub cmdOrdenar_Click()
    If cmdOrdenar.Caption = "A" Then
       cmdOrdenar.Caption = "C"
    Else
       cmdOrdenar.Caption = "A"
    End If
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
    On Error GoTo erro
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Seção"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstSecao, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando cadastro de Seção"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    sstSecao.TabEnabled(0) = False
    sstSecao.Tab = 1
        
    Call Reposicao
    
    Exit Sub
erro:
    Call erro.erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Finalizando cadastro de Seção"
        
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    strCombo = Empty
    
    Exit Sub
erro:
    Call erro.erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgSecao_Click()
    If hfgSecao.Col = 0 Then
        
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
           
       frmAguarde.Show
       DoEvents
       
       txtCodigo.Text = hfgSecao.TextArray((hfgSecao.Row * hfgSecao.Cols + hfgSecao.Col + 1))
       txtDescricao.Text = hfgSecao.TextArray((hfgSecao.Row * hfgSecao.Cols + hfgSecao.Col + 2))
            
       booAlterar = True
       txtConsulta.Text = Empty
       sstSecao.TabEnabled(0) = True
       sstSecao.Tab = 0
       Me.txtDescricao.SetFocus
   End If
   Unload frmAguarde
End Sub

Private Sub hfgSecao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgSecao_Click
    End If
End Sub

Private Sub sstSecao_Click(PreviousTab As Integer)
    If sstSecao.Tab = 0 Then
       txtDescricao.SetFocus
    ElseIf sstSecao.Tab = 1 Then
       If strCombo <> Empty And strCombo <> "Todos" Then
          cbbCampos.Text = strCombo
          txtConsulta.SetFocus
       ElseIf strCombo = "Todos" Then
          hfgSecao.Row = 1
          hfgSecao.Col = 0
          hfgSecao.SetFocus
       End If
    End If
End Sub

Private Sub tlbbotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
        End Select
End Sub

Function Gravar()
    On Error GoTo erro
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    
    If txtCodigo.Text = Empty Then
       MsgBox "Código não pode ser nulo.", vbInformation, "Only Tech"
       txtCodigo.SetFocus
       Exit Function
    End If
    
    Call Objetos.Maiusculo_TXT(Me)
    
    strCampo = "PKCodigo_TBSecao,DFDescricao_TBSecao"
    
    strValores = "" & txtCodigo & ",'" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & "'"
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFDescricao_TBSecao = '" & Funcoes_Gerais.Grava_String(txtDescricao.Text) & " '"
       Call funcoes_banco.Alterar("TBSecao", strSet, "PKCodigo_TBSecao", txtCodigo.Text, "OTICA", Me, "BDRetaguarda")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBSecao", strCampo, strValores, "OTICA", Me, "BDRetaguarda")
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
    
    If booPrivilegio_Consultar = False Then
      hfgSecao.Visible = False
    End If
    
    txtCodigo.Enabled = False
    sstSecao.TabEnabled(0) = False
    sstSecao.Tab = 1
    hfgSecao.Refresh
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo erro
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBSecao", "PKCodigo_TBSecao", Me.txtCodigo.Text, "OTICA", Me, "BDRetaguarda")
    
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
    
    If booPrivilegio_Consultar = False Then
       hfgSecao.Visible = False
    End If
            
    sstSecao.TabEnabled(0) = False
    sstSecao.Tab = 1
    
    Exit Function
erro:
     Call erro.erro(Me, "OTICA", "Excluir")
     Exit Function
End Function

Private Function Cancelar()
    On Error GoTo erro
    
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
    
    If booPrivilegio_Consultar = False Then
       hfgSecao.Visible = False
    End If
    
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Operação com Registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstSecao.TabEnabled(0) = False
    sstSecao.Tab = 1
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo erro
    
    Call Reposicao
    
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
    
    sstSecao.TabEnabled(0) = True
    sstSecao.Tab = 0
    
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    booAlterar = False
    
    Exit Function
erro:
    Call erro.erro(Me, "OTICA", "Novo")
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
       Movimentacoes.Verifica_Numero "PKCodigo_TBSecao", "TBSecao", txtCodigo, "OTICA", Me
    End If
End Sub

Private Function Reposicao()
    On Error GoTo erro
          
    Movimentacoes.Monta_HFlex_Grid hfgSecao, "1000,3100", "Código,Nome da Seção", 2, "OTICA", Me
    
    Call Monta_Combo
              
    hfgSecao.Refresh
    Exit Function
erro:
   Call erro.erro(Me, "OTICA", "Reposicao")
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
    
    strSql = "SELECT * FROM TBSecao"
        
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Código da Seção" Then
          strSql = strSql & " WHERE convert(nvarchar,PKCodigo_TBSecao) LIKE '" & txtConsulta.Text & "' "
       Else
          strSql = strSql & " WHERE convert(nvarchar,DFDescricao_TBSecao) LIKE '%" & txtConsulta.Text & "%' "
       End If
    End If
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgSecao, "1000,3100", "Código,Nome da Seção", "BDRetaguarda", "Otica", Me
           
    frmAguarde.Show
    DoEvents
    
    If cmdOrdenar.Caption = "C" Then
       strSql = strSql & " ORDER BY TBSecao.PKCodigo_TBSecao"
    ElseIf cmdOrdenar.Caption = "A" Then
       strSql = strSql & " ORDER BY TBSecao.DFDescricao_TBsecao"
    End If
        
    hfgSecao.Refresh
    
    Unload frmAguarde
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código da Seção")
    cbbCampos.AddItem ("Descrição da Seção")
End Function
