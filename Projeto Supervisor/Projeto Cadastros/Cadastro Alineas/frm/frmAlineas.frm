VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAlineas 
   Caption         =   "Alineas"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlineas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4605
      _ExtentX        =   8123
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
            Key             =   "alt + g"
            Description     =   "Confirmar"
            Object.ToolTipText     =   "Gravar registro - CTRL+G"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alt + c"
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alt + e"
            Description     =   "Excluir"
            Object.ToolTipText     =   "Excluir registro - CTRL+E"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alt + i"
            Description     =   "Imprimir"
            Object.ToolTipText     =   "Imprimir - CTRL+I"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alt + s"
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
      Left            =   4620
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
            Picture         =   "frmAlineas.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlineas.frx":0FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlineas.frx":12FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlineas.frx":1698
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlineas.frx":1A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlineas.frx":1D4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sstAlineas 
      Height          =   2235
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   3942
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483638
      OLEDropMode     =   1
      MouseIcon       =   "frmAlineas.frx":2066
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "frmAlineas.frx":2082
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCodigo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Listagem"
      TabPicture(1)   =   "frmAlineas.frx":209E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtConsulta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgAlineas"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgAlineas 
         Height          =   945
         Left            =   120
         TabIndex        =   8
         Top             =   1170
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   1667
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox txtConsulta 
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   4365
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   1500
         Width           =   4335
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   -74880
         TabIndex        =   6
         Top             =   1260
         Width           =   825
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Picture         =   "frmAlineas.frx":20BA
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   -74880
         TabIndex        =   5
         Top             =   600
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmAlineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Cadastrar de Alineas                                           '
' Data de Criação........: 30/04/2003                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião                        '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim conexao As New DLLConexao_Sistema.conexao
Dim log As New DLLSystemManager.log

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 78: Call Novo     'CTRL+N
                       Case 71: Call Gravar   'CTRL+G
                       Case 67: Call Cancelar 'CTRL+C
                       Case 69: Call Excluir  'CTRL+E
                       Case 83: Unload Me     'CTRL+S
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
   
    'Informações constantes para o log
    
    log.Data = Date
    'Ver
'   strEstacao_log = MDIPrincipal_Cadastro_Base.strEstação
'   strUsuario_log = MDIPrincipal_Cadastro_Base.UsuárioOCX.NomeReduzido
    log.Estacao = "INFO-888"
    log.Usuario = "Adão"
    log.Programa = "Cadastro de Alineas"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando o cadastro de alineas"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando o log
    log.Gravar_log "PDV", Me
    
    sstAlineas.Tab = 1
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = True
 
    Call Reposicao
    
    Exit Sub

Erro:
    Call Erro.Erro(Me, "PDV", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    Set log = Nothing
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "PDV", "Unload")
    Exit Sub
    
End Sub

Private Sub hfgAlineas_Click()

    If hfgAlineas.Col > 0 Then
        strCampo_consulta = hfgAlineas.DataField(0, hfgAlineas.ColSel - 1)
        txtConsulta.SetFocus
    End If
    If hfgAlineas.Col = 0 Then
        tlbBotoes.Buttons.Item(1).Enabled = False
        tlbBotoes.Buttons.Item(2).Enabled = True
        tlbBotoes.Buttons.Item(3).Enabled = True
        tlbBotoes.Buttons.Item(4).Enabled = True
        tlbBotoes.Buttons.Item(5).Enabled = False
        
        On Error Resume Next
        txtCodigo.Text = hfgAlineas.TextArray((hfgAlineas.Row * hfgAlineas.Cols + hfgAlineas.Col + 1))
        txtDescricao.Text = hfgAlineas.TextArray((hfgAlineas.Row * hfgAlineas.Cols + hfgAlineas.Col + 2))
    
        booAlterar = True
        txtConsulta.Text = Empty
        sstAlineas.Tab = 0
        Me.txtDescricao.SetFocus
   End If
End Sub

Private Sub tlbbotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           'Case 5: Call Imprimir
           Case 7: Unload Me
    End Select
End Sub

Function Gravar()

    On Error GoTo Erro
    
    Dim strSet As String
    Dim strCampo As String
    Dim strvalores As String
    
    strCampo = "PKCodigo_TBAlineas,DFDescricao_TBAlineas"
    strvalores = " " & txtCodigo & " , '" & txtDescricao.Text & "' "
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFDescricao_TBAlineas = '" & Me.txtDescricao.Text & "' "
       Call funcoes_banco.Alterar("TBAlineas", strSet, "PKCodigo_TBAlineas", txtCodigo.Text, "PDV", Me)
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "PDV", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBAlineas", strCampo, strvalores, "PDV", Me)
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "PDV", Me
    End If
    
    Call Limpar_Campos
    Call Reposicao
    
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    Exit Function
    
Erro:
    Call Erro.Erro(Me, "PDV", "Gravar")
    Exit Function
        
End Function

Private Function Excluir()

    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "PDV", Me
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBAlineas", "PKCodigo_TBAlineas", Me.txtCodigo.Text, "PDV", Me, "BDSupervisor")
    
    Call Limpar_Campos
    
    Call Reposicao
    
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = True
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    Call Reposicao
    
    Exit Function
    
Erro:

    Call Erro.Erro(Me, "PDV", "Excluir")
    Exit Function

End Function
Private Function Cancelar()
    On Error GoTo Erro
    Call Limpar_Campos
    'Inserir log
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = True
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "PDV", "Cancelar")
    Exit Function
End Function
Private Function Limpar_Campos()

    txtCodigo.Text = Empty
    txtDescricao.Text = Empty
    
End Function
Private Function Novo()
    On Error GoTo Erro
    
    sstAlineas.Tab = 0
    Call Limpar_Campos
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    tlbBotoes.Buttons.Item(1).Enabled = False
    tlbBotoes.Buttons.Item(2).Enabled = True
    tlbBotoes.Buttons.Item(3).Enabled = True
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = False
    txtCodigo.Enabled = True
    txtCodigo.SetFocus
    booAlterar = False
    Exit Function
Erro:
    Call Erro.Erro(Me, "PDV", "Novo")
    Exit Function
End Function

Private Sub txtCodigo_LostFocus()
    Movimentacoes.Verifica_Numero "PKCodigo_TBAlineas", "TBAlineas", txtCodigo, "PDV", Me
End Sub

Private Function Reposicao()
    On Error GoTo Erro

    Dim strSQL As String

    strSQL = "SELECT * FROM TBalineas"
    
    If txtConsulta.Text <> Empty Then
        strSQL = strSQL & " WHERE " & strCampo_consulta & " LIKE '" & txtConsulta.Text & "%' "
    End If

    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgAlineas, "800,3100", "Código,Descrição", "BDSupervisor", "PDV", Me

    Exit Function

Erro:
    Call Erro.Erro(Me, "PDV", "Reposicao")
    Resume Next

End Function

Private Sub txtConsulta_Change()
   Call Reposicao
End Sub
