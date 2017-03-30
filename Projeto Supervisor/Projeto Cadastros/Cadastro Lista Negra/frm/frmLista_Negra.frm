VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLista_Negra 
   Caption         =   "Lista Negra"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLista_Negra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
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
      Left            =   6480
      Top             =   300
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
            Picture         =   "frmLista_Negra.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Negra.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Negra.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Negra.frx":09DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Negra.frx":0D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Negra.frx":108E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sstLista_Negra 
      Height          =   2295
      Left            =   0
      TabIndex        =   4
      Top             =   330
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      OLEDropMode     =   1
      MouseIcon       =   "frmLista_Negra.frx":13A8
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
      TabPicture(0)   =   "frmLista_Negra.frx":13C4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCpfCnpj"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCodigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraPessoa"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCpfCnpj"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dtcAlineas"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Listagem"
      TabPicture(1)   =   "frmLista_Negra.frx":13E0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtConsulta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgLista_Negra"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgLista_Negra 
         Height          =   1005
         Left            =   -74880
         TabIndex        =   13
         Top             =   1140
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   1773
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo dtcAlineas 
         Height          =   360
         Left            =   1620
         TabIndex        =   11
         Top             =   870
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   635
         _Version        =   393216
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
      Begin VB.TextBox txtCpfCnpj 
         Height          =   375
         Left            =   2190
         TabIndex        =   2
         Top             =   1680
         Width           =   2835
      End
      Begin VB.Frame fraPessoa 
         Caption         =   "Pessoa"
         Height          =   705
         Left            =   120
         TabIndex        =   9
         Top             =   1350
         Width           =   1965
         Begin VB.OptionButton optJuridica 
            Caption         =   "Jurídica"
            Height          =   345
            Left            =   960
            TabIndex        =   1
            Top             =   270
            Width           =   975
         End
         Begin VB.OptionButton optFisica 
            Caption         =   "Física"
            Height          =   345
            Left            =   120
            TabIndex        =   0
            Top             =   270
            Width           =   795
         End
      End
      Begin VB.TextBox txtConsulta 
         Height          =   375
         Left            =   -74880
         TabIndex        =   6
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtCodigo 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   870
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alinea"
         Height          =   240
         Left            =   1620
         TabIndex        =   12
         Top             =   630
         Width           =   525
      End
      Begin VB.Label lblCpfCnpj 
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         Height          =   240
         Left            =   2220
         TabIndex        =   10
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   435
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Picture         =   "frmLista_Negra.frx":13FC
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   630
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmLista_Negra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Cadastrar de Lista Negra                                       '
' Data de Criação........: 30/04/2003                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião                        '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim strID As String
Dim booAlterar As Boolean
Dim conexao As DLLConexao_Sistema.conexao
Public log As New DLLSystemManager.log

Private Sub hfgLista_negra_Click()
    
    If hfgLista_Negra.Col > 0 Then
        strCampo_consulta = hfgLista_Negra.DataField(0, hfgLista_Negra.ColSel - 1)
        txtConsulta.SetFocus
    End If
    If hfgLista_Negra.Col = 0 Then
    
       On Error Resume Next

       tlbBotoes.Buttons.Item(1).Enabled = False
       tlbBotoes.Buttons.Item(2).Enabled = True
       tlbBotoes.Buttons.Item(3).Enabled = True
       tlbBotoes.Buttons.Item(4).Enabled = True
       tlbBotoes.Buttons.Item(5).Enabled = False
       
       strID = hfgLista_Negra.TextArray((hfgLista_Negra.Row * hfgLista_Negra.Cols + hfgLista_Negra.Col + 1))
       txtCodigo.Text = hfgLista_Negra.TextArray((hfgLista_Negra.Row * hfgLista_Negra.Cols + hfgLista_Negra.Col + 2))
       dtcAlineas.Text = hfgLista_Negra.TextArray((hfgLista_Negra.Row * hfgLista_Negra.Cols + hfgLista_Negra.Col + 3))
    
       If hfgLista_Negra.TextArray((hfgLista_Negra.Row * hfgLista_Negra.Cols + hfgLista_Negra.Col + 4)) = "Não" Then
          optFisica.Value = True
       Else
          optJuridica.Value = True
       End If
    
       txtCpfCnpj.Text = hfgLista_Negra.TextArray((hfgLista_Negra.Row * hfgLista_Negra.Cols + hfgLista_Negra.Col + 5))
           
       booAlterar = True
       txtConsulta.Text = Empty
       sstLista_Negra.Tab = 0
       txtCodigo_interno.SetFocus
    End If
    
End Sub

Private Sub dtcAlineas_Click(Area As Integer)
    
    txtCodigo.Text = dtcAlineas.BoundText
    
End Sub

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
    If KeyAscii = 13 Then
        KeyAscii = 0
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
    log.Programa = "Cadastro de Lista Negra"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando o cadastro de Lista Negra"
    log.Tipo = 1
    
    log.Gravar_log "PDV", Me
    
    sstLista_Negra.Tab = 1
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    Call Reposicao
    
    Exit Sub
    
Erro:

    log.Evento = "Load"
    log.Tipo = 3
    log.Descricao = Err.Description
    
    log.Gravar_log "PDV", Me
    
    Call Erro.Erro(Me, "PDV", "Load")
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Erro
    
    log.Evento = "Unload"
    
    log.Gravar_log "PDV", Me
    
    Exit Sub
    
Erro:

    log.Evento = "Unload"
    log.Tipo = 3
    log.Descricao = Err.Description
    
    log.Gravar_log "PDV", Me
    
    Call Erro.Erro(Me, "PDV", "Unload")
    
    Exit Sub

End Sub

Private Sub optFisica_Click()
    
    lblCpfCnpj.Caption = "CPF"
      
End Sub

Private Sub optJuridica_Click()

    lblCpfCnpj.Caption = "CNPJ"
    
End Sub

Private Sub txtCodigo_Change()
    
    dtcAlineas.BoundText = txtCodigo.Text
    
End Sub

Private Sub txtCpfCnpj_LostFocus()
    
    If Me.txtCpfCnpj.Text <> Empty Then
       If optFisica.Value = True Then
          Call CGC_CPF.FormatarCPF(txtCpfCnpj.Text, Me.txtCpfCnpj)
       Else
          Call CGC_CPF.FormatarCNPJ(txtCpfCnpj.Text, Me.txtCpfCnpj)
       End If
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
    Dim strPessoa As String
    
    If optFisica.Value = True Then
       strPessoa = 0
    Else
      If optJuridica.Value = True Then
         strPessoa = 1
      End If
    End If
    
    strCampo = "PKCodigo_TBAlineas,DFTipo_pessoa_TBLista_Negra,DFCnpj_TBLista_Negra"
    strvalores = " " & txtCodigo.Text & " , " & strPessoa & " , '" & txtCpfCnpj.Text & "' "
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET PKCodigo_TBAlineas = '" & txtCodigo.Text & "', " & _
                "    DFCnpj_TBLista_Negra = '" & Me.txtCpfCnpj.Text & "', " & _
                "    DFTipo_pessoa_TBLista_Negra =  " & strPessoa & " "
       Call funcoes_banco.Alterar("TBLista_Negra", strSet, "DFID_TBLista_Negra", strID, "PDV", Me, "BDSupervisor")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Gravar_log "PDV", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBLista_Negra", strCampo, strvalores, "PDV", Me, "VDSupervisor")
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Gravar_log "PDV", Me
    End If
    
    Call Objetos.Limpa_TXT(Me)
    Call Reposicao
    
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    Exit Function
    
Erro:

    log.Tipo = 3
    log.Descricao = Err.Description
    
    log.Gravar_log "PDV", Me
    Call Erro.Erro(Me, "PDV", "Gravar")
    Exit Function
    
End Function

Private Function Excluir()

    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    
    log.Gravar_log "PDV", Me
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBLista_Negra", "DFID_TBLista_Negra", strID, "PDV", Me, "BDSupervisor")
    
    Call Objetos.Limpa_TXT(Me)
    
    Call Reposicao
    
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    Call Reposicao
    
    Exit Function
    
Erro:
    log.Evento = "Excluir"
    log.Tipo = 3
    log.Descricao = Err.Description
    
    log.Gravar_log "PDV", Me
    Call Erro.Erro(Me, "PDV", "Excluir")
    Exit Function

End Function
Private Function Cancelar()

    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    
    'Inserir log
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Gravar_log "PDV", Me
    
    Exit Function
Erro:
    log.Evento = "Cancelar"
    log.Tipo = 3
    log.Descricao = Err.Description
    
    log.Gravar_log "PDV", Me
    Call Erro.Erro(Me, "PDV", "Cancelar")
    Exit Function

End Function

Private Function Novo()

    On Error GoTo Erro
    
    sstLista_Negra.Tab = 0
    
    Call Objetos.Limpa_TXT(Me)
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    
    log.Gravar_log "PDV", Me
    
    tlbBotoes.Buttons.Item(1).Enabled = False
    tlbBotoes.Buttons.Item(2).Enabled = True
    tlbBotoes.Buttons.Item(3).Enabled = True
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = False
    Me.optFisica.SetFocus
    booAlterar = False
    Exit Function
Erro:
    log.Evento = "Novo"
    log.Tipo = 3
    log.Descricao = Err.Description
    
    log.Gravar_log "PDV", Me
    
    Call Erro.Erro(Me, "PDV", "Novo")
    Exit Function

End Function

Private Sub txtCodigo_LostFocus()
    Movimentacoes.Verifica_Numero "PKCodigo_TBAlineas", "TBLista_Negra", txtCodigo, "PDV", Me
End Sub

Private Function Reposicao()

    On Error GoTo Erro

    Dim strSQL As String

    strSQL = "SELECT TBlista_Negra.DFid_TBLista_Negra, TBlista_Negra.PKCodigo_TBAlineas, TBAlineas.DFDescricao_TBAlineas, " & _
             "TBlista_Negra.DFTipo_pessoa_TBLista_Negra, TBlista_Negra.DFCnpj_TBLista_negra FROM TBLista_Negra " & _
             "INNER JOIN TBAlineas ON " & _
             "TBLista_Negra.PKCodigo_TBAlineas = TBAlineas.PKCodigo_TBAlineas"
             
    If txtConsulta.Text <> Empty Then
        strSQL = strSQL & " WHERE " & strCampo_consulta & " LIKE '" & txtConsulta.Text & "%' "
    End If

    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgLista_Negra, "0,800,2600,0,1600", "ID,Código,Alinea,Tipo,CPF/CNPJ", "BDSupervisor", "PDV", Me
     
    strSQL = "SELECT TBAlineas.PKCodigo_TBAlineas, TBAlineas.DFDescricao_TBAlineas FROM TBAlineas"
    Call Movimentacoes.Movimenta_DataCombo("PKCodigo_TBAlineas", "DFDescricao_TBAlineas", dtcAlineas, strSQL, "BDSupervisor", "PDV", Me)
    
    Exit Function

Erro:
    log.Evento = "Reposição"
    log.Tipo = 3
    log.Descricao = Err.Description
    log.Gravar_log "PDV", Me
    
    Call Erro.Erro(Me, "PDV", "Reposicao")
    Resume Next

End Function

Private Sub txtConsulta_Change()

   Call Reposicao
   
End Sub
