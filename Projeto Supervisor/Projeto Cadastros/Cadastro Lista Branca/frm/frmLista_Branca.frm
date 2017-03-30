VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLista_Branca 
   Caption         =   "Lista Branca"
   ClientHeight    =   2640
   ClientLeft      =   225
   ClientTop       =   615
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLista_Branca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5190
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5190
      _ExtentX        =   9155
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
            Picture         =   "frmLista_Branca.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Branca.frx":0FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Branca.frx":12FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Branca.frx":1698
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Branca.frx":1A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista_Branca.frx":1D4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sstLista_Branca 
      Height          =   2295
      Left            =   0
      TabIndex        =   5
      Top             =   330
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      OLEDropMode     =   1
      MouseIcon       =   "frmLista_Branca.frx":2066
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
      TabPicture(0)   =   "frmLista_Branca.frx":2082
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCpfCnpj"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCodigo_interno"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtLimite_credito"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraPessoa"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCpfCnpj"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Listagem"
      TabPicture(1)   =   "frmLista_Branca.frx":209E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtConsulta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgLista"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgLista 
         Height          =   1005
         Left            =   120
         TabIndex        =   13
         Top             =   1170
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1773
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtCpfCnpj 
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Frame fraPessoa 
         Caption         =   "Pessoa"
         Height          =   705
         Left            =   -74880
         TabIndex        =   11
         Top             =   540
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
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtLimite_credito 
         Height          =   375
         Left            =   -71490
         TabIndex        =   4
         Top             =   1560
         Width           =   1545
      End
      Begin VB.TextBox txtCodigo_interno 
         Height          =   375
         Left            =   -72990
         TabIndex        =   3
         Top             =   1560
         Width           =   1425
      End
      Begin VB.Label lblCpfCnpj 
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         Height          =   240
         Left            =   -74880
         TabIndex        =   12
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Limite de Crédito"
         Height          =   240
         Left            =   -71490
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Picture         =   "frmLista_Branca.frx":20BA
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Interno"
         Height          =   240
         Left            =   -72960
         TabIndex        =   8
         Top             =   1320
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmLista_Branca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Cadastrar de Lista Branca                                      '
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

Private Sub hfgLista_Click()
    
    If hfgLista.Col > 0 Then
        strCampo_consulta = hfgLista.DataField(0, hfgLista.ColSel - 1)
        txtConsulta.SetFocus
    End If
    If hfgLista.Col = 0 Then
    
       On Error Resume Next

       tlbBotoes.Buttons.Item(1).Enabled = False
       tlbBotoes.Buttons.Item(2).Enabled = True
       tlbBotoes.Buttons.Item(3).Enabled = True
       tlbBotoes.Buttons.Item(4).Enabled = True
       tlbBotoes.Buttons.Item(5).Enabled = False
        
       strID = hfgLista.TextArray((hfgLista.Row * hfgLista.Cols + hfgLista.Col + 1))
       
       If hfgLista.TextArray((hfgLista.Row * hfgLista.Cols + hfgLista.Col + 2)) = "Não" Then
          optFisica.Value = True
       Else
          optJuridica.Value = True
       End If
       
       txtCpfCnpj.Text = hfgLista.TextArray((hfgLista.Row * hfgLista.Cols + hfgLista.Col + 3))
       txtCodigo_interno.Text = hfgLista.TextArray((hfgLista.Row * hfgLista.Cols + hfgLista.Col + 4))
       txtLimite_credito.Text = Format(hfgLista.TextArray((hfgLista.Row * hfgLista.Cols + hfgLista.Col + 5)), "#,###0.00")
    
       booAlterar = True
       txtConsulta.Text = Empty
       sstLista_Branca.Tab = 0
       Me.txtCodigo_interno.SetFocus
       
    End If
    
End Sub

Private Sub adgLista_Branca_HeadClick(ByVal ColIndex As Integer)
    strCampo_consulta = adgLista_branca.Columns(ColIndex).DataField
    txtConsulta.SetFocus
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
        SendKeys "{TAB}"
    End If
End Sub
  
Private Sub Form_Load()
    On Error GoTo Erro
   
    'Informações constantes para o log
    
    'Ver
    log.Data = Date
    
    'Ver
'   strEstacao_log = MDIPrincipal_Cadastro_Base.strEstação
'   strUsuario_log = MDIPrincipal_Cadastro_Base.UsuárioOCX.NomeReduzido
    log.Estacao = "INFO-888"
    log.Usuario = "Adão"
    log.Programa = "Cadastro de Lista Branca"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando o cadastro de Lista Branca"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando o log
    log.Gravar_log "PDV", Me
    
    sstLista_Branca.Tab = 1
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
    
    Exit Sub
    
Erro:

    Call Erro.Erro(Me, "PDV", "Unload")
    Exit Sub

End Sub

Private Sub optFisica_Click()
    
    lblCpfCnpj.Caption = "CPF"
   ' txtCpfCnpj.Text = Empty
    
End Sub

Private Sub optJuridica_Click()
    lblCpfCnpj.Caption = "CNPJ"
    'txtCpfCnpj.Text = Empty
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
    ElseIf optJuridica.Value = True Then
       strPessoa = 1
    End If
    
    strCampo = "DFCodigo_interno_TBLista_branca,DFLimite_credito_TBLista_Branca,DFCnpj_TBLista_Branca,DFTipo_pessoa_TBLista_Branca"
    strvalores = " '" & txtCodigo_interno.Text & "' , " & Funcoes_Gerais.Grava_Moeda(txtLimite_credito) & " , '" & Me.txtCpfCnpj.Text & "'," & strPessoa & ""
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFCodigo_interno_TBLista_branca = '" & txtCodigo_interno.Text & "', " & _
                "    DFLimite_credito_TBLista_Branca = " & Funcoes_Gerais.Grava_Moeda(txtLimite_credito) & ", " & _
                "    DFCnpj_TBLista_Branca = '" & Me.txtCpfCnpj.Text & "', " & _
                "    DFTipo_pessoa_TBLista_Branca =  " & strPessoa & " "
       Call funcoes_banco.Alterar("TBLista_Branca", strSet, "DFID_TBLista_Branca", strID, "PDV", Me, "BDSupervisor")
       log.Descricao = "Alterando o registro: " + txtCpfCnpj.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       
       'Gravando log
       log.Gravar_log "PDV", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBLista_branca", strCampo, strvalores, "PDV", Me, "BDSupervisor")
       log.Descricao = "Gravando o registro: " + txtCpfCnpj.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       
       'Gravando log
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

    Call Erro.Erro(Me, "PDV", "Gravar")
    Exit Function
    
End Function

Private Function Excluir()

    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtCpfCnpj.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
        
    'Gravando log
    log.Gravar_log "PDV", Me
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBLista_Branca", "DFID_TBLista_branca", strID, "PDV", Me, "BDSupervisor")
    
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
    
    Call Erro.Erro(Me, "PDV", "Excluir")
    Exit Function

End Function
Private Function Cancelar()

    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    'Inserir log
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
Private Function Novo()

    On Error GoTo Erro
    
    sstLista_Branca.Tab = 0
    Call Objetos.Limpa_TXT(Me)
    
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
    Me.optFisica.SetFocus
    booAlterar = False
    Exit Function
Erro:
    Call Erro.Erro(Me, "PDV", "Novo")
    Exit Function

End Function

'Private Sub txtCodigo_interno_LostFocus()
'    Movimentacoes.Verifica_Numero "DFCodigo_interno_TBLista_Branca", "TBLista_Branca", txtCodigo_interno, "PDV", Me
'End Sub

Private Function Reposicao()

    On Error GoTo Erro

    Dim strSQL As String

    strSQL = "SELECT * FROM TBLista_Branca"
    
    If txtConsulta.Text <> Empty Then
        strSQL = strSQL & " WHERE " & strCampo_consulta & " LIKE '" & txtConsulta.Text & "%' "
    End If

    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgLista, "0,0,1500,1500,1600", "ID,Tipo,CPF/CNPJ,Código Interno,Limite de Crédito", "BDSupervisor", "PDV", Me
     
    Exit Function

Erro:
    
    Call Erro.Erro(Me, "PDV", "Reposição")
    Exit Function

End Function

Private Sub txtConsulta_Change()

   Call Reposicao
   
End Sub

Private Sub txtLimite_credito_LostFocus()

   txtLimite_credito.Text = Format(txtLimite_credito.Text, "#,###0.00")
   
End Sub
