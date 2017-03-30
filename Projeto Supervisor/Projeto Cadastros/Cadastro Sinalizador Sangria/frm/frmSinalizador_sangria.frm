VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSinalizador_sangria 
   Caption         =   "Sinalizador Sangria"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstSinalizador 
      Height          =   2505
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4419
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
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "frmSinalizador_sangria.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCodigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtValor"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtcFinalizadora"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Listagem"
      TabPicture(1)   =   "frmSinalizador_sangria.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "hfgSinalizador"
      Tab(1).Control(1)=   "txtConsulta"
      Tab(1).Control(2)=   "Label6"
      Tab(1).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSinalizador 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   10
         Top             =   1170
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   2143
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo dtcFinalizadora 
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.TextBox txtValor 
         Height          =   375
         Left            =   2370
         TabIndex        =   1
         Top             =   1560
         Width           =   2325
      End
      Begin VB.TextBox txtConsulta 
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   4
         TabIndex        =   0
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   240
         Left            =   2370
         TabIndex        =   7
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   1320
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
         TabIndex        =   4
         Top             =   600
         Width           =   585
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4830
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
            Picture         =   "frmSinalizador_sangria.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinalizador_sangria.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinalizador_sangria.frx":066C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinalizador_sangria.frx":0A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinalizador_sangria.frx":0DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSinalizador_sangria.frx":10BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
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
End
Attribute VB_Name = "frmSinalizador_sangria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Cadastrar de Sinalizador Sangria                               '
' Data de Criação........: 30/04/2003                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião                        '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim strID As String
Dim booAlterar As Boolean
Dim conexao As DLLConexao_Sistema.conexao
Dim log As New DLLSystemManager.log
Private Sub hfgSinalizador_Click()

    If hfgSinalizador.Col > 0 Then
       strCampo_consulta = hfgSinalizador.DataField(0, hfgSinalizador.ColSel - 1)
       txtConsulta.SetFocus
    End If
    If hfgSinalizador.Col = 0 Then
     
       On Error Resume Next

       tlbBotoes.Buttons.Item(1).Enabled = False
       tlbBotoes.Buttons.Item(2).Enabled = True
       tlbBotoes.Buttons.Item(3).Enabled = True
       tlbBotoes.Buttons.Item(4).Enabled = True
       tlbBotoes.Buttons.Item(5).Enabled = False
     
       strID = hfgSinalizador.TextArray((hfgSinalizador.Row * hfgSinalizador.Cols + hfgSinalizador.Col + 1))
       txtCodigo.Text = hfgSinalizador.TextArray((hfgSinalizador.Row * hfgSinalizador.Cols + hfgSinalizador.Col + 2))
       txtValor.Text = Format(hfgSinalizador.TextArray((hfgSinalizador.Row * hfgSinalizador.Cols + hfgSinalizador.Col + 3)), "#,###0.00")
       dtcFinalizadora.Text = hfgSinalizador.TextArray((hfgSinalizador.Row * hfgSinalizador.Cols + hfgSinalizador.Col + 4))
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstSinalizador.Tab = 0
       dtcFinalizadora.SetFocus
       
    End If
   
End Sub
Private Sub dtcFinalizadora_Click(Area As Integer)
    txtCodigo.Text = dtcFinalizadora.BoundText
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
    
    'Ver
    log.Data = Date
    
    'Ver
'   strEstacao_log = MDIPrincipal_Cadastro_Base.strEstação
'   strUsuario_log = MDIPrincipal_Cadastro_Base.UsuárioOCX.NomeReduzido
    log.Estacao = "INFO-888"
    log.Usuario = "Adão"
    log.Programa = "Cadastro de Sinalizador Sangria"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando o Cad. de Sinalizador"
    log.Tipo = 1
    
    log.Gravar_log "PDV", Me
    
    sstSinalizador.Tab = 1
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
    
    strCampo = "PKCodigo_TBFinalizadora,DFValor_TBSinalizador_sangria"
    strvalores = " " & txtCodigo.Text & " , " & Funcoes_Gerais.Grava_Moeda(txtValor) & " "
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFValor_TBSinalizador_sangria = " & Funcoes_Gerais.Grava_Moeda(txtValor) & " "
       Call funcoes_banco.Alterar("TBSinalizador_sangria", strSet, "PKID_TBSinalizador_sangria", strID, "PDV", Me, "BDSupervisor")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Gravar_log "PDV", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBSinalizador_sangria", strCampo, strvalores, "PDV", Me, "BDSupervisor")
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
    log.Descricao = "Exclusão do registro: " + txtCodigo.Text
    log.Tipo = 1
           
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBSinalizador_sangria", "PKID_TBSinalizador_sangria", srtid, "PDV", Me, "BDSupervisor")
    
    log.Gravar_log "PDV", Me
    
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
    
    sstSinalizador.Tab = 0
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
    
    txtCodigo.Enabled = True
    dtcFinalizadora.Enabled = True
    txtCodigo.SetFocus
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

Private Sub txtCodigo_Change()
    dtcFinalizadora.BoundText = txtCodigo.Text
End Sub

Private Sub txtCodigo_LostFocus()
    Movimentacoes.Verifica_Numero "FKCodigo_TBFinalizadora", "TBSinalizador_sangria", txtCodigo, "PDV", Me
End Sub
Private Function Reposicao()

    On Error GoTo Erro

    Dim strSQL As String
    
    strSQL = "SELECT TBSinalizador_sangria.PKId_TBSinalizador_sangria, TBSinalizador_sangria.PKCODIGO_TBFinalizadora, TBSinalizador_sangria.DFValor_TBSinalizador_sangria, TBFinalizadora.DFDescricao_TBFinalizadora " & _
             "FROM TBSinalizador_sangria " & _
             "INNER JOIN TBFinalizadora " & _
             "ON TBSinalizador_sangria.PKCODIGO_TBFinalizadora = TBFinalizadora.PKCODIGO_TBFinalizadora"
        
    If txtConsulta.Text <> Empty Then
        strSQL = strSQL & " WHERE " & strCampo_consulta & " LIKE '" & txtConsulta.Text & "%' "
    End If

    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgSinalizador, "0,800,1000,2450", "ID,Código,Valor,Finalizadora", "BDSupervisor", "PDV", Me
    
    strSQL = "SELECT TBFinalizadora.PKCodigo_TBFinalizadora, TBFinalizadora.DFDescricao_TBFinalizadora " & _
             "FROM TBFinalizadora"
             
    Call Movimentacoes.Movimenta_DataCombo("PKCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSQL, "BDSupervisor", "PDV", Me)
     
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

Private Sub txtValor_LostFocus()
    
   txtValor.Text = Format(txtValor.Text, "#,###0.00")
   
End Sub
