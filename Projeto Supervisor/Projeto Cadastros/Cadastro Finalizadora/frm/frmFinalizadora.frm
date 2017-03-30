VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFinalizadora 
   Caption         =   "Finalizadora"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
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
   ScaleHeight     =   2595
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstFinalizadora 
      Height          =   2265
      Left            =   0
      TabIndex        =   7
      Top             =   330
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   3995
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
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "frmFinalizadora.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(7)=   "txtCodigo"
      Tab(0).Control(8)=   "txtDescricao"
      Tab(0).Control(9)=   "txtAutenticacoes"
      Tab(0).Control(10)=   "cbbPermitir_Troco"
      Tab(0).Control(11)=   "cbbAbrir_Gaveta"
      Tab(0).Control(12)=   "cbbComprovante"
      Tab(0).Control(13)=   "cbbTipo_Pgto"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Listagem"
      TabPicture(1)   =   "frmFinalizadora.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtConsulta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgFinalizadora"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgFinalizadora 
         Height          =   1035
         Left            =   120
         TabIndex        =   18
         Top             =   1110
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   1826
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox cbbTipo_Pgto 
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmFinalizadora.frx":0038
         Left            =   -69930
         List            =   "frmFinalizadora.frx":003A
         TabIndex        =   6
         Top             =   1410
         Width           =   1245
      End
      Begin VB.ComboBox cbbComprovante 
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   -71130
         TabIndex        =   5
         Top             =   1410
         Width           =   1155
      End
      Begin VB.ComboBox cbbAbrir_Gaveta 
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmFinalizadora.frx":003C
         Left            =   -72300
         List            =   "frmFinalizadora.frx":003E
         TabIndex        =   4
         Top             =   1410
         Width           =   1125
      End
      Begin VB.ComboBox cbbPermitir_Troco 
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmFinalizadora.frx":0040
         Left            =   -73650
         List            =   "frmFinalizadora.frx":0042
         TabIndex        =   3
         Top             =   1410
         Width           =   1275
      End
      Begin VB.TextBox txtAutenticacoes 
         Height          =   360
         Left            =   -74880
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1410
         Width           =   1185
      End
      Begin VB.TextBox txtConsulta 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   6165
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   -73710
         MaxLength       =   50
         TabIndex        =   1
         Top             =   780
         Width           =   4995
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   0
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Pagto."
         Height          =   240
         Left            =   -69930
         TabIndex        =   17
         Top             =   1170
         Width           =   1245
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Comprovante"
         Height          =   240
         Left            =   -71130
         TabIndex        =   16
         Top             =   1170
         Width           =   1140
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Abrir Gaveta"
         Height          =   240
         Left            =   -72300
         TabIndex        =   15
         Top             =   1170
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Autenticações"
         Height          =   240
         Left            =   -74880
         TabIndex        =   13
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Permitir Troco"
         Height          =   240
         Left            =   -73620
         TabIndex        =   12
         Top             =   1170
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   -73680
         TabIndex        =   10
         Top             =   540
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
         Left            =   -74880
         TabIndex        =   9
         Top             =   540
         Width           =   585
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6420
      Top             =   360
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
            Picture         =   "frmFinalizadora.frx":0044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":0678
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":0DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinalizadora.frx":10C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
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
Attribute VB_Name = "frmFinalizadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Cadastrar de Finalizadoras                                     '
' Data de Criação........: 30/04/2003                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião                        '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim conexao As DLLConexao_Sistema.conexao
Public log As New DLLSystemManager.log

Private Sub hfgFinalizadora_Click()

    If hfgFinalizadora.Col > 0 Then
        strCampo_consulta = hfgFinalizadora.DataField(0, hfgFinalizadora.ColSel - 1)
        txtConsulta.SetFocus
    End If
    If hfgFinalizadora.Col = 0 Then
       
       On Error Resume Next
       
       tlbBotoes.Buttons.Item(1).Enabled = False
       tlbBotoes.Buttons.Item(2).Enabled = True
       tlbBotoes.Buttons.Item(3).Enabled = True
       tlbBotoes.Buttons.Item(4).Enabled = True
       tlbBotoes.Buttons.Item(5).Enabled = False
        
       txtCodigo.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 1))
       txtDescricao.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 2))
       cbbPermitir_Troco.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 3))
       txtAutenticacoes.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 4))
       cbbAbrir_Gaveta.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 5))
       cbbComprovante.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 6))
       cbbTipo_Pgto.Text = hfgFinalizadora.TextArray((hfgFinalizadora.Row * hfgFinalizadora.Cols + hfgFinalizadora.Col + 7))
                 
       booAlterar = True
       txtConsulta.Text = Empty
       sstFinalizadora.Tab = 0
       Me.txtDescricao.SetFocus
       
    End If
    
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
    log.Programa = "Cadastro de Finalizadora"
        
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando o Cadastro de Finalizadora"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando o log
    log.Gravar_log "PDV", Me
    
    sstFinalizadora.Tab = 1
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
    Dim strPermite_Troco As String
    Dim strAbre_Gaveta As String
    Dim strImprime_Comprovante As String
    Dim strTipo_Pgto As String
                
    If cbbPermitir_Troco.Text = "Não" Then
       strPermite_Troco = 1
    Else
      If cbbPermitir_Troco.Text = "Dinheiro" Then
         strPermite_Troco = 2
      Else
         strPermite_Troco = 3
      End If
    End If
    
    If cbbAbrir_Gaveta.Text = "Sim" Then
       strAbre_Gaveta = 1
    Else
       strAbre_Gaveta = 0
    End If
    
    If cbbComprovante.Text = "Sim" Then
       strImprime_Comprovante = 1
    Else
       strImprime_Comprovante = 0
    End If
    
    If cbbTipo_Pgto.Text = "Simples" Then
       strTipo_Pgto = 1
    Else
       If cbbTipo_Pgto.Text = "Imprime Cheque" Then
          strTipo_Pgto = 2
       Else
          strTipo_Pgto = 3
       End If
    End If
    
    strCampo = "PKCodigo_TBFinalizadora,DFDescricao_TBFinalizadora,DFPermite_troco_TBFinalizadora," & _
               "DFNumero_autenticacoes_TBFinalizadora,DFAbre_gaveta_TBFinalizadora," & _
               "DFImprime_comprovante_vinculado_TBFinalizadora,DFTipo_pagamento_TBFinalizadora"
               
    strvalores = " " & txtCodigo.Text & " , '" & txtDescricao.Text & "' , '" & strPermite_Troco & "' , " & _
                 " '" & txtAutenticacoes.Text & "' , '" & strAbre_Gaveta & "' , '" & strImprime_Comprovante & "' , " & _
                 " '" & strTipo_Pgto & "'"
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFDescricao_TBFinalizadora = '" & txtDescricao.Text & "'  , " & _
                "    DFPermite_troco_TBFinalizadora = '" & strPermite_Troco & "' , " & _
                "    DFNumero_autenticacoes_TBFinalizadora = '" & txtAutenticacoes.Text & "' , " & _
                "    DFAbre_gaveta_TBFinalizadora = '" & strAbre_Gaveta & "' , " & _
                "    DFImprime_Comprovante_Vinculado_TBFinalizadora = '" & strImprime_Comprovante & "' , " & _
                "    DFTipo_pagamento_TBFinalizadora = '" & strTipo_Pgto & "'"
                
       Call funcoes_banco.Alterar("TBFinalizadora", strSet, "PKCodigo_TBFinalizadora", txtCodigo.Text, "pdv", Me, "BDSupervisor")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "PDV", Me
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBFinalizadora", strCampo, strvalores, "pdv", Me, "BDSupervisor")
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       
       'Gravando log
       log.Gravar_log "PDV", Me
    End If
    
    Call Reposicao
    Call Objetos.Limpa_TXT(Me)
    
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
    Call funcoes_banco.Excluir("TBFinalizadora", "PKCodigo_TBFinalizadora", adgFinalizadora.Columns(0).Value, "pdv", Me, "BDSupervisor")
          
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
    
    'Inserir log
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    Call Objetos.Limpa_TXT(Me)
    cbbAbrir_Gaveta.Text = Empty
    cbbComprovante.Text = Empty
    cbbPermitir_Troco.Text = Empty
    cbbTipo_Pgto.Text = Empty
    
    sstFinalizadora.Tab = 1
    
            
    Exit Function
Erro:
    
    Call Erro.Erro(Me, "PDV", "Cancelar")
    Exit Function

End Function
Private Function Novo()

    On Error GoTo Erro
    
    sstFinalizadora.Tab = 0
    'Call Verifica_TXT.Limpa_TXT
    
    Call Abastece_Combos
       
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
    Movimentacoes.Verifica_Numero "PKCodigo_TBFinalizadora", "TBFinalizadora", txtCodigo, "PDV", Me
End Sub

Private Function Reposicao()

    On Error GoTo Erro

    Dim strSQL As String
    Dim strTeste As String

    strSQL = "SELECT * FROM TBFinalizadora"
    
    If txtConsulta.Text <> Empty Then
        strSQL = strSQL & " WHERE " & strCampo_consulta & " LIKE '" & txtConsulta.Text & "%' "
    End If

    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgFinalizadora, "800,1000,1325,1150,1200,1290,1350", "Código,Descrição,Permite Troco,Autenticações,Abrir Gaveta,Comprovante,Tipo de Pgto.", "BDSupervisor", "PDV", Me
        
    hfgFinalizadora.Col = 3
    If hfgFinalizadora.TopRow <> 0 Then
          hfgFinalizadora.Row = hfgFinalizadora.TopRow
          For I = 1 To hfgFinalizadora.Rows - 1
              hfgFinalizadora.Row = I
              If hfgFinalizadora.Text = "1" Then
                 hfgFinalizadora.Text = "Não"
              Else
                If hfgFinalizadora.Text = "2" Then
                   hfgFinalizadora.Text = "Dinheiro"
                Else
                   hfgFinalizadora.Text = "Contavale"
                End If
              End If
          Next I
    End If
    
    hfgFinalizadora.Col = 7
    If hfgFinalizadora.TopRow <> 0 Then
          hfgFinalizadora.Row = hfgFinalizadora.TopRow
          For I = 1 To hfgFinalizadora.Rows - 1
              hfgFinalizadora.Row = I
              If hfgFinalizadora.Text = "1" Then
                 hfgFinalizadora.Text = "Simples"
              Else
                If hfgFinalizadora.Text = "2" Then
                   hfgFinalizadora.Text = "Imprime Cheque"
                Else
                   hfgFinalizadora.Text = "TEF"
                End If
              End If
          Next I
    End If
    
    hfgFinalizadora.Refresh
    
    Call Abastece_Combos
    
    Exit Function

Erro:
    
    Call Erro.Erro(Me, "PDV", "Reposicao")
    Resume Next
    
End Function

Private Sub txtConsulta_Change()
   Call Reposicao
End Sub

Private Function Abastece_Combos()
    
   cbbPermitir_Troco.Clear
   cbbPermitir_Troco.AddItem ("Não")
   cbbPermitir_Troco.AddItem ("Dinheiro")
   cbbPermitir_Troco.AddItem ("ContaVale")
    
   cbbComprovante.Clear
   cbbComprovante.AddItem ("Sim")
   cbbComprovante.AddItem ("Não")
    
   cbbAbrir_Gaveta.Clear
   cbbAbrir_Gaveta.AddItem ("Sim")
   cbbAbrir_Gaveta.AddItem ("Não")
      
   cbbTipo_Pgto.Clear
   cbbTipo_Pgto.AddItem ("Simples")
   cbbTipo_Pgto.AddItem ("Imprime Cheque")
   cbbTipo_Pgto.AddItem ("TEF")

End Function
