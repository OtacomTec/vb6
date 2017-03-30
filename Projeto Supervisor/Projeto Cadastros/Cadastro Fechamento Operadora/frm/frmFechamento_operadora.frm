VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFechamento_operadora 
   Caption         =   "Fechamento Operadora"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
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
   ScaleHeight     =   4185
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtData 
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   900
      Visible         =   0   'False
      Width           =   1515
   End
   Begin TabDlg.SSTab sstFechamento_operadora 
      Height          =   3855
      Left            =   0
      TabIndex        =   7
      Top             =   330
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   6800
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
      TabPicture(0)   =   "frmFechamento_operadora.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCodigo_usuario"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtObservacao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtValor"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DTPicker1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCodigo_finalizadora"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dtcDescricao_usuario"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dtcDescricao_finalizadora"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Listagem"
      TabPicture(1)   =   "frmFechamento_operadora.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtConsulta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgFechamento"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgFechamento 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   19
         Top             =   1170
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   4471
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo dtcDescricao_finalizadora 
         Height          =   360
         Left            =   1890
         TabIndex        =   5
         Top             =   2160
         Width           =   3405
         _ExtentX        =   6006
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
      Begin MSDataListLib.DataCombo dtcDescricao_usuario 
         Height          =   360
         Left            =   1890
         TabIndex        =   3
         Top             =   1440
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   8388608
         Text            =   ""
      End
      Begin VB.TextBox txtCodigo_finalizadora 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   750
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   16777215
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   19660801
         CurrentDate     =   37769
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   1
         Top             =   750
         Width           =   3375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   375
         Left            =   -74880
         TabIndex        =   8
         Top             =   720
         Width           =   5145
      End
      Begin VB.TextBox txtObservacao 
         ForeColor       =   &H00800000&
         Height          =   840
         Left            =   120
         MaxLength       =   50
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2880
         Width           =   5145
      End
      Begin VB.TextBox txtCodigo_usuario 
         Height          =   375
         Left            =   120
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   1920
         TabIndex        =   17
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   240
         Left            =   1920
         TabIndex        =   16
         Top             =   1890
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Código Finalizadora"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   510
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   240
         Left            =   1890
         TabIndex        =   12
         Top             =   510
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   2610
         Width           =   1005
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
         Caption         =   "Código Usuário"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1290
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
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
            Picture         =   "frmFechamento_operadora.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_operadora.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_operadora.frx":066C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_operadora.frx":0A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_operadora.frx":0DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechamento_operadora.frx":10BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   5400
      _ExtentX        =   9525
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
Attribute VB_Name = "frmFechamento_operadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Cadastrar de Fechamento Operadora                              '
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

Private Sub hfgFechamento_Click()

    If hfgFechamento.Col > 0 Then
        strCampo_consulta = hfgFechamento.DataField(0, hfgFechamento.ColSel - 1)
        txtConsulta.SetFocus
    End If
    If hfgFechamento.Col = 0 Then
       
       On Error Resume Next
        
       tlbBotoes.Buttons.Item(1).Enabled = False
       tlbBotoes.Buttons.Item(2).Enabled = True
       tlbBotoes.Buttons.Item(3).Enabled = True
       tlbBotoes.Buttons.Item(4).Enabled = True
       tlbBotoes.Buttons.Item(5).Enabled = False
       
       strID = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 1))
       txtCodigo_usuario.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 2))
       dtcDescricao_usuario.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 3))
       txtCodigo_finalizadora.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 4))
       dtcDescricao_finalizadora.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 5))
       txtValor.Text = Format(hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 6)), "#,###0.00")
       DTPicker1.Value = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 7))
       txtObservacao.Text = hfgFechamento.TextArray((hfgFechamento.Row * hfgFechamento.Cols + hfgFechamento.Col + 8))
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstFechamento_operadora.Tab = 0
       Me.txtCodigo_usuario.SetFocus
    End If
    
End Sub
Private Sub dtcDescricao_usuario_Click(Area As Integer)
    txtCodigo_usuario.Text = Me.dtcDescricao_usuario.BoundText
End Sub
Private Sub dtcDescricao_finalizadora_Click(Area As Integer)
    txtCodigo_finalizadora.Text = dtcDescricao_finalizadora.BoundText
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
    log.Programa = "Cadastro de Fechamento Operadora"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando o Cadastro de Fechamento Operadora"
    log.Tipo = 1
    
    'Gravando o log
    log.Gravar_log "PDV", Me
    
    sstFechamento_operadora.Tab = 1
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
    
    'Gravando no Log
    log.Gravar_log "PDV", Me
    
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
    Dim strData As String
                
    strCampo = "FKCodigo_TBUsuario,FKCodigo_TBFinalizadora," & _
               "DFValor_TBFechamento_operadora,DFData_TBFechamento_operadora," & _
               "DFObeservacao_TBFechamento_operadora"
    
    strData = Format(DTPicker1.Value, "YYYYMMDD")

    strvalores = " " & txtCodigo_usuario.Text & " , " & txtCodigo_finalizadora & " , " & Funcoes_Gerais.Grava_Moeda(txtValor) & " , " & _
                 " '" & strData & "' ,'" & txtObservacao.Text & "'"
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFValor_TBFechamento_operadora = " & Funcoes_Gerais.Grava_Moeda(txtValor) & ", DFData_TBFechamento_operadora= '" & DTPicker1.Value & "'," & _
                "DFObeservacao_TBFechamento_operadora = '" & txtObservacao.Text & "'"
       Call funcoes_banco.Alterar("TBFechamento_operadora", strSet, "PKId_TBFechamento_operadora", strID, "PDV", Me, "BDSupervisor")
       log.Descricao = "Alterando o registro: " + txtCodigo_finalizadora.Text
       log.Tipo = 1
       log.Gravar_log "PDV", Me
    Else
       Call funcoes_banco.Gravar("TBFechamento_operadora", strCampo, strvalores, "PDV", Me, "BDSupervisor")
       log.Evento = "Incluir Novo"
       log.Descricao = "Gravando o registro: " + txtCodigo_finalizadora.Text
       log.Tipo = 1
       log.Gravar_log "PDV", Me
    End If
    
    Call Objetos.Limpa_TXT(Me)
    dtcDescricao_finalizadora.Text = Empty
    dtcDescricao_usuario.Text = Empty
    
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
    
    'Gravando log
    log.Gravar_log "PDV", Me
    Call Erro.Erro(Me, "PDV", "Gravar")
    Exit Function
    
End Function

Private Function Excluir()

    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtCodigo_finalizadora.Text
    log.Tipo = 1
           
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBFechamento_operadora", "PKId_TBFechamento_operadora", strID, "PDV", Me, "BDSupervisor")
    
    'Gravando log
    log.Gravar_log "PDV", Me
    
    Call Objetos.Limpa_TXT(Me)
    dtcDescricao_finalizadora.Text = Empty
    dtcDescricao_usuario.Text = Empty
    
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
    dtcDescricao_finalizadora.Text = Empty
    dtcDescricao_usuario.Text = Empty
    
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
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    Call Erro.Erro(Me, "PDV", "Cancelar")
    Exit Function

End Function
Private Function Novo()

    On Error GoTo Erro
    
    sstFechamento_operadora.Tab = 0
    Call Objetos.Limpa_TXT(Me)
    dtcDescricao_finalizadora.Text = Empty
    dtcDescricao_usuario.Text = Empty
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    
    log.Gravar_log "PDV", Me
    
    tlbBotoes.Buttons.Item(1).Enabled = False
    tlbBotoes.Buttons.Item(2).Enabled = True
    tlbBotoes.Buttons.Item(3).Enabled = True
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = False
    Me.DTPicker1.Enabled = True
    Me.DTPicker1.SetFocus
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

Private Sub dtpicker1_LostFocus()
    txtData.Text = DTPicker1.Value
    Movimentacoes.Verifica_Numero "DFData_TBFechamento_operadora", "TBFechamento_operadora", txtData, "PDV", Me
End Sub


Private Function Reposicao()
    On Error GoTo Erro

    Dim strSQL As String

    strSQL = "SELECT TBFechamento_operadora.PKId_TBFechamento_Operadora,TBFechamento_operadora.FKCodigo_TBUsuario," & _
             "TBUsuario.DFNome_TBUsuario,TBFechamento_operadora.FKCodigo_TBFinalizadora,TBFinalizadora.DFDescricao_TBFinalizadora," & _
             "TBFechamento_operadora.DFValor_TBFechamento_Operadora,TBFechamento_operadora.DFData_TBFechamento_Operadora," & _
             "TBFechamento_operadora.DFObeservacao_TBFechamento_operadora FROM TBFechamento_operadora " & _
             "INNER JOIN TBUsuario ON TBFechamento_operadora.FKCodigo_TBUsuario = TBUsuario.PKCodigo_TBUsuario " & _
             "INNER JOIN TBFinalizadora ON TBFechamento_operadora.FKCodigo_TBFinalizadora =  TBFinalizadora.PKCodigo_TBFinalizadora"
    
    If txtConsulta.Text <> Empty Then
        strSQL = strSQL & " WHERE " & strCampo_consulta & " LIKE '" & txtConsulta.Text & "%' "
    End If

    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgFechamento, "0,0,1500,0,1500,1000,1000,3500", "C,C,Nome,C,Descrição,Valor,Data,Observaçaõ.", "BDSupervisor", "PDV", Me
    
    strSQL = "SELECT * FROM TBUsuario"
    Call Movimentacoes.Movimenta_DataCombo("PKCodigo_TBUsuario", "DFNome_TBUsuario", dtcDescricao_usuario, strSQL, "BDSupervisor", "PDV", Me)
            
    strSQL = "SELECT * FROM TBFinalizadora"
    Call Movimentacoes.Movimenta_DataCombo("PKCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcDescricao_finalizadora, strSQL, "BDSupervisor", "PDV", Me)
    
    Exit Function

Erro:
    log.Evento = "Reposição"
    log.Tipo = 3
    log.Descricao = Err.Description
    log.Gravar_log "PDV", Me
    
    Call Erro.Erro(Me, "PDV", "Reposicao")
    Resume Next
    
End Function

Private Sub txtCodigo_finalizadora_LostFocus()
    Me.dtcDescricao_finalizadora.BoundText = txtCodigo_finalizadora.Text
End Sub

Private Sub txtCodigo_usuario_LostFocus()
    Me.dtcDescricao_usuario.BoundText = txtCodigo_usuario.Text
End Sub

Private Sub txtConsulta_Change()
   Call Reposicao
End Sub

Private Sub txtObservacao_LostFocus()
    
    txtObservacao.Text = UCase(txtObservacao.Text)
    
End Sub

Private Sub txtValor_LostFocus()
    
    txtValor.Text = Format(txtValor.Text, "#,###0.00")
    
End Sub
