VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCategoria 
   Caption         =   "Categoria"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "frmCategoria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
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
   Begin TabDlg.SSTab sstCategoria 
      Height          =   2925
      Left            =   30
      TabIndex        =   5
      Top             =   330
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   5159
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      OLEDropMode     =   1
      MouseIcon       =   "frmCategoria.frx":030A
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
      TabPicture(0)   =   "frmCategoria.frx":0326
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "txtCodigo"
      Tab(0).Control(4)=   "txtDescricao"
      Tab(0).Control(5)=   "txtCodigo_subgrupo"
      Tab(0).Control(6)=   "dtcSubgrupo"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Listagem"
      TabPicture(1)   =   "frmCategoria.frx":0342
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtConsulta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgCategoria"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgCategoria 
         Height          =   1635
         Left            =   120
         TabIndex        =   11
         Top             =   1140
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   2884
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo dtcSubgrupo 
         Height          =   360
         Left            =   -73920
         TabIndex        =   3
         Top             =   2250
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
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
      Begin VB.TextBox txtCodigo_subgrupo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   2250
         Width           =   915
      End
      Begin VB.TextBox txtConsulta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4665
      End
      Begin VB.TextBox txtDescricao 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   1
         Top             =   1500
         Width           =   4665
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   0
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Subgrupo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   10
         Top             =   2010
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   8
         Top             =   1260
         Width           =   825
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Picture         =   "frmCategoria.frx":035E
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   7
         Top             =   600
         Width           =   585
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4770
      Top             =   -90
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
            Picture         =   "frmCategoria.frx":2D3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategoria.frx":3057
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategoria.frx":3371
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategoria.frx":370B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategoria.frx":3AA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategoria.frx":3DBF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Cadastrar de Categoria                                         '
' Data de Criação........: 30/04/2003                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes,M. Delão  '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim strSQL As String
Public log As New DLLSystemManager.log


Private Sub hfgCategoria_Click()
         
    If hfgCategoria.Col > 0 Then
        strCampo_consulta = hfgCategoria.DataField(0, hfgCategoria.ColSel - 1)
        txtConsulta.SetFocus
    End If
    If hfgCategoria.Col = 0 Then
    
       On Error Resume Next
       
       tlbBotoes.Buttons.Item(1).Enabled = False
       tlbBotoes.Buttons.Item(2).Enabled = True
       tlbBotoes.Buttons.Item(3).Enabled = True
       tlbBotoes.Buttons.Item(4).Enabled = True
       tlbBotoes.Buttons.Item(5).Enabled = False
    
       txtCodigo.Text = hfgCategoria.TextArray((hfgCategoria.Row * hfgCategoria.Cols + hfgCategoria.Col + 1))
       txtDescricao.Text = hfgCategoria.TextArray((hfgCategoria.Row * hfgCategoria.Cols + hfgCategoria.Col + 2))
       txtCodigo_subgrupo.Text = hfgCategoria.TextArray((hfgCategoria.Row * hfgCategoria.Cols + hfgCategoria.Col + 3))
       dtcSubgrupo.Text = hfgCategoria.TextArray((hfgCategoria.Row * hfgCategoria.Cols + hfgCategoria.Col + 4))
    
       booAlterar = True
       txtConsulta.Text = Empty
       sstCategoria.Tab = 0
       Me.txtDescricao.SetFocus
    End If
    
End Sub

Private Sub dtcSubgrupo_Click(Area As Integer)
    txtCodigo_subgrupo = dtcSubgrupo.BoundText
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
    log.Programa = "Cadastro de Categoria"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando o Cadastro de Categoria"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando o log
    log.Gravar_log "PDV", Me
    
    sstCategoria.Tab = 1
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
    log.Programa = "Cadastro de Categoria"
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    Exit Sub
    
Erro:

    log.Evento = "Unload"
    log.Tipo = 3
    log.Descricao = Err.Description
    
    'Gravando log
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
    
    strCampo = "PKCodigo_TBCategoria,DFDescricao_TBCategoria,FKCodigo_TBSubgrupo"
    strvalores = " " & txtCodigo & " , '" & txtDescricao.Text & "'," & txtCodigo_subgrupo.Text & ""
    
    If booAlterar = True Then
       log.Evento = "Alterar"
       strSet = "SET DFDescricao_TBCategoria = '" & txtDescricao.Text & "', " & _
                "    FKCodigo_TBSubgrupo = " & txtCodigo_subgrupo.Text & " "
                
       Call funcoes_banco.Alterar("TBCategoria", strSet, "PKCodigo_TBCategoria", txtCodigo.Text, "PDV", Me, "BDSupervisor")
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       
       'Gravando log
       log.Gravar_log "PDV", Me
       
    Else
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBCategoria", strCampo, strvalores, "PDV", Me, "BDSupervisor")
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "PDV", Me
       
    End If
    
    Call Objetos.Limpa_TXT(Me)
    dtcSubgrupo.Text = Empty
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
    Call funcoes_banco.Excluir("TBCategoria", "PKCodigo_TBCategoria", Me.adgCategoria.Columns(0).Value, "PDV", Me, "BDSupervisor")
   
    Call Limpar_Campos
   
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
    dtcSubgrupo.Text = Empty
    
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
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "PDV", "Cancelar")
    Exit Function
    
End Function
Private Function Novo()

    On Error GoTo Erro
    
    sstCategoria.Tab = 0
    
    Call Objetos.Limpa_TXT(Me)
    dtcSubgrupo.Text = Empty
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    
     strSQL = "SELECT * FROM TBSubgrupo"
    'montando combo de seção'
    Call Movimentacoes.Movimenta_DataCombo("PKCodigo_TBSubgrupo", "DFDescricao_TBSubgrupo", dtcSubgrupo, strSQL, "BDSupervisor", "PDV", Me)
    
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

Private Sub txtCodigo_subgrupo_Change()
    dtcSubgrupo.BoundText = txtCodigo_subgrupo.Text
End Sub

Private Sub txtCodigo_LostFocus()
      Movimentacoes.Verifica_Numero "PKCodigo_TBCategoria", "TBCategoria", txtCodigo, "PDV", Me
End Sub

Private Function Reposicao()
    
    On Error GoTo Erro
     
    Dim strSQL As String
       
    strSQL = "SELECT TBCategoria.PKCodigo_TBCategoria,TBCategoria.DFDescricao_TBCategoria,TBCategoria.FKCodigo_TBSubgrupo,TBSubgrupo.DFDescricao_TBSubgrupo " & _
             "FROM TBCategoria " & _
             "INNER JOIN TBSubgrupo " & _
             "ON TBCategoria.FKCodigo_TBSubgrupo = TBSubgrupo.PKCodigo_TBSubgrupo"
    
    If txtConsulta.Text <> Empty Then
        strSQL = strSQL & " WHERE " & strCampo_consulta & " LIKE '" & txtConsulta.Text & "%' "
    End If
    
    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgCategoria, "800,3000,800,3000", "Categoria,Descrição,SubGrupo,Descrição", "BDSupervisor", "PDV", Me
    
    strSQL = "SELECT * FROM TBSubgrupo"
    'montando combo de seção'
    Call Movimentacoes.Movimenta_DataCombo("PKCodigo_TBSubgrupo", "DFDescricao_TBSubgrupo", dtcSubgrupo, strSQL, "BDSupervisor", "PDV", Me)
    
    Exit Function
     
Erro:
    Call Erro.Erro(Me, "PDV", "Reposicao")
    Resume Next
End Function


Private Sub txtConsulta_Change()
    Call Reposicao
End Sub
