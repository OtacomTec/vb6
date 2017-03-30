VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFornecedor 
   Caption         =   "Fornecedor"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFornecedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
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
   Begin TabDlg.SSTab sstFornecedor 
      Height          =   4275
      Left            =   0
      TabIndex        =   13
      Top             =   330
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7541
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
      TabPicture(0)   =   "frmFornecedor.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblCpfCnpj"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCodigo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtNome"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtEndereco"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtNumero"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtComplemento"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtBairro"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCidade"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCep"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCpfCnpj"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtInsestadual"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "fraPessoa"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Listagem"
      TabPicture(1)   =   "frmFornecedor.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtConsulta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "hfgFornecedor"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgFornecedor 
         Height          =   2925
         Left            =   -74880
         TabIndex        =   27
         Top             =   1170
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   5159
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame fraPessoa 
         Caption         =   "Pessoa"
         Height          =   705
         Left            =   120
         TabIndex        =   8
         Top             =   3450
         Width           =   1965
         Begin VB.OptionButton optFisica 
            Caption         =   "Física"
            Height          =   345
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   795
         End
         Begin VB.OptionButton optJuridica 
            Caption         =   "Jurídica"
            Height          =   345
            Left            =   960
            TabIndex        =   10
            Top             =   270
            Width           =   975
         End
      End
      Begin VB.TextBox txtInsestadual 
         Height          =   375
         Left            =   4230
         TabIndex        =   12
         Top             =   3660
         Width           =   1965
      End
      Begin VB.TextBox txtCpfCnpj 
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   3660
         Width           =   1995
      End
      Begin VB.TextBox txtCep 
         Height          =   375
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2940
         Width           =   1275
      End
      Begin VB.TextBox txtCidade 
         Height          =   375
         Left            =   120
         MaxLength       =   30
         TabIndex        =   5
         Top             =   2940
         Width           =   2505
      End
      Begin VB.TextBox txtBairro 
         Height          =   375
         Left            =   2700
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2940
         Width           =   2145
      End
      Begin VB.TextBox txtComplemento 
         Height          =   375
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2250
         Width           =   4740
      End
      Begin VB.TextBox txtNumero 
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2250
         Width           =   1275
      End
      Begin VB.TextBox txtEndereco 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1530
         Width           =   6105
      End
      Begin VB.TextBox txtConsulta 
         Height          =   375
         Left            =   -74880
         TabIndex        =   15
         Top             =   720
         Width           =   6105
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   4935
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estadual"
         Height          =   240
         Left            =   4230
         TabIndex        =   26
         Top             =   3420
         Width           =   1545
      End
      Begin VB.Label lblCpfCnpj 
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         Height          =   240
         Left            =   2160
         TabIndex        =   25
         Top             =   3420
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CEP"
         Height          =   240
         Left            =   4920
         TabIndex        =   24
         Top             =   2670
         Width           =   330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   2670
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   240
         Left            =   2700
         TabIndex        =   22
         Top             =   2670
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Complemento"
         Height          =   240
         Left            =   1470
         TabIndex        =   21
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   1980
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   240
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   495
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
            Picture         =   "frmFornecedor.frx":0902
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFornecedor.frx":0C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFornecedor.frx":0F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFornecedor.frx":12D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFornecedor.frx":166A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFornecedor.frx":1984
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Cadastrar de Fornecedores                                      '
' Data de Criação........: 30/04/2003                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião                        '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim conexao As DLLConexao_Sistema.conexao
Public log As New DLLSystemManager.log

Private Sub hfgFornecedor_Click()

    If hfgFornecedor.Col > 0 Then
        strCampo_consulta = hfgFornecedor.DataField(0, hfgFornecedor.ColSel - 1)
        txtConsulta.SetFocus
    End If
    If hfgFornecedor.Col = 0 Then
    
       On Error Resume Next
       
       tlbBotoes.Buttons.Item(1).Enabled = False
       tlbBotoes.Buttons.Item(2).Enabled = True
       tlbBotoes.Buttons.Item(3).Enabled = True
       tlbBotoes.Buttons.Item(4).Enabled = True
       tlbBotoes.Buttons.Item(5).Enabled = False
    
       txtCodigo.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 1))
       txtNome.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 2))
       txtEndereco.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 3))
       txtNumero.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 4))
       txtComplemento.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 5))
       txtBairro.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 6))
       txtCidade.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 7))
       txtCep.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 8))
       
       If hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 9)) = "Não" Then
          optFisica.Value = True
       Else
          optJuridica.Value = True
       End If
       
       txtCpfCnpj.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 10))
       txtInsestadual.Text = hfgFornecedor.TextArray((hfgFornecedor.Row * hfgFornecedor.Cols + hfgFornecedor.Col + 11))
       
       booAlterar = True
       txtConsulta.Text = Empty
       sstFornecedor.Tab = 0
       Me.txtNome.SetFocus
       
   End If
   
End Sub

Private Sub adgFornecedor_HeadClick(ByVal ColIndex As Integer)
    strCampo_consulta = adgFornecedor.Columns(ColIndex).DataField
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
    
    log.Data = Date
    
'   strEstacao_log = MDIPrincipal_Cadastro_Base.strEstação
'   strUsuario_log = MDIPrincipal_Cadastro_Base.UsuárioOCX.NomeReduzido
    log.Estacao = "INFO-888"
    log.Usuario = "Adão"
    log.Programa = "Cadastro de Fornecedor"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando o Cadastro de Fornecedor"
    log.Tipo = 1
    
    log.Gravar_log "PDV", Me
    
    sstFornecedor.Tab = 1
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
    Me.txtCpfCnpj.Text = Empty
            
End Sub

Private Sub optJuridica_Click()
    
    lblCpfCnpj.Caption = "CNPJ"
    Me.txtCpfCnpj.Text = Empty
            
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
    Dim strGuardatipo As String
    
    strCampo = "PKCodigo_TBFornecedor,DFNome_TBFornecedor,DFEndereco_TBFornecedor," & _
               "DFNumero_TBFornecedor,DFComplemento_TBFornecedor," & _
               "DFBairro_TBFornecedor,DFCidade_TBFornecedor,DFCep_TBFornecedor," & _
               "DFTipo_pessoa_TBFornecedor,DFCgc_TBFornecedor,DFInscricao_estadual_TBFornecedor "
               
    If optFisica.Value = True Then
       strGuardatipo = 0
    Else
       strGuardatipo = 1
    End If
    
    strvalores = " " & txtCodigo.Text & " , '" & txtNome.Text & "' , '" & txtEndereco.Text & "' , " & _
                 " '" & txtNumero.Text & "' , '" & txtComplemento.Text & "' , '" & txtBairro.Text & "' , " & _
                 " '" & txtCidade.Text & "' , '" & txtCep.Text & "' , " & strGuardatipo & " , " & _
                 " '" & txtCpfCnpj.Text & "' , '" & txtInsestadual.Text & "' "
    
    If booAlterar = True Then
       strSet = "SET DFNome_TBFornecedor = '" & Me.txtNome.Text & "' , " & _
                "    DFEndereco_TBFornecedor =  '" & txtEndereco.Text & "' , " & _
                "    DFNumero_TBFornecedor = '" & txtNumero.Text & "' , " & _
                "    DFComplemento_TBFornecedor = '" & txtComplemento.Text & "' , " & _
                "    DFBairro_TBFornecedor =  '" & txtBairro.Text & "', " & _
                "    DFCidade_TBFornecedor = '" & txtCidade.Text & "' , " & _
                "    DFCep_TBFornecedor = '" & txtCep.Text & "' , " & _
                "    DFTipo_pessoa_TBFornecedor = '" & strGuardatipo & "' , " & _
                "    DFCgc_TBFornecedor = '" & txtCpfCnpj.Text & "', " & _
                "    DFInscricao_estadual_TBFornecedor = '" & txtInsestadual.Text & "'"
       Call funcoes_banco.Alterar("TBFornecedor", strSet, "PKCodigo_TBFornecedor", txtCodigo.Text, "PDV", Me, "BDSupervisor")
       log.Evento = "Alterar"
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Gravar_log "PDV", Me
    Else
       Call funcoes_banco.Gravar("TBFornecedor", strCampo, strvalores, "PDV", Me, "BDSupervisor")
       log.Evento = "Incluir Novo"
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
    
    'Excluindo Registro
    Call funcoes_banco.Excluir("TBFornecedor", "PKCodigo_TBFornecedor", txtCodigo.Text, "PDV", Me, "BDSupervivor")
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtCodigo.Text
    log.Tipo = 1
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
    
    
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = True
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Gravar_log "PDV", Me
    
    sstFornecedor.Tab = 1
    
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
    
    sstFornecedor.Tab = 0
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

Private Sub txtBairro_LostFocus()
    
    txtBairro.Text = UCase(txtBairro.Text)
    
End Sub

Private Sub txtCep_LostFocus()
    
    txtCep.Text = Format(txtCep.Text, "#####-###")
    
End Sub

Private Sub txtCidade_LostFocus()
    
    txtCidade.Text = UCase(txtCidade.Text)
    
End Sub

Private Sub txtCodigo_LostFocus()
    Movimentacoes.Verifica_Numero "PKCodigo_TBFornecedor", "TBFornecedor", txtCodigo, "PDV", Me
End Sub


Private Function Reposicao()
    
    On Error GoTo Erro

    Dim strSQL As String

    strSQL = "SELECT * FROM TBFornecedor"
    
    If txtConsulta.Text <> Empty Then
        strSQL = strSQL & " WHERE " & strCampo_consulta & " LIKE '" & txtConsulta.Text & "%' "
    End If

    Movimentacoes.Movimenta_HFlex_Grid strSQL, hfgFornecedor, "800,1000,2500,900,1500,1000,1500,1200,0,1500,2000", "Código,Nome,Endereço,Número,Complemento,Bairro,Cidade,CEP,Pessoa,CNPJ/CPF,Incrição Estadual", "BDSupervisor", "PDV", Me
     
    Exit Function

Erro:
    log.Evento = "Reposição"
    log.Tipo = 3
    log.Descricao = Err.Description
    log.Gravar_log "PDV", Me
    
    Call Erro.Erro(Me, "PDV", "Reposicao")
    Resume Next

End Function


Private Sub txtComplemento_lostfocus()
    
    txtComplemento.Text = UCase(txtComplemento.Text)
    
End Sub

Private Sub txtConsulta_Change()
   Call Reposicao
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

Private Sub txtEndereco_LostFocus()
    
      txtEndereco.Text = UCase(txtEndereco.Text)
      
End Sub

Private Sub txtNome_LostFocus()
    
    txtNome.Text = UCase(txtNome.Text)
    
End Sub

