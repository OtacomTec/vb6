VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Flash de Vendas - Finalizadora"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7185
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12674
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Análise Geral"
      TabPicture(0)   =   "frmFlash_Finalizadora.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "dtcEcf"
      Tab(0).Control(2)=   "dtcFinalizadora"
      Tab(0).Control(3)=   "dtpFinal"
      Tab(0).Control(4)=   "dtpInicial"
      Tab(0).Control(5)=   "mshfgFinalizadora"
      Tab(0).Control(6)=   "chkData"
      Tab(0).Control(7)=   "chkNumero_Ecf"
      Tab(0).Control(8)=   "chkFinalizadora"
      Tab(0).Control(9)=   "txtCodigo_Finalizadora"
      Tab(0).Control(10)=   "Frame2"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Análise Gráfica"
      TabPicture(1)   =   "frmFlash_Finalizadora.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "mscPizza"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "mscBarra"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   " Classificar "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -71610
         TabIndex        =   14
         Top             =   1470
         Width           =   3495
         Begin VB.OptionButton optDecrescente 
            Caption         =   "Decrescente"
            Height          =   195
            Left            =   2070
            TabIndex        =   16
            Top             =   360
            Width           =   1245
         End
         Begin VB.OptionButton optCrescente 
            Caption         =   "Crescente"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.TextBox txtCodigo_Finalizadora 
         BackColor       =   &H8000000A&
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
         Left            =   -72870
         MaxLength       =   10
         TabIndex        =   9
         Top             =   920
         Width           =   1155
      End
      Begin VB.CheckBox chkFinalizadora 
         Caption         =   "Finalizadora"
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
         Left            =   -72870
         TabIndex        =   8
         Top             =   540
         Width           =   1365
      End
      Begin VB.CheckBox chkNumero_Ecf 
         Caption         =   "Número do ECF"
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
         Height          =   345
         Left            =   -74880
         TabIndex        =   7
         Top             =   540
         Width           =   1965
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Período"
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
         TabIndex        =   6
         Top             =   1380
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshfgFinalizadora 
         Height          =   4635
         Left            =   -74880
         TabIndex        =   5
         Top             =   2370
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8176
         _Version        =   393216
         FixedCols       =   0
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
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame Frame1 
         Caption         =   "Gráficos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   300
         TabIndex        =   2
         Top             =   420
         Width           =   6495
         Begin VB.OptionButton optBarra 
            Caption         =   "Barra"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   5250
            TabIndex        =   4
            Top             =   285
            Width           =   1065
         End
         Begin VB.OptionButton optPizza 
            Caption         =   "Pizza"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   300
            TabIndex        =   3
            Top             =   350
            Width           =   1725
         End
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   1770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19595265
         CurrentDate     =   37788
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   375
         Left            =   -73170
         TabIndex        =   11
         Top             =   1770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19595265
         CurrentDate     =   37788
      End
      Begin MSDataListLib.DataCombo dtcFinalizadora 
         Height          =   360
         Left            =   -71610
         TabIndex        =   12
         Top             =   915
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483638
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
      Begin MSDataListLib.DataCombo dtcEcf 
         Height          =   360
         Left            =   -74880
         TabIndex        =   17
         Top             =   915
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483638
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
      Begin MSChart20Lib.MSChart mscBarra 
         DragIcon        =   "frmFlash_Finalizadora.frx":0038
         Height          =   5865
         Left            =   210
         OleObjectBlob   =   "frmFlash_Finalizadora.frx":0342
         TabIndex        =   1
         Top             =   1230
         Width           =   6705
      End
      Begin MSChart20Lib.MSChart mscPizza 
         Height          =   5625
         Left            =   120
         OleObjectBlob   =   "frmFlash_Finalizadora.frx":209D
         TabIndex        =   19
         Top             =   1470
         Width           =   6705
      End
      Begin VB.Label Label1 
         Caption         =   "a"
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
         Left            =   -73350
         TabIndex        =   13
         Top             =   1920
         Width           =   105
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7020
      Top             =   600
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
            Picture         =   "frmFlash_Finalizadora.frx":43DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":46F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":4A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":4DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":5144
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":545E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "alt + N"
            Description     =   "Novo"
            Object.ToolTipText     =   "Nova Consulta - CTRL+N"
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
            Object.ToolTipText     =   "Consultar registro - CTRL+C"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar Consulta"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Imprimir"
            Object.ToolTipText     =   "Imprimir - CTRL+I"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Flash de Vendas - Finalizadora                                 '
' Data de Criação........: 01/07/2003                                                     '
' Equipe Responsável.....: Giordano Vilela                                                '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim conexao As New DLLConexao_Sistema.conexao
Public log As New DLLSystemManager.log
Dim strSQL As String
Private Sub chkData_Click()

    If chkData.Value = 1 Then
       dtpInicial.Enabled = True
       dtpFinal.Enabled = True
       dtpInicial.SetFocus
       dtpInicial.CalendarBackColor = &H80000005
       dtpFinal.CalendarBackColor = &H80000005
    Else
       dtpInicial.Enabled = False
       dtpFinal.Enabled = False
       dtpInicial.CalendarBackColor = &H8000000A
       dtpFinal.CalendarBackColor = &H8000000A
    End If
           
End Sub
Private Sub chkFinalizadora_Click()
    
    If chkFinalizadora.Value = 1 Then
       txtCodigo_Finalizadora.Enabled = True
       dtcFinalizadora.Enabled = True
       txtCodigo_Finalizadora.SetFocus
       txtCodigo_Finalizadora.BackColor = &H80000005
       dtcFinalizadora.BackColor = &H80000005
    Else
       txtCodigo_Finalizadora.Enabled = False
       txtCodigo_Finalizadora.BackColor = &H8000000A
       dtcFinalizadora.Enabled = False
       dtcFinalizadora.BackColor = &H8000000A
    End If
        
End Sub
Private Sub chkNumero_Ecf_Click()
    
    If chkNumero_Ecf.Value = 1 Then
       dtcEcf.Enabled = True
       dtcEcf.SetFocus
       dtcEcf.BackColor = &H80000005
    Else
       dtcEcf.Enabled = False
       dtcEcf.BackColor = &H8000000A
    End If
    
End Sub
Private Sub dtcFinalizadora_LostFocus()

    txtCodigo_Finalizadora.Text = dtcFinalizadora.BoundText
        
End Sub

Private Sub optPizza_Click()
    
    mscPizza.Visible = True
    
    Call Monta_Graficos
    
End Sub

Private Sub optBarra_Click()

    mscBarra.Visible = True
    
    Call Monta_Graficos
    
End Sub

Private Sub txtCodigo_Finalizadora_LostFocus()

    dtcFinalizadora.BoundText = txtCodigo_Finalizadora.Text
    
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
    log.Programa = "Consulta por Finalizadoras"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando a consulta por finalizadoras"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando o log
    log.Gravar_log "PDV", Me
    
    Me.SSTab1.Tab = 0
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    
    'Coloca o Grafico Invisivel
    mscBarra.Visible = False
    mscPizza.Visible = False
    
    strSQL = "SELECT TBFinalizadora.PKCodigo_TBFinalizadora,TBFinalizadora.DFDescricao_TBFinalizadora FROM TBFinalizadora"
    Call Movimentacoes.Movimenta_DataCombo("PKCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSQL, "BDSupervisor", "PDV", Me)
    
    Exit Sub
    
Erro:
    Call Erro.Erro(Me, "PDV", "Load")
    Exit Sub
    
    
End Sub
Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Consultar
           Case 3: Call Cancelar
           'Case 4: Call Imprimir
           Case 6: Unload Me
    End Select
    
End Sub
Private Function Novo()

    On Error GoTo Erro
    
    'este "if" serve para verificar o estado da conexao, pois como esta consulta utiliza-se de
    'arquivos temporarios esta verificacao se faz necessaria para garantir que esses arquivos
    'sejam criados do "zero" quanado solicitada uma nova consulta. Giordano
    
    If conexao.CNConexao.State = 1 Then
       conexao.Fechar_conexao
    End If
        
    Call Objetos.Limpa_TXT(Me)
    dtcFinalizadora.Text = Empty
    optPizza.Value = False
    optBarra.Value = False
        
    log.Evento = "Novo"
    log.Descricao = "Solicitação de uma nova consulta"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    tlbBotoes.Buttons.Item(1).Enabled = False
    tlbBotoes.Buttons.Item(2).Enabled = True
    tlbBotoes.Buttons.Item(3).Enabled = True
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = False
        
    mscBarra.Visible = False
    mscPizza.Visible = False
    mshfgFinalizadora.Clear
    
    chkData.Enabled = True
    chkFinalizadora.Enabled = True
    chkNumero_Ecf.Enabled = True
        
    Exit Function
    
Erro:

    Call Erro.Erro(Me, "PDV", "Novo")
    Exit Function

End Function
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

Private Function Cancelar()

    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    dtpInicial.Enabled = False
    dtpFinal.Enabled = False
    dtcFinalizadora.Text = Empty
    optPizza.Value = False
    optBarra.Value = False
    
    'Inserir log
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
            
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Consulta"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
       
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    chkData.Value = 0
    chkFinalizadora.Value = 0
    chkNumero_Ecf.Value = 0
        
    chkData.Enabled = False
    chkFinalizadora.Enabled = False
    chkNumero_Ecf.Enabled = False
       
    Exit Function
    
Erro:
    Call Erro.Erro(Me, "PDV", "Cancelar")
    Exit Function

End Function
Private Function Consultar()
    
    On Error GoTo Erro
             
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = True
    
    'Inserir log
    log.Evento = "Consulta"
    log.Descricao = "Consulta de registro(s)"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
       
    'Gravando Log
    log.Gravar_log "PDV", Me
    
    Call Reposicao
    
    Exit Function
    
Erro:
    Call Erro.Erro(Me, "PDV", "Consulta")
    Exit Function

End Function

Private Function Reposicao()

    On Error GoTo Erro

    Dim strCampos_Grid As String
    Dim strTamanhos_Campos_Grid As String
    Dim strData_Ini As String
    Dim strData_Fin As String
    
    strData_Ini = Format(dtpInicial.Value, "YYYYMMDD")
    strData_Fin = Format(dtpFinal.Value, "YYYYMMDD")
    
    strSQL = "SELECT TBFinalizadora.PKCodigo_TBFinalizadora,TBFinalizadora.DFDescricao_TBFinalizadora," & _
             "SUM(TBVenda.DFValor_TBVenda) AS Tvenda " & _
             "INTO ##TBTEMP_CALCULO " & _
             "FROM TBVenda " & _
             "INNER JOIN TBFinalizadora_venda ON TBVenda.PKId_TBVenda = TBFinalizadora_Venda.FKId_TBVenda " & _
             "INNER JOIN TBFinalizadora ON TBFinalizadora_Venda.FKCodigo_TBFinalizadora = TBFinalizadora.PKCodigo_TBFinalizadora " & _
             "INNER JOIN TBEcf ON TBVenda.FKId_TBEcf = TBEcf.PKId_TBEcf " & _
             "WHERE TBVenda.DFCupom_Cancelado_TBVenda = 0 "
             
    If dtcEcf.Text <> Empty Then
       strSQL = strSQL & " AND TBEcf.DFNumero_TBEcf = " & dtcEcf.Text & ""
    End If
    
    If txtCodigo_Finalizadora.Text <> Empty Then
       strSQL = strSQL & " AND TBFinalizadora_Venda.FKCodigo_TBFinalizadora = " & txtCodigo_Finalizadora.Text & ""
    End If
    
    If Me.chkData.Value = 1 And dtpInicial.Value <> Empty Then
       strSQL = strSQL & " AND TBVenda.DFData_TBVenda >= '" & strData_Ini & "' AND TBVenda.DFData_TBVenda <= '" & strData_Fin & "'"
    End If
    
    strSQL = strSQL & " GROUP BY TBFinalizadora.PKCodigo_TBFinalizadora,TBFinalizadora.DFDescricao_TBFinalizadora "
    
    conexao.Initial_Catalog = "BDSupervisor"
    
    conexao.Abrir_conexao ("PDV")
    
    conexao.CNConexao.Execute strSQL
    
    strSQL = Empty
    
    strSQL = "SELECT ##tbtemp_calculo.*," & _
             "SUM(TBFinalizadora_venda.DFValor_TBFinalizadora_venda * 100) / (SELECT SUM(Tvenda) AS TOTAL_VENDAS FROM ##TBTEMP_CALCULO) as Part " & _
             "INTO ##TBTEMP_CALCULO1 " & _
             "FROM ##TBTEMP_CALCULO " & _
             "INNER JOIN TBFinalizadora_Venda ON TBFinalizadora_Venda.FKCodigo_TBFinalizadora = ##TBTEMP_CALCULO.PKCodigo_TBFinalizadora " & _
             "GROUP BY ##TBTEMP_CALCULO.PKCodigo_TBFinalizadora," & _
             "##TBTEMP_CALCULO.DFDescricao_TBFinalizadora,##TBTEMP_CALCULO.Tvenda "
            
    conexao.CNConexao.Execute strSQL
    
    strSQL = Empty
    
    strSQL = "SELECT * FROM ##TBTEMP_CALCULO1"
    
    strCampos_Grid = "Código,Finalizadora,Valor Total,Participação %"
                                         
    strTamanhos_Campos_Grid = "1000,3000,1500,1500"
    
    Movimentacoes.Movimenta_HFlex_Grid strSQL, mshfgFinalizadora, strTamanhos_Campos_Grid, strCampos_Grid, "BDSupervisor", "PDV", Me
    
    'For I = 1 To rs.RecordCount
    '    Retorno = Me.mshfgCandidatos.TextArray(I * mshfgCandidatos.Cols + 3)
    '    If Retorno < 50 Then
    '       mshfgCandidatos.Row = I
    '       Me.mshfgCandidatos.Col = 3
    '       Me.mshfgCandidatos.CellForeColor = vbWhite
    '       Me.mshfgCandidatos.CellBackColor = vbRed
    '    End If
    'Next I
        
    Call Objetos.Limpa_TXT(Me)
    optPizza.Value = False
    optBarra.Value = False
    chkData.Value = 0
    chkFinalizadora.Value = 0
    chkNumero_Ecf.Value = 0
       
    chkData.Enabled = False
    chkFinalizadora.Enabled = False
    chkNumero_Ecf.Enabled = False
       
    Exit Function

Erro:
    Call Erro.Erro(Me, "PDV", "Reposicao")
    Resume Next

End Function
Private Function Monta_Graficos()

    Dim L As Integer
    Dim C As Integer
    Dim intValores As Integer
    Dim max_colunas As Integer
    Dim max_linhas As Integer
    Dim rstGrafico As New ADODB.Recordset
    Dim conexao_Grafico As New DLLConexao_Sistema.conexao
    Dim Values() As String
    
    conexao_Grafico.Initial_Catalog = "BDSupervisor"
    conexao_Grafico.Abrir_conexao ("PDV")
       
    rstGrafico.CursorLocation = adUseClient
    rstGrafico.Open strSQL, conexao_Grafico.CNConexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    'Impede a selecao das partes do grafico
    mscBarra.AllowSelections = False
    mscPizza.AllowSelections = False
    
    mscBarra.Enabled = False
    mscPizza.Enabled = False
    
    If optBarra.Value = True Then
        
        Dim intValor_Max As Integer
        Dim intLabel As Integer
        
        mscPizza.Visible = False
        
       'Abilita a aparicao dos "ROW's" no eixo x
       mscBarra.Plot.Axis(VtChAxisIdX).AxisScale.Hide = False
       
       max_colunas = rstGrafico.Fields.Count
       max_linhas = rstGrafico.RecordCount
       
       ReDim Values(1 To max_linhas, 1 To max_colunas)
       rstGrafico.MoveFirst
       
       L = 1
       C = 1
       
       Do While rstGrafico.EOF = False
          C = 1
          Values(L, C) = rstGrafico.Fields("DFDescricao_TBFinalizadora") & " " & rstGrafico.Fields("Tvenda")
          C = C + 1
          Values(L, C) = rstGrafico.Fields("Tvenda")
          If rstGrafico.Fields("Tvenda") > intValor_Max Then
             intValor_Max = rstGrafico.Fields("Tvenda")
          End If
          rstGrafico.MoveNext
          L = L + 1
       Loop
             
            
       mscBarra.chartType = VtChChartType2dBar
       mscBarra.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = intValor_Max
       mscBarra.ChartData = Values
       mscBarra.Title = "Finalizadoras por receita"
       mscBarra.Plot.Axis(VtChAxisIdX).AxisTitle.Text = "Finalizadoras"
       mscBarra.Plot.Axis(VtChAxisIdY).AxisTitle.Text = "Valores R$"
       mscBarra.Plot.Axis(VtChAxisIdY).AxisTitle.TextLayout.VertAlignment = VtVerticalAlignmentCenter
       
       mscBarra.Plot.SeriesCollection(1).LegendText = "Finalizadoras"
       mscBarra.Plot.SeriesCollection(2).Position.Hidden = True
       mscBarra.Plot.SeriesCollection(3).Position.Hidden = True
       
       conexao_Grafico.Fechar_conexao
       
    Else
        mscBarra.Visible = False
        'desabilita a aparicao dos "ROW's" no eixo x
        mscPizza.Plot.Axis(VtChAxisIdX).AxisScale.Hide = True
        
        max_linhas = rstGrafico.RecordCount
        
        ReDim Values(1 To max_linhas)
        rstGrafico.MoveFirst
        
        L = 1
        C = 1
        Do While rstGrafico.EOF = False
           Values(L) = rstGrafico.Fields("Part")
           rstGrafico.MoveNext
           L = L + 1
        Loop
        
        mscPizza.RowCount = rstGrafico.RecordCount
        
        mscPizza.ChartData = Values
        mscPizza.chartType = VtChChartType2dPie
        mscPizza.Title = "Finalizadora por percentual de participação"
        
        mscPizza.RowLabel = Empty
        
        rstGrafico.MoveFirst
        For L = 1 To mscPizza.Plot.SeriesCollection.Count
           If rstGrafico.EOF = False Then
              mscPizza.Plot.SeriesCollection(L).Position.Hidden = False
              mscPizza.Plot.SeriesCollection(L).LegendText = rstGrafico.Fields(1) & " " & "(" & rstGrafico.Fields(3) & "%" & ")"
              rstGrafico.MoveNext
           Else
              Exit For
           End If
        Next L
        
        conexao_Grafico.Fechar_conexao
               
    End If

End Function
