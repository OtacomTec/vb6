VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Flash de Vendas - Finalizadora"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   -30
      TabIndex        =   0
      Top             =   330
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
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
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dtcEcf"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dtcFinalizadora"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpFinal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtpInicial"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "mshfgFinalizadora"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkData"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkNumero_Ecf"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkFinalizadora"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCodigo_Finalizadora"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Análise Gráfica"
      TabPicture(1)   =   "frmFlash_Finalizadora.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSChart1"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
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
         Left            =   3390
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
         Left            =   2130
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
         Left            =   2130
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
         Left            =   120
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
         Left            =   120
         TabIndex        =   6
         Top             =   1380
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshfgFinalizadora 
         Height          =   2865
         Left            =   120
         TabIndex        =   5
         Top             =   2370
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5054
         _Version        =   393216
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
         Left            =   -74670
         TabIndex        =   2
         Top             =   570
         Width           =   6495
         Begin VB.OptionButton Option2 
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
         Begin VB.OptionButton Option1 
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
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   3255
         Left            =   -74850
         OleObjectBlob   =   "frmFlash_Finalizadora.frx":0038
         TabIndex        =   1
         Top             =   1650
         Width           =   6705
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   375
         Left            =   120
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
         Format          =   19529729
         CurrentDate     =   37788
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   375
         Left            =   1830
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
         Format          =   19529729
         CurrentDate     =   37788
      End
      Begin MSDataListLib.DataCombo dtcFinalizadora 
         Height          =   360
         Left            =   3390
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
         Left            =   120
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
         Left            =   1650
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
            Picture         =   "frmFlash_Finalizadora.frx":239F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":26B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":29D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":2D6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":3107
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlash_Finalizadora.frx":3421
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
      Width           =   7020
      _ExtentX        =   12383
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
Dim booAlterar As Boolean
Dim conexao As New DLLConexao_Sistema.conexao
Public log As New DLLSystemManager.log
Dim strSql As String

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
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
    log.Gravar_log ("PDV")
    
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    
    strSql = "SELECT TBFinalizadora.PKCodigo_TBFinalizadora,TBFinalizadora.DFDescricao_TBFinalizadora FROM TBFinalizadora"
    Call Movimentacoes.Movimenta_DataCombo("PKCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSql, "BDSupervisor", "PDV", Me)
    
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
    
    Call Objetos.Limpa_TXT(Me)
    dtcFinalizadora.Text = Empty
        
    log.Evento = "Novo"
    log.Descricao = "Solicitação de uma nova consulta"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log ("PDV")
    
    tlbBotoes.Buttons.Item(1).Enabled = False
    tlbBotoes.Buttons.Item(2).Enabled = True
    tlbBotoes.Buttons.Item(3).Enabled = True
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = False
        
    chkData.Enabled = True
    chkFinalizadora.Enabled = True
    chkNumero_Ecf.Enabled = True
        
    booAlterar = False
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
    log.Gravar_log ("PDV")
    
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
    log.Gravar_log ("PDV")
    
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
    log.Gravar_log ("PDV")
    
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
    
    strSql = "SELECT TBFinalizadora.PKCodigo_TBFinalizadora,TBFinalizadora.DFDescricao_TBFinalizadora," & _
             "SUM(TBVenda.DFValor_TBVenda) AS Tvenda,TBVenda.DFData_TBVenda " & _
             "INTO ##TBTEMP_CALCULO " & _
             "FROM TBVenda " & _
             "INNER JOIN TBFinalizadora_venda ON TBVenda.PKId_TBVenda = TBFinalizadora_Venda.FKId_TBVenda " & _
             "INNER JOIN TBFinalizadora ON TBFinalizadora_Venda.FKCodigo_TBFinalizadora = TBFinalizadora.PKCodigo_TBFinalizadora " & _
             "INNER JOIN TBEcf ON TBVenda.FKId_TBEcf = TBEcf.PKId_TBEcf " & _
             "WHERE TBVenda.DFCupom_Cancelado_TBVenda = 0 " & _
             "GROUP BY TBFinalizadora.PKCodigo_TBFinalizadora," & _
             "TBFinalizadora.DFDescricao_TBFinalizadora,TBVenda.DFData_TBVenda "
             
             
             
    conexao.Abrir_conexao ("PDV")
    
    conexao.CNConexao.Execute strSql
    
    strSql = Empty
    
    strSql = "SELECT ##tbtemp_calculo.*,TBFinalizadora_venda.DFValor_TBFinalizadora_venda," & _
             "SUM(TBFinalizadora_venda.DFValor_TBFinalizadora_venda * 100) / (SELECT SUM(Tvenda) AS TOTAL_VENDAS FROM ##TBTEMP_CALCULO) as Part " & _
             "INTO ##TBTEMP_CALCULO1 " & _
             "FROM ##TBTEMP_CALCULO " & _
             "INNER JOIN TBFinalizadora_Venda ON TBFinalizadora_Venda.FKCodigo_TBFinalizadora = ##TBTEMP_CALCULO.PKCodigo_TBFinalizadora " & _
             "GROUP BY ##TBTEMP_CALCULO.PKCodigo_TBFinalizadora," & _
             "##TBTEMP_CALCULO.Tvenda,##TBTEMP_CALCULO.DFDescricao_TBFinalizadora," & _
             "##TBTEMP_CALCULO.DFData_TBVenda,TBFinalizadora_venda.DFValor_TBFinalizadora_venda"
             
    conexao.CNConexao.Execute strSql
    
    strSql = Empty
    strSql = "SELECT * FROM ##TBTEMP_CALCULO1"
    
    If Me.chkNumero_Ecf.Value = 1 Then
       strSql = strSql & " AND TBEcf.DFNumero_TBEcf = " & txtNumero_Ecf.Text & ""
    End If
    
    If Me.chkData.Value = 1 Then
       strSql = strSql & " AND TBVenda.DFData_TBVenda >= '" & strData_Ini & "' AND TBVenda.DFData_TBVenda <= '" & strData_Fin & "'"
    End If
    
    If Me.chkFinalizadora.Value = 1 Then
       strSql = strSql & " AND TBFinalizadora_Venda.FKCodigo_TBFinalizadora = " & txtCodigo_Finalizadora.Text & ""
    End If
       
    strCampos_Grid = "COLUNA 1,Código,Finalizadora,Valor Total,Data,Valor,Participação %"
                                         
    strTamanhos_Campos_Grid = "0,1000,3000,1500,1100,0,1500"
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, mshfgFinalizadora, strTamanhos_Campos_Grid, strCampos_Grid, "BDSupervisor", "PDV", Me
    
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


