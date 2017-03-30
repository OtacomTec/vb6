VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmConsulta_Cupons 
   Caption         =   "Consulta de Cupons"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Consultar Por "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5355
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
         Left            =   2190
         TabIndex        =   11
         Top             =   1440
         Width           =   1905
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
         Left            =   2190
         TabIndex        =   10
         Top             =   540
         Width           =   1965
      End
      Begin VB.CheckBox chkNumero_Cupom 
         Caption         =   "Número do Cupom"
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
         Height          =   285
         Left            =   210
         TabIndex        =   9
         Top             =   570
         Width           =   1905
      End
      Begin VB.TextBox txtNumero_Cupom 
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
         Left            =   210
         MaxLength       =   10
         TabIndex        =   8
         Top             =   960
         Width           =   1875
      End
      Begin VB.TextBox txtNumero_Ecf 
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
         Left            =   2190
         MaxLength       =   10
         TabIndex        =   7
         Top             =   960
         Width           =   1875
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
         Left            =   210
         TabIndex        =   6
         Top             =   2310
         Width           =   1365
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
         Height          =   360
         Left            =   210
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2730
         Width           =   1155
      End
      Begin VB.TextBox txtValor 
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
         Left            =   210
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1830
         Width           =   1875
      End
      Begin VB.CheckBox chkValor 
         Caption         =   "Valor da Venda"
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
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   1500
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   375
         Left            =   2190
         TabIndex        =   12
         Top             =   1830
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   19660801
         CurrentDate     =   37788
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   1830
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   19660801
         CurrentDate     =   37788
      End
      Begin MSDataListLib.DataCombo dtcFinalizadora 
         Height          =   360
         Left            =   1470
         TabIndex        =   14
         Top             =   2730
         Width           =   3735
         _ExtentX        =   6588
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
         Left            =   3630
         TabIndex        =   15
         Top             =   1980
         Width           =   105
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5610
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
            Picture         =   "frmConsulta_Cupons.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsulta_Cupons.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsulta_Cupons.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsulta_Cupons.frx":09CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsulta_Cupons.frx":0D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsulta_Cupons.frx":1082
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5595
      _ExtentX        =   9869
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
   Begin MSDataGridLib.DataGrid adgConsulta 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   4260
      _Version        =   393216
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsulta_Cupons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logicx                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Supervisor de PDV                                              '
' Módulo.................: Cadastros                                                      '
' Objetivo...............: Consunta de Cupons                                             '
' Data de Criação........: 30/04/2003                                                     '
' Equipe Responsável.....: Giordano Vilela                                                '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim conexao As DLLConexao_Sistema.conexao
Public log As New DLLSystemManager.log
Dim strSql As String

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

Private Sub chkNumero_Cupom_Click()
    
    If chkNumero_Cupom.Value = 1 Then
       txtNumero_Cupom.Enabled = True
       txtNumero_Cupom.SetFocus
       txtNumero_Cupom.BackColor = &H80000005
    Else
       txtNumero_Cupom.Enabled = False
       txtNumero_Cupom.BackColor = &H8000000A
    End If
           
End Sub

Private Sub chkNumero_Ecf_Click()
    
    If chkNumero_Ecf.Value = 1 Then
       txtNumero_Ecf.Enabled = True
       txtNumero_Ecf.SetFocus
       txtNumero_Ecf.BackColor = &H80000005
    Else
       txtNumero_Ecf.Enabled = False
       txtNumero_Ecf.BackColor = &H8000000A
    End If
    
End Sub

Private Sub chkValor_Click()
    
    If chkValor.Value = 1 Then
       txtValor.Enabled = True
       txtValor.SetFocus
       txtValor.BackColor = &H80000005
    Else
       txtValor.Enabled = False
       txtValor.BackColor = &H8000000A
    End If
       
End Sub

Private Sub dtcFinalizadora_LostFocus()

    txtCodigo_Finalizadora.Text = dtcFinalizadora.BoundText
        
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
    log.Programa = "Consulta de Cupons"
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Descricao = "Inicializando a consulta de cupons"
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
    
    Me.Height = 4350
    Me.Width = 5700
       
    Call Objetos.Limpa_TXT(Me)
    dtcFinalizadora.Text = Empty
        
    log.Evento = "Novo"
    log.Descricao = "Solicitação de uma nova Consulta"
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
    chkNumero_Cupom.Enabled = True
    chkNumero_Ecf.Enabled = True
    chkValor.Enabled = True
    
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
    Me.dtpFinal.Enabled = False
    dtcFinalizadora.Text = Empty
    
    'Inserir log
    tlbBotoes.Buttons.Item(1).Enabled = True
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
            
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
       
    'Gravando Log
    log.Gravar_log ("PDV")
    
    chkData.Value = 0
    chkFinalizadora.Value = 0
    chkNumero_Cupom.Value = 0
    chkNumero_Ecf.Value = 0
    chkValor.Value = 0
    
    chkData.Enabled = False
    chkFinalizadora.Enabled = False
    chkNumero_Cupom.Enabled = False
    chkNumero_Ecf.Enabled = False
    chkValor.Enabled = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "PDV", "Cancelar")
    Exit Function

End Function

Private Sub txtCodigo_Finalizadora_LostFocus()

    dtcFinalizadora.BoundText = txtCodigo_Finalizadora.Text
    
End Sub
Private Function Consultar()
    
    On Error GoTo Erro
             
    Me.Width = 5715
    Me.Height = 6855
    
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
    
    strSql = "SELECT TBEcf.PKId_TBEcf,TBEcf.DFNumero_TBEcf,TBVenda.PKId_TBVenda,TBVenda.FKId_TBEcf," & _
             "TBVenda.DFNumero_TBCupom,TBVenda.DFData_TBVenda,TBVenda.DFValor_TBVenda," & _
             "TBFinalizadora_Venda.FKCodigo_TBFinalizadora, TBFinalizadora.PKCodigo_TBFinalizadora," & _
             "TBFinalizadora.DFDescricao_TBFinalizadora FROM TBEcf " & _
             "INNER JOIN TBVenda ON TBEcf.PKId_TBEcf = TBVenda.FKId_TBEcf " & _
             "INNER JOIN TBFinalizadora_Venda ON TBVenda.PKId_TBVenda = TBFinalizadora_Venda.FKId_TBVenda " & _
             "INNER JOIN TBFinalizadora ON TBFinalizadora_venda.FKCodigo_TBFinalizadora = TBFinalizadora.PKCodigo_TBFinalizadora"
             
                         
    If Me.chkNumero_Cupom.Value = 1 Then
       strSql = strSql & " WHERE TBVenda.DFNumero_TBCupom = " & txtNumero_Cupom.Text & ""
    End If
    
    If Me.chkNumero_Ecf.Value = 1 Then
       strSql = strSql & " AND TBEcf.DFNumero_TBEcf = " & txtNumero_Ecf.Text & ""
    End If
    
    If Me.chkValor.Value = 1 Then
       strSql = strSql & " AND TBVenda.DFValor_TBVenda = " & txtValor.Text & ""
    End If
    
    If Me.chkData.Value = 1 Then
       strSql = strSql & " AND TBVenda.DFData_TBVenda >= '" & strData_Ini & "' AND TBVenda.DFData_TBVenda <= '" & strData_Fin & "'"
    End If
    
    If Me.chkFinalizadora.Value = 1 Then
       strSql = strSql & " AND TBFinalizadora_Venda.FKCodigo_TBFinalizadora = " & txtCodigo_Finalizadora.Text & ""
    End If
       
    strCampos_Grid = "ID,Número do ECF,ID,ID,Número do Cupom,Data da Venda,Valor da Venda," & _
                     "Código,COD,Finalizadora"
                                          
    strTamanhos_Campos_Grid = "0,2000,0,0,2000,2000,2000,1000,0,3000"
    
    Movimentacoes.Movimenta_Data_Grid strSql, adgConsulta, strTamanhos_Campos_Grid, strCampos_Grid, "BDSupervisor", "PDV", Me
    
    Call Objetos.Limpa_TXT(Me)
    chkData.Value = 0
    chkFinalizadora.Value = 0
    chkNumero_Cupom.Value = 0
    chkNumero_Ecf.Value = 0
    chkValor.Value = 0
    
    chkData.Enabled = False
    chkFinalizadora.Enabled = False
    chkNumero_Cupom.Enabled = False
    chkNumero_Ecf.Enabled = False
    chkValor.Enabled = False
    
    Exit Function

Erro:
    Call Erro.Erro(Me, "PDV", "Reposicao")
    Resume Next

End Function

