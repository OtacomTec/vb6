VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMovimentacoes_Emissao_Nota_Cupom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Nota Cupom"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimentacoes_Emissao_Nota_Cupom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   11385
   Begin VB.Frame Frame4 
      Caption         =   "Informações da Nota"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2355
      Left            =   6810
      TabIndex        =   24
      Top             =   990
      Visible         =   0   'False
      Width           =   2955
      Begin VB.Label lblNotas_restantes 
         AutoSize        =   -1  'True
         Caption         =   "9999999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1980
         TabIndex        =   31
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Notas restantes..:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   1260
         Width           =   1770
      End
      Begin VB.Label lblContador_notas 
         AutoSize        =   -1  'True
         Caption         =   "lblContador_notas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1050
         TabIndex        =   29
         Top             =   1995
         Width           =   1785
      End
      Begin VB.Label lblNumero_Nota 
         AutoSize        =   -1  'True
         Caption         =   "9999999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1980
         TabIndex        =   28
         Top             =   330
         Width           =   840
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "N° Próxima Nota...:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   27
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número da Nota...:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   330
         Width           =   1800
      End
      Begin VB.Label lblProxima_Nota 
         AutoSize        =   -1  'True
         Caption         =   "9999999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1980
         TabIndex        =   25
         Top             =   780
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impressora Padrão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   90
      TabIndex        =   22
      Top             =   3390
      Width           =   6705
      Begin VB.CommandButton cmdImpressora_padrao 
         Height          =   465
         Left            =   6000
         Picture         =   "frmMovimentacoes_Emissao_Nota_Cupom.frx":1782
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione outra impressora"
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblImpressora_padrao 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   180
         TabIndex        =   23
         Top             =   330
         Width           =   5610
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   90
      TabIndex        =   10
      Top             =   990
      Width           =   6705
      Begin VB.TextBox txtCidade 
         Height          =   360
         Left            =   120
         MaxLength       =   5
         TabIndex        =   7
         ToolTipText     =   "Código da cidade"
         Top             =   1815
         Width           =   1245
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Código do Cliente"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txtLetra 
         Height          =   360
         Left            =   2310
         TabIndex        =   2
         Top             =   570
         Width           =   855
      End
      Begin VB.TextBox txtSerie 
         Height          =   360
         Left            =   1410
         TabIndex        =   1
         Top             =   570
         Width           =   855
      End
      Begin VB.TextBox txtNumero 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   570
         Width           =   1245
      End
      Begin MSDataListLib.DataCombo dtcCidade 
         Height          =   360
         Left            =   1410
         TabIndex        =   8
         Top             =   1815
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   360
         Left            =   1410
         TabIndex        =   6
         Top             =   1200
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   360
         Left            =   3210
         TabIndex        =   3
         Top             =   570
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
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
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   58720257
         CurrentDate     =   38229
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   360
         Left            =   5130
         TabIndex        =   4
         Top             =   570
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
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
         CustomFormat    =   "dd/MM/YYYY"
         Format          =   58720257
         CurrentDate     =   38229
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cidade"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1590
         Width           =   585
      End
      Begin VB.Label Label6 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label7 
         Caption         =   "Letra"
         Height          =   255
         Left            =   2310
         TabIndex        =   15
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Série"
         Height          =   255
         Left            =   1410
         TabIndex        =   14
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Número"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Período Emissão"
         Height          =   240
         Left            =   3210
         TabIndex        =   12
         Top             =   330
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "até"
         Height          =   240
         Left            =   4740
         TabIndex        =   11
         Top             =   690
         Width           =   270
      End
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   360
      Left            =   90
      TabIndex        =   18
      ToolTipText     =   "Empresa"
      Top             =   600
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   8940
      Top             =   4380
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
            Picture         =   "frmMovimentacoes_Emissao_Nota_Cupom.frx":1B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Cupom.frx":1E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Cupom.frx":2140
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Cupom.frx":24DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Cupom.frx":2874
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Emissao_Nota_Cupom.frx":2B8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1(0)"
      HotImageList    =   "ImageList1(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Confirmar"
            Object.ToolTipText     =   "Gravar registro - CTRL+G"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin CRVIEWER9LibCtl.CRViewer9 crvFiltrar 
      Height          =   510
      Left            =   90
      TabIndex        =   20
      Top             =   4380
      Width           =   7815
      lastProp        =   500
      _cx             =   13785
      _cy             =   900
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      Height          =   270
      Left            =   90
      TabIndex        =   21
      Top             =   360
      Width           =   1290
   End
End
Attribute VB_Name = "frmMovimentacoes_Emissao_Nota_Cupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador de Vendas                                         '
' Objetivo...............: Movimentação de Emissão de Nota Cupom                          '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Sá Peixoto                                               '
' Data de Criação........: 10/06/2006                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public strSql As String
Dim strImprime As Integer
Dim Conexao As New DLLConexao_Sistema.Conexao
Dim rstAplicacao As New ADODB.Recordset
Dim conexao_relatorio As New DLLConexao_Sistema.Conexao
Dim log As New DLLSystemManager.log
Option Explicit

Private Sub dtpFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos no dataPicker pelo ENTER
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos no dataPicker pelo ENTER
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    
    On Error GoTo erro
    
    'INFORMAÇÕES CONSTANTES PARA O LOG
    log.Usuario = MDIPrincipal.ocxUsuario.Nome
    log.Programa = "Emissão de Nota Cupom"
    log.Estacao = MDIPrincipal.ocxUsuario.Estacao
    
    'INFORMAÇÕES VARIAVEIS PARA O LOG
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.ocxUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    strSql = "SELECT * FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    dtcEmpresa.BoundText = MDIPrincipal.ocxUsuario.Empresa
    
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente WHERE IXCodigo_TBEmpresa = " & dtcEmpresa.BoundText & ""
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_Correios_TBCidade_otica,DFNome_TBCidade_otica FROM TBCidade_otica"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_Correios_TBCidade_otica", "DFNome_TBCidade_otica", dtcCidade, strSql, "BDRetaguarda", "Otica", Me

    dtpInicial.Value = Date
    dtpFinal.Value = Date
    
    log.Descricao = "Inicializando Emissão de Nota Cupom"
    'GRAVANDO LOG
    log.Gravar_log "Otica", Me
    
    Exit Sub
    
erro:
    Call erro.erro(Me, "Otica", "Load")
    Exit Sub

End Sub
