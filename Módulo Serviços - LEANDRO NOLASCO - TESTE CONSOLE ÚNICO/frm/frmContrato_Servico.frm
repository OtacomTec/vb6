VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContrato_Servico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contrato de Serviços"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContrato_Servico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7800
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Integração"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8490
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrato_Servico.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrato_Servico.frx":1A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrato_Servico.frx":1DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrato_Servico.frx":2150
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrato_Servico.frx":24EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrato_Servico.frx":2804
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrato_Servico.frx":2B1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sstContrato_servico 
      Height          =   3915
      Left            =   0
      TabIndex        =   20
      Top             =   330
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   6906
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
      TabPicture(0)   =   "frmContrato_Servico.frx":2EB8
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label37"
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(7)=   "Label16"
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(11)=   "lblObservacao"
      Tab(0).Control(12)=   "dtpData_Contrato"
      Tab(0).Control(13)=   "dtcCliente"
      Tab(0).Control(14)=   "dtpData_Envio"
      Tab(0).Control(15)=   "dtpData_Validade"
      Tab(0).Control(16)=   "dtcPlano_servico"
      Tab(0).Control(17)=   "dtcBanco"
      Tab(0).Control(18)=   "cbbTabela"
      Tab(0).Control(19)=   "txtCliente"
      Tab(0).Control(20)=   "txtNumero_contrato"
      Tab(0).Control(21)=   "txtDesconto"
      Tab(0).Control(22)=   "txtBanco"
      Tab(0).Control(23)=   "txtPlano_Servico"
      Tab(0).Control(24)=   "txtValor_Plano"
      Tab(0).Control(25)=   "txtValor_Contrato"
      Tab(0).Control(26)=   "cmdInformacaoes_Adicionais"
      Tab(0).Control(27)=   "txtObservacao"
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "&Listagem"
      TabPicture(1)   =   "frmContrato_Servico.frx":2ED4
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblA"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dtpFim_Consulta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "dtpInicio_Consulta"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cbbCampos"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "hfgContrato_servico"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtConsulta"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdRefresh"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdConsulta"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtObservacao 
         Height          =   360
         Left            =   -71880
         MaxLength       =   100
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         ToolTipText     =   "Observação"
         Top             =   3330
         Width           =   4485
      End
      Begin VB.CommandButton cmdInformacaoes_Adicionais 
         Height          =   360
         Left            =   -67740
         Picture         =   "frmContrato_Servico.frx":2EF0
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Informações Adicionais"
         Top             =   2070
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   6900
         Picture         =   "frmContrato_Servico.frx":327A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   7290
         Picture         =   "frmContrato_Servico.frx":4F74
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   2370
         TabIndex        =   15
         Top             =   780
         Width           =   4485
      End
      Begin VB.TextBox txtValor_Contrato 
         Height          =   360
         Left            =   -73290
         TabIndex        =   12
         ToolTipText     =   "Valor do Contrato"
         Top             =   3330
         Width           =   1365
      End
      Begin VB.TextBox txtValor_Plano 
         Enabled         =   0   'False
         Height          =   360
         Left            =   -69030
         TabIndex        =   36
         ToolTipText     =   "Valor do Plano de Serviços"
         Top             =   2070
         Width           =   1245
      End
      Begin VB.TextBox txtPlano_Servico 
         Height          =   360
         Left            =   -74880
         MaxLength       =   5
         TabIndex        =   7
         ToolTipText     =   "Código Plano de Serviços"
         Top             =   2070
         Width           =   1545
      End
      Begin VB.TextBox txtBanco 
         Height          =   360
         Left            =   -74880
         MaxLength       =   5
         TabIndex        =   9
         ToolTipText     =   "Código Banco"
         Top             =   2700
         Width           =   1545
      End
      Begin VB.TextBox txtDesconto 
         Height          =   360
         Left            =   -74880
         TabIndex        =   11
         ToolTipText     =   "Percentual de Desconto"
         Top             =   3330
         Width           =   1545
      End
      Begin VB.TextBox txtNumero_contrato 
         Enabled         =   0   'False
         Height          =   360
         Left            =   -74880
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Número Contrato"
         Top             =   780
         Width           =   1545
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Left            =   -74880
         TabIndex        =   5
         ToolTipText     =   "Código do Cliente"
         Top             =   1410
         Width           =   1545
      End
      Begin AutoCompletar.CbCompleta cbbTabela 
         Height          =   360
         Left            =   -68580
         TabIndex        =   4
         ToolTipText     =   "Tabela"
         Top             =   780
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
      End
      Begin MSDataListLib.DataCombo dtcBanco 
         Height          =   360
         Left            =   -73290
         TabIndex        =   10
         ToolTipText     =   "Descrição Banco"
         Top             =   2700
         Width           =   5925
         _ExtentX        =   10451
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
      Begin MSDataListLib.DataCombo dtcPlano_servico 
         Height          =   360
         Left            =   -73290
         TabIndex        =   8
         ToolTipText     =   "Descrição Plano de Serviços"
         Top             =   2070
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSComCtl2.DTPicker dtpData_Validade 
         Height          =   360
         Left            =   -70140
         TabIndex        =   3
         ToolTipText     =   "Data de Validade"
         Top             =   780
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20381697
         CurrentDate     =   37858
      End
      Begin MSComCtl2.DTPicker dtpData_Envio 
         Height          =   360
         Left            =   -71790
         TabIndex        =   2
         ToolTipText     =   "Data de Envio do Certificado"
         Top             =   780
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20381697
         CurrentDate     =   37858
      End
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   360
         Left            =   -73290
         TabIndex        =   6
         ToolTipText     =   "Razão Social do Cliente"
         Top             =   1410
         Width           =   5925
         _ExtentX        =   10451
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
      Begin MSComCtl2.DTPicker dtpData_Contrato 
         Height          =   360
         Left            =   -73290
         TabIndex        =   1
         ToolTipText     =   "Data de Contrato"
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20381697
         CurrentDate     =   37858
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgContrato_servico 
         Height          =   2565
         Left            =   120
         TabIndex        =   33
         Top             =   1230
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   4524
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   120
         TabIndex        =   14
         Top             =   780
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
      End
      Begin MSComCtl2.DTPicker dtpInicio_Consulta 
         Height          =   360
         Left            =   2370
         TabIndex        =   16
         Top             =   780
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20381697
         CurrentDate     =   38386
      End
      Begin MSComCtl2.DTPicker dtpFim_Consulta 
         Height          =   360
         Left            =   5250
         TabIndex        =   17
         Top             =   780
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         Format          =   20381697
         CurrentDate     =   38386
      End
      Begin VB.Label lblObservacao 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   240
         Left            =   -71880
         TabIndex        =   38
         Top             =   3090
         Width           =   1005
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         Height          =   240
         Left            =   4530
         TabIndex        =   35
         Top             =   900
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data Contrato"
         Height          =   240
         Left            =   -73290
         TabIndex        =   31
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor Contrato"
         Height          =   240
         Left            =   -73290
         TabIndex        =   30
         Top             =   3090
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor Plano"
         Height          =   240
         Left            =   -69030
         TabIndex        =   29
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Data Validade"
         Height          =   240
         Left            =   -70140
         TabIndex        =   28
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data Envio Cert."
         Height          =   240
         Left            =   -71790
         TabIndex        =   27
         Top             =   540
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco Cobrança"
         Height          =   240
         Left            =   -74880
         TabIndex        =   26
         Top             =   2460
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plano de Serviço"
         Height          =   240
         Left            =   -74880
         TabIndex        =   25
         Top             =   1830
         Width           =   1425
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "% Desconto"
         Height          =   240
         Left            =   -74880
         TabIndex        =   24
         Top             =   3090
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   240
         Left            =   -74880
         TabIndex        =   23
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tabela"
         Height          =   240
         Left            =   -68580
         TabIndex        =   22
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   240
         Left            =   -74880
         TabIndex        =   21
         Top             =   1170
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmContrato_Servico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro Contrato Serviços                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Peixoto                                                  '
' Data de Criação........: 04/03/2005                                                     '
' Desenvolvedor..........: Jones Sá Peixoto                                               '
' Data última manutenção.: 06/10/2005                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strNomes As String
Dim strTamanho As String
Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Dim booAlterar As Boolean
Dim conexao As New DLLConexao_Sistema.conexao
'Declaração das variaveis da acessibilidade
Public strSql As String
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
Dim CNConexao As New DLLConexao_Sistema.conexao
Dim rstVerifica_Titulo As New ADODB.Recordset
Option Explicit

Function Imprimir()
    On Error GoTo Erro
    'Tratamento de erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Contrato_Servico.Show
    
    Unload frmAguarde
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty
    
    If cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       dtpFim_Consulta.Visible = False
       dtpInicio_Consulta.Visible = False
       lblA.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Data Validade" Or cbbCampos.Text = "Data Envio" Then
       txtConsulta.Visible = False
       dtpFim_Consulta.Visible = True
       dtpInicio_Consulta.Visible = True
       lblA.Visible = True
    Else
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If

End Sub

Private Sub cbbTabela_LostFocus()
    If txtPlano_Servico.Text <> Empty Then
       Call dtcPlano_servico_LostFocus
    End If
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdInformacaoes_Adicionais_Click()
    Unload frmContrato_Servico_Informacoes_Adicionais
    If txtPlano_Servico.Text <> Empty Then
       frmContrato_Servico_Informacoes_Adicionais.Show
    End If
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub dtcPlano_servico_GotFocus()
    If cbbTabela.Text = Empty Then
       MsgBox "Tabela de Preços inválida. Verifique.", vbInformation, "Only Tech"
       txtPlano_Servico.Text = Empty
       cbbTabela.SetFocus
       Exit Sub
    Else
       If txtPlano_Servico.Text = Empty Then
          Call Movimentacoes.Verifica_DataCombo(dtcPlano_servico.Text)
       End If
    End If
End Sub

Private Sub dtcPlano_servico_LostFocus()

    txtPlano_Servico.Text = dtcPlano_servico.BoundText
    If IsNumeric(txtPlano_Servico.Text) = False Or dtcPlano_servico.Text = Empty Then
       txtPlano_Servico.Text = Empty
       txtValor_Plano.Text = Empty
       Exit Sub
    End If
    
    strSql = "SELECT SUM(DFPreco1_conveniado_TBServico_laboratorio) AS SOMA_1," & _
             "SUM(DFPreco2_conveniado_TBServico_laboratorio) AS SOMA_2," & _
             "SUM(DFPreco3_conveniado_TBServico_laboratorio) AS SOMA_3 " & _
             "FROM TBPlano_servico_servico_laboratorio " & _
             "INNER JOIN TBPlano_servico " & _
             "ON TBPlano_servico_servico_laboratorio.FKCodigo_TBPlano_servico = TBPlano_servico.PKCodigo_TBPlano_servico " & _
             "INNER JOIN TBServico_laboratorio ON TBPlano_servico_servico_laboratorio.FKCodigo_TBServico_laboratorio = TBServico_laboratorio.PKCodigo_TBServico_laboratorio " & _
             "WHERE PKCodigo_TBPlano_servico = " & txtPlano_Servico.Text & " "
             
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       If Not IsNull(rstAplicacao.Fields("SOMA_1")) And cbbTabela.Text = 1 Then
          txtValor_Plano.Text = Format(rstAplicacao.Fields("SOMA_1"), "#,###0.00")
       ElseIf Not IsNull(rstAplicacao.Fields("SOMA_2")) And cbbTabela.Text = 2 Then
          txtValor_Plano.Text = Format(rstAplicacao.Fields("SOMA_2"), "#,###0.00")
       ElseIf Not IsNull(rstAplicacao.Fields("SOMA_3")) And cbbTabela.Text = 3 Then
          txtValor_Plano.Text = Format(rstAplicacao.Fields("SOMA_3"), "#,###0.00")
       Else
          txtValor_Plano.Text = Empty
       End If
    Else
      txtValor_Plano.Text = Empty
    End If
    
    Set rstAplicacao = Nothing
    
End Sub

Private Sub dtcBanco_GotFocus()
    If txtBanco.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcBanco.Text)
    End If
End Sub

Private Sub dtcBanco_LostFocus()
    txtBanco.Text = dtcBanco.BoundText
    If IsNumeric(txtBanco.Text) = False Or dtcBanco.Text = Empty Then txtBanco.Text = Empty: Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Teclas de Atalho da TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 78: If booPrivilegio_Incluir = True Then Call Novo     'CTRL+N
                       Case 71: If booPrivilegio_Incluir = True Then Call Gravar   'CTRL+G
                       Case 67: If booPrivilegio_Incluir = True Then Call Cancelar 'CTRL+C
                       Case 69: If booPrivilegio_Excluir = True Then Call Excluir  'CTRL+E
                       Case 83: Unload Me  'CTRL+S
                End Select
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
  
Private Sub Form_Load()
    On Error GoTo Erro
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Contrato de Veículo"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstContrato_servico, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "OTICA", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o Cadastro de Contratos de Veículos"
    'Gravando o log
    log.Gravar_log "OTICA", Me
    
    Call Reposicao
    
    dtpData_Contrato.Value = Date
    dtpData_Validade.Value = Date
    dtpData_Envio.Value = Date
    
    dtpInicio_Consulta.Value = Date
    dtpFim_Consulta.Value = Date + 7
    
    sstContrato_servico.TabEnabled(0) = False
    sstContrato_servico.Tab = 1
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    log.Descricao = "Finalizando o Cadastro de Contratos"
    
    Set log = Nothing
    
    strCombo = Empty
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgContrato_servico_Click()

    If hfgContrato_servico.Col = 0 And hfgContrato_servico.Text <> Empty Then
    
       'Novo
       tlbBotoes.Buttons.Item(1).Enabled = False
       'Gravar
       tlbBotoes.Buttons.Item(2).Enabled = booPrivilegio_Alterar
       'Cancelar
       tlbBotoes.Buttons.Item(3).Enabled = booPrivilegio_Alterar
       'Excluir
       tlbBotoes.Buttons.Item(4).Enabled = booPrivilegio_Excluir
       'Imprimir
       tlbBotoes.Buttons.Item(5).Enabled = False
       'Integração
       If booIntegra_Portal = True Then
          tlbBotoes.Buttons.Item(9).Enabled = True
       End If
        
       frmAguarde.Show
       DoEvents
        On Error Resume Next
        
       txtNumero_contrato.Text = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 1))
       txtCliente.Text = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 2))
       txtPlano_Servico.Text = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 4))
       txtValor_Contrato.Text = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 6))
       txtDesconto.Text = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 7))
       cbbTabela.Text = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 8))
       txtBanco.Text = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 9))
       dtpData_Contrato.Value = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 11))
       dtpData_Envio.Value = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 12))
       dtpData_Validade.Value = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 13))
       txtObservacao.Text = hfgContrato_servico.TextArray((hfgContrato_servico.Row * hfgContrato_servico.Cols + hfgContrato_servico.Col + 14))
       
       Call dtcPlano_servico_LostFocus
       
       booAlterar = True
       txtConsulta.Text = Empty
       Unload frmAguarde
       txtNumero_contrato.Enabled = False
       sstContrato_servico.Tab = 0
       sstContrato_servico.TabEnabled(0) = True
       
       dtpData_Contrato.SetFocus
   End If

End Sub

Private Sub hfgContrato_servico_DblClick()
    hfgContrato_servico.Sort = 1
End Sub

Private Sub hfgContrato_servico_KeyPress(KeyAscii As Integer)
    'Retorno do grid com espaço
    If KeyAscii = 32 Then
       Call hfgContrato_servico_Click
    End If
End Sub

Private Sub sstContrato_servico_Click(PreviousTab As Integer)
    If sstContrato_servico.Tab = 0 Then
       If txtNumero_contrato.Enabled = True Then
          txtNumero_contrato.SetFocus
       Else
          dtpData_Validade.SetFocus
       End If
    ElseIf sstContrato_servico.Tab = 1 Then
       If frmIntegracao.Visible = True Then
          Unload frmIntegracao
       End If
       If strCombo <> Empty And strCombo <> "Todos" Then
          cbbCampos.Text = strCombo
          If txtConsulta.Visible = True Then
             txtConsulta.SetFocus
          Else
             dtpInicio_Consulta.SetFocus
          End If
       ElseIf strCombo = "Todos" Then
            hfgContrato_servico.Row = 1
            hfgContrato_servico.Col = 0
            hfgContrato_servico.SetFocus
       End If
    End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstContrato_servico.Tab = 0: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
    End Select
End Sub

Function Gravar()
    On Error GoTo Erro
    
    'Verifica se os campos necessarios para gravar não estão nulos
    If txtNumero_contrato.Text = Empty Then
       MsgBox "O campo número do contrato não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtNumero_contrato.SetFocus
       Exit Function
    ElseIf txtCliente.Text = Empty Then
       MsgBox "O campo código do cliente não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtCliente.SetFocus
       Exit Function
    ElseIf txtPlano_Servico.Text = Empty Then
       MsgBox "O campo código do plano de serviço não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtPlano_Servico.SetFocus
       Exit Function
    ElseIf txtBanco.Text = Empty Then
       MsgBox "O campo código do banco não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtBanco.SetFocus
       Exit Function
    End If

    Dim intContador As Integer
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim strId_Cliente As String

    Call Objetos.Maiusculo_TXT(Me)

    strId_Cliente = Funcoes_Gerais.Localiza_ID("PKId_TBCliente", "IXCodigo_TBCliente", txtCliente.Text, "TBCliente", "OTICA", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", MDIPrincipal.OCXUsuario.Empresa)
   
    strCampo = "PKCodigo_TBContrato_cliente,FKId_TBCliente,FKCodigo_TBPlano_servico,DFValor_TBContrato_cliente, " & _
               "DFDesconto_TBContrato_cliente,DFTabela_preco_TBContrato_cliente,DFBanco_TBContrato_cliente, " & _
               "DFData_envio_certificado_TBContrato_cliente,DFData_validade_TBContrato_cliente,DFObservacao_TBContrato_cliente, " & _
               "DFData_contrato_TBContato_cliente,DFData_alteracao_TBContrato_cliente," & _
               "DFIntegrado_filiais_TBContrato_cliente "
               
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBContrato_cliente "
    End If
    
    strValores = "" & txtNumero_contrato.Text & "," & strId_Cliente & "," & txtPlano_Servico.Text & "," & Funcoes_Gerais.Grava_Moeda(txtValor_Contrato.Text) & "," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtDesconto.Text) & "," & cbbTabela.Text & ", " & _
                 "" & txtBanco.Text & ",'" & Format(dtpData_Envio.Value, "YYYYMMDD") & "','" & Format(dtpData_Validade.Value, "YYYYMMDD") & "', " & _
                 "'" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "', " & _
                 "'" & Format(dtpData_Contrato.Value, "YYYYMMDD") & "'," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0 "
    
    If booIntegra_Portal = True Then
        strValores = strValores & ",0 "
    End If
    
    If booAlterar = True Then
        
       On Error GoTo Erro_alteracao
       
       CNConexao.Abrir_conexao "OTICA"
       CNConexao.CNConexao.BeginTrans
        
       log.Evento = "Alterar"
       strSql = "UPDATE TBContrato_cliente SET FKId_TBCliente = " & strId_Cliente & "," & _
                "FKCodigo_TBPlano_servico = " & txtPlano_Servico.Text & "," & _
                "DFValor_TBContrato_cliente = " & Funcoes_Gerais.Grava_Moeda(txtValor_Contrato.Text) & ", " & _
                "DFDesconto_TBContrato_cliente = " & Funcoes_Gerais.Grava_Moeda(txtDesconto.Text) & "," & _
                "DFTabela_preco_TBContrato_cliente = " & cbbTabela.Text & "," & _
                "DFBanco_TBContrato_cliente = " & txtBanco.Text & ", " & _
                "DFData_envio_certificado_TBContrato_cliente = '" & Format(dtpData_Envio.Value, "YYYYMMDD") & "'," & _
                "DFData_validade_TBContrato_cliente = '" & Format(dtpData_Validade.Value, "YYYYMMDD") & "'," & _
                "DFObservacao_TBContrato_cliente = '" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "', " & _
                "DFData_contrato_TBContato_cliente = '" & Format(dtpData_Contrato.Value, "YYYYMMDD") & "'," & _
                "DFData_alteracao_TBContrato_cliente = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBContrato_cliente = 0 "
                
       If booIntegra_Portal = True Then
          strSql = strSql & ",DFIntegrado_portal_TBContrato_cliente = 0 "
       End If
       
       strSql = strSql & "WHERE PKCodigo_TBContrato_cliente = " & txtNumero_contrato.Text & ""
           
       CNConexao.CNConexao.Execute strSql

       CNConexao.CNConexao.CommitTrans
       'fim da gravação das alterações nas guias
       CNConexao.Fechar_conexao
       
       log.Descricao = "Alterando o registro: " + txtNumero_contrato.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       On Error GoTo Erro
       
       log.Evento = "Incluir Novo"
       Call funcoes_banco.Gravar("TBContrato_cliente", strCampo, strValores, "OTICA", Me, "BDRetaguarda")
       
       log.Descricao = "Gravando o registro: " + txtNumero_contrato.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
       
    End If
    
    Call Objetos.Limpa_TXT(Me)
           
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgContrato_servico.Visible = False
    End If
    
    sstContrato_servico.TabEnabled(0) = False
    sstContrato_servico.Tab = 1
    
    Exit Function
    
Erro_alteracao:
    CNConexao.CNConexao.RollbackTrans
    
    CNConexao.Fechar_conexao
 
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()

    'Excluindo Registro
    Call funcoes_banco.Excluir("TBContrato_cliente", "PKCodigo_TBContrato_cliente", txtNumero_contrato.Text, "Otica", Me, "BDRetaguarda")
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtNumero_contrato.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    Call Objetos.Limpa_TXT(Me)
    cbbTabela.Text = Empty
    
    'Novo
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = False
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = False
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = False
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       Me.hfgContrato_servico.Visible = False
    End If
    
    sstContrato_servico.TabEnabled(0) = False
    sstContrato_servico.Tab = 1
    txtNumero_contrato.Enabled = False
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Excluir")
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    
    'Novo
     tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = booPrivilegio_Excluir
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       Me.hfgContrato_servico.Visible = False
    End If
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    txtNumero_contrato.Enabled = False
    
    dtpData_Contrato.Value = Date
    dtpData_Validade.Value = Date
    dtpData_Envio.Value = Date
    
    sstContrato_servico.TabEnabled(0) = False
    sstContrato_servico.Tab = 1

    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro

    Call Objetos.Limpa_TXT(Me)
    cbbTabela.Text = Empty
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    'Novo
    tlbBotoes.Buttons.Item(1).Enabled = False
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = booPrivilegio_Incluir
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = booPrivilegio_Incluir
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = False
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = False
     
    booAlterar = False
        
    txtNumero_contrato.Enabled = True

    sstContrato_servico.TabEnabled(0) = True
    sstContrato_servico.TabEnabled(1) = True
    sstContrato_servico.Tab = 0
    
    dtpData_Contrato.Value = Date
    dtpData_Validade.Value = Date
    dtpData_Envio.Value = Date
    

    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    On Error GoTo Erro
    
    strTamanho = "1200,1200,3500,1200,2600,1550,1000," & _
                 "1000,1150,2200,1200,1200,1200,2600"
    
    strNomes = "Nº Contrato,Cod. Cliente,Cliente,Cod. Plano,Plano,Valor," & _
               "Desconto,Tabela,Cod. Banco,Banco,Dt. Contrato,Dt. Envio," & _
               "Dt. Validade,Observação"
    
    Movimentacoes.Monta_HFlex_Grid hfgContrato_servico, strTamanho, strNomes, 14, "OTICA", Me
        
    Call Monta_Combo
    
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "OTICA", Me
    
    strSql = "SELECT PKCodigo_TBBancos,DFNome_TBBancos FROM TBBancos"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBBancos", "DFNome_TBBancos", dtcBanco, strSql, "BDRetaguarda", "OTICA", Me

    strSql = "SELECT PKCodigo_TBPlano_servico,DFDescricao_TBPlano_servico FROM TBPlano_servico"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBPlano_servico", "DFDescricao_TBPlano_servico", dtcPlano_servico, strSql, "BDRetaguarda", "OTICA", Me
    
    strSql = Empty
    Exit Function

Erro:
    Call Erro.Erro(Me, "OTICA", "Reposicao")
    Resume Next
End Function

Private Function Consulta()

    If cbbCampos.Text = Empty Or cbbCampos.Text <> "Todos" And dtpInicio_Consulta.Visible = False Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbInformation, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
    
    If cbbCampos.Text = "Data Envio" Or cbbCampos.Text = "Data Validade" Then
       If dtpInicio_Consulta.Value > dtpFim_Consulta.Value Then
          MsgBox "Data final menor que data Inicial. Verifique.", vbInformation, "Only Tech"
          dtpInicio_Consulta.SetFocus
          Exit Function
       End If
    End If
   
    strSql = "SELECT PKCodigo_TBContrato_cliente,IXCodigo_TBCliente,DFNome_TBCliente," & _
             "FKCodigo_TBPlano_servico," & _
             "DFDescricao_TBPlano_servico," & _
             "DFValor_TBContrato_cliente," & _
             "DFDesconto_TBContrato_cliente," & _
             "DFTabela_preco_TBContrato_cliente," & _
             "DFBanco_TBContrato_cliente," & _
             "DFNome_TBBancos," & _
             "DFData_contrato_TBContato_cliente," & _
             "DFData_envio_certificado_TBContrato_cliente," & _
             "DFData_validade_TBContrato_cliente," & _
             "DFObservacao_TBContrato_cliente " & _
             "FROM TBContrato_cliente " & _
             "INNER JOIN TBPlano_servico " & _
             "ON TBContrato_cliente.FKCodigo_TBPlano_servico = TBPlano_servico.PKCodigo_TBPlano_servico " & _
             "INNER JOIN TBBancos " & _
             "ON TBContrato_cliente.DFBanco_TBContrato_cliente = TBBancos.PKCodigo_TBBancos " & _
             "INNER JOIN TBCliente " & _
             "ON TBContrato_cliente.FKId_TBCliente = TBCliente.PKId_TBCliente "
                              
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    txtConsulta.Text = Funcoes_Gerais.Grava_String(txtConsulta.Text)
    
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Número Contrato" Then
          strSql = strSql & " WHERE convert(nvarchar,PKCodigo_TBContrato_cliente) = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Cod.Cliente" Then
          strSql = strSql & " WHERE convert(nvarchar,IXCodigo_TBCliente) = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Cliente" Then
          strSql = strSql & " WHERE convert(nvarchar,DFNome_TBCliente) LIKE '%" & Funcoes_Gerais.Grava_String(txtConsulta) & "%' "
       ElseIf cbbCampos.Text = "Cod.Plano" Then
          strSql = strSql & " WHERE convert(nvarchar,FKCodigo_TBPlano_servico) = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Plano" Then
          strSql = strSql & " WHERE convert(nvarchar,DFDescricao_TBPlano_servico) = " & Funcoes_Gerais.Grava_String(txtConsulta) & " "
       ElseIf cbbCampos.Text = "Valor do Contrato" Then
          strSql = strSql & " WHERE convert(money,DFValor_TBContrato_cliente) = " & Funcoes_Gerais.Grava_Moeda(txtConsulta) & " "
       ElseIf cbbCampos.Text = "Desconto" Then
          strSql = strSql & " WHERE convert(money,DFDesconto_TBContrato_cliente) = " & Funcoes_Gerais.Grava_Moeda(txtConsulta) & " "
       ElseIf cbbCampos.Text = "Tabela" Then
          strSql = strSql & " WHERE DFTabela_preco_TBContrato_cliente = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Cod.Banco" Then
          strSql = strSql & " WHERE convert(nvarchar,DFBanco_TBContrato_cliente) = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Banco" Then
          strSql = strSql & " WHERE convert(nvarchar,DFNome_TBBancos) LIKE '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Data Envio" Then
          strSql = strSql & " WHERE DFData_envio_certificado_TBContrato_cliente >= '" & Format(dtpInicio_Consulta.Value, "YYYYMMDD") & "' " & _
                  " AND DFData_envio_certificado_TBContrato_cliente <= '" & Format(dtpFim_Consulta.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Data Validade" Then
          strSql = strSql & " WHERE DFData_validade_TBContrato_cliente >= '" & Format(dtpInicio_Consulta.Value, "YYYYMMDD") & "' " & _
                  " AND DFData_validade_TBContrato_cliente <= '" & Format(dtpFim_Consulta.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Observação" Then
          strSql = strSql & " WHERE convert(nvarchar,DFObservacao_TBContrato_cliente) LIKE '%" & txtConsulta.Text & "%' "
       End If
    End If
    
    frmAguarde.Show
    DoEvents
                                           
    strSql = strSql & "ORDER BY PKCodigo_TBContrato_cliente"
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgContrato_servico, strTamanho, strNomes, "BDRetaguarda", "OTICA", Me
    
    hfgContrato_servico.Row = 1
    hfgContrato_servico.Col = 0
    If hfgContrato_servico.Text = Empty Then
       hfgContrato_servico.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgContrato_servico, strTamanho, strNomes, 14, "OTICA", Me
    End If
    
    hfgContrato_servico.Refresh
    
    Unload frmAguarde
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Número Contrato")
    cbbCampos.AddItem ("Cod.Cliente")
    cbbCampos.AddItem ("Cliente")
    cbbCampos.AddItem ("Cod.Plano")
    cbbCampos.AddItem ("Plano")
    cbbCampos.AddItem ("Valor do Contrato")
    cbbCampos.AddItem ("Desconto")
    cbbCampos.AddItem ("Cod.Banco")
    cbbCampos.AddItem ("Banco")
    cbbCampos.AddItem ("Data Validade")
    cbbCampos.AddItem ("Data Envio")
    cbbCampos.AddItem ("Observação")

    cbbTabela.Clear
    cbbTabela.AddItem ("1")
    cbbTabela.AddItem ("2")
    cbbTabela.AddItem ("3")
    
End Function

Private Sub txtBanco_Change()
    dtcBanco.BoundText = txtBanco.Text
    If IsNumeric(txtBanco.Text) = False Then txtBanco.Text = Empty: Exit Sub
End Sub

Private Sub txtBanco_LostFocus()
    If dtcBanco.Text = Empty Then txtBanco.Text = Empty
End Sub

Private Sub txtCliente_Change()
    dtcCliente.BoundText = txtCliente.Text
    If IsNumeric(txtCliente.Text) = False Then txtCliente.Text = Empty: Exit Sub
End Sub

Private Sub txtCliente_LostFocus()
    If dtcCliente.Text = Empty Then txtCliente.Text = Empty
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
    cmdConsulta.SetFocus
End Sub

Private Sub txtPlano_servico_Change()
    dtcPlano_servico.BoundText = txtPlano_Servico.Text
    If IsNumeric(txtPlano_Servico.Text) = False Then txtPlano_Servico.Text = Empty: Exit Sub
End Sub

Private Sub txtPlano_servico_GotFocus()
    If cbbTabela.Text = Empty Then
       MsgBox "Tabela de Preços inválida. Verifique.", vbInformation, "Only Tech"
       txtPlano_Servico.Text = Empty
       cbbTabela.SetFocus
       Exit Sub
    End If
    
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPlano_servico_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtBanco_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtCliente_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtNumero_contrato_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_contrato_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtDesconto_GotFocus()
    If txtValor_Plano.Text = Empty Then
       MsgBox "Plano de Serviços inválido. Verifique.", vbInformation, "Only Tech"
       txtDesconto.Text = Empty
       txtPlano_Servico.SetFocus
       Exit Sub
    End If
    
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtDesconto_LostFocus()
    txtDesconto.Text = Format(txtDesconto.Text, "#,###0.00")
    If txtDesconto.Text <> Empty Then
       If CDbl(txtDesconto.Text) > 100 Then
          MsgBox "Desconto superior ao valor do contrato. Verifique.", vbInformation, "Only Tech"
          txtDesconto.SetFocus
       Else
          txtValor_Contrato.Text = Format(CDbl(txtValor_Plano.Text) - CDbl(txtValor_Plano.Text) * CDbl(txtDesconto.Text) / 100, "#,###0.00")
       End If
    End If
End Sub

Private Sub txtObservacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtObservacao_LostFocus()
    txtObservacao.Text = UCase(txtObservacao.Text)
End Sub

Private Sub txtNumero_contrato_LostFocus()
    Movimentacoes.Verifica_Numero "PKCodigo_TBContrato_cliente", "TBContrato_cliente", txtNumero_contrato, "Otica", Me
End Sub

Private Sub txtPlano_Servico_LostFocus()
    Call dtcPlano_servico_LostFocus
End Sub

Private Sub dtcCliente_GotFocus()
    If txtCliente.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCliente.Text)
    End If
End Sub

Private Sub dtcCliente_LostFocus()
    txtCliente.Text = dtcCliente.BoundText
    If IsNumeric(txtCliente.Text) = False Or dtcCliente.Text = Empty Then txtCliente.Text = Empty: Exit Sub
End Sub

Private Sub txtValor_Contrato_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtValor_Contrato_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Contrato_LostFocus()
    txtValor_Contrato.Text = Format(txtValor_Contrato.Text, "#,###0.00")
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKCodigo_TBContrato_cliente", txtNumero_contrato.Text, "DFIntegrado_filiais_TBContrato_cliente", "TBContrato_cliente", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBContrato_cliente", Me.Top, Me.Left, Me.Width, Me.Height, "Contrato Serviço")
    
End Function

