VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSolicitacao_Visita 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitação Visita"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSolicitacao_Visita.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7680
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
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
   Begin TabDlg.SSTab sstSolicitacao_Visita 
      Height          =   5805
      Left            =   0
      TabIndex        =   32
      Top             =   330
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   10239
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "&Geral"
      TabPicture(0)   =   "frmSolicitacao_Visita.frx":1782
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label20"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label13"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbbStatus"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtpHora_Solicitacao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtpData_Solicitacao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCodigo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtContato_Solicitacao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtValor_Orcamento"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCondicao_Pagamento"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtObservacao"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "&Agenda"
      TabPicture(1)   =   "frmSolicitacao_Visita.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtTelefone_Agenda"
      Tab(1).Control(1)=   "txtContato_Agenda"
      Tab(1).Control(2)=   "cmdRemover_Agenda"
      Tab(1).Control(3)=   "cmdIncluir_Agenda"
      Tab(1).Control(4)=   "hfgAgenda"
      Tab(1).Control(5)=   "Label25"
      Tab(1).Control(6)=   "Label26"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "A&tendimento"
      TabPicture(2)   =   "frmSolicitacao_Visita.frx":17BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdIncluir_Atendimento"
      Tab(2).Control(1)=   "cmdRemover_Atendimento"
      Tab(2).Control(2)=   "txtMotivo_Atendimento"
      Tab(2).Control(3)=   "txtVendedor"
      Tab(2).Control(4)=   "dtpData_Previsao"
      Tab(2).Control(5)=   "hfgAtendimento"
      Tab(2).Control(6)=   "dtcVendedor"
      Tab(2).Control(7)=   "dtpData_Atendimento"
      Tab(2).Control(8)=   "Label12"
      Tab(2).Control(9)=   "Label29"
      Tab(2).Control(10)=   "Label28"
      Tab(2).Control(11)=   "Label27"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "&Listagem"
      TabPicture(3)   =   "frmSolicitacao_Visita.frx":17D6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblA"
      Tab(3).Control(1)=   "Label6"
      Tab(3).Control(2)=   "dtpHora_Final"
      Tab(3).Control(3)=   "cbbConsulta"
      Tab(3).Control(4)=   "cbbCampos"
      Tab(3).Control(5)=   "dtpInicial"
      Tab(3).Control(6)=   "dtpFinal"
      Tab(3).Control(7)=   "hfgSolicitacao_Visita"
      Tab(3).Control(8)=   "cmdOrdenar"
      Tab(3).Control(9)=   "cmdConsulta"
      Tab(3).Control(10)=   "cmdRefresh"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "txtConsulta"
      Tab(3).Control(12)=   "dtpHora_Inicial"
      Tab(3).ControlCount=   13
      Begin MSComCtl2.DTPicker dtpHora_Inicial 
         Height          =   360
         Left            =   -72660
         TabIndex        =   66
         Top             =   780
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   51445762
         CurrentDate     =   37923
      End
      Begin VB.TextBox txtObservacao 
         Height          =   375
         Left            =   120
         MaxLength       =   100
         TabIndex        =   14
         ToolTipText     =   "Limite de Crédito"
         Top             =   5250
         Width           =   7440
      End
      Begin VB.TextBox txtCondicao_Pagamento 
         Height          =   375
         Left            =   120
         MaxLength       =   100
         TabIndex        =   13
         ToolTipText     =   "Limite de Crédito"
         Top             =   4620
         Width           =   7410
      End
      Begin VB.TextBox txtValor_Orcamento 
         Height          =   360
         Left            =   5280
         TabIndex        =   12
         ToolTipText     =   "Limite de Atraso (em dias)"
         Top             =   3960
         Width           =   2265
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   -72660
         TabIndex        =   27
         Top             =   780
         Width           =   3945
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   -67860
         Picture         =   "frmSolicitacao_Visita.frx":17F2
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   -68250
         Picture         =   "frmSolicitacao_Visita.frx":2834
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdOrdenar 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68640
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Ordenar: (A) Alfabética/ (C) Código"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtTelefone_Agenda 
         Height          =   375
         Left            =   -71460
         TabIndex        =   16
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtContato_Agenda 
         Height          =   375
         Left            =   -74880
         MaxLength       =   40
         TabIndex        =   15
         Top             =   780
         Width           =   3375
      End
      Begin VB.CommandButton cmdRemover_Agenda 
         Caption         =   "Remover"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68670
         TabIndex        =   18
         Top             =   780
         Width           =   1185
      End
      Begin VB.CommandButton cmdIncluir_Agenda 
         Caption         =   "Incluir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69930
         TabIndex        =   17
         Top             =   780
         Width           =   1185
      End
      Begin VB.TextBox txtContato_Solicitacao 
         Height          =   375
         Left            =   1095
         MaxLength       =   40
         TabIndex        =   3
         ToolTipText     =   "Razão Social"
         Top             =   780
         Width           =   3630
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Código do Cliente(Informado Automaticamente)"
         Top             =   780
         Width           =   915
      End
      Begin VB.CommandButton cmdIncluir_Atendimento 
         Caption         =   "Incluir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69930
         TabIndex        =   24
         Top             =   1410
         Width           =   1185
      End
      Begin VB.CommandButton cmdRemover_Atendimento 
         Caption         =   "Remover"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68670
         TabIndex        =   25
         Top             =   1410
         Width           =   1185
      End
      Begin VB.TextBox txtMotivo_Atendimento 
         Height          =   360
         Left            =   -74880
         MaxLength       =   50
         TabIndex        =   23
         Top             =   1410
         Width           =   4875
      End
      Begin VB.TextBox txtVendedor 
         Height          =   360
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   19
         Top             =   780
         Width           =   915
      End
      Begin MSComCtl2.DTPicker dtpData_Solicitacao 
         Height          =   375
         Left            =   4785
         TabIndex        =   0
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   51445761
         CurrentDate     =   37858
      End
      Begin MSComCtl2.DTPicker dtpData_Previsao 
         Height          =   360
         Left            =   -70290
         TabIndex        =   21
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
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
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   51445761
         CurrentDate     =   37858
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgAtendimento 
         Height          =   3825
         Left            =   -74880
         TabIndex        =   38
         Top             =   1860
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   6747
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame4 
         Caption         =   "Endereçamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   39
         Top             =   1230
         Width           =   7410
         Begin VB.TextBox txtUf 
            Enabled         =   0   'False
            Height          =   360
            Left            =   5130
            MaxLength       =   10
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Unidade Federativa"
            Top             =   1890
            Width           =   405
         End
         Begin VB.CommandButton cmdLogradouro 
            Height          =   300
            Left            =   6870
            Picture         =   "frmSolicitacao_Visita.frx":452E
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1920
            Width           =   375
         End
         Begin VB.TextBox txtEndereco 
            Height          =   375
            Left            =   120
            MaxLength       =   40
            TabIndex        =   4
            ToolTipText     =   "Endereço"
            Top             =   570
            Width           =   7125
         End
         Begin VB.TextBox txtNumero 
            Height          =   375
            Left            =   120
            MaxLength       =   10
            TabIndex        =   5
            ToolTipText     =   "Número"
            Top             =   1230
            Width           =   1275
         End
         Begin VB.TextBox txtComplemento 
            Height          =   375
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   6
            ToolTipText     =   "Complemento do Endereço"
            Top             =   1230
            Width           =   2505
         End
         Begin VB.TextBox txtBairro 
            Height          =   375
            Left            =   3990
            MaxLength       =   30
            TabIndex        =   7
            ToolTipText     =   "Bairro"
            Top             =   1230
            Width           =   3255
         End
         Begin VB.TextBox txtCep 
            Height          =   360
            Left            =   5580
            MaxLength       =   10
            TabIndex        =   10
            ToolTipText     =   "CEP"
            Top             =   1890
            Width           =   1695
         End
         Begin VB.TextBox txtCodigo_Cidade 
            Height          =   360
            Left            =   120
            MaxLength       =   5
            TabIndex        =   8
            ToolTipText     =   "Código da Cidade"
            Top             =   1890
            Width           =   1005
         End
         Begin MSDataListLib.DataCombo dtcCidade 
            Height          =   360
            Left            =   1170
            TabIndex        =   9
            ToolTipText     =   "Nome da Cidade"
            Top             =   1890
            Width           =   3930
            _ExtentX        =   6932
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
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   240
            Left            =   5130
            TabIndex        =   29
            Top             =   1650
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   240
            Left            =   120
            TabIndex        =   30
            Top             =   330
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   240
            Left            =   150
            TabIndex        =   31
            Top             =   990
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Complemento"
            Height          =   240
            Left            =   1470
            TabIndex        =   43
            Top             =   990
            Width           =   1185
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   240
            Left            =   3990
            TabIndex        =   42
            Top             =   990
            Width           =   510
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   240
            Left            =   120
            TabIndex        =   41
            Top             =   1650
            Width           =   585
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   240
            Left            =   5580
            TabIndex        =   40
            Top             =   1650
            Width           =   330
         End
      End
      Begin MSComCtl2.DTPicker dtpHora_Solicitacao 
         Height          =   375
         Left            =   6225
         TabIndex        =   1
         ToolTipText     =   "Data de Recadastramento"
         Top             =   780
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   51445762
         CurrentDate     =   38477
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgAgenda 
         Height          =   4425
         Left            =   -74880
         TabIndex        =   49
         Top             =   1260
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   7805
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSolicitacao_Visita 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   54
         Top             =   1230
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   7858
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   360
         Left            =   -70320
         TabIndex        =   55
         Top             =   780
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   51445761
         CurrentDate     =   37923
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   360
         Left            =   -72660
         TabIndex        =   56
         Top             =   780
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   51445761
         CurrentDate     =   37923
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   -74880
         TabIndex        =   26
         Top             =   780
         Width           =   2175
         _ExtentX        =   3836
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
      Begin AutoCompletar.CbCompleta cbbStatus 
         Height          =   360
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Bloqueado"
         Top             =   3960
         Width           =   5115
         _ExtentX        =   9022
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
      Begin MSDataListLib.DataCombo dtcVendedor 
         Height          =   360
         Left            =   -73920
         TabIndex        =   20
         ToolTipText     =   "Nome da Cidade"
         Top             =   780
         Width           =   3600
         _ExtentX        =   6350
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
      Begin MSComCtl2.DTPicker dtpData_Atendimento 
         Height          =   360
         Left            =   -68850
         TabIndex        =   22
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
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
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   51445761
         CurrentDate     =   37858
      End
      Begin AutoCompletar.CbCompleta cbbConsulta 
         Height          =   360
         Left            =   -72660
         TabIndex        =   64
         Top             =   780
         Visible         =   0   'False
         Width           =   3945
         _ExtentX        =   6959
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
      Begin MSComCtl2.DTPicker dtpHora_Final 
         Height          =   360
         Left            =   -70320
         TabIndex        =   65
         Top             =   780
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         Format          =   51445762
         CurrentDate     =   37923
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Dt. Atendimento"
         Height          =   240
         Left            =   -68850
         TabIndex        =   63
         Top             =   540
         Width           =   1380
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   240
         Left            =   120
         TabIndex        =   62
         Top             =   5010
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Condição Pagamento"
         Height          =   240
         Left            =   120
         TabIndex        =   61
         Top             =   4380
         Width           =   1800
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Valor Orçamento"
         Height          =   240
         Left            =   5295
         TabIndex        =   60
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   240
         Left            =   120
         TabIndex        =   59
         Top             =   3720
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   -74880
         TabIndex        =   58
         Top             =   540
         Width           =   435
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "até"
         Height          =   240
         Left            =   -70800
         TabIndex        =   57
         Top             =   930
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Contato"
         Height          =   240
         Left            =   -74880
         TabIndex        =   51
         Top             =   540
         Width           =   660
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   240
         Left            =   -71460
         TabIndex        =   50
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contato Solicitação"
         Height          =   240
         Left            =   1095
         TabIndex        =   47
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   120
         TabIndex        =   46
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Hora Solicitação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6225
         TabIndex        =   45
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Data Previsão"
         Height          =   240
         Left            =   -70290
         TabIndex        =   37
         Top             =   540
         Width           =   1170
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
         Height          =   240
         Left            =   -74880
         TabIndex        =   36
         Top             =   1170
         Width           =   555
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         Height          =   240
         Left            =   -74880
         TabIndex        =   35
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Data Solicitação"
         Height          =   240
         Left            =   4800
         TabIndex        =   34
         Top             =   540
         Width           =   1365
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7770
      Top             =   360
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
            Picture         =   "frmSolicitacao_Visita.frx":54F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolicitacao_Visita.frx":580A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolicitacao_Visita.frx":5B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolicitacao_Visita.frx":5EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolicitacao_Visita.frx":6258
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolicitacao_Visita.frx":6572
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSolicitacao_Visita.frx":688C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSolicitacao_Visita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro de Solicitação de Visita                              '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rafael Gomes                                                   '
' Data de Criação........: 05/05/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strTamanho As String
Dim strNomes As String
Public strCombo As String
Dim strConsulta As String
Dim strSolicita As String
Dim intAgenda As Integer
Dim intAtendimento As Integer
Dim strID_Agenda As String
Dim strID_Atendimento As String
Dim I As Integer
Dim Conexao_Visita As New DLLConexao_Sistema.conexao
''''''''''''''''''''''''''''''''''''''''''''
Public strSql As String
Public booAlterar As Boolean
'''''''''''''''' Vetores '''''''''''''''''
'Declaração das variaveis da acessibilidade
Dim conexao As New DLLConexao_Sistema.conexao
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log
''''''''''''''''''''''''''''''''''''''''''''''
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Option Explicit

Function Imprimir()
    On Error GoTo Erro
    'Tratamento de Erro
    If strSql = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique!", vbInformation, "Only Tech"
       cbbCampos.SetFocus
       Exit Function
    End If
    
    frmAguarde.Show
    DoEvents
    
    Call frmConsole_Relatorio_Solicitacao_Visita.Show
    
    Unload frmAguarde
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Sub cbbCampos_Click()
    
    'se Mudou de campo
    If cbbCampos.Text <> strCombo Then
       strCombo = cbbCampos.Text
    End If
            
    'se o campo "strCombo" for nulo
    If strCombo <> Empty Then
       cbbCampos.Text = strCombo
    End If
    
    'se o campo "cbbcampos" for nulo
    If cbbCampos.Text = Empty Then
       cbbCampos.SetFocus
       Exit Sub
    End If
        
    txtConsulta.Text = Empty
    cbbConsulta.Text = Empty
   
    If cbbCampos.Text = "Todos" Then
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       dtpHora_Inicial.Visible = False
       dtpHora_Final.Visible = False
       lblA.Visible = False
       txtConsulta.Visible = False
       cbbConsulta.Visible = False
       cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Status" Then
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       dtpHora_Inicial.Visible = False
       dtpHora_Final.Visible = False
       lblA.Visible = False
       txtConsulta.Visible = False
       cbbConsulta.Visible = True
       cbbConsulta.SetFocus
    ElseIf cbbCampos.Text = "Data Solicitação" Then
       dtpInicial.Visible = True
       dtpFinal.Visible = True
       dtpHora_Inicial.Visible = False
       dtpHora_Final.Visible = False
       dtpInicial.Value = Date
       dtpFinal.Value = Date + 7
       lblA.Visible = True
       txtConsulta.Visible = False
       cbbConsulta.Visible = False
       dtpInicial.SetFocus
    ElseIf cbbCampos.Text = "Hora Solicitação" Then
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       dtpHora_Inicial.Visible = True
       dtpHora_Final.Visible = True
       lblA.Visible = True
       txtConsulta.Visible = False
       cbbConsulta.Visible = False
       dtpHora_Inicial.SetFocus
    Else
       dtpInicial.Visible = False
       dtpFinal.Visible = False
       dtpHora_Inicial.Visible = False
       dtpHora_Final.Visible = False
       lblA.Visible = False
       cbbConsulta.Visible = False
       txtConsulta.Visible = True
       txtConsulta.SetFocus
    End If
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdIncluir_Agenda_Click()
    If txtContato_Agenda.Text = Empty Or txtTelefone_Agenda.Text = Empty Then
       MsgBox "Digite uma agenda válida. Verifique!", vbInformation, "Only Tech"
       Exit Sub
    End If
    
    hfgAgenda.Row = 1
    hfgAgenda.Col = 0
    
    If hfgAgenda.Text <> Empty Then
       intAgenda = hfgAgenda.Rows
    ElseIf intAgenda = 0 Then
       intAgenda = 1
    End If
    
    hfgAgenda.ColWidth(0) = 500
    hfgAgenda.Col = 0
    hfgAgenda.Rows = intAgenda + 1
    hfgAgenda.Row = intAgenda

    hfgAgenda.Col = 0

    hfgAgenda.TextArray((hfgAgenda.Row * hfgAgenda.Cols + hfgAgenda.Col + 1)) = Empty
    hfgAgenda.TextArray((hfgAgenda.Row * hfgAgenda.Cols + hfgAgenda.Col + 2)) = Empty
    hfgAgenda.TextArray((hfgAgenda.Row * hfgAgenda.Cols + hfgAgenda.Col + 3)) = txtContato_Agenda.Text
    hfgAgenda.TextArray((hfgAgenda.Row * hfgAgenda.Cols + hfgAgenda.Col + 4)) = txtTelefone_Agenda.Text
    
    hfgAgenda.Row = hfgAgenda.Rows - 1
    hfgAgenda.Col = 0
    hfgAgenda.CellBackColor = &H80FFFF
    hfgAgenda.CellFontBold = False
    hfgAgenda.CellFontSize = 7
    hfgAgenda.Text = hfgAgenda.Rows - 1
    
    intAgenda = intAgenda + 1

    txtContato_Agenda.Text = Empty
    txtTelefone_Agenda.Text = Empty

    txtContato_Agenda.SetFocus
End Sub

Private Sub cmdIncluir_Atendimento_Click()
    If txtVendedor.Text = Empty Or dtcVendedor.Text = Empty Then
       MsgBox "Digite um atendimento válido. Verifique!", vbInformation, "Only Tech"
       Exit Sub
    End If
    
    hfgAtendimento.Row = 1
    hfgAtendimento.Col = 0
    
    If hfgAtendimento.Text <> Empty Then
       intAtendimento = hfgAtendimento.Rows
    ElseIf intAtendimento = 0 Then
       intAtendimento = 1
    End If
    
    hfgAtendimento.ColWidth(0) = 500
    hfgAtendimento.Col = 0
    hfgAtendimento.Rows = intAtendimento + 1
    hfgAtendimento.Row = intAtendimento

    hfgAtendimento.Col = 0

    hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 1)) = Empty
    hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 2)) = txtVendedor.Text
    hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 3)) = dtcVendedor.Text
    hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 4)) = Empty
    hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 5)) = dtpData_Previsao.Value
    hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 6)) = dtpData_Atendimento.Value
    hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 7)) = txtMotivo_Atendimento.Text
    
    hfgAtendimento.Col = 0
    hfgAtendimento.CellBackColor = &H80FFFF
    hfgAtendimento.CellFontBold = False
    hfgAtendimento.CellFontSize = 7
    hfgAtendimento.Text = intAtendimento
    
    intAtendimento = intAtendimento + 1

    txtVendedor.Text = Empty
    dtcVendedor.Text = Empty
    dtpData_Previsao.Value = Date
    dtpData_Atendimento.Value = Date + 7
    txtMotivo_Atendimento.Text = Empty

    txtVendedor.SetFocus
End Sub

Private Sub cmdLogradouro_Click()
    frmAguarde.Show
    DoEvents
    frmLogradouro_Solicitacao_Visita.Show
    Unload frmAguarde
End Sub

Private Sub cmdOrdenar_Click()
    If cmdOrdenar.Caption = "C" Then
       cmdOrdenar.Caption = "A"
    Else
       cmdOrdenar.Caption = "C"
    End If
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub

Private Sub cmdRemover_Agenda_Click()
    If intAgenda = 0 Then intAgenda = hfgAgenda.Rows
    
    hfgAgenda.Col = 0
    If hfgAgenda.Text <> Empty Then
       If hfgAgenda.Text = Empty Then
          MsgBox "Não há um contato selecionada. Verifique!", vbInformation, "OnlyTech"
          Exit Sub
       End If
       
       If hfgAgenda.Rows <= 2 Then
          hfgAgenda.Text = Empty
          hfgAgenda.Clear
          Movimentacoes.Monta_HFlex_Grid hfgAgenda, "0,0,4800,2000", "ID,ID,Contato,Telefone", 4, "Otica", Me
       Else
          hfgAgenda.RemoveItem (hfgAgenda.Row)
       End If
    
       intAgenda = intAgenda - 1
    
       hfgAgenda.Col = 0
       hfgAgenda.Row = 1
       For I = 1 To hfgAgenda.Rows - 1
          hfgAgenda.Text = I
          If hfgAgenda.Row + 1 <= hfgAgenda.Rows - 1 Then
             hfgAgenda.Row = hfgAgenda.Row + 1
          End If
       Next I
       hfgAgenda.Row = 0
       hfgAgenda.Text = Empty
    End If
        
    hfgAgenda.Col = 1
    hfgAgenda.Row = 1
    If hfgAgenda.Text = Empty Then
       hfgAgenda.Col = 0
       hfgAgenda.Row = 1
       hfgAgenda.Text = Empty
    End If
End Sub

Private Sub cmdRemover_Atendimento_Click()
    If intAtendimento = 0 Then intAtendimento = hfgAtendimento.Rows
    
    hfgAtendimento.Col = 0
    If hfgAtendimento.Text <> Empty Then
       If hfgAtendimento.Text = Empty Then
          MsgBox "Não há Condutor selecionada. Verifique!", vbInformation, "OnlyTech"
          Exit Sub
       End If
       
       If hfgAtendimento.Rows <= 2 Then
          hfgAtendimento.Text = Empty
          hfgAtendimento.Clear
          Movimentacoes.Monta_HFlex_Grid hfgAtendimento, "0,1000,4000,0,1400,1600,3000", "ID,Vendedor,Nome,ID,Data Previsão,Data Atendimento,Motivo", 7, "Otica", Me
       Else
          hfgAtendimento.RemoveItem (hfgAtendimento.Row)
       End If
    
       intAtendimento = intAtendimento - 1
    
       hfgAtendimento.Col = 0
       hfgAtendimento.Row = 1
       For I = 1 To hfgAtendimento.Rows - 1
          hfgAtendimento.Text = I
          If hfgAtendimento.Row + 1 <= hfgAtendimento.Rows - 1 Then
             hfgAtendimento.Row = hfgAtendimento.Row + 1
          End If
       Next I
       hfgAtendimento.Row = 0
       hfgAtendimento.Text = Empty
    End If
        
    hfgAtendimento.Col = 1
    hfgAtendimento.Row = 1
    If hfgAtendimento.Text = Empty Then
       hfgAtendimento.Col = 0
       hfgAtendimento.Row = 1
       hfgAtendimento.Text = Empty
    End If
End Sub

Private Sub dtcCidade_GotFocus()
    If Me.txtCodigo_Cidade.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCidade.Text)
    End If
End Sub

Private Sub dtcCidade_LostFocus()
    Dim rstCidade As New ADODB.Recordset
    
    txtCodigo_Cidade.Text = dtcCidade.BoundText
    
    If IsNumeric(txtCodigo_Cidade.Text) = False Or dtcCidade.Text = Empty Then txtCodigo_Cidade.Text = Empty: Exit Sub
    
    dtcCidade.BoundText = txtCodigo_Cidade.Text
    
    strSql = "Select TBCidade_Otica.DFUf_TBCidade_Otica FROM TBCidade_Otica " & _
             "WHERE TBCidade_Otica.IXCodigo_Correios_TBCidade_otica = '" & txtCodigo_Cidade.Text & "'"
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstCidade, "Otica", Me)
    
    If rstCidade.RecordCount <> 0 Then
       txtUf.Text = rstCidade.Fields("DFUf_TBCidade_Otica")
    Else
       txtUf.Text = Empty
    End If
    
    Set rstCidade = Nothing
End Sub

Private Sub dtcVendedor_GotFocus()
    If txtVendedor.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcVendedor.Text)
    End If
End Sub

Private Sub dtcVendedor_LostFocus()
    txtVendedor.Text = dtcVendedor.BoundText
    If IsNumeric(txtVendedor.Text) = False Or dtcVendedor.Text = Empty Then txtVendedor.Text = Empty: Exit Sub
End Sub

Private Sub hfgSolicitacao_Visita_Click()
    If hfgSolicitacao_Visita.Col = 0 And hfgSolicitacao_Visita.Text <> Empty Then
       On Error Resume Next
       
       Call Objetos.Limpa_TXT(Me)
       
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
       
       frmAguarde.Show
       DoEvents

       txtCodigo.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 1))
       txtContato_Solicitacao.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 2))
       dtpData_Solicitacao.Value = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 3))
       dtpHora_Solicitacao.Value = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 4))
       txtEndereco.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 5))
       txtNumero.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 6))
       txtComplemento.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 7))
       txtBairro.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 8))
       txtCodigo_Cidade.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 9))
       dtcCidade.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 10))
       txtUf.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 11))
       txtCep.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 12))
       cbbStatus.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 13))
       txtValor_Orcamento.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 14))
       txtCondicao_Pagamento.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 15))
       txtObservacao.Text = hfgSolicitacao_Visita.TextArray((hfgSolicitacao_Visita.Row * hfgSolicitacao_Visita.Cols + hfgSolicitacao_Visita.Col + 16))
            
''''''''''''''''''''''''''''''''''''''''''''ABASTECENDO AS OUTRAS 2 GUIAS''''''''''''''''''''''''''''''''''''''''''

'''''''Abastecendo itens da guia Agenda''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strSql = Empty
       strSql = "SELECT TBAgenda_solicitacao_visita.PKId_TBAgenda_solicitacao_visita," & _
                "TBAgenda_solicitacao_visita.FKId_TBSolicitacao_visita," & _
                "TBAgenda_solicitacao_visita.DFContato_TBAgenda_solicitacao_visita," & _
                "TBAgenda_solicitacao_visita.DFTelefone_TBAgenda_solicitacao_visita " & _
                "FROM TBAgenda_solicitacao_visita " & _
                "WHERE TBAgenda_solicitacao_visita.FKId_TBSolicitacao_visita = " & txtCodigo.Text & " "
                
       Movimentacoes.Movimenta_HFlex_Grid strSql, hfgAgenda, "0,0,4800,2000", "ID,ID,Contato,Telefone", "BDRetaguarda", "Otica", Me

'''''''Abastecendo itens da guia Atendimento'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       strSql = Empty
       strSql = "SELECT TBAtendimento_solicitacao_visita.PKId_TBAtendimento_solicitacao_visita," & _
                "TBVendedor.IXCodigo_TBVendedor," & _
                "TBVendedor.DFNome_TBVendedor," & _
                "TBAtendimento_solicitacao_visita.FKId_TBSolicitacao_visita," & _
                "TBAtendimento_solicitacao_visita.DFData_previsao_TBAtendimento_solicitacao_visita," & _
                "TBAtendimento_solicitacao_visita.DFData_atendimento_TBAtendimento_solicitacao_visita," & _
                "TBAtendimento_solicitacao_visita.DFMotivo_TBAtendimento_solicitacao_visita " & _
                "FROM TBAtendimento_solicitacao_visita " & _
                "INNER JOIN TBVendedor " & _
                "ON TBAtendimento_solicitacao_visita.FKId_TBVendedor = TBVendedor.PKId_TBVendedor " & _
                "WHERE TBAtendimento_solicitacao_visita.FKId_TBSolicitacao_visita = " & txtCodigo.Text & " " & _
                "ORDER BY TBAtendimento_solicitacao_visita.DFData_previsao_TBAtendimento_solicitacao_visita DESC"
       
       Movimentacoes.Movimenta_HFlex_Grid strSql, hfgAtendimento, "0,1000,4000,0,1400,1600,3000", "ID,Vendedor,Nome,ID,Data Previsão,Data Atendimento,Motivo", "BDRetaguarda", "Otica", Me
          
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

       hfgAgenda.Col = 1: hfgAgenda.Row = 0
       If hfgAgenda.Text = Empty Then Call Limpa_HFGAgenda

       hfgAtendimento.Col = 1: hfgAtendimento.Row = 0
       If hfgAtendimento.Text = Empty Then Call Limpa_HFGAtendimento
       
       booAlterar = True
       txtConsulta.Text = Empty
       
       sstSolicitacao_Visita.Tab = 0
       sstSolicitacao_Visita.TabEnabled(0) = True
       sstSolicitacao_Visita.TabEnabled(1) = True
       sstSolicitacao_Visita.TabEnabled(2) = True
       
       txtContato_Solicitacao.SetFocus
     End If
    Unload frmAguarde
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
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Erro
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Solicitação Visita"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstSolicitacao_Visita, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o cadastro de Solicitação de Visita"
    'Gravando o log
    log.Gravar_log "Otica", Me
        
    dtpData_Solicitacao.Value = Date
    dtpHora_Solicitacao.Value = Now
    dtpData_Previsao.Value = Date
    dtpData_Atendimento.Value = Date + 7
    
    sstSolicitacao_Visita.TabEnabled(0) = False
    sstSolicitacao_Visita.TabEnabled(1) = False
    sstSolicitacao_Visita.TabEnabled(2) = False
    sstSolicitacao_Visita.Tab = 3
    
    Call Reposicao
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Load")
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    
    strCombo = Empty
       
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Unload")
    Exit Sub
End Sub

Private Sub hfgSolicitacao_Visita_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgSolicitacao_Visita_Click
    End If
End Sub

Private Sub sstSolicitacao_Visita_Click(PreviousTab As Integer)
    If sstSolicitacao_Visita.Tab = 0 Then
       txtContato_Solicitacao.SetFocus
    ElseIf sstSolicitacao_Visita.Tab = 1 Then
       txtContato_Agenda.SetFocus
    ElseIf sstSolicitacao_Visita.Tab = 2 Then
       txtVendedor.SetFocus
    ElseIf sstSolicitacao_Visita.Tab = 3 Then
        If strCombo <> Empty And strCombo <> "Todos" And txtConsulta.Visible = True Then
           cbbCampos.Text = strCombo
           txtConsulta.SetFocus
        ElseIf strCombo = "Todos" Then
           hfgSolicitacao_Visita.Row = 1
           hfgSolicitacao_Visita.Col = 0
           hfgSolicitacao_Visita.SetFocus
        ElseIf strCombo = "Status" Then
               cbbConsulta.SetFocus
        End If
    End If
End Sub

Private Sub tlbbotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2: Call Gravar
           Case 3: Call Cancelar
           Case 4: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
    End Select
End Sub

Function Gravar()
    On Error GoTo Erro

    'Verifica se os campos necessarios para gravar não estão nulos
    If txtEndereco.Text = Empty Then
       MsgBox "O campo endereço não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtEndereco.SetFocus
       Exit Function
    ElseIf txtCodigo_Cidade.Text = Empty Then
       MsgBox "O campo código da cidade não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtCodigo_Cidade.SetFocus
       Exit Function
    ElseIf dtcCidade.Text = Empty Then
       MsgBox "O campo descrição da cidade não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtCodigo_Cidade.Text = Empty
       txtCodigo_Cidade.SetFocus
       Exit Function
    ElseIf txtCep.Text = Empty Then
       MsgBox "O campo CEP não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtCep.SetFocus
       Exit Function
    ElseIf cbbStatus.Text = Empty Then
       MsgBox "O campo status não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       cbbStatus.SetFocus
       Exit Function
    End If

    If tlbBotoes.Buttons.Item(2).Enabled = False Then
       Exit Function
    End If
    
    Dim strSet As String
    Dim strCampo As String
    Dim strID_Cidade As String
    Dim strValores As String
    Dim strStatus As String
                    
    Call Objetos.Maiusculo_TXT(Me)
    
    If cbbStatus.Text = "1 - A Visitar" Then
       strStatus = "1"
    ElseIf cbbStatus.Text = "2 - Visitado - Aguardando Confirmação" Then
       strStatus = "2"
    ElseIf cbbStatus.Text = "3 - Visitado - Fechado" Then
       strStatus = "3"
    ElseIf cbbStatus.Text = "4 - Visitado - Não Finalizado" Then
       strStatus = "4"
    End If
    
    strID_Cidade = Funcoes_Gerais.Localiza_ID("PKId_TBCidade_Otica", "IXCodigo_Correios_TBCidade_Otica", txtCodigo_Cidade.Text, "TBCidade_Otica", "Otica", Me, "BDRetaguarda")
    
    strCampo = "DFData_TBSolicitacao_visita," & _
               "DFHora_TBSolicitacao_visita," & _
               "DFContato_TBSolicitacao_visita," & _
               "DFEndereco_TBSolicitacao_visita," & _
               "DFNumero_TBSolicitacao_visita," & _
               "DFComplemento_TBSolicitacao_visita," & _
               "DFBairro_TBSolicitacao_visita," & _
               "FKId_TBCidade_otica," & _
               "DFCep_TBSolicitacao_visita," & _
               "DFStatus_TBSolicitacao_visita," & _
               "DFValor_Orcamento_TBSolicitacao_visita," & _
               "DFCondicao_pagamento_TBSolicitacao_visita, " & _
               "DFObservacao_TBSolicitacao_visita"
    
    strValores = "'" & Format(dtpData_Solicitacao.Value, "YYYYMMDD") & "'," & _
                 "'" & Format(dtpHora_Solicitacao.Value, "HH:MM:SS") & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtContato_Solicitacao.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtEndereco.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtNumero.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtComplemento.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtBairro.Text) & "'," & _
                 "'" & strID_Cidade & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtCep.Text) & "'," & _
                 "'" & strStatus & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtValor_Orcamento) & "," & _
                 "'" & Funcoes_Gerais.Grava_String(txtCondicao_Pagamento.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "'"
    
    'Indicando o banco à conectar-se
    Conexao_Visita.Initial_Catalog = "BDRetaguarda"

    'Estabelecendo conexão com o banco
    Conexao_Visita.Abrir_conexao ("Otica")

    'Indica o inicio da transação junto o banco
    Conexao_Visita.CNConexao.BeginTrans

    On Error GoTo Erro_Transacao

    If booAlterar = True Then
       log.Evento = "Alterar"
       
       Call Alterar_Agenda
       Call Alterar_Atendimento
       
       strSql = "UPDATE TBSolicitacao_visita SET DFData_TBSolicitacao_visita = '" & Format(dtpData_Solicitacao.Value, "YYYYMMDD") & "', " & _
                "       DFHora_TBSolicitacao_visita = '" & Format(dtpHora_Solicitacao.Value, "HH:MM:SS") & "'," & _
                "       DFContato_TBSolicitacao_visita  = '" & Funcoes_Gerais.Grava_String(txtContato_Solicitacao.Text) & " '," & _
                "       DFEndereco_TBSolicitacao_visita = '" & Funcoes_Gerais.Grava_String(txtEndereco.Text) & "'," & _
                "       DFNumero_TBSolicitacao_visita = '" & Funcoes_Gerais.Grava_String(txtNumero.Text) & "'," & _
                "       DFComplemento_TBSolicitacao_visita  = '" & Funcoes_Gerais.Grava_String(txtComplemento.Text) & "'," & _
                "       DFBairro_TBSolicitacao_visita  = '" & Funcoes_Gerais.Grava_String(txtBairro.Text) & "'," & _
                "       FKId_TBCidade_otica = '" & strID_Cidade & "'," & _
                "       DFCep_TBSolicitacao_visita = '" & Funcoes_Gerais.Grava_String(txtCep.Text) & "'," & _
                "       DFStatus_TBSolicitacao_visita = '" & strStatus & "'," & _
                "       DFValor_Orcamento_TBSolicitacao_visita = " & Funcoes_Gerais.Grava_Moeda(txtValor_Orcamento) & "," & _
                "       DFCondicao_pagamento_TBSolicitacao_visita = '" & Funcoes_Gerais.Grava_String(txtCondicao_Pagamento.Text) & "'," & _
                "       DFObservacao_TBSolicitacao_visita = '" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "' " & _
                "       WHERE PKId_TBSolicitacao_visita = " & txtCodigo.Text & ""
       
       'Gravando Alteração na TBSolicitacao_visita
       Conexao_Visita.CNConexao.Execute strSql
       
       log.Descricao = "Alterando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "Otica", Me
    Else
       log.Evento = "Incluir Novo"
       
       'Inserindo novo registro na TBSolicitacao_visita
       strSql = Empty
       strSql = "INSERT INTO TBSolicitacao_visita " & _
                "(DFData_TBSolicitacao_visita," & _
                "DFHora_TBSolicitacao_visita," & _
                "DFContato_TBSolicitacao_visita," & _
                "DFEndereco_TBSolicitacao_visita," & _
                "DFNumero_TBSolicitacao_visita," & _
                "DFComplemento_TBSolicitacao_visita," & _
                "DFBairro_TBSolicitacao_visita," & _
                "FKId_TBCidade_otica," & _
                "DFCep_TBSolicitacao_visita," & _
                "DFStatus_TBSolicitacao_visita," & _
                "DFValor_Orcamento_TBSolicitacao_visita," & _
                "DFCondicao_pagamento_TBSolicitacao_visita, " & _
                "DFObservacao_TBSolicitacao_visita) "
       strSql = strSql & "SELECT '" & Format(dtpData_Solicitacao.Value, "YYYYMMDD") & "'," & _
                         "'" & Format(dtpHora_Solicitacao.Value, "HH:MM:SS") & "'," & _
                         "'" & Funcoes_Gerais.Grava_String(txtContato_Solicitacao.Text) & "'," & _
                         "'" & Funcoes_Gerais.Grava_String(txtEndereco.Text) & "'," & _
                         "'" & Funcoes_Gerais.Grava_String(txtNumero.Text) & "'," & _
                         "'" & Funcoes_Gerais.Grava_String(txtComplemento.Text) & "'," & _
                         "'" & Funcoes_Gerais.Grava_String(txtBairro.Text) & "'," & _
                         "'" & strID_Cidade & "'," & _
                         "'" & Funcoes_Gerais.Grava_String(txtCep.Text) & "'," & _
                         "'" & strStatus & "'," & _
                         "" & Funcoes_Gerais.Grava_Moeda(txtValor_Orcamento) & "," & _
                         "'" & Funcoes_Gerais.Grava_String(txtCondicao_Pagamento.Text) & "'," & _
                         "'" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "'"

       'Gravando Inclusão TBSolicitacao_visita
       Conexao_Visita.CNConexao.Execute strSql
       
       'Comitando a transação
       Conexao_Visita.CNConexao.CommitTrans

       'Fechando a conexão
       Conexao_Visita.CNConexao.Close

''''''''''''''''''''INCLUINDO OUTRAS 2 GUIAS USANDO PKID QUE ACABOU DE SER GERADO ANTERIORMENTE''''''''''''''''''''
       
       'Indicando o banco à conectar-se
       Conexao_Visita.Initial_Catalog = "BDRetaguarda"
    
       'Estabelecendo conexão com o banco
       Conexao_Visita.Abrir_conexao ("Otica")
    
       'Indica o inicio da transação junto o banco
       Conexao_Visita.CNConexao.BeginTrans

       On Error GoTo Erro_inclusao
       
       Call Incluir_Agenda
       Call Incluir_Atendimento
       
       log.Descricao = "Gravando o registro: " + txtCodigo.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "Otica", Me
    End If
    
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
    Call Limpa_HFGAgenda
    Call Limpa_HFGAtendimento

          
    tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    tlbBotoes.Buttons.Item(2).Enabled = False
    tlbBotoes.Buttons.Item(3).Enabled = False
    tlbBotoes.Buttons.Item(4).Enabled = False
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    
    If booPrivilegio_Consultar = False Then
       Me.hfgSolicitacao_Visita.Visible = False
    End If
    
    If booAlterar = False Then
       Dim rstSolicita As New ADODB.Recordset
       
       strSql = Empty
       strSql = "SELECT MAX(PKId_TBSolicitacao_visita) as ID FROM TBSolicitacao_visita "
       
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstSolicita, "Otica", Me
       
       If rstSolicita.RecordCount <> 0 Then
          MsgBox "** O código dessa Solicitação Visita é: " & rstSolicita.Fields("ID") & "", vbOKOnly, "Only Tech"
       End If
       
       Set rstSolicita = Nothing
    End If
    
    sstSolicitacao_Visita.TabEnabled(0) = False
    sstSolicitacao_Visita.TabEnabled(1) = False
    sstSolicitacao_Visita.TabEnabled(2) = False
    sstSolicitacao_Visita.Tab = 3

    'Comitando a transação
    Conexao_Visita.CNConexao.CommitTrans

    'Fechando a conexão
    Conexao_Visita.CNConexao.Close

    Exit Function
    
Erro_inclusao:
    
    'ROOLBACK NA TRANSAÇÃO
    Conexao_Visita.CNConexao.RollbackTrans
        
   'Excluindo na TBSolicitacao_visita
    strSql = "DELETE FROM TBSolicitacao_visita WHERE PKId_TBSolicitacao_visita = " & txtCodigo.Text & " "
    
    conexao.CNConexao.Execute strSql

    'Comitando a transação
    Conexao_Visita.CNConexao.CommitTrans

    'Fechando a conexão
    Conexao_Visita.CNConexao.Close

    Call Erro.Erro(Me, "Otica", "Gravar")
    
    Exit Function

Erro_Transacao:
    
    'ROOLBACK NA TRANSAÇÃO
    Conexao_Visita.CNConexao.RollbackTrans

    'Comitando a transação
    Conexao_Visita.CNConexao.CommitTrans

    'Fechando a conexão
    Conexao_Visita.CNConexao.Close

    Call Erro.Erro(Me, "Otica", "Gravar")
    
    Exit Function
Erro:
    
    Call Erro.Erro(Me, "Otica", "Gravar")
    
    Exit Function
End Function

Private Function Incluir_Agenda()
    'Inclusão na TBAgenda_solicitacao_visita
    hfgAgenda.Row = 1
    hfgAgenda.Col = 3
       
    If hfgAgenda.Text <> Empty Then
       intAgenda = hfgAgenda.Rows - 1
    End If

    Do While intAgenda <> 0
       hfgAgenda.Row = intAgenda
       hfgAgenda.Col = 0
       
       Dim rstSolicita As New ADODB.Recordset
       
       strSql = Empty
       strSql = "SELECT MAX(PKId_TBSolicitacao_visita) as ID FROM TBSolicitacao_visita "
       
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstSolicita, "Otica", Me
       
       If rstSolicita.RecordCount <> 0 Then
          txtCodigo.Text = rstSolicita.Fields("ID")
       End If
       
       Set rstSolicita = Nothing
       
       strSql = Empty
       strSql = "INSERT INTO TBAgenda_solicitacao_visita (FKId_TBSolicitacao_visita," & _
                "DFContato_TBAgenda_solicitacao_visita," & _
                "DFTelefone_TBAgenda_solicitacao_visita) " & _
                "SELECT " & txtCodigo.Text & "," & _
                "'" & Funcoes_Gerais.Grava_String(hfgAgenda.TextArray((hfgAgenda.Row * hfgAgenda.Cols + hfgAgenda.Col + 3))) & "'," & _
                "'" & Funcoes_Gerais.Grava_String(hfgAgenda.TextArray((hfgAgenda.Row * hfgAgenda.Cols + hfgAgenda.Col + 4))) & "'"
                
       'Gravando Inclusão TBSolicitacao_visita
       Conexao_Visita.CNConexao.Execute strSql

       intAgenda = intAgenda - 1
    Loop
End Function

Private Function Incluir_Atendimento()
    hfgAtendimento.Row = 1
    hfgAtendimento.Col = 2
       
    Dim strID_Vendedor As String
    
    If hfgAtendimento.Text <> Empty Then
       intAtendimento = hfgAtendimento.Rows - 1
    End If
    
    Do While intAtendimento <> 0
       hfgAtendimento.Col = 0: hfgAtendimento.Row = intAtendimento
       
       strID_Vendedor = Funcoes_Gerais.Localiza_ID("PKId_TBVendedor", "IXCodigo_TBVendedor", hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 2)), "TBVendedor", "Otica", Me, "BDRetaguarda")
       
       Dim rstSolicita As New ADODB.Recordset
       
       strSql = Empty
       strSql = "SELECT MAX(PKId_TBSolicitacao_visita) as ID FROM TBSolicitacao_visita "
       
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstSolicita, "Otica", Me
       
       If rstSolicita.RecordCount <> 0 Then
          txtCodigo.Text = rstSolicita.Fields("ID")
       End If
       
       Set rstSolicita = Nothing
       
       strSql = Empty
       strSql = "INSERT INTO TBAtendimento_solicitacao_visita (FKId_TBVendedor," & _
                "FKId_TBSolicitacao_visita," & _
                "DFData_previsao_TBAtendimento_solicitacao_visita," & _
                "DFData_atendimento_TBAtendimento_solicitacao_visita," & _
                "DFMotivo_TBAtendimento_solicitacao_visita) " & _
                "SELECT " & strID_Vendedor & "," & _
                " " & txtCodigo.Text & "," & _
                "'" & Format(hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 5)), "YYYYMMDD") & "'," & _
                "'" & Format(hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 6)), "YYYYMMDD") & "'," & _
                "'" & Funcoes_Gerais.Grava_String(hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 7))) & "'"
                
       'Gravando Inclusão TBSolicitacao_visita
       Conexao_Visita.CNConexao.Execute strSql

       intAtendimento = intAtendimento - 1
    Loop
End Function

Private Function Alterar_Agenda()
    'Alteração na TBAgenda_solicitacao_visita
    strSql = "DELETE FROM TBAgenda_solicitacao_visita WHERE FKId_TBSolicitacao_visita = " & txtCodigo.Text & " "
    
    'Excluindo TBAgenda_solicitacao_visita
    Conexao_Visita.CNConexao.Execute strSql
    
    'Alteração na TBAgenda_solicitacao_visita
    hfgAgenda.Row = 1
    hfgAgenda.Col = 3
       
    If hfgAgenda.Text <> Empty Then
       intAgenda = hfgAgenda.Rows - 1
    End If

    Do While intAgenda <> 0
       hfgAgenda.Col = 0: hfgAgenda.Row = intAgenda
       
       strSql = Empty
       strSql = "INSERT INTO TBAgenda_solicitacao_visita (FKId_TBSolicitacao_visita," & _
                "DFContato_TBAgenda_solicitacao_visita," & _
                "DFTelefone_TBAgenda_solicitacao_visita) " & _
                "SELECT " & txtCodigo.Text & "," & _
                "'" & Funcoes_Gerais.Grava_String(hfgAgenda.TextArray((hfgAgenda.Row * hfgAgenda.Cols + hfgAgenda.Col + 3))) & "'," & _
                "'" & Funcoes_Gerais.Grava_String(hfgAgenda.TextArray((hfgAgenda.Row * hfgAgenda.Cols + hfgAgenda.Col + 4))) & "'"
                
       'Gravando Inclusão TBSolicitacao_visita
       Conexao_Visita.CNConexao.Execute strSql

       intAgenda = intAgenda - 1
    Loop
End Function

Private Function Alterar_Atendimento()
    'Alteração na TBAtendimento_solicitacao_visita
    strSql = "DELETE FROM TBAtendimento_solicitacao_visita WHERE FKId_TBSolicitacao_visita = " & txtCodigo.Text & " "
    
    'Excluindo TBAtendimento_solicitacao_visita
    Conexao_Visita.CNConexao.Execute strSql

    'Alteração na TBAtendimento_solicitacao_visita
    hfgAtendimento.Row = 1
    hfgAtendimento.Col = 2
       
    Dim strID_Vendedor As String
    
    If hfgAtendimento.Text <> Empty Then
       intAtendimento = hfgAtendimento.Rows - 1
    End If
      
    Do While intAtendimento <> 0
       hfgAtendimento.Col = 0: hfgAtendimento.Row = intAtendimento
       
       strID_Vendedor = Funcoes_Gerais.Localiza_ID("PKId_TBVendedor", "IXCodigo_TBVendedor", hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 2)), "TBVendedor", "Otica", Me, "BDRetaguarda")
       
       strSql = Empty
       strSql = "INSERT INTO TBAtendimento_solicitacao_visita (FKId_TBVendedor," & _
                "FKId_TBSolicitacao_visita," & _
                "DFData_previsao_TBAtendimento_solicitacao_visita," & _
                "DFData_atendimento_TBAtendimento_solicitacao_visita," & _
                "DFMotivo_TBAtendimento_solicitacao_visita) " & _
                "SELECT " & strID_Vendedor & "," & _
                " " & txtCodigo.Text & "," & _
                "'" & Format(hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 5)), "YYYYMMDD") & "'," & _
                "'" & Format(hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 6)), "YYYYMMDD") & "'," & _
                "'" & Funcoes_Gerais.Grava_String(hfgAtendimento.TextArray((hfgAtendimento.Row * hfgAtendimento.Cols + hfgAtendimento.Col + 7))) & "'"
                
       'Gravando Inclusão TBSolicitacao_visita
       Conexao_Visita.CNConexao.Execute strSql

       intAtendimento = intAtendimento - 1
    Loop
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    If tlbBotoes.Buttons.Item(4).Enabled = False Then
       Exit Function
    End If
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + Me.txtCodigo.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "Otica", Me
    
    'Indicando o banco à conectar-se
    conexao.Initial_Catalog = "BDRetaguarda"

    'Estabelecendo Conexão com Banco
    conexao.Abrir_conexao ("Otica")
    
    'Indica o Inicio da Transação Junto ao Banco
    conexao.CNConexao.BeginTrans

    'Excluindo na TBAgenda_solicitacao_visita
    strSql = "DELETE FROM TBAgenda_solicitacao_visita WHERE FKId_TBSolicitacao_visita = " & txtCodigo.Text & " "
    
    conexao.CNConexao.Execute strSql
    
    'Excluindo na TBAtendimento_solicitacao_visita
    strSql = "DELETE FROM TBAtendimento_solicitacao_visita WHERE FKId_TBSolicitacao_visita = " & txtCodigo.Text & " "
    
    conexao.CNConexao.Execute strSql
    
   'Excluindo na TBSolicitacao_visita
    strSql = "DELETE FROM TBSolicitacao_visita WHERE PKId_TBSolicitacao_visita = " & txtCodigo.Text & " "
    
    conexao.CNConexao.Execute strSql

    'Indica o Sucesso da Transação do Banco
    conexao.CNConexao.CommitTrans

    'Fechando a Conexão
    conexao.Fechar_conexao
    
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
    Call Limpa_HFGAgenda
    Call Limpa_HFGAtendimento

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
    
    If booPrivilegio_Consultar = False Then
       Me.hfgSolicitacao_Visita.Visible = False
    End If
                
    sstSolicitacao_Visita.TabEnabled(0) = False
    sstSolicitacao_Visita.TabEnabled(1) = False
    sstSolicitacao_Visita.TabEnabled(2) = False
    sstSolicitacao_Visita.Tab = 3
        
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Excluir")
    
    'Indica o Fracasso da Transação do Banco
    conexao.CNConexao.RollbackTrans
    
    'Fecha a Conexão com o Banco
    conexao.Fechar_conexao

    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
    Call Limpa_HFGAgenda
    Call Limpa_HFGAtendimento
       
    'Novo
     tlbBotoes.Buttons.Item(1).Enabled = booPrivilegio_Incluir
    'Gravar
    tlbBotoes.Buttons.Item(2).Enabled = False
    'Cancelar
    tlbBotoes.Buttons.Item(3).Enabled = False
    'Excluir
    tlbBotoes.Buttons.Item(4).Enabled = booPrivilegio_Excluir
    'Imprimir
    tlbBotoes.Buttons.Item(5).Enabled = booPrivilegio_Consultar
    
    If booPrivilegio_Consultar = False Then
       Me.hfgSolicitacao_Visita.Visible = False
    End If
    
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    
    sstSolicitacao_Visita.TabEnabled(0) = False
    sstSolicitacao_Visita.TabEnabled(1) = False
    sstSolicitacao_Visita.TabEnabled(2) = False
    sstSolicitacao_Visita.Tab = 3
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
    Call Monta_DataCombo
     
    dtpData_Solicitacao.Value = Date
    dtpHora_Solicitacao.Value = Now
    dtpData_Previsao.Value = Date
    dtpData_Atendimento.Value = Date + 7

    Call Objetos.Limpa_TXT(Me)
    Call Limpa_Combos
    Call Limpa_HFGAgenda
    Call Limpa_HFGAtendimento
       
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
        
    'Gravando Log
    log.Gravar_log "Otica", Me
            
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
        
    sstSolicitacao_Visita.TabEnabled(0) = True
    sstSolicitacao_Visita.TabEnabled(1) = True
    sstSolicitacao_Visita.TabEnabled(2) = True
    sstSolicitacao_Visita.Tab = 0
    
    txtContato_Solicitacao.SetFocus
    
    booAlterar = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Novo")
    Exit Function
End Function

Private Sub txtBairro_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtBairro_LostFocus()
    txtBairro.Text = UCase(txtBairro.Text)
End Sub

Private Sub txtCep_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtcep_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtCep_LostFocus()
    txtCep.Text = Format(txtCep.Text, "#####-###")
End Sub

Private Sub txtCodigo_Cidade_Change()
    Dim rstCidade As New ADODB.Recordset
    
    dtcCidade.BoundText = txtCodigo_Cidade.Text
    
    If IsNumeric(txtCodigo_Cidade.Text) = False Then txtCodigo_Cidade.Text = Empty: Exit Sub
    
    strSql = "Select TBCidade_Otica.DFUf_TBCidade_Otica FROM TBCidade_Otica " & _
             "WHERE TBCidade_Otica.IXCodigo_Correios_TBCidade_otica = '" & txtCodigo_Cidade.Text & "'"
    Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstCidade, "Otica", Me)
    
    If rstCidade.RecordCount <> 0 Then
       txtUf.Text = rstCidade.Fields("DFUf_TBCidade_Otica")
    Else
       txtUf.Text = Empty
    End If
    
    Set rstCidade = Nothing
End Sub

Private Sub txtCodigo_Cidade_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_Cidade_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    Movimentacoes.Verifica_Numero "IXCodigo_TBCliente", "TBCliente", txtCodigo, "Otica", Me
End Sub

Private Function Reposicao()
    
    strTamanho = "1600,3000,1500,1500,2000,800,1400," & _
                 "1000,800,3000,500,1500,3500,1600,2500,3000"
                            
    strNomes = "Solicitação Visita,Contato,Data Solicitação,Hora Solicitação,Endereço,Número,Complemento," & _
               "Bairro,Cidade,Nome,UF,CEP,Status,Valor Orçamento,Condição Pagamento,Observação"
    
    On Error GoTo Erro
               
    Movimentacoes.Monta_HFlex_Grid hfgSolicitacao_Visita, strTamanho, strNomes, 16, "Otica", Me
       
    Call Monta_DataCombo
    Call Monta_Combos
    Call Limpa_HFGAgenda
    Call Limpa_HFGAtendimento
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Reposicao")
    Resume Next
End Function

Private Sub txtComplemento_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtComplemento_LostFocus()
    txtComplemento.Text = UCase(txtComplemento.Text)
End Sub

Private Sub txtCondicao_Pagamento_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtCondicao_Pagamento_LostFocus()
    txtCondicao_Pagamento.Text = UCase(txtCondicao_Pagamento.Text)
End Sub

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtContato_Agenda_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtContato_Agenda_LostFocus()
    txtContato_Agenda.Text = UCase(txtContato_Agenda.Text)
End Sub

Private Sub txtContato_Solicitacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtContato_Solicitacao_LostFocus()
    txtContato_Solicitacao.Text = UCase(txtContato_Solicitacao.Text)
End Sub

Private Function Monta_Combos()
    cbbStatus.Clear
    cbbStatus.AddItem ("1 - A Visitar")
    cbbStatus.AddItem ("2 - Visitado - Aguardando Confirmação")
    cbbStatus.AddItem ("3 - Visitado - Fechado")
    cbbStatus.AddItem ("4 - Visitado - Não Finalizado")
            
    cbbConsulta.Clear
    cbbConsulta.AddItem ("1 - A Visitar")
    cbbConsulta.AddItem ("2 - Visitado - Aguardando Confirmação")
    cbbConsulta.AddItem ("3 - Visitado - Fechado")
    cbbConsulta.AddItem ("4 - Visitado - Não Finalizado")
                        
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Solicitação Visita")
    cbbCampos.AddItem ("Contato")
    cbbCampos.AddItem ("Data Solicitação")
    cbbCampos.AddItem ("Hora Solicitação")
    cbbCampos.AddItem ("Endereço")
    cbbCampos.AddItem ("Número")
    cbbCampos.AddItem ("Complemento")
    cbbCampos.AddItem ("Bairro")
    cbbCampos.AddItem ("CEP")
    cbbCampos.AddItem ("Código Cidade")
    cbbCampos.AddItem ("Nome Cidade")
    cbbCampos.AddItem ("Status")
    cbbCampos.AddItem ("Valor Orçamento")
    cbbCampos.AddItem ("Condição Pagamento")
    cbbCampos.AddItem ("Observação")
End Function

Private Sub txtEndereco_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtEndereco_LostFocus()
    txtEndereco.Text = UCase(txtEndereco.Text)
End Sub

Private Sub txtMotivo_Atendimento_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtMotivo_Atendimento_LostFocus()
    txtMotivo_Atendimento.Text = UCase(txtMotivo_Atendimento.Text)
End Sub

Private Sub txtNumero_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Limpa_Combos()
    cbbStatus.Text = Empty
    dtcCidade.Text = Empty
End Function

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
    
    If cbbCampos.Text = "Valor Orçamento" Then
       txtConsulta.Text = Format(txtConsulta.Text, "#,###0.00")
    End If
End Sub

Public Function Consulta()
    Dim strStatus As String
    Dim I As Integer
    
    If cbbCampos.Text <> "Todos" And txtConsulta.Visible = True Or cbbCampos.Text = Empty Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    ElseIf cbbCampos.Text = "Status" Then
           If cbbConsulta.Text = Empty Then
              MsgBox "Selecione uma opção para consulta.", vbCritical, "Only Tech"
              cbbConsulta.SetFocus
              Exit Function
           End If
    ElseIf cbbCampos.Text = "Data solicitação" Then
           If dtpInicial.Value > dtpFinal.Value Then
              MsgBox "Data final menor que data Inicial.Verifique!", vbInformation, "Only Tech"
              dtpFinal.SetFocus
              Exit Function
           End If
    ElseIf cbbCampos.Text = "Hora Solicitação" Then
           If dtpHora_Inicial.Value > dtpHora_Final.Value Then
              MsgBox "Hora Final menor que data Hora Inicial.Verifique!", vbInformation, "Only Tech"
              dtpHora_Final.SetFocus
              Exit Function
           End If
    End If
    
    If cbbConsulta.Text = "1 - A Visitar" Then
       strStatus = "1"
    ElseIf cbbConsulta.Text = "2 - Visitado - Aguardando Confirmação" Then
       strStatus = "2"
    ElseIf cbbConsulta.Text = "3 - Visitado - Fechado" Then
       strStatus = "3"
    ElseIf cbbConsulta.Text = "4 - Visitado - Não Finalizado" Then
       strStatus = "4"
    End If
    
    strSql = "SELECT PKId_TBSolicitacao_visita,DFContato_TBSolicitacao_visita,DFData_TBSolicitacao_visita," & _
             "DFHora_TBSolicitacao_visita,DFEndereco_TBSolicitacao_visita,DFNumero_TBSolicitacao_visita," & _
             "DFComplemento_TBSolicitacao_visita,DFBairro_TBSolicitacao_visita,IXCodigo_Correios_TBCidade_otica," & _
             "DFNome_TBCidade_otica,DFUf_TBCidade_otica,DFCep_TBSolicitacao_visita,DFStatus_TBSolicitacao_visita," & _
             "DFValor_Orcamento_TBSolicitacao_visita,DFCondicao_pagamento_TBSolicitacao_visita," & _
             "DFObservacao_TBSolicitacao_visita " & _
             "FROM TBSolicitacao_visita " & _
             "INNER JOIN TBCidade_Otica ON TBSolicitacao_visita.FKId_TBCidade_Otica = TBCidade_Otica.PKId_TBCidade_Otica "
             
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    
    txtConsulta.Text = Funcoes_Gerais.Grava_String(txtConsulta.Text)
    
    If cbbCampos.Text <> "Todos" Then
        If cbbCampos.Text = "Solicitação Visita" Then
           strSql = strSql & " WHERE convert(nvarchar,PKId_TBSolicitacao_visita) = '" & txtConsulta.Text & "' "
        ElseIf cbbCampos.Text = "Contato" Then
            strSql = strSql & " WHERE convert(nvarchar,DFContato_TBSolicitacao_visita) LIKE '%" & txtConsulta.Text & "%' "
        ElseIf cbbCampos.Text = "Endereço" Then
            strSql = strSql & " WHERE convert(nvarchar,DFEndereco_TBSolicitacao_visita) LIKE '%" & txtConsulta.Text & "%' "
        ElseIf cbbCampos.Text = "Número" Then
            strSql = strSql & " WHERE convert(nvarchar,DFNumero_TBSolicitacao_visita) LIKE '" & txtConsulta.Text & "%' "
        ElseIf cbbCampos.Text = "Complemento" Then
            strSql = strSql & " WHERE convert(nvarchar,DFComplemento_TBSolicitacao_visita) LIKE '%" & txtConsulta.Text & "%' "
        ElseIf cbbCampos.Text = "Bairro" Then
            strSql = strSql & " WHERE convert(nvarchar,DFBairro_TBSolicitacao_visita) LIKE '%" & txtConsulta.Text & "%' "
        ElseIf cbbCampos.Text = "Código Cidade" Then
            strSql = strSql & " WHERE convert(nvarchar,IXCodigo_Correios_TBCidade_Otica) = '" & txtConsulta.Text & "'"
        ElseIf cbbCampos.Text = "Nome Cidade" Then
            strSql = strSql & " WHERE TBCidade_Otica.DFNome_TBCidade_Otica LIKE '%" & txtConsulta.Text & "%'"
        ElseIf cbbCampos.Text = "UF" Then
            strSql = strSql & " WHERE TBCidade_Otica.DFUf_TBCidade_Otica LIKE '%" & txtConsulta.Text & "%'"
        ElseIf cbbCampos.Text = "CEP" Then
            txtConsulta.Text = Format(txtConsulta.Text, "#####-###")
            strSql = strSql & " WHERE convert(nvarchar,DFCep_TBSolicitacao_visita) = '" & txtConsulta.Text & "'"
        ElseIf cbbCampos.Text = "Status" Then
            strSql = strSql & " WHERE convert(nvarchar,DFStatus_TBSolicitacao_visita) = '" & strStatus & "'"
        ElseIf cbbCampos.Text = "Data Solicitação" Then
            strSql = strSql & " WHERE TBSolicitacao_visita.DFData_TBSolicitacao_visita >= '" & Format(dtpInicial.Value, "YYYYMMDD") & "' AND" & _
                              " TBSolicitacao_visita.DFData_TBSolicitacao_visita <= '" & Format(dtpFinal.Value, "YYYYMMDD") & "'"
        ElseIf cbbCampos.Text = "Hora Solicitação" Then
            strSql = strSql & " WHERE TBSolicitacao_visita.DFHora_TBSolicitacao_visita >= '" & Format(dtpHora_Inicial.Value, "HH:MM:SS") & "' AND" & _
                              " TBSolicitacao_visita.DFHora_TBSolicitacao_visita <= '" & Format(dtpHora_Final.Value, "HH:MM:SS") & "'"
        ElseIf cbbCampos.Text = "Condição Pagamento" Then
            strSql = strSql & " WHERE TBSolicitacao_visita.DFCondicao_pagamento_TBSolicitacao_visita LIKE '%" & txtConsulta.Text & "%'"
        ElseIf cbbCampos.Text = "Observação" Then
            strSql = strSql & " WHERE TBSolicitacao_visita.DFObservacao_TBSolicitacao_visita LIKE '%" & txtConsulta.Text & "%'"
        ElseIf cbbCampos.Text = "Valor Orçamento" Then
            strSql = strSql & " WHERE convert(nvarchar,DFValor_Orcamento_TBSolicitacao_visita) = '" & txtConsulta.Text & "'"
        End If
    End If
       
    frmAguarde.Show
    DoEvents
        
    If cmdOrdenar.Caption = "C" Then
       strSql = strSql & " ORDER BY TBSolicitacao_visita.PKId_TBSolicitacao_visita"
    ElseIf cmdOrdenar.Caption = "A" Then
       strSql = strSql & " ORDER BY TBSolicitacao_visita.DFContato_TBSolicitacao_visita"
    End If
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgSolicitacao_Visita, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    I = hfgSolicitacao_Visita.Rows - 1
        
    Do While I <> 0
       hfgSolicitacao_Visita.Col = 13: hfgSolicitacao_Visita.Row = I
       If hfgSolicitacao_Visita.Text = "1" Then
          hfgSolicitacao_Visita.Text = "1 - A Visitar"
       ElseIf hfgSolicitacao_Visita.Text = "2" Then
          hfgSolicitacao_Visita = "2 - Visitado - Aguardando Confirmação"
       ElseIf hfgSolicitacao_Visita.Text = "3" Then
          hfgSolicitacao_Visita = "3 - Visitado - Fechado"
       ElseIf hfgSolicitacao_Visita.Text = "4" Then
          hfgSolicitacao_Visita = "4 - Visitado - Não Finalizado"
       End If
       
       hfgSolicitacao_Visita.Col = 14: hfgSolicitacao_Visita.Row = I
       hfgSolicitacao_Visita.Text = Format(hfgSolicitacao_Visita.Text, "#,###0.00")

       I = I - 1
    Loop
    
    hfgSolicitacao_Visita.Row = 1: hfgSolicitacao_Visita.Col = 0
    
    If hfgSolicitacao_Visita.Text = Empty Then
       hfgSolicitacao_Visita.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgSolicitacao_Visita, strTamanho, strNomes, 42, "Otica", Me
    End If
    
    Unload frmAguarde
    
    hfgSolicitacao_Visita.SetFocus
End Function

Private Sub txtObservacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtObservacao_LostFocus()
    txtObservacao.Text = UCase(txtObservacao.Text)
End Sub

Private Sub txtTelefone_Agenda_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtTelefone_Agenda_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtTelefone_Agenda_LostFocus()
    If Len(txtTelefone_Agenda.Text) <= 10 Then
       txtTelefone_Agenda.Text = Format(txtTelefone_Agenda.Text, "(00) 0000-0000")
    Else
       txtTelefone_Agenda.Text = Format(txtTelefone_Agenda.Text, "0000-0000-000")
    End If
End Sub

Private Sub txtUf_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Monta_DataCombo()
    strSql = "SELECT * FROM TBCidade_Otica"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_Correios_TBCidade_Otica", "DFNome_TBCidade_Otica", dtcCidade, strSql, "BDRetaguarda", "Otica", Me
        
    strSql = "SELECT IXCodigo_TBVendedor,DFNome_TBVendedor FROM TBVendedor"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBVendedor", "DFNome_TBVendedor", dtcVendedor, strSql, "BDRetaguarda", "Otica", Me
End Function

Private Function Limpa_HFGAgenda()
    'Removendo linhas do grid, evitando assim que fiquem linhas em branco.
    hfgAgenda.ClearStructure
    Do While hfgAgenda.Rows <= hfgAgenda.Rows + 1
       hfgAgenda.Col = 1
       If hfgAgenda.Text = "" And hfgAgenda.Rows = 2 Then
          Exit Do
       End If
       hfgAgenda.Row = hfgAgenda.Rows - 1
       hfgAgenda.RemoveItem hfgAgenda.Rows - 1
    Loop
    
    Movimentacoes.Monta_HFlex_Grid hfgAgenda, "0,0,4800,2000", "ID,ID,Contato,Telefone", 4, "Otica", Me
End Function

Private Function Limpa_HFGAtendimento()
    'Removendo linhas do grid, evitando assim que fiquem linhas em branco.
    hfgAtendimento.ClearStructure
    Do While hfgAtendimento.Rows <= hfgAtendimento.Rows + 1
       hfgAtendimento.Col = 1
       If hfgAtendimento.Text = "" And hfgAtendimento.Rows = 2 Then
          Exit Do
       End If
       hfgAtendimento.Row = hfgAtendimento.Rows - 1
       hfgAtendimento.RemoveItem hfgAtendimento.Rows - 1
    Loop

    Movimentacoes.Monta_HFlex_Grid hfgAtendimento, "0,1000,4000,0,1400,1600,3000", "ID,Vendedor,Nome,ID,Data Previsão,Data Atendimento,Motivo", 7, "Otica", Me
End Function

Private Sub txtValor_Orcamento_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtValor_Orcamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = "44" Or KeyAscii = "46" Then
       Exit Sub
    ElseIf (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtValor_Orcamento_LostFocus()
    txtValor_Orcamento.Text = Format(txtValor_Orcamento.Text, "#,###0.00")
End Sub

Private Sub txtVendedor_Change()
    dtcVendedor.BoundText = txtVendedor.Text
    If IsNumeric(txtVendedor.Text) = False Then txtVendedor.Text = Empty: Exit Sub
End Sub

Private Sub txtVendedor_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
