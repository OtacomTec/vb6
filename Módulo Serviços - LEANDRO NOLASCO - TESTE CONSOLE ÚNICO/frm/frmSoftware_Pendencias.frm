VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSoftware_Pendencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendências"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSoftware_Pendencias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10365
   Begin TabDlg.SSTab sstPendencias 
      Height          =   6645
      Left            =   0
      TabIndex        =   51
      Top             =   330
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   11721
      _Version        =   393216
      Tab             =   2
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
      TabPicture(0)   =   "frmSoftware_Pendencias.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtNumero"
      Tab(0).Control(1)=   "txtCodigo"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "txtObservacao"
      Tab(0).Control(4)=   "txtPrioridade"
      Tab(0).Control(5)=   "txtStatus"
      Tab(0).Control(6)=   "txtTipo_Servico"
      Tab(0).Control(7)=   "txtPrograma"
      Tab(0).Control(8)=   "txtFuncionario"
      Tab(0).Control(9)=   "txtCliente"
      Tab(0).Control(10)=   "txtMenu"
      Tab(0).Control(11)=   "dtcEmpresa"
      Tab(0).Control(12)=   "dtcMenu"
      Tab(0).Control(13)=   "dtcCliente"
      Tab(0).Control(14)=   "dtcFuncionario"
      Tab(0).Control(15)=   "dtcPrograma"
      Tab(0).Control(16)=   "dtcTipo_servico"
      Tab(0).Control(17)=   "dtcStatus"
      Tab(0).Control(18)=   "dtcPrioridade"
      Tab(0).Control(19)=   "Label12"
      Tab(0).Control(20)=   "Label11"
      Tab(0).Control(21)=   "Label44"
      Tab(0).Control(22)=   "Label9"
      Tab(0).Control(23)=   "Label8"
      Tab(0).Control(24)=   "Label5"
      Tab(0).Control(25)=   "Label4"
      Tab(0).Control(26)=   "Label2"
      Tab(0).Control(27)=   "Label1"
      Tab(0).Control(28)=   "Label3"
      Tab(0).Control(29)=   "Label7"
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "&Estatistica"
      TabPicture(1)   =   "frmSoftware_Pendencias.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "&Listagem"
      TabPicture(2)   =   "frmSoftware_Pendencias.frx":17BA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblAte"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "dtpConsulta_Data_Inicio"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "dtpConsulta_Data_Fim"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cbbCampos"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "hfgPendencia"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtConsulta"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdRefresh"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdConsulta"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdParametros_Consulta_Empresa"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdStatus"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.CommandButton cmdStatus 
         Caption         =   "EA"
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
         Left            =   9060
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Status: (EA) Em Aberto/ (C) Concluída / (T) Todas"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox txtNumero 
         Height          =   360
         Left            =   -73320
         MaxLength       =   9
         TabIndex        =   2
         ToolTipText     =   "Número Formulário"
         Top             =   1470
         Width           =   1500
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   360
         Left            =   -74880
         TabIndex        =   1
         ToolTipText     =   "Código Pendência"
         Top             =   1470
         Width           =   1500
      End
      Begin VB.CommandButton cmdParametros_Consulta_Empresa 
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
         Left            =   8670
         Picture         =   "frmSoftware_Pendencias.frx":17D6
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   780
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data e Hora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Left            =   -68250
         TabIndex        =   39
         Top             =   2010
         Width           =   3450
         Begin MSComCtl2.DTPicker dtpFim 
            Height          =   360
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "Data Fim"
            Top             =   1995
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            _Version        =   393216
            CalendarForeColor=   8388608
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   20316163
            CurrentDate     =   2
         End
         Begin MSComCtl2.DTPicker dtpData_Cadastro 
            Height          =   360
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Data Cadastro"
            Top             =   585
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            CalendarForeColor=   8388608
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   8388608
            Format          =   20316161
            CurrentDate     =   2
         End
         Begin MSComCtl2.DTPicker dtpData_Inicio 
            Height          =   360
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Data Início"
            Top             =   1290
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            _Version        =   393216
            CalendarForeColor=   8388608
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   20316163
            CurrentDate     =   2
         End
         Begin MSComCtl2.DTPicker dtpHora_Cadastro 
            Height          =   360
            Left            =   1770
            TabIndex        =   47
            ToolTipText     =   "Hora Cadastro"
            Top             =   585
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            CalendarForeColor=   8388608
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   8388608
            CustomFormat    =   "hh:mm"
            Format          =   20316162
            UpDown          =   -1  'True
            CurrentDate     =   37881
         End
         Begin MSComCtl2.DTPicker dtpHora_inicio 
            Height          =   360
            Left            =   1770
            TabIndex        =   19
            ToolTipText     =   "Hora Início"
            Top             =   1290
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   635
            _Version        =   393216
            CalendarForeColor=   8388608
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   8388608
            CustomFormat    =   "hh:mm"
            Format          =   20316162
            UpDown          =   -1  'True
            CurrentDate     =   37923
         End
         Begin MSComCtl2.DTPicker dtpHora_Fim 
            Height          =   360
            Left            =   1770
            TabIndex        =   21
            ToolTipText     =   "Hora Fim"
            Top             =   1995
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   635
            _Version        =   393216
            CalendarForeColor=   8388608
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   8388608
            CustomFormat    =   "hh:mm"
            Format          =   20316162
            UpDown          =   -1  'True
            CurrentDate     =   37923
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   1030
            Width           =   450
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   240
            Left            =   120
            TabIndex        =   42
            Top             =   1740
            Width           =   315
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Cadastro"
            Height          =   240
            Left            =   120
            TabIndex        =   40
            Top             =   330
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdConsulta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9450
         Picture         =   "frmSoftware_Pendencias.frx":2818
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9840
         Picture         =   "frmSoftware_Pendencias.frx":4512
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
         TabIndex        =   36
         Top             =   780
         Width           =   6255
      End
      Begin VB.TextBox txtObservacao 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   -74880
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         ToolTipText     =   "Observação"
         Top             =   5490
         Width           =   10065
      End
      Begin VB.TextBox txtPrioridade 
         Height          =   360
         Left            =   -74880
         TabIndex        =   7
         ToolTipText     =   "Código Prioridade Serviço"
         Top             =   2790
         Width           =   1500
      End
      Begin VB.TextBox txtStatus 
         Height          =   360
         Left            =   -69810
         TabIndex        =   15
         ToolTipText     =   "Código Status"
         Top             =   4830
         Width           =   1500
      End
      Begin VB.TextBox txtTipo_Servico 
         Height          =   360
         Left            =   -74880
         TabIndex        =   13
         ToolTipText     =   "Código Tipo Serviço"
         Top             =   4830
         Width           =   1500
      End
      Begin VB.TextBox txtPrograma 
         Height          =   360
         Left            =   -74880
         TabIndex        =   11
         ToolTipText     =   "Código Programa"
         Top             =   4140
         Width           =   1500
      End
      Begin VB.TextBox txtFuncionario 
         Height          =   360
         Left            =   -71760
         TabIndex        =   3
         ToolTipText     =   "Código Funcionário"
         Top             =   1470
         Width           =   1500
      End
      Begin VB.TextBox txtCliente 
         Height          =   360
         Left            =   -74880
         TabIndex        =   5
         ToolTipText     =   "Código Cliente"
         Top             =   2130
         Width           =   1500
      End
      Begin VB.TextBox txtMenu 
         Height          =   360
         Left            =   -74880
         TabIndex        =   9
         ToolTipText     =   "Código Menu"
         Top             =   3450
         Width           =   1500
      End
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   360
         Left            =   -74880
         TabIndex        =   0
         Top             =   810
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcMenu 
         Height          =   360
         Left            =   -73320
         TabIndex        =   10
         ToolTipText     =   "Descrição Menu"
         Top             =   3450
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
         Left            =   -73320
         TabIndex        =   6
         ToolTipText     =   "Nome Cliente"
         Top             =   2130
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcFuncionario 
         Height          =   360
         Left            =   -70170
         TabIndex        =   4
         ToolTipText     =   "Nome Funcionário"
         Top             =   1470
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcPrograma 
         Height          =   360
         Left            =   -73320
         TabIndex        =   12
         ToolTipText     =   "Descrição Programa"
         Top             =   4140
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcTipo_servico 
         Height          =   360
         Left            =   -73320
         TabIndex        =   14
         ToolTipText     =   "Descrição Tipo Serviço"
         Top             =   4830
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcStatus 
         Height          =   360
         Left            =   -68250
         TabIndex        =   16
         ToolTipText     =   "Descrição Status"
         Top             =   4830
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSDataListLib.DataCombo dtcPrioridade 
         Height          =   360
         Left            =   -73320
         TabIndex        =   8
         ToolTipText     =   "Descrição da Prioridade do Serviço"
         Top             =   2790
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   16777215
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgPendencia 
         Height          =   5295
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   9340
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
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
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   120
         TabIndex        =   34
         Top             =   780
         Width           =   2235
         _ExtentX        =   3942
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
      Begin MSComCtl2.DTPicker dtpConsulta_Data_Fim 
         Height          =   360
         Left            =   6810
         TabIndex        =   38
         Top             =   780
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20316163
         CurrentDate     =   37923
      End
      Begin MSComCtl2.DTPicker dtpConsulta_Data_Inicio 
         Height          =   360
         Left            =   2400
         TabIndex        =   37
         Top             =   780
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   8388608
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20316163
         CurrentDate     =   37923
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   240
         Left            =   -73320
         TabIndex        =   49
         Top             =   1230
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   240
         Left            =   -74880
         TabIndex        =   48
         Top             =   1230
         Width           =   585
      End
      Begin VB.Label lblAte 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "até"
         Height          =   240
         Left            =   5340
         TabIndex        =   45
         Top             =   930
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Observações"
         Height          =   240
         Left            =   -74880
         TabIndex        =   31
         Top             =   5250
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prioridade"
         Height          =   240
         Left            =   -74880
         TabIndex        =   30
         Top             =   2550
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   240
         Left            =   -69810
         TabIndex        =   29
         Top             =   4590
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Serviço"
         Height          =   240
         Left            =   -74880
         TabIndex        =   28
         Top             =   4590
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programa"
         Height          =   240
         Left            =   -74880
         TabIndex        =   27
         Top             =   3900
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário"
         Height          =   240
         Left            =   -71760
         TabIndex        =   26
         Top             =   1230
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   240
         Left            =   -74880
         TabIndex        =   25
         Top             =   1890
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa [ F2 ]"
         Height          =   375
         Left            =   -74880
         TabIndex        =   24
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
         Height          =   240
         Left            =   -74880
         TabIndex        =   23
         Top             =   3210
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10620
      Top             =   390
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
            Picture         =   "frmSoftware_Pendencias.frx":5554
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware_Pendencias.frx":586E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware_Pendencias.frx":5B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware_Pendencias.frx":5F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware_Pendencias.frx":62BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware_Pendencias.frx":65D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSoftware_Pendencias.frx":68F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   10365
      _ExtentX        =   18283
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
End
Attribute VB_Name = "frmSoftware_Pendencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Cadastro de Pendências                                         '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rodrigo Santos                                                 '
' Data de Criação........: 01/03/2006                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strCombo As String
Dim strConsulta As String
Dim strCampo_consulta As String
Public strSql As String
Dim strId As String
Dim booAlterar As Boolean
Dim conexao As New DLLConexao_Sistema.conexao
Dim log As New DLLSystemManager.log
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
Dim strTamanho As String
Dim strNomes As String
Dim intClique_Analise As Integer
Dim intClique_Especificacao As Integer
Dim intContador As Integer
Dim strAnalise_Antiga As String
Public strCodigo_Empresa_Consulta As String
Dim intId As Integer
    
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
    
    Call frmConsole_Software_Pendencias.Show
    
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
       dtpConsulta_Data_Fim.Visible = False
       dtpConsulta_Data_Inicio.Visible = False
       lblAte.Visible = False
       If booPrivilegio_Consultar = True Then: cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Data Cadastro" Or cbbCampos.Text = "Data Início" Or _
       cbbCampos.Text = "Data Fim" Then
       txtConsulta.Visible = False
       dtpConsulta_Data_Fim.Visible = True
       dtpConsulta_Data_Inicio.Visible = True
       lblAte.Visible = True
       dtpConsulta_Data_Inicio.Value = Date
       dtpConsulta_Data_Fim.Value = Date + 15
    Else
       txtConsulta.Visible = True
       txtConsulta.SetFocus
       dtpConsulta_Data_Fim.Visible = False
       dtpConsulta_Data_Inicio.Visible = False
       lblAte.Visible = False
    End If
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdParametros_Consulta_Empresa_Click()
'STRING QUE COLETA DADOS RELATIVOS A ACESSIBILIDADE DO USUARIO
    Dim rstAcesso_Consulta_Empresa As New ADODB.Recordset

    strSql = Empty
    strSql = "SELECT  DFNivel_TBUsuario FROM TBUsuario " & _
             "WHERE DFNome_TBUsuario = '" & MDIPrincipal.OCXUsuario.Nome & "'"

    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstAcesso_Consulta_Empresa, "Otica", Me

    If rstAcesso_Consulta_Empresa!DFNivel_TBUsuario < 5 Then
       Exit Sub
    End If

    Set rstAcesso_Consulta_Empresa = Nothing

    Unload frmSoftware_Pendencias_Consulta_Empresa
    frmAguarde.Show
    DoEvents
    frmSoftware_Pendencias_Consulta_Empresa.Show
    Unload frmAguarde
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    Call Consulta
End Sub

Private Sub cmdStatus_Click()
    If cmdStatus.Caption = "EA" Then
       cmdStatus.Caption = "C"
    ElseIf cmdStatus.Caption = "C" Then
       cmdStatus.Caption = "T"
    ElseIf cmdStatus.Caption = "T" Then
       cmdStatus.Caption = "EA"
    End If
End Sub

Private Sub dtcCliente_GotFocus()
    If Me.txtCliente.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCliente.Text)
    End If
End Sub

Private Sub dtcCliente_LostFocus()
    txtCliente.Text = dtcCliente.BoundText
End Sub

Private Sub dtcEmpresa_Change()
    Call Monta_Data_Combos
End Sub

Private Sub dtcEmpresa_LostFocus()
    dtcEmpresa.Enabled = False
End Sub

Private Sub dtcFuncionario_GotFocus()
    If Me.txtFuncionario.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcFuncionario.Text)
    End If
End Sub

Private Sub dtcFuncionario_LostFocus()
    txtFuncionario.Text = dtcFuncionario.BoundText
End Sub

Private Sub dtcMenu_GotFocus()
    If Me.txtMenu.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcMenu.Text)
    End If
End Sub

Private Sub dtcMenu_LostFocus()
    txtMenu.Text = dtcMenu.BoundText
    
    If dtcMenu.BoundText <> Empty Then
       strSql = "SELECT PKId_TBProgramas,DFDescricao_TBProgramas FROM TBProgramas " & _
                "WHERE FKID_Menu = '" & dtcMenu.BoundText & "'"
       Movimentacoes.Movimenta_DataCombo "PKId_TBProgramas", "DFDescricao_TBProgramas", dtcPrograma, strSql, "BDRetaguarda", "Otica", Me
    Else
       strSql = "SELECT PKId_TBProgramas,DFDescricao_TBProgramas FROM TBProgramas "
       Movimentacoes.Movimenta_DataCombo "PKId_TBProgramas", "DFDescricao_TBProgramas", dtcPrograma, strSql, "BDRetaguarda", "Otica", Me
    End If
    
    dtcPrograma.BoundText = txtPrograma.Text
    
    If Not IsNumeric(dtcPrograma.BoundText) Then txtPrograma.Text = Empty
    If txtMenu.Text = Empty Then txtPrograma.Text = Empty
End Sub

Private Sub dtcPrioridade_GotFocus()
    If Me.txtPrioridade.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcPrioridade.Text)
    End If
End Sub

Private Sub dtcPrioridade_LostFocus()
    txtPrioridade.Text = dtcPrioridade.BoundText
End Sub

Private Sub dtcPrograma_GotFocus()
    If Me.txtPrograma.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcPrograma.Text)
    End If
End Sub

Private Sub dtcPrograma_LostFocus()
    txtPrograma.Text = dtcPrograma.BoundText
    
    If dtcMenu.BoundText = Empty And txtPrograma.Text <> Empty Then
       Dim rstBusca_Modulo As New ADODB.Recordset
    
       strSql = "SELECT PKId_TBMenu " & _
                "FROM TBMenu, TBProgramas " & _
                "WHERE TBProgramas.FKId_Menu = TBMenu.PKId_TBMenu " & _
                "AND PKId_TBProgramas = '" & txtPrograma.Text & "' "
       
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstBusca_Modulo, "Otica", Me
       
       If rstBusca_Modulo.EOF = False Then
          txtMenu.Text = rstBusca_Modulo!PKId_TBMenu
       End If
       
       Set rstBusca_Modulo = Nothing
    End If
End Sub

Private Sub dtcStatus_GotFocus()
    If Me.txtStatus.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcStatus.Text)
    End If
End Sub

Private Sub dtcStatus_LostFocus()
    txtStatus.Text = dtcStatus.BoundText
End Sub

Private Sub dtcTipo_servico_GotFocus()
    If Me.txtTipo_Servico.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcTipo_servico.Text)
    End If
End Sub

Private Sub dtcTipo_servico_LostFocus()
    txtTipo_Servico.Text = dtcTipo_servico.BoundText
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
    If KeyCode = "113" Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
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
    
    'Carrega data combo Empresa
    strSql = "SELECT PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
   
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Cadastro de Pendências"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstPendencias, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o cadastro de Pendências"
       
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    strCodigo_Empresa_Consulta = Empty
    
    sstPendencias.TabEnabled(0) = False
    sstPendencias.TabEnabled(1) = False
    sstPendencias.Tab = 2
    
    dtpConsulta_Data_Inicio.Value = "1/1/1900"
    
    Call Reposicao
    
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
    log.Descricao = "Finalizando o cadastro de Pendências"
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Set log = Nothing
    
    strCombo = Empty
    strCodigo_Empresa_Consulta = Empty
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
End Sub

Private Sub hfgPendencia_Click()

 If hfgPendencia.Col = 0 And hfgPendencia.Text <> Empty Then
      On Error Resume Next
       
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
      
      intId = Empty
      
      dtcEmpresa.BoundText = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 24))
      intId = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 1))
      txtCodigo.Text = intId
      txtNumero.Text = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 2))
      dtpData_Cadastro.Value = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 3))
      txtFuncionario.Text = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 4))
      txtPrioridade.Text = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 6))
      txtCliente.Text = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 8))
      txtMenu.Text = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 10))
      txtPrograma.Text = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 12))
      txtTipo_Servico.Text = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 14))
      txtStatus.Text = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 16))
      dtpData_Inicio.Value = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 18))
      dtpFim.Value = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 19))
      dtpHora_inicio.Value = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 20))
      dtpHora_Fim.Value = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 21))
      dtpHora_Cadastro.Value = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 22))
      txtObservacao.Text = hfgPendencia.TextArray((hfgPendencia.Row * hfgPendencia.Cols + hfgPendencia.Col + 23))
      
                
      booAlterar = True
      txtConsulta.Text = Empty
      sstPendencias.TabEnabled(0) = True
      sstPendencias.Tab = 0
                            
   End If
   
   Unload frmAguarde

End Sub

Private Sub hfgPendencia_DblClick()
    hfgPendencia.Sort = 1
End Sub

Private Sub hfgPendencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgPendencia_Click
    End If
End Sub

Private Sub sstPendencias_Click(PreviousTab As Integer)
    If sstPendencias.Tab = 0 Then
       txtNumero.SetFocus
    ElseIf sstPendencias.Tab = 1 Then
       'txtAnalise.SetFocus
    ElseIf sstPendencias.Tab = 2 Then
      If frmIntegracao.Visible = True Then
         Unload frmIntegracao
      End If
      If strCombo <> Empty And strCombo <> "Todos" Then
         If strCombo = "Data Cadastro" Or strCombo = "Data Início" Or strCombo = "Data Fim" Then
            dtpConsulta_Data_Inicio.SetFocus
         Else
            cbbCampos.Text = strCombo
            txtConsulta.SetFocus
         End If
      ElseIf strCombo = "Todos" Then
         hfgPendencia.Row = 1
         hfgPendencia.Col = 0
         hfgPendencia.SetFocus
      End If
   End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstPendencias.Tab <> 2: Call Gravar
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
    If txtFuncionario.Text = Empty Then
       MsgBox "O campo funcionário não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtFuncionario.SetFocus
       Exit Function
    End If
    If txtCliente.Text = Empty Then
       MsgBox "O campo cliente não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtCliente.SetFocus
       Exit Function
    End If
    If txtPrograma.Text = Empty Then
       MsgBox "O campo cliente não pode ser nulo. Verifique.", vbInformation, "Only Tech"
       txtCliente.SetFocus
       Exit Function
    End If
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim intID_Cliente As Integer
    Dim rstVerifica As New ADODB.Recordset
        
    intID_Cliente = Funcoes_Gerais.Localiza_ID("PKId_TBCliente", "IXCodigo_TBCliente", dtcCliente.BoundText, "TBCliente", _
                    "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcEmpresa.BoundText)
    
    strCampo = "FKCodigo_TBFuncionario,FKID_TBCliente," & _
               "FKCodigo_TBEmpresa,FKID_TBProgramas,DFData_Cadastro_TBPendencia_servico," & _
               "DFData_Inicio_TBPendencia_servico,DFData_fim_TBPendencia_servico," & _
               "FKCodigo_TBStatus_Pendencia_Servico,DFHora_Cadastro_TBPendencia_servico," & _
               "DFHora_Inicio_TBPendencia_servico,DFHora_Fim_TBPendencia_servico," & _
               "DFObservacao_TBPendencia_servico,FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico," & _
               "TBPendencia_Servicos.FKCodigo_TBTipo_servico_Pendencia_Servico,FKID_TBMenu," & _
               "DFNumero_Relatorio_TBPendencia_servico,DFData_alteracao_TBPendencia_Servicos," & _
               "DFIntegrado_filiais_TBPendencia_Servicos "
               
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBPendencia_Servicos "
    End If
    
    strValores = "" & dtcFuncionario.BoundText & "," & intID_Cliente & "," & dtcEmpresa.BoundText & "," & _
                 "" & dtcPrograma.BoundText & ",'" & Format(dtpData_Cadastro.Value, "YYYYMMDD") & "'," & _
                 "'" & Format(dtpData_Inicio.Value, "YYYYMMDD") & "','" & Format(dtpFim, "YYYYMMDD") & "'," & _
                 "" & dtcStatus.BoundText & ",'" & Format(dtpHora_Cadastro, "HH:MM:SS") & "'," & _
                 "'" & Format(dtpHora_inicio.Value, "HH:MM:SS") & "','" & Format(dtpHora_Fim.Value, "HH:MM:SS") & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "'," & dtcPrioridade.BoundText & "," & _
                 "'" & dtcTipo_servico.BoundText & "','" & dtcMenu.BoundText & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtNumero.Text) & "'," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0 "
                 
    If booIntegra_Portal = True Then
        strValores = strValores & ",0 "
    End If
                 
    'Abrindo conexao
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    On Error GoTo Erro_transacao
    
    If booAlterar = True Then
       log.Evento = "Alterar"

       strSet = "UPDATE TBPendencia_Servicos " & _
                "SET FKCodigo_TBFuncionario = '" & dtcFuncionario.BoundText & "'," & _
                "FKID_TBCliente = '" & intID_Cliente & "'," & _
                "FKCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'," & _
                "FKID_TBProgramas = '" & dtcPrograma.BoundText & "'," & _
                "DFData_Cadastro_TBPendencia_servico = '" & Format(dtpData_Cadastro.Value, "YYYYMMDD") & "'," & _
                "DFData_Inicio_TBPendencia_servico = '" & Format(dtpData_Inicio.Value, "YYYYMMDD") & "'," & _
                "DFData_fim_TBPendencia_servico = '" & Format(dtpFim, "YYYYMMDD") & "'," & _
                "FKCodigo_TBStatus_Pendencia_Servico = '" & dtcStatus.BoundText & "'," & _
                "DFHora_Cadastro_TBPendencia_servico = '" & Format(dtpHora_Cadastro, "HH:MM:SS") & "'," & _
                "DFHora_Inicio_TBPendencia_servico = '" & Format(dtpHora_inicio.Value, "HH:MM:SS") & "'," & _
                "DFHora_Fim_TBPendencia_servico = '" & Format(dtpHora_Fim.Value, "HH:MM:SS") & "'," & _
                "DFObservacao_TBPendencia_servico = '" & Funcoes_Gerais.Grava_String(txtObservacao.Text) & "'," & _
                "FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico = '" & dtcPrioridade.BoundText & "'," & _
                "FKCodigo_TBTipo_servico_Pendencia_Servico = '" & dtcTipo_servico.BoundText & "'," & _
                "FKID_TBMenu = '" & dtcMenu.BoundText & "'," & _
                "DFNumero_Relatorio_TBPendencia_servico = '" & txtNumero.Text & "'," & _
                "DFData_alteracao_TBPendencia_Servicos = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBPendencia_Servicos = 0 "
                
       If booIntegra_Portal = True Then
          strSet = strSet & ",DFIntegrado_portal_TBPendencia_Servicos = 0 "
       End If
       
       strSet = strSet & "WHERE PKID_TBPendencia_servico = '" & intId & "'"
              
       conexao.CNConexao.Execute strSet
                     
       log.Descricao = "Alterando o registro: " & " & intId & "
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"

       strSql = "INSERT INTO TBPendencia_Servicos(" & strCampo & ") VALUES(" & strValores & ")"
       
       conexao.CNConexao.Execute strSql
                     
       log.Descricao = "Gravando o registro: " & intId
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
       
    End If
    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
    
    
    
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
       Me.hfgPendencia.Visible = False
    End If
    
    sstPendencias.TabEnabled(0) = False
    sstPendencias.TabEnabled(1) = False
    sstPendencias.Tab = 2
    
    Exit Function
    
Erro_transacao:
    
    conexao.CNConexao.RollbackTrans
    conexao.Fechar_conexao
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " & intId
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "OTICA", Me
    
    'Abrindo conexao
    conexao.Abrir_conexao "Otica"
     
    'Excluindo Registro
    strSql = "DELETE TBPendencia_Servicos WHERE PKID_TBPendencia_servico = " & intId & ""
    
    conexao.CNConexao.Execute strSql
    
    Call Objetos.Limpa_TXT(Me)
    dtpHora_Fim.Value = "00:00:00"
    dtpHora_inicio = Now
    dtpData_Inicio.Value = Date
    
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
       hfgPendencia.Visible = False
    End If
           
    sstPendencias.TabEnabled(0) = False
    sstPendencias.TabEnabled(1) = False
    sstPendencias.Tab = 2
    
    'fechando conexao
    conexao.Fechar_conexao
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Excluir")
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    
    dtpHora_Fim.Value = "00:00:00"
    dtpHora_inicio.Value = "00:00:00"
    dtpHora_Cadastro.Value = Now
    dtpData_Cadastro.Value = Date
    dtpData_Inicio.Value = "01/01/1900"
    dtpFim.Value = "01/01/1900"
    
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
    'Integração
    tlbBotoes.Buttons.Item(9).Enabled = False
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
    
    If booPrivilegio_Consultar = False Then
       hfgPendencia.Visible = False
    End If
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstPendencias.TabEnabled(0) = False
    sstPendencias.TabEnabled(1) = False
    sstPendencias.Tab = 2
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function

Private Function Novo()
    On Error GoTo Erro
    
    Dim rstBusca_Parametro As New ADODB.Recordset
    Dim strCodigo_Insumo As String
    
    Call Objetos.Limpa_TXT(Me)
    
    log.Evento = "Novo"
    log.Descricao = "Solicitação de um novo registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
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
    
    sstPendencias.TabEnabled(0) = True
    sstPendencias.TabEnabled(1) = True
    sstPendencias.Tab = 0
    
    txtNumero.SetFocus
    dtpData_Cadastro.Value = Date
    dtpData_Inicio.Value = "01/01/1900"
    dtpHora_inicio.Value = "00:00:00"
    dtpFim.Value = "01/01/1900"
    dtpHora_Fim.Value = "00:00:00"
    dtpHora_Cadastro.Value = Now
    Call Monta_Data_Combos
       
    booAlterar = False
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Novo")
    Exit Function
End Function

Private Function Reposicao()
    On Error GoTo Erro
    
    strTamanho = "800,1000,1300,1550,4000,1400,3000," & _
                 "1300,4000,1100,2200,1600," & _
                 "4000,1600,3000,1200,3000," & _
                 "0,0,0,0,0," & _
                 "7000,1000,4000,"
                 
    strNomes = "Código,Número,Data Cadastro,Cod. Funcionário,Nome Funcionário,Cod. Prioridade,Descrição Prioridade," & _
               "Cod. Cliente,Nome Cliente,Cod. Menu,Descrição Menu,Cod. Programa," & _
               "Descrição Programa,Cod. Tipo Serviço,Tipo Serviço,Cod. Status,Descrição Status," & _
               "Data Início,Data Fim,Hora Início,Hora Fim,Hora Cadastro," & _
               "Observação,Empresa,Razão Social"
    
    Movimentacoes.Monta_HFlex_Grid hfgPendencia, strTamanho, strNomes, 25, "Otica", Me

    Call Monta_Combo
    Call Monta_Data_Combos
        
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Reposicao")
    Resume Next
End Function

Private Sub txtCliente_Change()
    dtcCliente.BoundText = txtCliente.Text
    If IsNumeric(txtCliente.Text) = False Then txtCliente.Text = Empty: Exit Sub
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCliente_LostFocus()
    If dtcCliente.Text = Empty Then txtCliente.Text = Empty
End Sub

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConsulta_LostFocus()
    txtConsulta.Text = UCase(txtConsulta.Text)
End Sub

Private Function Consulta()
    If cbbCampos.Text <> "Todos" And cbbCampos.Text <> "Data Cadastro" And cbbCampos.Text <> "Data Início" And _
       cbbCampos.Text <> "Data Fim" Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbInformation, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    End If
    
    strSql = "SELECT PKID_TBPendencia_servico,DFNumero_Relatorio_TBPendencia_servico,DFData_Cadastro_TBPendencia_servico," & _
             "TBPendencia_Servicos.FKCodigo_TBFuncionario,DFNome_TBFuncionario," & _
             "TBPendencia_Servicos.FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico,DFDescricao__TBPrioridade_pendencia_servico," & _
             "IXCodigo_TBCliente,DFNome_TBCliente," & _
             "TBPendencia_Servicos.FKID_TBMenu,DFDescricao_TBMenu,FKID_TBProgramas,DFDescricao_TBProgramas," & _
             "TBPendencia_Servicos.FKCodigo_TBTipo_servico_Pendencia_Servico,DFDescricao_TBTipo_Pendencia_servico," & _
             "TBPendencia_Servicos.FKCodigo_TBStatus_Pendencia_Servico,DFDescricao_TBStatus_pendencia_servico," & _
             "DFData_Inicio_TBPendencia_servico,DFData_fim_TBPendencia_servico," & _
             "CONVERT(char,DFHora_Inicio_TBPendencia_servico,108)Hora_Inicio,CONVERT(char,DFHora_Fim_TBPendencia_servico,108)Hora_Fim," & _
             "CONVERT(char,DFHora_Cadastro_TBPendencia_servico,108)Hora_Cadastrado,DFObservacao_TBPendencia_servico,PKCodigo_TBEmpresa," & _
             "DFRazao_Social_TBEmpresa " & _
             "FROM TBPendencia_Servicos " & _
             "INNER JOIN TBFuncionario ON TBPendencia_Servicos.FKCodigo_TBFuncionario = TBFuncionario.PKCodigo_TBFuncionario " & _
             "INNER JOIN TBCliente ON TBPendencia_Servicos.FKID_TBCliente = TBCliente.PKId_TBCliente "
    
    strSql = strSql & "INNER JOIN TBEmpresa ON TBPendencia_Servicos.FKCodigo_TBEmpresa = TBEmpresa.PKCodigo_TBEmpresa " & _
                      "INNER JOIN TBProgramas ON TBPendencia_Servicos.FKID_TBProgramas = TBProgramas.PKId_TBProgramas " & _
                      "INNER JOIN TBStatus_Pendencia_servico " & _
                      "ON TBPendencia_Servicos.FKCodigo_TBStatus_Pendencia_Servico = TBStatus_Pendencia_servico.PKCodigo_TBStatus_pendencia_servico " & _
                      "INNER JOIN TBPrioridade_Pendencia_Servico " & _
                      "ON TBPendencia_Servicos.FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico = TBPrioridade_Pendencia_Servico.PKCodigo__TBPrioridade_pendencia_servico " & _
                      "INNER JOIN TBTipo_servico_Pendencia_Servico " & _
                      "ON TBPendencia_Servicos.FKCodigo_TBTipo_servico_Pendencia_Servico = TBTipo_servico_Pendencia_Servico.PKCodigo_Prioridade_TBTipo_Pendencia_servico " & _
                      "INNER JOIN TBMenu ON TBPendencia_Servicos.FKID_TBMenu = TBMenu.PKId_TBMenu "

    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text
    Funcoes_Gerais.Grava_String (txtConsulta.Text)
       
    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Data Cadastro" Then
          strSql = strSql & " WHERE DFData_Cadastro_TBPendencia_servico BETWEEN '" & Format(dtpConsulta_Data_Inicio.Value, "YYYYMMDD") & "' AND '" & Format(dtpConsulta_Data_Fim.Value, "YYYYMMDD") & "'"
       ElseIf cbbCampos.Text = "Cod. Funcionário" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE FKCodigo_TBFuncionario = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Nome Funcionário" Then
          strSql = strSql & " WHERE DFNome_TBFuncionario LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Cod. Prioridade" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE FKCodigo_Prioridade_TBPrioridade_Pendencia_Servico = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Descrição Prioridade" Then
          strSql = strSql & " WHERE DFDescricao__TBPrioridade_pendencia_servico LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Cod. Cliente" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE IXCodigo_TBCliente = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Nome Cliente" Then
          strSql = strSql & " WHERE DFNome_TBCliente LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Cod. Menu" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE FKID_TBMenu = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Descrição Menu" Then
          strSql = strSql & " WHERE DFDescricao_TBMenu LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Cod. Programa" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE FKID_TBProgramas = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Descrição Programa" Then
          strSql = strSql & " WHERE DFDescricao_TBProgramas LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Cod. Tipo Serviço" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE FKCodigo_TBTipo_servico_Pendencia_Servico = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Tipo Serviço" Then
          strSql = strSql & " WHERE DFDescricao_TBTipo_Pendencia_servico LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Cod. Status" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " WHERE FKCodigo_TBStatus_Pendencia_Servico = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Descrição Status" Then
          strSql = strSql & " WHERE DFDescricao_TBStatus_pendencia_servico LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Data Início" Then
          strSql = strSql & " WHERE DFData_Inicio_TBPendencia_servico BETWEEN '" & Format(dtpConsulta_Data_Inicio.Value, "YYYYMMDD") & "' AND '" & Format(dtpConsulta_Data_Fim.Value, "YYYYMMDD") & "'"
       ElseIf cbbCampos.Text = "Data Fim" Then
          strSql = strSql & " WHERE DFData_fim_TBPendencia_servico BETWEEN '" & Format(dtpConsulta_Data_Inicio.Value, "YYYYMMDD") & "' AND '" & Format(dtpConsulta_Data_Fim.Value, "YYYYMMDD") & "'"
       ElseIf cbbCampos.Text = "Observação" Then
          strSql = strSql & " WHERE DFObservacao_TBPendencia_servico LIKE '%" & txtConsulta.Text & "%' "
       ElseIf cbbCampos.Text = "Formulário" Then
          strSql = strSql & " WHERE DFNumero_Relatorio_TBPendencia_servico = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Código" Then
          strSql = strSql & " WHERE PKID_TBPendencia_servico = '" & txtConsulta.Text & "' "
       End If
       If IsNumeric(strCodigo_Empresa_Consulta) = False Then
           strSql = strSql & " AND TBPendencia_Servicos.FKCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' "
       ElseIf IsNumeric(strCodigo_Empresa_Consulta) = True And CDbl(strCodigo_Empresa_Consulta) <> 0 Then
           strSql = strSql & " AND TBPendencia_Servicos.FKCodigo_TBEmpresa = '" & strCodigo_Empresa_Consulta & "' "
       End If
    Else
       If IsNumeric(strCodigo_Empresa_Consulta) = False Then
          strSql = strSql & " WHERE TBPendencia_Servicos.FKCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' "
       ElseIf IsNumeric(strCodigo_Empresa_Consulta) = True And CDbl(strCodigo_Empresa_Consulta) <> 0 Then
          strSql = strSql & " WHERE TBPendencia_Servicos.FKCodigo_TBEmpresa = '" & strCodigo_Empresa_Consulta & "' "
       End If
    End If
    
    If Me.cmdStatus.Caption = "EA" Then
       strSql = strSql & " AND FKCodigo_TBStatus_Pendencia_Servico = 1"
    End If
    
    If Me.cmdStatus.Caption = "C" Then
       strSql = strSql & " AND FKCodigo_TBStatus_Pendencia_Servico = 2 or FKCodigo_TBStatus_Pendencia_Servico = 4"
    End If
    
    frmAguarde.Show
    DoEvents
            
    strSql = strSql & " ORDER BY PKID_TBPendencia_servico"
       
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgPendencia, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    hfgPendencia.Col = 0
    hfgPendencia.Row = 1
    If hfgPendencia.Text = Empty Then
       hfgPendencia.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgPendencia, strTamanho, strNomes, 25, "Otica", Me
    End If
    
    Unload frmAguarde
End Function

Private Function Monta_Combo()
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Código")
    cbbCampos.AddItem ("Formulário")
    cbbCampos.AddItem ("Data Cadastro")
    cbbCampos.AddItem ("Cod. Funcionário")
    cbbCampos.AddItem ("Nome Funcionário")
    cbbCampos.AddItem ("Cod. Prioridade")
    cbbCampos.AddItem ("Descrição Prioridade")
    cbbCampos.AddItem ("Cod. Cliente")
    cbbCampos.AddItem ("Nome Cliente")
    cbbCampos.AddItem ("Cod. Menu")
    cbbCampos.AddItem ("Descrição Menu")
    cbbCampos.AddItem ("Cod. Programa")
    cbbCampos.AddItem ("Descrição Programa")
    cbbCampos.AddItem ("Cod. Tipo Serviço")
    cbbCampos.AddItem ("Tipo Serviço")
    cbbCampos.AddItem ("Cod. Status")
    cbbCampos.AddItem ("Descrição Status")
    cbbCampos.AddItem ("Data Início")
    cbbCampos.AddItem ("Data Fim")
    cbbCampos.AddItem ("Observação")
End Function

Private Function Monta_Data_Combos()
    
    strSql = "SELECT PKCodigo_TBFuncionario,DFNome_TBFuncionario FROM TBFuncionario " & _
             "WHERE FKCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBFuncionario", "DFNome_TBFuncionario", dtcFuncionario, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo__TBPrioridade_pendencia_servico,DFDescricao__TBPrioridade_pendencia_servico " & _
             "FROM TBPrioridade_Pendencia_servico"
    Movimentacoes.Movimenta_DataCombo "PKCodigo__TBPrioridade_pendencia_servico", "DFDescricao__TBPrioridade_pendencia_servico", dtcPrioridade, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente " & _
             "WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKId_TBMenu,DFDescricao_TBMenu FROM TBMenu"
    Movimentacoes.Movimenta_DataCombo "PKId_TBMenu", "DFDescricao_TBMenu", dtcMenu, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKId_TBProgramas,DFDescricao_TBProgramas FROM TBProgramas"
    Movimentacoes.Movimenta_DataCombo "PKId_TBProgramas", "DFDescricao_TBProgramas", dtcPrograma, strSql, "BDRetaguarda", "Otica", Me

    strSql = "SELECT PKCodigo_Prioridade_TBTipo_Pendencia_servico,DFDescricao_TBTipo_Pendencia_servico " & _
             "FROM TBTipo_servico_Pendencia_servico"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_Prioridade_TBTipo_Pendencia_servico", "DFDescricao_TBTipo_Pendencia_servico", dtcTipo_servico, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBStatus_pendencia_servico,DFDescricao_TBStatus_pendencia_servico " & _
             "FROM TBStatus_Pendencia_servico"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBStatus_pendencia_servico", "DFDescricao_TBStatus_pendencia_servico", dtcStatus, strSql, "BDRetaguarda", "Otica", Me

End Function

Private Sub txtFuncionario_Change()
    dtcFuncionario.BoundText = txtFuncionario.Text
    If IsNumeric(txtFuncionario.Text) = False Then txtFuncionario.Text = Empty: Exit Sub
End Sub

Private Sub txtFuncionario_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFuncionario_LostFocus()
    If dtcFuncionario.Text = Empty Then txtFuncionario.Text = Empty
End Sub

Private Sub txtMenu_Change()
    dtcMenu.BoundText = txtMenu.Text
    If IsNumeric(txtMenu.Text) = False Then txtMenu.Text = Empty: Exit Sub
End Sub

Private Sub txtMenu_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMenu_LostFocus()
    If dtcMenu.Text = Empty Then txtMenu.Text = Empty
End Sub

Private Sub txtNumero_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtObservacao_LostFocus()
    txtObservacao.Text = UCase(txtObservacao.Text)
End Sub

Private Sub txtPrioridade_Change()
    dtcPrioridade.BoundText = txtPrioridade.Text
    If IsNumeric(txtPrioridade.Text) = False Then txtPrioridade.Text = Empty: Exit Sub
End Sub

Private Sub txtPrioridade_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrioridade_LostFocus()
    If dtcPrioridade.Text = Empty Then txtPrioridade.Text = Empty
End Sub

Private Sub txtPrograma_Change()
    dtcPrograma.BoundText = txtPrograma.Text
    If IsNumeric(txtPrograma.Text) = False Then txtPrograma.Text = Empty: Exit Sub
End Sub

Private Sub txtPrograma_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrograma_LostFocus()
    If dtcPrograma.Text = Empty Then txtPrograma.Text = Empty
End Sub

Private Sub txtStatus_Change()
    dtcStatus.BoundText = txtStatus.Text
    If IsNumeric(txtStatus.Text) = False Then txtStatus.Text = Empty: Exit Sub
End Sub

Private Sub txtStatus_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtStatus_LostFocus()
    If dtcStatus.Text = Empty Then txtStatus.Text = Empty
End Sub

Private Sub txtTipo_Servico_Change()
    dtcTipo_servico.BoundText = txtTipo_Servico.Text
    If IsNumeric(txtTipo_Servico.Text) = False Then txtTipo_Servico.Text = Empty: Exit Sub
End Sub

Private Sub txtTipo_Servico_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTipo_Servico_LostFocus()
    If dtcTipo_servico.Text = Empty Then txtTipo_Servico.Text = Empty
End Sub

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKID_TBPendencia_servico", txtCodigo.Text, "DFIntegrado_filiais_TBPendencia_Servicos", "TBPendencia_Servicos", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBPendencia_Servicos", Me.Top, Me.Left, Me.Width, Me.Height, "Pendência Serviços")
    
End Function
