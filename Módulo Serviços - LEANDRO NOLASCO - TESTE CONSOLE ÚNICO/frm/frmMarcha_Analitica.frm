VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMarcha_Analitica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Marcha Analítica"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   Icon            =   "frmMarcha_Analitica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11790
   Begin TabDlg.SSTab sstMarcha 
      Height          =   7275
      Left            =   0
      TabIndex        =   41
      Top             =   330
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   12832
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
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
      TabPicture(0)   =   "frmMarcha_Analitica.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtTipo_Marcha"
      Tab(0).Control(1)=   "Frame18"
      Tab(0).Control(2)=   "cmdInformacoes_Liberacao"
      Tab(0).Control(3)=   "txtNumero_Sequencial"
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(7)=   "txtObservacao"
      Tab(0).Control(8)=   "dtpData_Previsao"
      Tab(0).Control(9)=   "dtcEmpresa"
      Tab(0).Control(10)=   "cbbPrioridade"
      Tab(0).Control(11)=   "dtcTipo_Marcha"
      Tab(0).Control(12)=   "dtpCompetencia"
      Tab(0).Control(13)=   "Label12"
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(16)=   "Label18"
      Tab(0).Control(17)=   "Label1"
      Tab(0).Control(18)=   "Label13"
      Tab(0).Control(19)=   "Label58"
      Tab(0).Control(20)=   "lblLiberado"
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Pesagem"
      TabPicture(1)   =   "frmMarcha_Analitica.frx":179E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "hfgteste"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Avaliação"
      TabPicture(2)   =   "frmMarcha_Analitica.frx":17BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11"
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(2)=   "Frame9"
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(4)=   "hfgteste2"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Cálculos"
      TabPicture(3)   =   "frmMarcha_Analitica.frx":17D6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame17"
      Tab(3).Control(1)=   "Frame16"
      Tab(3).Control(2)=   "Frame15"
      Tab(3).Control(3)=   "Frame14"
      Tab(3).Control(4)=   "Frame13"
      Tab(3).Control(5)=   "Frame12"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "&Gráficos"
      TabPicture(4)   =   "frmMarcha_Analitica.frx":17F2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "CRViewer91"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Listagem"
      TabPicture(5)   =   "frmMarcha_Analitica.frx":180E
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Label6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "lblA"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "dtpFim_Consulta"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "dtpInicio_Consulta"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "cbbConsulta"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "cbbCampos"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "hfgMarcha"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "cmdRefresh"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "cmdConsulta"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "cmdMarcha_Consulta_Empresa"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "txtConsulta"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).ControlCount=   11
      Begin VB.TextBox txtTipo_Marcha 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   0
         ToolTipText     =   "Código do Tipo de Marcha"
         Top             =   1320
         Width           =   1150
      End
      Begin VB.Frame Frame18 
         Caption         =   "Matéria-Prima"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2730
         Left            =   -65610
         TabIndex        =   169
         Top             =   3840
         Width           =   2265
         Begin VB.TextBox txtLote_Materia 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   27
            ToolTipText     =   "Lote"
            Top             =   1080
            Width           =   2025
         End
         Begin VB.TextBox txtFornecedor_Materia 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   26
            ToolTipText     =   "Fornecedor / Fabricante"
            Top             =   480
            Width           =   2025
         End
         Begin MSComCtl2.DTPicker dtpFabricacao_Materia 
            Height          =   315
            Left            =   120
            TabIndex        =   28
            ToolTipText     =   "Data de Fabricação"
            Top             =   1680
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20381697
            CurrentDate     =   38481
         End
         Begin MSComCtl2.DTPicker dtpValidade_Materia 
            Height          =   315
            Left            =   120
            TabIndex        =   29
            ToolTipText     =   "Data de Validade"
            Top             =   2310
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20381697
            CurrentDate     =   38481
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Data Validade"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   197
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Data Fabricação"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   196
            Top             =   1470
            Width           =   1170
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   171
            Top             =   870
            Width           =   315
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor / Fabricante"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   170
            Top             =   270
            Width           =   1740
         End
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
         Height          =   360
         Left            =   3000
         TabIndex        =   32
         Top             =   780
         Width           =   7425
      End
      Begin VB.CommandButton cmdMarcha_Consulta_Empresa 
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
         Left            =   10500
         Picture         =   "frmMarcha_Analitica.frx":182A
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdConsulta 
         Height          =   360
         Left            =   10890
         Picture         =   "frmMarcha_Analitica.frx":286C
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Consultar"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   360
         Left            =   11280
         Picture         =   "frmMarcha_Analitica.frx":4566
         Style           =   1  'Graphical
         TabIndex        =   160
         TabStop         =   0   'False
         ToolTipText     =   "Recarregar Grid"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton cmdInformacoes_Liberacao 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70050
         Picture         =   "frmMarcha_Analitica.frx":55A8
         Style           =   1  'Graphical
         TabIndex        =   159
         ToolTipText     =   "Informações Cancelamento"
         Top             =   750
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Frame Frame17 
         Caption         =   "Unif. de Doses Unitárias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4980
         Left            =   -69210
         TabIndex        =   153
         Top             =   540
         Width           =   5865
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgteste4 
            Height          =   4515
            Left            =   120
            TabIndex        =   155
            Top             =   330
            Width           =   5625
            _ExtentX        =   9922
            _ExtentY        =   7964
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Peso Médio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4980
         Left            =   -74880
         TabIndex        =   152
         Top             =   540
         Width           =   5625
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgteste3 
            Height          =   4515
            Left            =   120
            TabIndex        =   154
            Top             =   330
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   7964
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Legenda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   -65250
         TabIndex        =   147
         Top             =   5610
         Width           =   1905
         Begin VB.Label Label100 
            AutoSize        =   -1  'True
            Caption         =   "Normal Divergente:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   151
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label99 
            AutoSize        =   -1  'True
            Caption         =   "Normal .............:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   150
            Top             =   660
            Width           =   1380
         End
         Begin VB.Label Label98 
            AutoSize        =   -1  'True
            Caption         =   "Negativo...........:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   149
            Top             =   1260
            Width           =   1365
         End
         Begin VB.Shape Shape12 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1560
            Shape           =   1  'Square
            Top             =   360
            Width           =   225
         End
         Begin VB.Shape Shape11 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1560
            Shape           =   1  'Square
            Top             =   660
            Width           =   225
         End
         Begin VB.Shape Shape10 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1560
            Shape           =   1  'Square
            Top             =   960
            Width           =   225
         End
         Begin VB.Shape Shape9 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1560
            Shape           =   1  'Square
            Top             =   1260
            Width           =   225
         End
         Begin VB.Label Label97 
            AutoSize        =   -1  'True
            Caption         =   "Normal .............:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   148
            Top             =   960
            Width           =   1380
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Parecer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   -69180
         TabIndex        =   140
         Top             =   5610
         Width           =   3885
         Begin VB.Label Label96 
            AutoSize        =   -1  'True
            Caption         =   "Não Conforme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2580
            TabIndex        =   146
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label Label95 
            AutoSize        =   -1  'True
            Caption         =   "Não Conforme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2580
            TabIndex        =   145
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label94 
            AutoSize        =   -1  'True
            Caption         =   "Conforme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   2580
            TabIndex        =   144
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label93 
            AutoSize        =   -1  'True
            Caption         =   "Unif. de doses unitárias 30 UN...:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   143
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label92 
            AutoSize        =   -1  'True
            Caption         =   "Unif. de doses unitárias 10 UN...:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   142
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label91 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio..........................:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   141
            Top             =   360
            Width           =   2430
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   -71790
         TabIndex        =   131
         Top             =   5610
         Width           =   2565
         Begin VB.TextBox ee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   135
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox ww 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   134
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   1155
         End
         Begin VB.TextBox rr 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   133
            ToolTipText     =   "Código do Cliente"
            Top             =   1080
            Width           =   1155
         End
         Begin VB.TextBox tt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   132
            ToolTipText     =   "Descrição do Cliente"
            Top             =   1080
            Width           =   1125
         End
         Begin VB.Label Label90 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   139
            Top             =   870
            Width           =   810
         End
         Begin VB.Label Label89 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   138
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label88 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   137
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label87 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   136
            Top             =   870
            Width           =   495
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Forma Farmacêutica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   -74880
         TabIndex        =   125
         Top             =   5610
         Width           =   3045
         Begin VB.TextBox qq 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1110
            MaxLength       =   50
            TabIndex        =   128
            ToolTipText     =   "Descrição do Cliente"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox qw 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   127
            ToolTipText     =   "Código do Cliente"
            Top             =   1080
            Width           =   945
         End
         Begin VB.TextBox wq 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   126
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   2805
         End
         Begin VB.Label Label81 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   870
            Width           =   495
         End
         Begin VB.Label Label79 
            AutoSize        =   -1  'True
            Caption         =   "Histórico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   129
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Forma Farmacêutica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   -74880
         TabIndex        =   118
         Top             =   5610
         Width           =   3045
         Begin VB.TextBox aaas 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   121
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   2805
         End
         Begin VB.TextBox sa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   120
            ToolTipText     =   "Código do Cliente"
            Top             =   1080
            Width           =   945
         End
         Begin VB.TextBox sdf 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1110
            MaxLength       =   50
            TabIndex        =   119
            ToolTipText     =   "Descrição do Cliente"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label86 
            AutoSize        =   -1  'True
            Caption         =   "Histórico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   123
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label85 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   122
            Top             =   870
            Width           =   495
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   -71790
         TabIndex        =   109
         Top             =   5610
         Width           =   2565
         Begin VB.TextBox ffff 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   113
            ToolTipText     =   "Descrição do Cliente"
            Top             =   1080
            Width           =   1125
         End
         Begin VB.TextBox dd 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   112
            ToolTipText     =   "Código do Cliente"
            Top             =   1080
            Width           =   1155
         End
         Begin VB.TextBox fds 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   111
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   1155
         End
         Begin VB.TextBox ff 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   110
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label Label84 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   117
            Top             =   870
            Width           =   495
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   115
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   114
            Top             =   870
            Width           =   810
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Parecer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   -69180
         TabIndex        =   102
         Top             =   5610
         Width           =   3885
         Begin VB.Label Label78 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio..........................:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   2430
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "Unif. de doses unitárias 10 UN...:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            Caption         =   "Unif. de doses unitárias 30 UN...:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   106
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            Caption         =   "Conforme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   2580
            TabIndex        =   105
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            Caption         =   "Não Conforme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2580
            TabIndex        =   104
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            Caption         =   "Não Conforme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2580
            TabIndex        =   103
            Top             =   720
            Width           =   1185
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Legenda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   -65250
         TabIndex        =   98
         Top             =   5610
         Width           =   1905
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Normal .............:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   124
            Top             =   960
            Width           =   1380
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1560
            Shape           =   1  'Square
            Top             =   1260
            Width           =   225
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1560
            Shape           =   1  'Square
            Top             =   960
            Width           =   225
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1560
            Shape           =   1  'Square
            Top             =   660
            Width           =   225
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1560
            Shape           =   1  'Square
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "Negativo...........:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   101
            Top             =   1260
            Width           =   1365
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Normal .............:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   100
            Top             =   660
            Width           =   1380
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "Normal Divergente:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   99
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Forma Farmacêutica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2190
         Left            =   -67530
         TabIndex        =   80
         Top             =   540
         Width           =   4185
         Begin VB.TextBox t4 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   156
            ToolTipText     =   "Histórico do Transporte"
            Top             =   1680
            Width           =   3945
         End
         Begin VB.TextBox t3 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1110
            MaxLength       =   50
            TabIndex        =   83
            ToolTipText     =   "Descrição do Cliente"
            Top             =   1080
            Width           =   2955
         End
         Begin VB.TextBox t2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   82
            ToolTipText     =   "Código do Cliente"
            Top             =   1080
            Width           =   945
         End
         Begin VB.TextBox t1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   81
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   3945
         End
         Begin VB.Label Label101 
            AutoSize        =   -1  'True
            Caption         =   "Histórico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   157
            Top             =   1470
            Width           =   615
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   870
            Width           =   495
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "Histórico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1590
         Left            =   -67530
         TabIndex        =   74
         Top             =   2760
         Width           =   4185
         Begin VB.TextBox aaaa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2850
            MaxLength       =   50
            TabIndex        =   94
            ToolTipText     =   "Descrição do Cliente"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox t6 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2850
            MaxLength       =   50
            TabIndex        =   92
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox t5 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            MaxLength       =   50
            TabIndex        =   90
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   1305
         End
         Begin VB.TextBox txtHistorico 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   77
            ToolTipText     =   "Histórico do Transporte"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox f 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   76
            ToolTipText     =   "Código do Cliente"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox aa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            MaxLength       =   50
            TabIndex        =   75
            ToolTipText     =   "Descrição do Cliente"
            Top             =   1080
            Width           =   1305
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   96
            Top             =   870
            Width           =   810
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1500
            TabIndex        =   95
            Top             =   870
            Width           =   810
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2850
            TabIndex        =   93
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1500
            TabIndex        =   91
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   870
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Parecer"
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
         Height          =   1545
         Left            =   -67530
         TabIndex        =   70
         Top             =   4380
         Width           =   4185
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "Não Conforme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2700
            TabIndex        =   88
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            Caption         =   "Não Conforme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2700
            TabIndex        =   87
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Conforme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   2700
            TabIndex        =   86
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Unif. de doses unitárias 30 UN...:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Unif. de doses unitárias 10 UN...:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "Peso Médio..........................:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   360
            Width           =   2430
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Legenda"
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
         Height          =   1095
         Left            =   -67530
         TabIndex        =   65
         Top             =   6000
         Width           =   4185
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Normal Divergente..:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Normal ...............:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Negativo Normal.....:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2220
            TabIndex        =   67
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Negativo Divergente:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2220
            TabIndex        =   66
            Top             =   360
            Width           =   1545
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1680
            Shape           =   1  'Square
            Top             =   360
            Width           =   225
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1680
            Shape           =   1  'Square
            Top             =   750
            Width           =   225
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   3840
            Shape           =   1  'Square
            Top             =   750
            Width           =   225
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   3840
            Shape           =   1  'Square
            Top             =   360
            Width           =   225
         End
      End
      Begin VB.TextBox txtNumero_Sequencial 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         TabIndex        =   2
         ToolTipText     =   "Número Sequencial da Marcha"
         Top             =   1890
         Width           =   1545
      End
      Begin VB.Frame Frame3 
         Caption         =   "Insumo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2730
         Left            =   -74880
         TabIndex        =   52
         Top             =   3840
         Width           =   9225
         Begin VB.TextBox txtMedida 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3780
            MaxLength       =   10
            TabIndex        =   23
            ToolTipText     =   "Medida"
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtFabricante 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4680
            MaxLength       =   50
            TabIndex        =   24
            ToolTipText     =   "Fabricante"
            Top             =   1080
            Width           =   2190
         End
         Begin VB.TextBox txtFuncao 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   183
            ToolTipText     =   "Função"
            Top             =   2310
            Width           =   4515
         End
         Begin VB.CommandButton cmdInformacaoes_Adicionais_Insumo 
            Height          =   315
            Left            =   6030
            Picture         =   "frmMarcha_Analitica.frx":5932
            Style           =   1  'Graphical
            TabIndex        =   182
            ToolTipText     =   "Informações Adicionais"
            Top             =   480
            Width           =   405
         End
         Begin VB.TextBox txtPeso 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6480
            TabIndex        =   17
            ToolTipText     =   "Peso"
            Top             =   480
            Width           =   1005
         End
         Begin VB.TextBox txtConservacao 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4680
            TabIndex        =   178
            ToolTipText     =   "Conservação"
            Top             =   1680
            Width           =   4425
         End
         Begin VB.TextBox txtObservacao_Insumo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4680
            MaxLength       =   300
            TabIndex        =   177
            ToolTipText     =   "Observação"
            Top             =   2310
            Width           =   4425
         End
         Begin VB.TextBox txtQuantidade 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2820
            MaxLength       =   50
            TabIndex        =   22
            ToolTipText     =   "Quantidade"
            Top             =   1080
            Width           =   915
         End
         Begin VB.TextBox txtEmbalagem_Insumo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6915
            MaxLength       =   20
            TabIndex        =   25
            ToolTipText     =   "Embalagem"
            Top             =   1080
            Width           =   2190
         End
         Begin VB.TextBox txtNome_Cientifico 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   172
            ToolTipText     =   "Nome Científico"
            Top             =   1680
            Width           =   4515
         End
         Begin VB.TextBox txtInsumo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   15
            ToolTipText     =   "Código do Insumo"
            Top             =   480
            Width           =   1155
         End
         Begin VB.TextBox txtUnidade 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7530
            MaxLength       =   3
            TabIndex        =   18
            ToolTipText     =   "Unidade"
            Top             =   480
            Width           =   510
         End
         Begin VB.TextBox txtLote_Insumo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8085
            MaxLength       =   10
            TabIndex        =   19
            ToolTipText     =   "Lote"
            Top             =   480
            Width           =   1005
         End
         Begin MSDataListLib.DataCombo dtcInsumo 
            Height          =   315
            Left            =   1320
            TabIndex        =   16
            ToolTipText     =   "Descrição do Insumo"
            Top             =   480
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ForeColor       =   8388608
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker dtpFabricacao 
            Height          =   315
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "Data de Fabricação"
            Top             =   1080
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20381697
            CurrentDate     =   38481
         End
         Begin MSComCtl2.DTPicker dtpValidade 
            Height          =   315
            Left            =   1470
            TabIndex        =   21
            ToolTipText     =   "Data de Validade"
            Top             =   1080
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20381697
            CurrentDate     =   38481
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Data Validade"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1470
            TabIndex        =   195
            Top             =   870
            Width           =   990
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Data Fabricação"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   194
            Top             =   870
            Width           =   1170
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Medida"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3780
            TabIndex        =   193
            Top             =   870
            Width           =   510
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Fabricante"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4680
            TabIndex        =   192
            Top             =   870
            Width           =   765
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Função"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   184
            Top             =   2100
            Width           =   525
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6480
            TabIndex        =   180
            Top             =   270
            Width           =   345
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conservação"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4680
            TabIndex        =   179
            Top             =   1470
            Width           =   945
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Observação Insumo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4680
            TabIndex        =   176
            Top             =   2100
            Width           =   1440
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2820
            TabIndex        =   175
            Top             =   870
            Width           =   840
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Embalagem"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6915
            TabIndex        =   174
            Top             =   870
            Width           =   810
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome Científico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   173
            Top             =   1470
            Width           =   1110
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lote"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8085
            TabIndex        =   89
            Top             =   270
            Width           =   315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1320
            TabIndex        =   54
            Top             =   270
            Width           =   690
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7530
            TabIndex        =   53
            Top             =   270
            Width           =   210
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cliente X Serviço"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   -74880
         TabIndex        =   43
         Top             =   2280
         Width           =   11535
         Begin VB.TextBox txtDescricao_Plano 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7050
            MaxLength       =   50
            TabIndex        =   181
            ToolTipText     =   "Descrição do Plano de Serviços"
            Top             =   480
            Width           =   4365
         End
         Begin VB.TextBox txtServicos_restantes 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8340
            MaxLength       =   50
            TabIndex        =   168
            ToolTipText     =   "Serviços restantes no período"
            Top             =   1080
            Width           =   1515
         End
         Begin VB.TextBox txtServico 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Código do Serviço"
            Top             =   1080
            Width           =   1140
         End
         Begin VB.TextBox txtNumero_Servicos 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7050
            MaxLength       =   50
            TabIndex        =   166
            ToolTipText     =   "Limite do Serviço"
            Top             =   1080
            Width           =   1245
         End
         Begin VB.CommandButton cmdConsulta_Cliente 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5220
            Picture         =   "frmMarcha_Analitica.frx":5CBC
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Consulta Detalhada do Histórico"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtCliente 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Código do Cliente"
            Top             =   480
            Width           =   1140
         End
         Begin VB.TextBox txtPlano 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5640
            TabIndex        =   45
            ToolTipText     =   "Código do Plano de Serviços"
            Top             =   480
            Width           =   1365
         End
         Begin VB.TextBox txtContrato 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9900
            MaxLength       =   50
            TabIndex        =   44
            ToolTipText     =   "Número do Contrato"
            Top             =   1080
            Width           =   1515
         End
         Begin VB.TextBox txtControle 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5640
            MaxLength       =   50
            TabIndex        =   40
            ToolTipText     =   "Tipo de Controle Plano"
            Top             =   1080
            Width           =   1365
         End
         Begin MSDataListLib.DataCombo dtcCliente 
            Height          =   315
            Left            =   1300
            TabIndex        =   12
            ToolTipText     =   "Descrição do Cliente"
            Top             =   480
            Width           =   3875
            _ExtentX        =   6826
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ForeColor       =   8388608
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtcServico 
            Height          =   315
            Left            =   1305
            TabIndex        =   14
            ToolTipText     =   "Descrição do Serviço"
            Top             =   1080
            Width           =   4310
            _ExtentX        =   7594
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            ForeColor       =   8388608
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serviço"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   167
            Top             =   870
            Width           =   525
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plano de Serviços"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5640
            TabIndex        =   50
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Limite"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7050
            TabIndex        =   49
            Top             =   870
            Width           =   405
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Número do Contrato"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9900
            TabIndex        =   48
            Top             =   870
            Width           =   1470
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Controle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5640
            TabIndex        =   47
            Top             =   870
            Width           =   615
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Restantes / Período"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8340
            TabIndex        =   46
            Top             =   870
            Width           =   1425
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Acompanhamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1560
         Left            =   -69540
         TabIndex        =   42
         Top             =   660
         Width           =   6195
         Begin VB.CommandButton cmdAcompanhamento 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3990
            Picture         =   "frmMarcha_Analitica.frx":6046
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Informações Adicionais de Acompanhamento"
            Top             =   1140
            Width           =   375
         End
         Begin VB.TextBox txtUsuario 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4410
            MaxLength       =   50
            TabIndex        =   38
            ToolTipText     =   "Usúario"
            Top             =   1140
            Width           =   1665
         End
         Begin MSComCtl2.DTPicker dtpData_Inicio 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Data Chegada Estágio"
            Top             =   540
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20381697
            CurrentDate     =   38481
         End
         Begin MSComCtl2.DTPicker dtpHora_Inicio 
            Height          =   315
            Left            =   1620
            TabIndex        =   7
            ToolTipText     =   "Hora Chegada Estágio"
            Top             =   540
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20381698
            CurrentDate     =   38481
         End
         Begin MSComCtl2.DTPicker dtpData_Fim 
            Height          =   315
            Left            =   3120
            TabIndex        =   8
            ToolTipText     =   "Data Saída Estágio"
            Top             =   540
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20381697
            CurrentDate     =   38481
         End
         Begin MSComCtl2.DTPicker dtpHora_Fim 
            Height          =   315
            Left            =   4620
            TabIndex        =   9
            ToolTipText     =   "Hora Saída Estágio"
            Top             =   540
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20381698
            CurrentDate     =   38481
         End
         Begin AutoCompletar.CbCompleta cbbEstagio 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Estágio do Acompanhamento"
            Top             =   1140
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388608
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Usuário"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4410
            TabIndex        =   190
            Top             =   930
            Width           =   540
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Estágio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   189
            Top             =   930
            Width           =   525
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Data Chegada"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   188
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Hora Chegada"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1620
            TabIndex        =   187
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Data Saída"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3120
            TabIndex        =   186
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Hora Saída"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4620
            TabIndex        =   185
            Top             =   330
            Width           =   780
         End
      End
      Begin VB.TextBox txtObservacao 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74880
         MaxLength       =   50
         TabIndex        =   30
         ToolTipText     =   "Observação Marcha"
         Top             =   6810
         Width           =   11535
      End
      Begin MSComCtl2.DTPicker dtpData_Previsao 
         Height          =   315
         Left            =   -72180
         TabIndex        =   4
         ToolTipText     =   "Data Previsão"
         Top             =   1890
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20381697
         CurrentDate     =   38481
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   315
         Left            =   -74880
         TabIndex        =   56
         ToolTipText     =   "Empresa"
         Top             =   750
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin AutoCompletar.CbCompleta cbbPrioridade 
         Height          =   315
         Left            =   -73290
         TabIndex        =   3
         ToolTipText     =   "Prioridade"
         Top             =   1890
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgteste 
         Height          =   6555
         Left            =   -74880
         TabIndex        =   64
         Top             =   540
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   11562
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgteste2 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   97
         Top             =   540
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8811
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgMarcha 
         Height          =   5895
         Left            =   120
         TabIndex        =   162
         Top             =   1230
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10398
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
      End
      Begin AutoCompletar.CbCompleta cbbCampos 
         Height          =   360
         Left            =   120
         TabIndex        =   31
         Top             =   780
         Width           =   2835
         _ExtentX        =   5001
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
      Begin AutoCompletar.CbCompleta cbbConsulta 
         Height          =   360
         Left            =   3000
         TabIndex        =   33
         Top             =   780
         Width           =   7425
         _ExtentX        =   13097
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
         Left            =   3000
         TabIndex        =   34
         Top             =   780
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   20381697
         CurrentDate     =   38386
      End
      Begin MSComCtl2.DTPicker dtpFim_Consulta 
         Height          =   360
         Left            =   8820
         TabIndex        =   35
         Top             =   780
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   20381697
         CurrentDate     =   38386
      End
      Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
         Height          =   6585
         Left            =   -74880
         TabIndex        =   165
         Top             =   540
         Width           =   11535
         lastProp        =   500
         _cx             =   20346
         _cy             =   11615
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   0   'False
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
         DisplayTabs     =   0   'False
         DisplayBackgroundEdge=   0   'False
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   -1  'True
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
         LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      End
      Begin MSDataListLib.DataCombo dtcTipo_Marcha 
         Height          =   315
         Left            =   -73680
         TabIndex        =   1
         ToolTipText     =   "Descrição do Tipo de Marcha"
         Top             =   1320
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpCompetencia 
         Height          =   315
         Left            =   -70800
         TabIndex        =   5
         ToolTipText     =   "Competência (MM/AAAA)"
         Top             =   1890
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/yyyy"
         Format          =   20381699
         CurrentDate     =   38481
         MinDate         =   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Competência"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70800
         TabIndex        =   191
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "até"
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
         Left            =   6420
         TabIndex        =   164
         Top             =   900
         Width           =   270
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
         TabIndex        =   163
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Marcha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   63
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Sequencial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   62
         Top             =   1680
         Width           =   990
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   -3360
         Top             =   -4200
         Width           =   5145
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa [ F2 ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   61
         Top             =   540
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Previsão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72180
         TabIndex        =   60
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Prioridade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73290
         TabIndex        =   59
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   58
         Top             =   6600
         Width           =   870
      End
      Begin VB.Label lblLiberado 
         AutoSize        =   -1  'True
         Caption         =   "Liberado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -71055
         TabIndex        =   57
         Top             =   840
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11850
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcha_Analitica.frx":63D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcha_Analitica.frx":66EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcha_Analitica.frx":6A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcha_Analitica.frx":6D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcha_Analitica.frx":7138
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcha_Analitica.frx":7452
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcha_Analitica.frx":776C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMarcha_Analitica.frx":85BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   158
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
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
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMarcha_Analitica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Transporte                                                     '
' Objetivo...............: Cadastro de Marcha Analítica                                   '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Sá Peixoto                                               '
' Data de Criação........: 29/12/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strTamanho As String
Dim strNomes As String
Dim booAlterar As Boolean
Dim conexao As New DLLConexao_Sistema.conexao
'Declaração das variaveis da acessibilidade
Public strSql As String
Dim rstAplicacao As New ADODB.Recordset
Dim rstServico As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log
Dim booPrivilegio_Incluir As Boolean
Dim booPrivilegio_Alterar As Boolean
Dim booPrivilegio_Excluir As Boolean
Dim booPrivilegio_Consultar As Boolean
Public strCodigo_Empresa_Consulta As String
Dim strUsuario_Digitacao As String
Dim strCombo As String
Dim strConsulta As String
Dim intContador As Integer
Public strID_Marcha As String
Dim strId_Cliente As String
'Variável para controle de click para abastecimento de verificacoes de servico
Dim booClick_Grid As Boolean
'Variável para controle de modificacao de estágio
Dim strAcompanhamento As String
Dim booIntegracao As Boolean
Dim booIntegra_Portal As Boolean
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
    
    Call frmConsole_Marcha_Analitica.Show
    
    Unload frmAguarde
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Sub cbbCampos_Click()
    txtConsulta.Text = Empty
    cbbConsulta.Text = Empty
    dtpInicio_Consulta.Value = Date
    dtpFim_Consulta.Value = Date
    
    If cbbCampos.Text = "Todos" Then
       txtConsulta.Visible = False
       cbbConsulta.Visible = False
       lblA.Visible = False
       dtpInicio_Consulta.Visible = False
       dtpFim_Consulta.Visible = False
       If booPrivilegio_Consultar = True Then cmdConsulta.SetFocus
    ElseIf cbbCampos.Text = "Previsão" Or cbbCampos.Text = "Data Fabricação" Or cbbCampos.Text = "Data Validade" Or _
       cbbCampos.Text = "Data Fabricação Matéria" Or cbbCampos.Text = "Data Validade Matéria" Or cbbCampos.Text = "Competência" Then
       txtConsulta.Visible = False
       cbbConsulta.Visible = False
       lblA.Visible = True
       dtpInicio_Consulta.Visible = True
       dtpFim_Consulta.Visible = True
    ElseIf cbbCampos.Text = "Prioridade" Then
       txtConsulta.Visible = False
       cbbConsulta.Visible = True
       lblA.Visible = False
       dtpInicio_Consulta.Visible = False
       dtpFim_Consulta.Visible = False
    Else
       txtConsulta.Visible = True
       cbbConsulta.Visible = False
       lblA.Visible = False
       dtpInicio_Consulta.Visible = False
       dtpFim_Consulta.Visible = False
    End If
    
End Sub

Private Sub cmdAcompanhamento_Click()
    If booAlterar = True Then
       Unload frmMarcha_Analitica_Acompanhamento
       frmMarcha_Analitica_Acompanhamento.Show
    End If
End Sub

Private Sub cmdConsulta_Click()
    Call Consulta
End Sub

Private Sub cmdRefresh_Click()
    cbbCampos.Text = strCombo
    txtConsulta.Text = strConsulta
    
    Call Consulta
End Sub
Private Sub cmdConsulta_Cliente_Click()
    Unload frmMarcha_Consulta_Cliente
    frmMarcha_Consulta_Cliente.Show
End Sub

Private Sub cmdInformacaoes_Adicionais_Insumo_Click()
    If txtInsumo.Text <> Empty And dtcInsumo.Text <> Empty Then
       Unload frmMarcha_Informacoes_Insumo
       frmMarcha_Informacoes_Insumo.Show
    End If
End Sub

Private Sub cmdMarcha_Consulta_Empresa_Click()
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
    
    'Unload frmOperacao_Transporte_Consulta_Empresa
    frmAguarde.Show
    DoEvents
    'frmOperacao_Transporte_Consulta_Empresa.Show
    Unload frmAguarde
End Sub

Private Sub dtcInsumo_LostFocus()
    txtInsumo.Text = dtcInsumo.BoundText
    If IsNumeric(txtInsumo.Text) = False Or dtcInsumo.Text = Empty Then
       txtInsumo.Text = Empty
       txtInsumo.Text = Empty
       txtNome_Cientifico.Text = Empty
       txtConservacao.Text = Empty
       txtFuncao.Text = Empty
       txtObservacao_Insumo.Text = Empty
       Exit Sub
    Else
       If txtInsumo.Text <> Empty And txtInsumo.Text <> " " And dtcInsumo.Text <> Empty Then
          strSql = "SELECT DFNome_cientifico_TBInsumo,DFObservacao_TBInsumo,DFConservacao_TBInsumo," & _
                   "DFDescricao_TBFuncao_insumo " & _
                   "FROM TBInsumo " & _
                   "INNER JOIN TBFuncao_insumo " & _
                   "ON TBInsumo.FKCodigo_TBFuncao_insumo = TBFuncao_insumo.PKCodigo_TBFuncao_insumo " & _
                   "WHERE PKCodigo_TBInsumo = " & txtInsumo.Text & ""
                    
          Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
          
          If rstAplicacao.RecordCount <> 0 Then
             txtNome_Cientifico.Text = rstAplicacao.Fields("DFNome_cientifico_TBInsumo")
             txtConservacao.Text = rstAplicacao.Fields("DFConservacao_TBInsumo")
             txtFuncao.Text = rstAplicacao.Fields("DFDescricao_TBFuncao_insumo")
             txtObservacao_Insumo.Text = rstAplicacao.Fields("DFObservacao_TBInsumo")
          End If
          Set rstAplicacao = Nothing
       End If
    End If
End Sub

Private Sub dtpCompetencia_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpCompetencia_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpData_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpData_Inicio_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpData_Inicio_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpData_Previsao_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpData_Previsao_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpData_Fim_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpFabricacao_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpFabricacao_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpFabricacao_Materia_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpFabricacao_Materia_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpHora_Fim_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpHora_Inicio_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpHora_Inicio_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpHora_Fim_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpValidade_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpValidade_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpValidade_Materia_KeyDown(KeyCode As Integer, Shift As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyCode = 13 Then
        KeyCode = vbKeyTab
    End If
End Sub

Private Sub dtpValidade_Materia_KeyPress(KeyAscii As Integer)
    'Habilita a troca do campo pelo espaço
    If KeyAscii = 32 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
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
    log.Programa = "Cadastro de Marcha Analítica"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio(Me.Caption, cmdConsulta, cmdRefresh, Me.sstMarcha, booPrivilegio_Incluir, booPrivilegio_Alterar, booPrivilegio_Excluir, booPrivilegio_Consultar, MDIPrincipal.OCXUsuario.Codigo, tlbBotoes, Me, "Otica", "BDRetaguarda")
    Else
       booPrivilegio_Incluir = True
       booPrivilegio_Alterar = True
       booPrivilegio_Excluir = True
       booPrivilegio_Consultar = True
    End If
    
    log.Descricao = "Inicializando o cadastro de Marcha Analítica"
    'Gravando o log
    log.Gravar_log "Otica", Me
       
    strUsuario_Digitacao = MDIPrincipal.OCXUsuario.Nome
    
    Call Reposicao
    
    sstMarcha.TabEnabled(0) = False
    sstMarcha.TabEnabled(1) = False
    sstMarcha.TabEnabled(2) = False
    sstMarcha.TabEnabled(3) = False
    sstMarcha.TabEnabled(4) = False
    sstMarcha.Tab = 5
    
    'INTEGRAÇÃO PORTAL E FILIAIS
    booIntegracao = Movimentacoes.Acessibilidade_nivel_usuario(Me, CLng(MDIPrincipal.OCXUsuario.Codigo), "Otica", "BDRetaguarda", CLng(MDIPrincipal.OCXUsuario.Empresa))
    booIntegra_Portal = Funcoes_Gerais.Verifica_integracao_portal(MDIPrincipal.OCXUsuario.Empresa, Me)
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Load")
    Exit Sub
End Sub

Private Function Reposicao()
    
    strTamanho = "0,0,1500,900,3500,900,2500," & _
                 "900,2500,900,2500,1200," & _
                 "1200,1300,1200,900,900," & _
                 "1200,1200," & _
                 "2500,1500,1500,2100,2100," & _
                 "2000,1500"
    
    strNomes = "ID,ID_Cliente,Nº Sequencial,Código,Cliente,Código,Serviço," & _
               "Código,Insumo,Código,Tipo de Marcha,Previsão," & _
               "Prioridade,Competência,Lote Insumo,Peso,Unidade," & _
               "Quantidade,Medida," & _
               "Embalagem,Dt. Fabricação,Dt. Validade,Fabricante,Dt. Fabricação Matéria," & _
               "Dt. Validade Matéria,Lote Matéria"
               
    On Error GoTo Erro

    Movimentacoes.Monta_HFlex_Grid hfgMarcha, strTamanho, strNomes, 26, "Otica", Me

    Call Monta_Data_Combos
    Call Monta_Combos

    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Reposicao")
    Resume Next
End Function

Private Function Novo()
    On Error GoTo Erro
        
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
    
    sstMarcha.TabEnabled(0) = True
    sstMarcha.Tab = 0
    
    dtpData_Inicio.Value = Date
    dtpHora_inicio.Value = Now
    dtpData_Fim.Value = "1/1/1900"
    dtpHora_Fim.Value = "00:00:00"
    
    dtpCompetencia.Value = Format(Date, "MM/YYYY")
    dtpData_Previsao.Value = Date
    dtpValidade.Value = Date
    dtpValidade_Materia.Value = Date
    dtpFabricacao.Value = Date
    dtpFabricacao_Materia.Value = Date
    
    txtUsuario.Text = MDIPrincipal.OCXUsuario.Nome
    
    dtcEmpresa.Width = 5265
    cmdInformacoes_Liberacao.Visible = False
    
    cbbPrioridade.Text = Empty
    cbbEstagio.Text = Empty
    
    sstMarcha.TabEnabled(0) = True
    sstMarcha.TabEnabled(1) = False
    sstMarcha.TabEnabled(2) = False
    sstMarcha.TabEnabled(3) = False
    sstMarcha.TabEnabled(4) = False
    sstMarcha.Tab = 0
    
    'retorno da variavel de controle de grid para false
    booClick_Grid = False
    booAlterar = False
    
    txtTipo_Marcha.SetFocus
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "Otica", "Novo")
    Exit Function
End Function

Private Sub hfgMarcha_Click()
    If hfgMarcha.Col = 0 And hfgMarcha.Text <> Empty Then
        
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
        
        'variavel de controle de click para abastecimento das verificacoes de servico
        booClick_Grid = True
        
        'A competencia é abastecida antes pra efeitos de verificação
        dtpCompetencia.Value = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 14))

        strID_Marcha = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 1))
        strId_Cliente = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 2))
        txtNumero_Sequencial.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 3))
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Abastecendo informacoes do cliente e plano
        txtCliente.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 4))
        Call txtCliente_LostFocus
        'Abastecendo informacoes dos servicos
        txtServico.Enabled = True
        txtServico.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 6))
        Call Verifica_Servico
        'Abastecendo informacoes do insumo
        txtInsumo.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 8))
        Call txtInsumo_LostFocus
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        txtTipo_Marcha.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 10))
        dtpData_Previsao.Value = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 12))
        cbbPrioridade.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 13))
        txtLote_Insumo.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 15))
        txtPeso.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 16))
        txtUnidade.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 17))
        txtQuantidade.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 18))
        txtMedida.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 19))
        txtEmbalagem_Insumo.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 20))
        dtpFabricacao.Value = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 21))
        dtpValidade.Value = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 22))
        txtFabricante.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 23))
        dtpFabricacao_Materia.Value = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 24))
        dtpValidade_Materia.Value = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 25))
        txtLote_Materia.Text = hfgMarcha.TextArray((hfgMarcha.Row * hfgMarcha.Cols + hfgMarcha.Col + 26))


        'Abastecendo o último acompanhamento
        strSql = "SELECT DFEstagio_TBAcompanhamento_marcha," & _
                 "DFData_inicio_TBAcompanhamento_marcha," & _
                 "DFHora_inicio_TBAcompanhamento_marcha," & _
                 "DFData_fim_TBAcompanhamento_marcha," & _
                 "DFHora_fim_TBAcompanhamento_marcha," & _
                 "DFUsuario_DFHora_inicio_TBAcompanhamento_marcha " & _
                 "FROM TBAcompanhamento_marcha " & _
                 "WHERE FKId_TBMarcha = " & strID_Marcha & " " & _
                 "ORDER BY PKId_TBAcompanhamento_marcha,DFData_fim_TBAcompanhamento_marcha "
                 
        Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
        If rstAplicacao.RecordCount <> 0 Then
           rstAplicacao.MoveLast
           If rstAplicacao.Fields("DFEstagio_TBAcompanhamento_marcha") = 1 Then
              cbbEstagio.Text = "Recebimento"
           ElseIf rstAplicacao.Fields("DFEstagio_TBAcompanhamento_marcha") = 2 Then
              cbbEstagio.Text = "Triagem"
           ElseIf rstAplicacao.Fields("DFEstagio_TBAcompanhamento_marcha") = 3 Then
              cbbEstagio.Text = "Amostragem"
           ElseIf rstAplicacao.Fields("DFEstagio_TBAcompanhamento_marcha") = 4 Then
              cbbEstagio.Text = "Laboratório"
           ElseIf rstAplicacao.Fields("DFEstagio_TBAcompanhamento_marcha") = 5 Then
              cbbEstagio.Text = "Micro"
           ElseIf rstAplicacao.Fields("DFEstagio_TBAcompanhamento_marcha") = 6 Then
              cbbEstagio.Text = "Físico-Químico"
           ElseIf rstAplicacao.Fields("DFEstagio_TBAcompanhamento_marcha") = 7 Then
              cbbEstagio.Text = "Digitação"
           End If
           'Variável para controle de modificacao de estágio
           strAcompanhamento = cbbEstagio.Text
        Else
           dtpData_Inicio.Value = "1/1/1900"
           dtpHora_inicio.Value = "00:00:00"
           dtpData_Fim.Value = "1/1/1900"
           dtpHora_Fim.Value = "00:00:00"
           cbbEstagio.Text = Empty
           strAcompanhamento = Empty
        End If
        
        Set rstAplicacao = Nothing

        booAlterar = True
        txtConsulta.Text = Empty
        
        sstMarcha.TabEnabled(0) = True
        sstMarcha.Tab = 0
        'passando o Usuário do momento
        txtUsuario.Text = MDIPrincipal.OCXUsuario.Nome
        'Retorno da variavel de controle de click para abastecimento das verificacoes de servico
        booClick_Grid = False
   End If
   Unload frmAguarde
End Sub

Private Sub hfgMarcha_DblClick()
    hfgMarcha.Sort = 1
End Sub

Private Sub hfgMarcha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
       Call hfgMarcha_Click
    End If
End Sub

Private Sub sstMarcha_Click(PreviousTab As Integer)
   If sstMarcha.Tab = 0 Then
      txtTipo_Marcha.SetFocus
   ElseIf sstMarcha.Tab = 5 Then
      If frmIntegracao.Visible = True Then
         Unload frmIntegracao
      End If
      If strCombo <> Empty And strCombo <> "Todos" Then
         cbbCampos.Text = strCombo
         If txtConsulta.Visible = True Then
            txtConsulta.SetFocus
         ElseIf dtpInicio_Consulta.Visible = True Then
            dtpInicio_Consulta.SetFocus
         ElseIf cbbConsulta.Visible = True Then
            cbbConsulta.SetFocus
         End If
      ElseIf strCombo = "Todos" Then
         hfgMarcha.Row = 1
         hfgMarcha.Col = 0
         hfgMarcha.SetFocus
      End If
   End If
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Novo
           Case 2 And sstMarcha.Tab <> 5: Call Gravar
           Case 3: Call Cancelar
           Case 4 And sstMarcha.Tab <> 5: Call Excluir
           Case 5: Call Imprimir
           Case 7: Unload Me
           Case 9: Call Integracao
    End Select
End Sub

Function Gravar()
    
    On Error GoTo Erro
    
    'Verifica se os campos necessarios para gravar não estão nulos
    If txtTipo_Marcha.Text = Empty Then
       MsgBox "O campo código do Tipo de Marcha não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtTipo_Marcha.SetFocus
       Exit Function
    End If
    If txtCliente.Text = Empty Then
       MsgBox "O campo código do Cliente não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtCliente.SetFocus
       Exit Function
    End If
    If txtServico.Text = Empty Then
       MsgBox "O campo código do Serviço não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       If txtServico.Enabled = True Then
          txtServico.SetFocus
       End If
       Exit Function
    End If
    If txtInsumo.Text = Empty Then
       MsgBox "O campo código do Insumo não pode ser nulo. Verifique!", vbInformation, "Only Tech"
       txtInsumo.SetFocus
       Exit Function
    End If
    
    Dim strSet As String
    Dim strCampo As String
    Dim strValores As String
    Dim strCliente As String
    Dim intPrioridade As Integer
    Dim intEstagio As Integer
    
    'Call Objetos.Maiusculo_TXT(Me)

    strCampo = "FKId_TBCliente,FKCodigo_TBInsumo," & _
               "FKCodigo_TBTipo_marcha,DFPrevisao_entrega_TBMarcha," & _
               "DFPrioridade_TBMarcha,DFCompetencia_TBMarcha," & _
               "DFLote_TBMarcha,DFQuantidade_TBMarcha," & _
               "DFMedida_TBMarcha,DFEmbalagem_TBMarcha," & _
               "DFData_fabricacao_TBMarcha,DFData_validade_TBMarcha," & _
               "DFFabricante_TBMarcha,DFData_fabricacao_material_TBMarcha," & _
               "DFData_validade_material_TBMarcha,DFLote_material_TBMarcha," & _
               "DFNumero_sequencia_TBMarcha,DFPeso_insumo_TBMarcha," & _
               "DFUnidade_insumo_TBMarcha,FKCodigo_TBServico_laboratorio," & _
               "DFData_alteracao_TBMarcha,DFIntegrado_filiais_TBMarcha "
    
    If booIntegra_Portal = True Then
        strCampo = strCampo & ",DFIntegrado_portal_TBMarcha "
    End If
        
    strCliente = Funcoes_Gerais.Localiza_ID("PKId_TBCliente", "IXCodigo_TBCliente", txtCliente.Text, "TBCliente", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcEmpresa.BoundText)
    
    If cbbPrioridade.Text = "Baixa" Then
       intPrioridade = 1
    ElseIf cbbPrioridade.Text = "Média" Then
       intPrioridade = 2
    ElseIf cbbPrioridade.Text = "Alta" Then
       intPrioridade = 3
    End If
    
    strValores = "" & strCliente & "," & txtInsumo.Text & "," & txtTipo_Marcha.Text & "," & _
                 "'" & Format(dtpData_Previsao.Value, "YYYYMMDD") & "'," & _
                 "'" & intPrioridade & "','" & Format(dtpCompetencia.Value, "YYYYMM01") & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtLote_Insumo.Text) & "','" & txtQuantidade.Text & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtMedida.Text) & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtEmbalagem_Insumo.Text) & "','" & Format(dtpFabricacao.Value, "YYYYMMDD") & "'," & _
                 "'" & Format(dtpValidade.Value, "YYYYMMDD") & "','" & Funcoes_Gerais.Grava_String(txtFabricante.Text) & "'," & _
                 "'" & Format(dtpFabricacao_Materia.Value, "YYYYMMDD") & "'," & _
                 "'" & Format(dtpValidade_Materia.Value, "YYYYMMDD") & "'," & _
                 "'" & Funcoes_Gerais.Grava_String(txtLote_Materia.Text) & "','" & Funcoes_Gerais.Grava_String(txtNumero_Sequencial.Text) & "'," & _
                 "" & Funcoes_Gerais.Grava_Moeda(txtPeso.Text) & "," & _
                 "'" & Funcoes_Gerais.Grava_String(txtUnidade.Text) & "'," & txtServico.Text & "," & _
                 "'" & Format(Date, "YYYYMMDD") & "',0 "
    
    If booIntegra_Portal = True Then
        strValores = strValores & ",0 "
    End If
    
    On Error GoTo Erro_transacao
    
    conexao.Initial_Catalog = "BDRetaguarda"
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    If booAlterar = True Then
       
       log.Evento = "Alterar"
       
       strSql = "UPDATE TBMarcha SET FKId_TBCliente = " & strCliente & ",FKCodigo_TBInsumo = " & txtInsumo.Text & "," & _
                "FKCodigo_TBTipo_marcha = " & txtTipo_Marcha.Text & "," & _
                "DFPrevisao_entrega_TBMarcha = '" & Format(dtpData_Previsao.Value, "YYYYMMDD") & "'," & _
                "DFPrioridade_TBMarcha = '" & intPrioridade & "'," & _
                "DFCompetencia_TBMarcha = '" & Format(dtpCompetencia.Value, "YYYYMM01") & "'," & _
                "DFLote_TBMarcha = '" & Funcoes_Gerais.Grava_String(txtLote_Insumo.Text) & "'," & _
                "DFQuantidade_TBMarcha = '" & txtQuantidade.Text & "'," & _
                "DFMedida_TBMarcha = '" & Funcoes_Gerais.Grava_String(txtMedida.Text) & "'," & _
                "DFEmbalagem_TBMarcha = '" & Funcoes_Gerais.Grava_String(txtEmbalagem_Insumo.Text) & "'," & _
                "DFData_fabricacao_TBMarcha = '" & Format(dtpFabricacao.Value, "YYYYMMDD") & "'," & _
                "DFData_validade_TBMarcha = '" & Format(dtpValidade.Value, "YYYYMMDD") & "'," & _
                "DFFabricante_TBMarcha = '" & Funcoes_Gerais.Grava_String(txtFabricante.Text) & "'," & _
                "DFData_fabricacao_material_TBMarcha = '" & Format(dtpFabricacao_Materia.Value, "YYYYMMDD") & "'," & _
                "DFData_validade_material_TBMarcha = '" & Format(dtpValidade_Materia.Value, "YYYYMMDD") & "'," & _
                "DFLote_material_TBMarcha = '" & Funcoes_Gerais.Grava_String(txtLote_Materia.Text) & "'," & _
                "DFNumero_sequencia_TBMarcha = '" & Funcoes_Gerais.Grava_String(txtNumero_Sequencial.Text) & "'," & _
                "DFPeso_insumo_TBMarcha = " & Funcoes_Gerais.Grava_Moeda(txtPeso.Text) & "," & _
                "DFUnidade_insumo_TBMarcha = '" & Funcoes_Gerais.Grava_String(txtUnidade.Text) & "'," & _
                "FKCodigo_TBServico_laboratorio = " & txtServico.Text & "," & _
                "DFData_alteracao_TBMarcha = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBMarcha = 0 "
                
       If booIntegra_Portal = True Then
          strSql = strSql & ",DFIntegrado_portal_TBMarcha = 0 "
       End If

       strSql = strSql & "WHERE PKId_TBMarcha = " & strID_Marcha & ""
               
       conexao.CNConexao.Execute strSql
       
       'Gravando as informações de Acompanhamento
       If strAcompanhamento <> cbbEstagio.Text Then

          If cbbEstagio.Text = "Recebimento" Then
             intEstagio = 1
          ElseIf cbbEstagio.Text = "Triagem" Then
             intEstagio = 2
          ElseIf cbbEstagio.Text = "Amostragem" Then
             intEstagio = 3
          ElseIf cbbEstagio.Text = "Laboratório" Then
             intEstagio = 4
          ElseIf cbbEstagio.Text = "Micro" Then
             intEstagio = 5
          ElseIf cbbEstagio.Text = "Físico-Químico" Then
             intEstagio = 6
          ElseIf cbbEstagio.Text = "Digitação" Then
             intEstagio = 7
          End If
          
          strSql = "INSERT INTO TBAcompanhamento_marcha(FKId_TBMarcha," & _
                   "DFEstagio_TBAcompanhamento_marcha," & _
                   "DFData_inicio_TBAcompanhamento_marcha," & _
                   "DFHora_inicio_TBAcompanhamento_marcha," & _
                   "DFData_fim_TBAcompanhamento_marcha," & _
                   "DFHora_fim_TBAcompanhamento_marcha," & _
                   "DFUsuario_DFHora_inicio_TBAcompanhamento_marcha," & _
                   "DFData_alteracao_TBAcompanhamento_marcha," & _
                   "DFIntegrado_filiais_TBAcompanhamento_marcha"
                   
          If booIntegra_Portal = True Then
             strSql = strSql & ",DFIntegrado_portal_TBAcompanhamento_marcha) "
          Else
             strSql = strSql & ") "
          End If

          strSql = strSql & "VALUES(" & strID_Marcha & "," & _
                   "" & intEstagio & "," & _
                   "'" & Format(dtpData_Inicio.Value, "YYYYMMDD") & "'," & _
                   "'" & Format(dtpHora_inicio.Value, "hh:mm:ss") & "'," & _
                   "'" & Format(dtpData_Fim.Value, "YYYYMMDD") & "'," & _
                   "'" & Format(dtpHora_Fim.Value, "hh:mm:ss") & "'," & _
                   "'" & Funcoes_Gerais.Grava_String(txtUsuario.Text) & "'," & _
                   "'" & Format(Date, "YYYYMMDD") & "',0) "
                   
          If booIntegra_Portal = True Then
             strSql = strSql & ",0) "
          End If
           
          conexao.CNConexao.Execute strSql
       End If
       
       log.Descricao = "Alterando o registro: " + txtNumero_Sequencial.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    Else
       log.Evento = "Incluir Novo"
       
       strSql = "INSERT INTO TBMarcha(" & strCampo & ") VALUES (" & strValores & ")"
       
       conexao.CNConexao.Execute strSql
       
       'Gravando as informações de Acompanhamento
       If cbbEstagio.Text <> Empty Then
          
          If cbbEstagio.Text = "Recebimento" Then
             intEstagio = 1
          ElseIf cbbEstagio.Text = "Triagem" Then
             intEstagio = 2
          ElseIf cbbEstagio.Text = "Amostragem" Then
             intEstagio = 3
          ElseIf cbbEstagio.Text = "Laboratório" Then
             intEstagio = 4
          ElseIf cbbEstagio.Text = "Micro" Then
             intEstagio = 5
          ElseIf cbbEstagio.Text = "Físico-Químico" Then
             intEstagio = 6
          ElseIf cbbEstagio.Text = "Digitação" Then
             intEstagio = 7
          End If
          
          strSql = "INSERT INTO TBAcompanhamento_marcha(FKId_TBMarcha," & _
                   "DFEstagio_TBAcompanhamento_marcha," & _
                   "DFData_inicio_TBAcompanhamento_marcha," & _
                   "DFHora_inicio_TBAcompanhamento_marcha," & _
                   "DFData_fim_TBAcompanhamento_marcha," & _
                   "DFHora_fim_TBAcompanhamento_marcha," & _
                   "DFUsuario_DFHora_inicio_TBAcompanhamento_marcha," & _
                   "DFData_alteracao_TBAcompanhamento_marcha," & _
                   "DFIntegrado_filiais_TBAcompanhamento_marcha"
                   
          If booIntegra_Portal = True Then
             strSql = strSql & ",DFIntegrado_portal_TBAcompanhamento_marcha) "
          Else
             strSql = strSql & ") "
          End If
                   
          strSql = strSql & "SELECT MAX(PKId_TBMarcha)," & _
                   "" & intEstagio & "," & _
                   "'" & Format(dtpData_Inicio.Value, "YYYYMMDD") & "'," & _
                   "'" & Format(dtpHora_inicio.Value, "hh:mm:ss") & "'," & _
                   "'" & Format(dtpData_Fim.Value, "YYYYMMDD") & "'," & _
                   "'" & Format(dtpHora_Fim.Value, "hh:mm:ss") & "'," & _
                   "'" & Funcoes_Gerais.Grava_String(txtUsuario.Text) & "'," & _
                   "'" & Format(Date, "YYYYMMDD") & "',0 "

          If booIntegra_Portal = True Then
             strSql = strSql & ",0 "
          End If
          
          strSql = strSql & "FROM TBMarcha"
                       
          conexao.CNConexao.Execute strSql
       End If
       
       'Somando 1 ao número sequencial do tipo de marcha
       strSql = "UPDATE TBTipo_marcha " & _
                "SET DFNumero_sequencial_TBTipo_marcha = (DFNumero_sequencial_TBTipo_marcha + 1), " & _
                "DFData_alteracao_TBTipo_marcha = '" & Format(Date, "YYYYMMDD") & "'," & _
                "DFIntegrado_filiais_TBTipo_marcha = 0 "
                
       If booIntegra_Portal = True Then
          strSql = strSql & ",DFIntegrado_portal_TBTipo_marcha = 0 "
       End If
       
       strSql = strSql & "WHERE PKCodigo_TBTipo_marcha = " & txtTipo_Marcha.Text & ""
                
       conexao.CNConexao.Execute strSql

       log.Descricao = "Gravando o registro: " + txtNumero_Sequencial.Text
       log.Tipo = 1
       log.Hora = Format(Now, "hh:mm:ss")
       'Gravando log
       log.Gravar_log "OTICA", Me
    End If
    
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
       hfgMarcha.Visible = False
    End If
    
    sstMarcha.TabEnabled(0) = False
    
    sstMarcha.Tab = 5
        
    Exit Function
    
Erro_transacao:
    conexao.CNConexao.RollbackTrans
    conexao.Fechar_conexao
Erro:
    Call Erro.Erro(Me, "OTICA", "Gravar")
    Exit Function
End Function

Private Function Excluir()
    
'    strSql = "SELECT PKCodigo_TBContrato_cliente FROM TBContrato_cliente " & _
'             "WHERE FKCodigo_TBPlano_servico = " & txtCodigo.Text & ""
'
'    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
'
'    If rstAplicacao.RecordCount <> 0 Then
'       MsgBox "Este Plano de Serviço está vinculado ao Contrato de Código " & rstAplicacao.Fields("PKCodigo_TBContrato_cliente") & " e não pode ser excluído. Verifique.", vbInformation, "Only Tech"
'       Set rstAplicacao = Nothing
'       Exit Function
'    End If
'    Set rstAplicacao = Nothing
'
    On Error GoTo Erro
    
    log.Evento = "Excluir"
    log.Descricao = "Exclusão do registro: " + txtNumero_Sequencial.Text
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando log
    log.Gravar_log "Otica", Me
    
    'abrindo conexao
    conexao.Abrir_conexao "Otica"
    conexao.CNConexao.BeginTrans
    
    'Excluindo Registro Secundário
    strSql = "DELETE FROM TBAcompanhamento_marcha WHERE FKId_TBMarcha = '" & strID_Marcha & "'"
    conexao.CNConexao.Execute strSql
    
    'Excluindo Registro Principal
    strSql = "DELETE FROM TBMarcha WHERE PKId_TBMarcha = '" & strID_Marcha & "'"
    
    conexao.CNConexao.Execute strSql
    
    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
       
    Call Objetos.Limpa_TXT(Me)
    cbbPrioridade.Text = Empty
    cbbEstagio.Text = Empty
    
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
       Me.hfgMarcha.Visible = False
    End If
    
    sstMarcha.TabEnabled(0) = False
    sstMarcha.Tab = 5
        
    Exit Function
Erro:
    conexao.CNConexao.RollbackTrans
    'fechando conexao
    conexao.CNConexao.CommitTrans
    conexao.Fechar_conexao
    
    Call Erro.Erro(Me, "Otica", "Excluir")
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
    cbbPrioridade.Text = Empty
    cbbEstagio.Text = Empty
    
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
       hfgMarcha.Visible = False
    End If
        
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de operação com registro"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    sstMarcha.TabEnabled(0) = False
    sstMarcha.TabEnabled(1) = False
    sstMarcha.TabEnabled(2) = False
    sstMarcha.TabEnabled(3) = False
    sstMarcha.TabEnabled(4) = False
    sstMarcha.Tab = 5
    
    Exit Function
Erro:
    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Function
End Function
Public Function Consulta()
    Dim booRetorno As Boolean
    Dim intConsulta_Combos As String

    If cbbCampos.Text = Empty Or cbbCampos.Text <> "Todos" And txtConsulta.Visible = True Then
       If cbbCampos.Text = Empty Or txtConsulta.Text = Empty Then
          MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
          cbbCampos.SetFocus
          Exit Function
       End If
    ElseIf cbbCampos.Text = "Prioridade" Then
       If cbbCampos.Text = Empty Or cbbConsulta.Text = Empty Then
          MsgBox "Selecione uma opção para consulta.", vbCritical, "Only Tech"
          cbbConsulta.SetFocus
          Exit Function
       End If
    End If
           
    If cbbCampos.Text = "Prioridade" Then
       If cbbConsulta.Text = "Baixa" Then
          intConsulta_Combos = "1"
       ElseIf cbbConsulta.Text = "Média" Then
          intConsulta_Combos = "2"
       ElseIf cbbConsulta.Text = "Alta" Then
          intConsulta_Combos = "3"
       End If
    End If
    
    strSql = Empty
    strSql = "SELECT PKId_TBMarcha," & _
             "FKId_TBCliente,DFNumero_sequencia_TBMarcha,IXCodigo_TBCliente," & _
             "DFNome_TBCliente,FKCodigo_TBServico_laboratorio,DFDescricao_TBServico_laboratorio," & _
             "FKCodigo_TBInsumo,DFDescricao_TBInsumo," & _
             "FKCodigo_TBTipo_marcha,DFDescricao_TBTipo_marcha," & _
             "DFPrevisao_entrega_TBMarcha," & _
             "DFPrioridade_TBMarcha,DFCompetencia_TBMarcha ," & _
             "DFLote_TBMarcha,DFPeso_insumo_TBMarcha,DFUnidade_insumo_TBMarcha,DFQuantidade_TBMarcha ," & _
             "DFMedida_TBMarcha,DFEmbalagem_TBMarcha ," & _
             "DFData_fabricacao_TBMarcha," & _
             "DFData_validade_TBMarcha," & _
             "DFFabricante_TBMarcha," & _
             "DFData_fabricacao_material_TBMarcha," & _
             "DFData_validade_material_TBMarcha," & _
             "DFLote_material_TBMarcha " & _
             "FROM  TBMarcha " & _
             "INNER JOIN TBCliente ON TBMarcha.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
             "INNER JOIN TBInsumo ON TBMarcha.FKCodigo_TBInsumo = TBInsumo.PKCodigo_TBInsumo " & _
             "INNER JOIN TBTipo_Marcha ON TBMarcha.FKCodigo_TBTipo_marcha = TBTipo_marcha.PKCodigo_TBTipo_marcha " & _
             "INNER JOIN TBServico_laboratorio ON TBMarcha.FKCodigo_TBServico_laboratorio = TBServico_laboratorio.PKCodigo_TBServico_laboratorio "
    
    'Essas variaveis sao abastecidas com o intuito de fazer a atualizacao do grid posteriormente
    strCombo = cbbCampos.Text
    strConsulta = txtConsulta.Text

    If cbbCampos.Text <> "Todos" Then
       If cbbCampos.Text = "Número Sequencial" Then
          strSql = strSql & " AND DFNumero_sequencia_TBMarcha = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Código Cliente" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " AND convert(nvarchar,IXCodigo_TBCliente) = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Cliente" Then
          strSql = strSql & " AND DFNome_TBCliente = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Código Serviço" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " AND FKCodigo_TBServico_laboratorio = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Serviço" Then
          strSql = strSql & " AND DFDescricao_TBServico_laboratorio LIKE '%" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Código Insumo" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " AND FKCodigo_TBInsumo = '" & intConsulta_Combos & "' "
       ElseIf cbbCampos.Text = "Insumo" Then
          strSql = strSql & " AND convert(nvarchar,DFDescricao_TBInsumo) = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Código Tipo Marcha" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " AND FKCodigo_TBTipo_marcha = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Tipo Marcha" Then
          strSql = strSql & " AND DFDescricao_TBTipo_marcha = '" & txtConsulta.Text & "' "
       ElseIf cbbCampos.Text = "Previsão" Then
          strSql = strSql & " AND DFPrevisao_entrega_TBMarcha >= '" & Format(dtpInicio_Consulta.Value, "YYYYMMDD") & "' " & _
                            " AND DFPrevisao_entrega_TBMarcha <= '" & Format(dtpFim_Consulta.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Prioridade" Then
          strSql = strSql & " AND DFPrioridade_TBMarcha = '" & intConsulta_Combos & "'"
       ElseIf cbbCampos.Text = "Competência" Then
          strSql = strSql & " AND DFCompetencia_TBMarcha >= '" & Format(dtpInicio_Consulta.Value, "YYYYMMDD") & "' " & _
                            " AND DFCompetencia_TBMarcha <= '" & Format(dtpFim_Consulta.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Lote Insumo" Then
          strSql = strSql & " AND convert(nvarchar,DFLote_TBMarcha) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Peso" Then
          strSql = strSql & " AND convert(money,DFPeso_insumo_TBMarcha) = " & Funcoes_Gerais.Grava_Moeda(txtConsulta.Text) & " "
       ElseIf cbbCampos.Text = "Unidade" Then
          strSql = strSql & " AND convert(nvarchar,DFUnidade_insumo_TBMarcha) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Quantidade" Then
          If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
          strSql = strSql & " AND DFQuantidade_TBMarcha = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Medida" Then
          strSql = strSql & " AND convert(nvarchar,DFMedida_TBMarcha) = '" & txtConsulta.Text & "'"
       ElseIf cbbCampos.Text = "Embalagem" Then
          strSql = strSql & " AND convert(nvarchar,DFEmbalagem_TBMarcha) LIKE '%" & txtConsulta.Text & "%'"
       ElseIf cbbCampos.Text = "Data Fabricação" Then
          strSql = strSql & " AND DFData_fabricacao_TBMarcha >= '" & Format(dtpInicio_Consulta.Value, "YYYYMMDD") & "' " & _
                            " AND DFData_fabricacao_TBMarcha <= '" & Format(dtpFim_Consulta.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Data Validade" Then
          strSql = strSql & " AND DFData_validade_TBMarcha >= '" & Format(dtpInicio_Consulta.Value, "YYYYMMDD") & "' " & _
                            " AND DFData_validade_TBMarcha <= '" & Format(dtpFim_Consulta.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Fabricante" Then
          strSql = strSql & " AND convert(nvarchar,DFFabricante_TBMarcha) LIKE '%" & txtConsulta.Text & "%'"
       ElseIf cbbCampos.Text = "Data Fabricação Matéria" Then
          strSql = strSql & " AND DFData_fabricacao_material_TBMarcha >= '" & Format(dtpInicio_Consulta.Value, "YYYYMMDD") & "' " & _
                            " AND DFData_fabricacao_material_TBMarcha <= '" & Format(dtpFim_Consulta.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Data Validade Matéria" Then
          strSql = strSql & " AND DFData_validade_material_TBMarcha >= '" & Format(dtpInicio_Consulta.Value, "YYYYMMDD") & "' " & _
                            " AND DFData_validade_material_TBMarcha <= '" & Format(dtpFim_Consulta.Value, "YYYYMMDD") & "' "
       ElseIf cbbCampos.Text = "Lote Matéria" Then
          strSql = strSql & " AND convert(nvarchar,DFLote_material_TBMarcha) = '" & txtConsulta.Text & "'"
       End If
'       If IsNumeric(strCodigo_Empresa_Consulta) = False Then
'          strSql = strSql & " AND TBMarcha.FKCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' "
'       ElseIf IsNumeric(strCodigo_Empresa_Consulta) = True And CDbl(strCodigo_Empresa_Consulta) <> 0 Then
'          strSql = strSql & " AND TBMarcha.FKCodigo_TBEmpresa = '" & strCodigo_Empresa_Consulta & "' "
'       End If
    Else
'       If IsNumeric(strCodigo_Empresa_Consulta) = False Then
'          strSql = strSql & " AND TBMarcha.FKCodigo_TBEmpresa = '" & MDIPrincipal.OCXUsuario.Empresa & "' "
'       ElseIf IsNumeric(strCodigo_Empresa_Consulta) = True And CDbl(strCodigo_Empresa_Consulta) <> 0 Then
'          strSql = strSql & " AND TBMarcha.FKCodigo_TBEmpresa = '" & strCodigo_Empresa_Consulta & "' "
'       End If
    End If
    
    strSql = strSql & " ORDER BY TBMarcha.PKId_TBMarcha"
    
    frmAguarde.Show
    DoEvents
    
    Movimentacoes.Movimenta_HFlex_Grid strSql, hfgMarcha, strTamanho, strNomes, "BDRetaguarda", "Otica", Me
    
    hfgMarcha.Row = 1
    hfgMarcha.Col = 0
    If hfgMarcha.Text = Empty Then
       hfgMarcha.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgMarcha, strTamanho, strNomes, 26, "Otica", Me
    Else
       intContador = 1
       hfgMarcha.Col = 13
       Do While intContador <= hfgMarcha.Rows - 1
          hfgMarcha.Row = intContador
          If hfgMarcha.Text = "1" Then
             hfgMarcha.Text = "Baixa"
          ElseIf hfgMarcha.Text = "2" Then
             hfgMarcha.Text = "Média"
          ElseIf hfgMarcha.Text = "3" Then
             hfgMarcha.Text = "Alta"
          End If
          intContador = intContador + 1
       Loop
    End If
    hfgMarcha.Col = 0
    hfgMarcha.Row = 1
    
    Unload frmAguarde
    
End Function

Private Function Monta_Combos()
    cbbPrioridade.Clear
    cbbPrioridade.AddItem ("Baixa")
    cbbPrioridade.AddItem ("Média")
    cbbPrioridade.AddItem ("Alta")
    
    cbbConsulta.Clear
    cbbConsulta.AddItem ("Baixa")
    cbbConsulta.AddItem ("Média")
    cbbConsulta.AddItem ("Alta")

    cbbEstagio.Clear
    cbbEstagio.AddItem ("Recebimento")
    cbbEstagio.AddItem ("Triagem")
    cbbEstagio.AddItem ("Amostragem")
    cbbEstagio.AddItem ("Laboratório")
    cbbEstagio.AddItem ("Micro")
    cbbEstagio.AddItem ("Físico-Químico")
    cbbEstagio.AddItem ("Digitação")
    
    cbbCampos.Clear
    cbbCampos.AddItem ("Todos")
    cbbCampos.AddItem ("Número Sequencial")
    cbbCampos.AddItem ("Código Cliente")
    cbbCampos.AddItem ("Cliente")
    cbbCampos.AddItem ("Código Serviço")
    cbbCampos.AddItem ("Serviço")
    cbbCampos.AddItem ("Código Insumo")
    cbbCampos.AddItem ("Insumo")
    cbbCampos.AddItem ("Código Tipo Marcha")
    cbbCampos.AddItem ("Tipo Marcha")
    cbbCampos.AddItem ("Previsão")
    cbbCampos.AddItem ("Prioridade")
    cbbCampos.AddItem ("Competência")
    cbbCampos.AddItem ("Lote Insumo")
    cbbCampos.AddItem ("Peso")
    cbbCampos.AddItem ("Unidade")
    cbbCampos.AddItem ("Quantidade")
    cbbCampos.AddItem ("Medida")
    cbbCampos.AddItem ("Embalagem")
    cbbCampos.AddItem ("Data Fabricação")
    cbbCampos.AddItem ("Data Validade")
    cbbCampos.AddItem ("Fabricante")
    cbbCampos.AddItem ("Data Fabricação Matéria")
    cbbCampos.AddItem ("Data Validade Matéria")
    cbbCampos.AddItem ("Lote Matéria")
    
End Function

Private Function Monta_Data_Combos()
    strSql = "SELECT PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
    
    strSql = "SELECT PKCodigo_TBTipo_marcha,DFDescricao_TBTipo_marcha FROM TBTipo_marcha"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBTipo_marcha", "DFDescricao_TBTipo_marcha", dtcTipo_Marcha, strSql, "BDRetaguarda", "Otica", Me
    
    strSql = "SELECT PKCodigo_TBInsumo,DFDescricao_TBInsumo FROM TBInsumo"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBInsumo", "DFDescricao_TBInsumo", dtcInsumo, strSql, "BDRetaguarda", "Otica", Me

    strSql = "SELECT IXCodigo_TBCliente,DFNome_TBCliente FROM TBCliente " & _
             "INNER JOIN TBContrato_cliente " & _
             "ON TBCliente.PKId_TBCliente = TBContrato_cliente.FKId_TBCliente " & _
             "WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "'"
             
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBCliente", "DFNome_TBCliente", dtcCliente, strSql, "BDRetaguarda", "OTICA", Me
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Erro
    
    log.Evento = "Unload"
    
    log.Hora = Format(Now, "hh:mm:ss")

    strCodigo_Empresa_Consulta = Empty
    strCombo = Empty
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    
    If frmIntegracao.Visible = True Then
       Unload frmIntegracao
    End If
       
    Exit Sub
Erro:
    Call Erro.Erro(Me, "Otica", "Unload")
    Exit Sub
End Sub

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtEmbalagem_Insumo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFabricante_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtFornecedor_Materia_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtInsumo_Change()
    dtcInsumo.BoundText = txtInsumo.Text
    If IsNumeric(txtInsumo.Text) = False Then txtInsumo.Text = Empty: Exit Sub
End Sub

Private Sub txtInsumo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtInsumo_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub dtcInsumo_GotFocus()
    If Me.txtInsumo.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcInsumo.Text)
    End If
End Sub

Private Sub txtInsumo_LostFocus()
    If IsNumeric(txtInsumo.Text) = False Or dtcInsumo.Text = Empty Then
       txtInsumo.Text = Empty
       txtNome_Cientifico.Text = Empty
       txtConservacao.Text = Empty
       txtFuncao.Text = Empty
       txtObservacao_Insumo.Text = Empty
       Exit Sub
    Else
       If txtInsumo.Text <> Empty And txtInsumo.Text <> " " And dtcInsumo.Text <> Empty Then
          strSql = "SELECT DFNome_cientifico_TBInsumo,DFObservacao_TBInsumo,DFConservacao_TBInsumo," & _
                   "DFDescricao_TBFuncao_insumo " & _
                   "FROM TBInsumo " & _
                   "INNER JOIN TBFuncao_insumo " & _
                   "ON TBInsumo.FKCodigo_TBFuncao_insumo = TBFuncao_insumo.PKCodigo_TBFuncao_insumo " & _
                   "WHERE PKCodigo_TBInsumo = " & txtInsumo.Text & ""
                    
          Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
          
          If rstAplicacao.RecordCount <> 0 Then
             txtNome_Cientifico.Text = rstAplicacao.Fields("DFNome_cientifico_TBInsumo")
             txtConservacao.Text = rstAplicacao.Fields("DFConservacao_TBInsumo")
             txtFuncao.Text = rstAplicacao.Fields("DFDescricao_TBFuncao_insumo")
             txtObservacao_Insumo.Text = rstAplicacao.Fields("DFObservacao_TBInsumo")
          End If
          Set rstAplicacao = Nothing
       End If
    End If
End Sub

Private Sub txtLote_Insumo_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtLote_Materia_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtMedida_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtObservacao_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPeso_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtPeso_LostFocus()
    txtPeso.Text = Format(txtPeso.Text, "#,###0.000")
End Sub

Private Sub txtQuantidade_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub txtServico_GotFocus()
    If txtCliente.Text = Empty Or txtPlano.Text = Empty Then txtCliente.SetFocus: Exit Sub
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub
Private Sub txtServico_Change()
    dtcServico.BoundText = txtServico.Text
    If IsNumeric(txtServico.Text) = False Then txtServico.Text = Empty: Exit Sub
End Sub

Private Sub txtServico_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub dtcServico_GotFocus()
    If dtcCliente.Text = Empty Or txtPlano.Text = Empty Then txtCliente.SetFocus: Exit Sub
    If Me.txtServico.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcServico.Text)
    End If
End Sub

Private Sub dtcServico_LostFocus()
    txtServico.Text = dtcServico.BoundText
    If IsNumeric(txtServico.Text) = False Or dtcServico.Text = Empty Then
       txtServico.Text = Empty
       txtControle.Text = Empty
       txtNumero_Servicos.Text = Empty
       txtServicos_restantes.Text = Empty
       Exit Sub
    Else
       If txtServico.Text <> Empty And dtcServico.Text <> Empty And txtPlano.Text <> Empty Then
          Call Verifica_Servico
       End If
    End If
End Sub

Private Sub txtServico_LostFocus()
    'Esta modificacao foi inserida para manter os dados após a montagem do combo
    dtcServico.BoundText = txtServico.Text
    
    If IsNumeric(txtServico.Text) = False Or dtcServico.Text = Empty Then
       txtServico.Text = Empty
       txtControle.Text = Empty
       txtNumero_Servicos.Text = Empty
       txtServicos_restantes.Text = Empty
       Exit Sub
    Else
       If txtServico.Text <> Empty And dtcServico.Text <> Empty And txtPlano.Text <> Empty Then
          Call Verifica_Servico
       End If
    End If
End Sub

Private Sub txtTipo_Marcha_Change()
    dtcTipo_Marcha.BoundText = txtTipo_Marcha.Text
    If IsNumeric(txtTipo_Marcha.Text) = False Then txtTipo_Marcha.Text = Empty: Exit Sub
End Sub

Private Sub txtTipo_Marcha_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtTipo_Marcha_KeyPress(KeyAscii As Integer)
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And KeyAscii <> 8 Then
          KeyAscii = 0
    End If
End Sub

Private Sub dtcTipo_Marcha_GotFocus()
    If Me.txtTipo_Marcha.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcTipo_Marcha.Text)
    End If
End Sub

Private Sub dtcTipo_Marcha_LostFocus()
    txtTipo_Marcha.Text = dtcTipo_Marcha.BoundText
    If IsNumeric(txtTipo_Marcha.Text) = False Or dtcTipo_Marcha.Text = Empty Then
       txtTipo_Marcha.Text = Empty
       Exit Sub
    Else
       If txtTipo_Marcha.Text <> Empty And txtTipo_Marcha.Text <> " " And dtcTipo_Marcha.Text <> Empty Then
          strSql = "SELECT DFDescricao_resumida_TBTipo_marcha,DFNumero_sequencial_TBTipo_marcha " & _
                   "FROM TBTipo_marcha WHERE PKCodigo_TBTipo_marcha = " & txtTipo_Marcha.Text & ""

          Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me

          If rstAplicacao.RecordCount <> 0 Then
             txtNumero_Sequencial = "" & rstAplicacao.Fields("DFDescricao_resumida_TBTipo_marcha") & "" & rstAplicacao.Fields("DFNumero_sequencial_TBTipo_marcha") & "/" & Format(Date, "YY") & ""
          End If
          Set rstAplicacao = Nothing
       End If
    End If
End Sub

Private Sub txtCliente_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

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
    If IsNumeric(txtCliente.Text) = False Or dtcCliente.Text = Empty Then
       txtCliente.Text = Empty
       txtPlano = Empty
       txtDescricao_Plano = Empty
       txtContrato.Text = Empty
       txtControle.Text = Empty
       txtNumero_Servicos.Text = Empty
       txtServicos_restantes.Text = Empty
       txtServico.Enabled = False
       dtcServico.Enabled = False
       Exit Sub
    Else
       If txtCliente.Text <> Empty And txtCliente.Text <> " " And dtcCliente.Text <> Empty Then
          strSql = "SELECT PKCodigo_TBContrato_cliente," & _
                   "FKCodigo_TBPlano_servico," & _
                   "DFDescricao_TBPlano_servico " & _
                   "FROM TBContrato_cliente " & _
                   "INNER JOIN TBCliente " & _
                   "ON TBContrato_cliente.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
                   "INNER JOIN TBPlano_servico " & _
                   "ON TBContrato_cliente.FKCodigo_TBPlano_servico = TBPlano_servico.PKCodigo_TBPlano_servico " & _
                   "WHERE IXCodigo_TBCliente = " & txtCliente.Text & ""
                    
          Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
          
          If rstAplicacao.RecordCount <> 0 Then
             txtContrato = rstAplicacao.Fields("PKCodigo_TBContrato_cliente")
             txtPlano = rstAplicacao.Fields("FKCodigo_TBPlano_servico")
             txtDescricao_Plano = rstAplicacao.Fields("DFDescricao_TBPlano_servico")
             
             txtServico.Enabled = True
             dtcServico.Enabled = True
             
             'montando o data combo de servicos com apenas os pertencentes ao plano
             strSql = "SELECT PKCodigo_TBServico_laboratorio,DFDescricao_TBServico_laboratorio " & _
                      "FROM TBServico_laboratorio " & _
                      "INNER JOIN TBPlano_servico_servico_laboratorio " & _
                      "ON TBServico_laboratorio.PKCodigo_TBServico_laboratorio = TBPlano_servico_servico_laboratorio.FKCodigo_TBServico_laboratorio " & _
                      "WHERE FKCodigo_TBPlano_servico = " & txtPlano.Text & ""
             
             Movimenta_DataCombo "PKCodigo_TBServico_laboratorio", "DFDescricao_TBServico_laboratorio", dtcServico, strSql, "BDRetaguarda", "Otica", Me

          End If
          Set rstAplicacao = Nothing
       End If
    End If
End Sub
Private Sub dtcCliente_GotFocus()
    If Me.txtCliente.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcCliente.Text)
    End If
End Sub

Private Sub dtcCliente_LostFocus()

    txtCliente.Text = dtcCliente.BoundText
    If IsNumeric(txtCliente.Text) = False Or dtcCliente.Text = Empty Then
       txtCliente.Text = Empty
       txtPlano = Empty
       txtDescricao_Plano = Empty
       txtContrato.Text = Empty
       txtControle.Text = Empty
       txtNumero_Servicos.Text = Empty
       txtServicos_restantes.Text = Empty
       txtServico.Enabled = False
       dtcServico.Enabled = False
       Exit Sub
    Else
       If txtCliente.Text <> Empty And dtcCliente.Text <> Empty Then
          strSql = "SELECT PKCodigo_TBContrato_cliente," & _
                   "FKCodigo_TBPlano_servico," & _
                   "DFDescricao_TBPlano_servico " & _
                   "FROM TBContrato_cliente " & _
                   "INNER JOIN TBCliente " & _
                   "ON TBContrato_cliente.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
                   "INNER JOIN TBPlano_servico " & _
                   "ON TBContrato_cliente.FKCodigo_TBPlano_servico = TBPlano_servico.PKCodigo_TBPlano_servico " & _
                   "WHERE IXCodigo_TBCliente = " & txtCliente.Text & ""
                    
          Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
          
          If rstAplicacao.RecordCount <> 0 Then
             txtContrato.Text = rstAplicacao.Fields("PKCodigo_TBContrato_cliente")
             txtPlano.Text = rstAplicacao.Fields("FKCodigo_TBPlano_servico")
             txtDescricao_Plano.Text = rstAplicacao.Fields("DFDescricao_TBPlano_servico")
             
             txtServico.Enabled = True
             dtcServico.Enabled = True
             
             'Montando o data combo de servicos com apenas os pertencentes ao plano
             strSql = "SELECT PKCodigo_TBServico_laboratorio,DFDescricao_TBServico_laboratorio " & _
                      "FROM TBServico_laboratorio " & _
                      "INNER JOIN TBPlano_servico_servico_laboratorio " & _
                      "ON TBServico_laboratorio.PKCodigo_TBServico_laboratorio = TBPlano_servico_servico_laboratorio.FKCodigo_TBServico_laboratorio " & _
                      "WHERE FKCodigo_TBPlano_servico = " & txtPlano.Text & ""
             
             Movimenta_DataCombo "PKCodigo_TBServico_laboratorio", "DFDescricao_TBServico_laboratorio", dtcServico, strSql, "BDRetaguarda", "Otica", Me
             
             txtServico.SetFocus
          End If
          Set rstAplicacao = Nothing
       End If
    End If
    
End Sub

Private Sub txtTipo_Marcha_LostFocus()
    If txtTipo_Marcha.Text <> Empty And txtTipo_Marcha.Text <> " " And dtcTipo_Marcha.Text <> Empty And booAlterar = False Then
       strSql = "SELECT DFDescricao_resumida_TBTipo_marcha,DFNumero_sequencial_TBTipo_marcha " & _
                "FROM TBTipo_marcha WHERE PKCodigo_TBTipo_marcha = " & txtTipo_Marcha.Text & ""
       
       Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
       
       If rstAplicacao.RecordCount <> 0 Then
          txtNumero_Sequencial = "" & rstAplicacao.Fields("DFDescricao_resumida_TBTipo_marcha") & "" & rstAplicacao.Fields("DFNumero_sequencial_TBTipo_marcha") & "/" & Format(Date, "YY") & ""
       End If
       Set rstAplicacao = Nothing
    End If
End Sub

Private Sub txtUnidade_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Verifica_Servico()
    Dim intTabela As Integer
    Dim intResp As Integer
    Dim Data_inicio As Date

    If txtServico.Text = Empty Or txtPlano.Text = Empty Or txtCliente.Text = Empty Then Exit Function
    strSql = Empty
    strSql = "SELECT DFQuantidade_TBPlano_servico_servico_laboratorio as Limite," & _
             "DFControle_TBPlano_servico_servico_laboratorio," & _
             "DFPeriodo_TBPlano_servico_servico_laboratorio as Periodo,DFTabela_preco_TBContrato_cliente " & _
             "FROM TBPlano_servico_servico_laboratorio " & _
             "INNER JOIN TBContrato_Cliente " & _
             "ON TBPlano_servico_servico_laboratorio.FKCodigo_TBPlano_servico = TBContrato_Cliente.FKCodigo_TBPlano_servico " & _
             "INNER JOIN TBCliente " & _
             "ON TBContrato_Cliente.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
             "WHERE TBPlano_servico_servico_laboratorio.FKCodigo_TBServico_laboratorio = " & txtServico.Text & " " & _
             "AND TBPlano_servico_servico_laboratorio.FKCodigo_TBPlano_servico = " & txtPlano.Text & " " & _
             "AND IXCodigo_TBCliente = " & txtCliente.Text & ""
             
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me

    If rstAplicacao.RecordCount <> 0 Then
 
       If rstAplicacao.Fields("DFControle_TBPlano_servico_servico_laboratorio") = 1 Then
          
          'Verifica se todos os servicos usados no periodo mais o preco do servico escolhido sao maiores que o contrato
          intTabela = rstAplicacao.Fields("DFTabela_preco_TBContrato_cliente")
          
          strSql = "SELECT DFPreco" & intTabela & "_TBServico_laboratorio as Preco_servico," & _
                   "SUM(DFPreco" & intTabela & "_TBServico_laboratorio) as Soma_servico," & _
                   "TBContrato_Cliente.DFValor_TBContrato_cliente as Valor_Contrato " & _
                   "FROM TBPlano_servico_servico_laboratorio " & _
                   "INNER JOIN TBServico_laboratorio " & _
                   "ON TBPlano_servico_servico_laboratorio.FKCodigo_TBServico_laboratorio = TBServico_laboratorio.PKCodigo_TBServico_laboratorio " & _
                   "LEFT JOIN TBMarcha " & _
                   "ON TBServico_laboratorio.PKCodigo_TBServico_laboratorio = TBMarcha.FKCodigo_TBServico_laboratorio " & _
                   "INNER JOIN TBCliente " & _
                   "ON TBMarcha.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
                   "INNER JOIN TBContrato_Cliente " & _
                   "ON TBCliente.PKId_TBCliente = TBContrato_Cliente.FKId_TBCliente " & _
                   "WHERE TBPlano_servico_servico_laboratorio.FKCodigo_TBServico_laboratorio = " & txtServico.Text & " " & _
                   "AND TBPlano_servico_servico_laboratorio.FKCodigo_TBplano_Servico = " & txtPlano.Text & " " & _
                   "AND TBCliente.IXCodigo_TBCliente = " & txtCliente.Text & " " & _
                   "AND YEAR(TBMarcha.DFCompetencia_TBMarcha) = '" & Format(dtpCompetencia.Value, "YYYY") & "' " & _
                   "AND MONTH(TBMarcha.DFCompetencia_TBMarcha) = '" & Format(dtpCompetencia.Value, "MM") & "' " & _
                   "GROUP BY DFPreco" & intTabela & "_TBServico_laboratorio,TBContrato_cliente.DFValor_TBContrato_cliente "
             
          Select_geral strSql, "BDRetaguarda", rstServico, "Otica", Me
          
          'Rotina para verificação da disponibilidade de serviços no período
          If rstServico.RecordCount <> 0 And booClick_Grid = False Then
             If CDbl(rstServico.Fields("Preco_servico")) + CDbl(rstServico.Fields("Soma_servico")) > CDbl(rstServico.Fields("Valor_Contrato")) Then
                intResp = MsgBox("O Valor deste serviço está acima do limite de contrato deste Cliente. Os dados de liberação serão salvos. Deseja prosseguir?", vbYesNo, "Only Tech")
                If intResp = 7 Then
                   txtServico.Text = Empty
                   Set rstAplicacao = Nothing
                   Set rstServico = Nothing
                   Exit Function
                End If
             End If
          End If
          
          txtControle.Text = "Valor Contrato"
          txtNumero_Servicos.Text = "Variável"
          
          If rstServico.RecordCount <> 0 Then
             txtServicos_restantes.Text = CDbl(rstServico.Fields("Valor_Contrato") - rstServico.Fields("Soma_servico")) / CDbl(rstServico.Fields("Preco_servico"))
          Else
             txtServicos_restantes.Text = "Variável"
          End If
          
       ElseIf rstAplicacao.Fields("DFControle_TBPlano_servico_servico_laboratorio") = 2 Then
          
          Data_inicio = DateAdd("m", -rstAplicacao.Fields("Periodo") + 1, dtpCompetencia.Value)
          
          strSql = "SELECT PKId_TBMarcha " & _
                   "FROM TBMarcha " & _
                   "INNER JOIN TBCliente " & _
                   "ON TBMarcha.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
                   "WHERE TBMarcha.FKCodigo_TBServico_laboratorio = " & txtServico.Text & " " & _
                   "AND TBCliente.IXCodigo_TBCliente = " & txtCliente.Text & " " & _
                   "AND TBMarcha.DFCompetencia_TBMarcha >= '" & Format(Data_inicio, "YYYYMM01") & "' " & _
                   "AND TBMarcha.DFCompetencia_TBMarcha <= '" & Format(dtpCompetencia.Value, "YYYYMM01") & "' "

          Select_geral strSql, "BDRetaguarda", rstServico, "Otica", Me
          
          'Rotina para verificação da disponibilidade de serviços no período
          If rstServico.RecordCount <> 0 And booClick_Grid = False Then
             If CDbl(rstServico.RecordCount) >= CDbl(rstAplicacao.Fields("Limite")) Then
                intResp = MsgBox("Este serviço já ultrapassou o limite previsto em contrato para este Cliente. Os dados de liberação serão salvos. Deseja prosseguir?", vbYesNo, "Only Tech")
                If intResp = 7 Then
                   txtServico.Text = Empty
                   Set rstAplicacao = Nothing
                   Set rstServico = Nothing
                   Exit Function
                End If
             End If
          End If
          
          txtControle.Text = "Serviços"
          txtNumero_Servicos.Text = rstAplicacao.Fields("Limite")
          txtServicos_restantes.Text = rstAplicacao.Fields("Limite") - CDbl(rstServico.RecordCount)
          
       ElseIf rstAplicacao.Fields("DFControle_TBPlano_servico_servico_laboratorio") = 3 Then

          Data_inicio = DateAdd("m", -rstAplicacao.Fields("Periodo") + 1, dtpCompetencia.Value)
          
          strSql = "SELECT PKId_TBMarcha " & _
                   "FROM TBMarcha " & _
                   "INNER JOIN TBCliente " & _
                   "ON TBMarcha.FKId_TBCliente = TBCliente.PKId_TBCliente " & _
                   "WHERE TBMarcha.FKCodigo_TBServico_laboratorio " & _
                   "IN (SELECT FKCodigo_TBServico_laboratorio FROM TBPlano_servico_servico_laboratorio " & _
                   "WHERE DFQuantidade_TBPlano_servico_servico_laboratorio = '" & rstAplicacao.Fields("Limite") & "' " & _
                   "AND DFControle_TBPlano_servico_servico_laboratorio = '3' " & _
                   "AND DFPeriodo_TBPlano_servico_servico_laboratorio = '" & rstAplicacao.Fields("Periodo") & "' " & _
                   "AND FKCodigo_TBPlano_Servico = " & txtPlano.Text & ") " & _
                   "AND TBCliente.IXCodigo_TBCliente = " & txtCliente.Text & " " & _
                   "AND TBMarcha.DFCompetencia_TBMarcha >= '" & Format(Data_inicio, "YYYYMM01") & "' " & _
                   "AND TBMarcha.DFCompetencia_TBMarcha <= '" & Format(dtpCompetencia.Value, "YYYYMM01") & "' "

          Select_geral strSql, "BDRetaguarda", rstServico, "Otica", Me
          
          'Rotina para verificação da disponibilidade de serviços no período
          If rstServico.RecordCount <> 0 And booClick_Grid = False Then
             If CDbl(rstServico.RecordCount) >= CDbl(rstAplicacao.Fields("Limite")) Then
                intResp = MsgBox("Este serviço já ultrapassou o limite do grupo previsto em contrato para este Cliente. Os dados de liberação serão salvos. Deseja prosseguir?", vbYesNo, "Only Tech")
                If intResp = 7 Then
                   txtServico.Text = Empty
                   Set rstAplicacao = Nothing
                   Set rstServico = Nothing
                   Exit Function
                End If
             End If
          End If
          txtControle.Text = "Grupo Serviços"
          txtNumero_Servicos.Text = rstAplicacao.Fields("Limite")
          txtServicos_restantes.Text = CDbl(rstAplicacao.Fields("Limite")) - CDbl(rstServico.RecordCount)
       End If
    Else
       txtControle.Text = Empty
       txtNumero_Servicos.Text = Empty
       txtServicos_restantes.Text = Empty
    End If

    Set rstAplicacao = Nothing
    Set rstServico = Nothing
    
End Function

Private Function Integracao()

    Call frmIntegracao.Verifica_Integracao("PKId_TBMarcha", strID_Marcha, "DFIntegrado_filiais_TBMarcha", "TBMarcha", "Otica", "BDRetaguarda", "DFIntegrado_portal_TBMarcha", Me.Top, Me.Left, Me.Width, Me.Height, "Marcha Analítica")
    
End Function
