VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnalise_checkouts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ánalise de acompanhamento de check outs"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnalise_checkouts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   13020
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1905
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   12855
      Begin VB.CommandButton Command1 
         Caption         =   "Auditoria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5760
         Picture         =   "frmAnalise_checkouts.frx":1782
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1140
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   4560
         Picture         =   "frmAnalise_checkouts.frx":1B0C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1140
         Width           =   1065
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3420
         Picture         =   "frmAnalise_checkouts.frx":1E96
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1140
         Width           =   1065
      End
      Begin AutoCompletar.CbCompleta cbbVisoes 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   570
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   345
         Left            =   4200
         TabIndex        =   2
         Top             =   570
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
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
         Format          =   20316161
         CurrentDate     =   37923
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   345
         Left            =   2340
         TabIndex        =   3
         Top             =   570
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   609
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
         Format          =   20316161
         CurrentDate     =   37923
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Período"
         Height          =   240
         Left            =   2370
         TabIndex        =   7
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Visões"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "até"
         Height          =   240
         Left            =   3855
         TabIndex        =   4
         Top             =   675
         Width           =   270
      End
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   6540
      Left            =   90
      TabIndex        =   0
      Top             =   2490
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   11536
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageListGeral 
      Left            =   11250
      Top             =   7860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":2220
            Key             =   "ico_produto"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":253C
            Key             =   "icoPDV"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":2856
            Key             =   "ico_Pedidos"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":2B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":3BC2
            Key             =   "ico_cliente"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":449C
            Key             =   "ico_transportador"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":47B6
            Key             =   "ico_Vendedor"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":7148
            Key             =   "ico_total"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":74E2
            Key             =   "ico_preco"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":77FC
            Key             =   "ico_preco2"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":7B16
            Key             =   "ico_quantidade"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":9820
            Key             =   "ico_perc"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":9BBA
            Key             =   "ico_Vendedor22"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":A00E
            Key             =   "icoOnline"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnalise_checkouts.frx":A328
            Key             =   "IcoOff"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgCupom 
      Height          =   3315
      Left            =   5100
      TabIndex        =   10
      Top             =   2460
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   5847
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3315
      Left            =   9480
      TabIndex        =   13
      Top             =   2460
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   5847
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   2835
      Left            =   5100
      TabIndex        =   16
      Top             =   6180
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5001
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Operação de Caixa"
      Height          =   375
      Left            =   5130
      TabIndex        =   17
      Top             =   5820
      Width           =   7815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Itens Cupom"
      Height          =   375
      Left            =   8940
      TabIndex        =   15
      Top             =   2130
      Width           =   4005
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Cupom"
      Height          =   345
      Left            =   5100
      TabIndex        =   14
      Top             =   2130
      Width           =   3795
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Informações na base do concentrador"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2130
      Width           =   4935
   End
End
Attribute VB_Name = "frmAnalise_checkouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Dim Rede As New DLLSystemManager.Rede
Dim booOnline As Boolean

Private Sub cmdConsultar_Click()

    If Me.cbbVisoes.Text <> "" Then
       Call Abastece_treeview
    End If
    
End Sub

Private Sub Form_Load()

    Me.cbbVisoes.AddItem ("Cargas")
    Me.cbbVisoes.AddItem ("Fechamento")
    Me.cbbVisoes.AddItem ("Log de eventos")
    Me.cbbVisoes.AddItem ("T.E.F")
    
End Sub

Private Sub tvTreeView_DragDrop(Source As Control, X As Single, Y As Single)
    If Not (tvTreeView.DropHighlight Is Nothing) Then
        Set SourceNode.Parent = tvTreeView.DropHighlight
        'SourceNode.Key
        Set SourceNode = tvTreeView.DropHighlight
        'tvTreeView.Nodes.Item
        Set tvTreeView.DropHighlight = Nothing
    End If
    Set SourceNode = Nothing
    SourceType = tp_TNÓ_Nenhum
End Sub

Private Sub tvTreeView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set SourceNode = tvTreeView.HitTest(X, Y)
End Sub

Private Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    frmAguarde.Show
    
    If Me.cbbVisoes.Text = "Fechamento" Then
       If "PDV" = Mid(Node.Key, 1, 3) Then
            strSql = Empty
            strSql = "SELECT DFData_Saida_TBCupom,DFNumero_TBCupom, DFSerie_TBCupom,DFTotal_cupom_TBCupom,DFCodigo_cupom_impressora_TBCupom FROM TBCUPOM " & _
                     "WHERE DFData_Saida_TBCupom BETWEEN '" & Format(Me.dtpInicial.Value, "YYYYMMDD") & "' AND '" & Format(Me.dtpFinal.Value, "YYYYMMDD") & "' " & _
                     "AND PKCodigo_TBPdv = " & InStr(1, Node.Text, "-") & ""
            Movimentacoes.Movimenta_HFlex_Grid strSql, Me.hfgCupom, "800,800,500,800,800,400", "Data,N°Cupom,Serie,Total Cupom,COO", "BDRetaguarda", "Otica", Me
       End If
    End If
    
    Unload frmAguarde
    
End Sub
Private Function Abastece_treeview()
    
    Dim rstPDV As New ADODB.Recordset
    Dim rstCupom As New ADODB.Recordset
    Dim rstCupom_Individual As New ADODB.Recordset
    
    tvTreeView.Nodes.Clear
    cont = 0
    cont2 = 0
    cont3 = 0
    
    frmAguarde.Show
    
'''    On Error GoTo Erro
    
    'Query para identificar todos os pedidos de todos os vendedores
    strSql = "SELECT * FROM TBPDV"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstPDV, "Otica", Me
    
    strSql = Empty
    strSql = "SELECT PKCodigo_TBPdv, SUM(DFTotal_cupom_TBCupom) AS TOTAL FROM TBCUPOM " & _
             "WHERE DFData_Saida_TBCupom BETWEEN '" & Format(Me.dtpInicial.Value, "YYYYMMDD") & "' AND '" & Format(Me.dtpFinal.Value, "YYYYMMDD") & "' " & _
             "GROUP BY PKCodigo_TBPdv"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCupom, "Otica", Me
    
    tvTreeView.ImageList = ImageListGeral
    
    rstPDV.MoveLast: rstPDV.MoveFirst
    
    Do While rstPDV.EOF = False And rstPDV.BOF = False
    
       Dim strKey As String
       
       strKey = "PDV" & rstPDV!PKCodigo_TBPdv
       
       Set nodX = tvTreeView.Nodes.Add(, , strKey, rstPDV!PKCodigo_TBPdv & " - (" & rstPDV!DFEndereco_ip_TBPdv & ")", "icoPDV", "icoPDV")
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'Montando o nó da TBOperação_Caixa
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       'Acessando o PDV pelo IP
       Call PING(rstPDV!DFEndereco_ip_TBPdv)
       
       If booOnline = True Then
          strIco = "icoOnline"
       Else
          strIco = "IcoOff"
       End If
       
       strKey2 = "TOT" & rstCupom.AbsolutePosition
       
       rstCupom.MoveFirst
       rstCupom.Find ("PKCodigo_TBPdv = " & rstPDV!PKCodigo_TBPdv & "")
       
       If rstCupom.EOF = True Then
          GoTo proximo:
       End If
       
       Set nodX = tvTreeView.Nodes.Add(strKey, tvwChild, strKey2, "Total vendido no check out:  " & Format(rstCupom!TOTAL, "#,###0.00"), strIco, strIco)
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
proximo:
       rstPDV.MoveNext
       
    Loop
    
    Set rstPDV = Nothing
    Set rstCupom = Nothing
    
    Unload frmAguarde
    
    Exit Function
    
Erro:
    Call Erro.Erro(Me, "Otica", "Visão Vendedor")
    Exit Function

End Function

Private Function PING(Ip As String) As String

   Rede.PING (Ip)
   
   If Rede.ECHO_Status = 0 Then
      PING = True
      booOnline = True
   Else
      PING = False
      booOnline = False
   End If
      
End Function

