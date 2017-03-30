VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMovimentacoes_exportacao_balancas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportação para Balanças"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimentacoes_exportacao_balancas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5505
   Begin TabDlg.SSTab sstTipo_Tabela 
      Height          =   5625
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   9922
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
      TabCaption(0)   =   "&Arquivos"
      TabPicture(0)   =   "frmMovimentacoes_exportacao_balancas.frx":1782
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdConsulta_Empresa_Rx"
      Tab(0).Control(1)=   "txtCaminho"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(4)=   "cmdExportar"
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(7)=   "dtcEmpresa"
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(9)=   "Label18"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "&Críticas dos arquivos gerados"
      TabPicture(1)   =   "frmMovimentacoes_exportacao_balancas.frx":179E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
      Begin VB.CommandButton cmdConsulta_Empresa_Rx 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70125
         Picture         =   "frmMovimentacoes_exportacao_balancas.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2160
         Width           =   435
      End
      Begin VB.TextBox txtCaminho 
         Height          =   375
         Left            =   -74880
         MaxLength       =   100
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2130
         Width           =   5220
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   -74880
         TabIndex        =   22
         ToolTipText     =   "Código"
         Top             =   780
         Width           =   825
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Enabled         =   0   'False
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
         Left            =   -70620
         TabIndex        =   21
         Top             =   780
         Width           =   1035
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   "Remover"
         Enabled         =   0   'False
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
         Left            =   -69540
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   780
         Width           =   1035
      End
      Begin VB.Frame Frame2 
         Caption         =   "Seções da Tabela"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   -74880
         TabIndex        =   16
         Top             =   1200
         Width           =   6405
         Begin VB.TextBox txtCodigo_Secao 
            Enabled         =   0   'False
            Height          =   360
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   825
         End
         Begin MSDataListLib.DataCombo dtcSecao 
            Height          =   360
            Left            =   1005
            TabIndex        =   18
            Top             =   480
            Width           =   5280
            _ExtentX        =   9313
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
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Seção"
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Modelos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -74880
         TabIndex        =   1
         Top             =   1140
         Width           =   5205
         Begin VB.CheckBox chkAcacia 
            Caption         =   "Filizola"
            Height          =   240
            Left            =   120
            TabIndex        =   15
            Top             =   330
            Value           =   2  'Grayed
            Width           =   960
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Toledo"
            Enabled         =   0   'False
            Height          =   240
            Left            =   3870
            TabIndex        =   14
            Top             =   330
            Width           =   1290
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Opções de Exportação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   -74880
         TabIndex        =   3
         Top             =   2610
         Width           =   5235
         Begin VB.CheckBox chkTodos 
            Caption         =   "Todos"
            Height          =   240
            Left            =   3990
            TabIndex        =   13
            Top             =   660
            Width           =   840
         End
         Begin VB.CheckBox chkSecao 
            Caption         =   "Seção"
            Height          =   240
            Left            =   120
            TabIndex        =   12
            Top             =   330
            Width           =   990
         End
         Begin VB.CheckBox chkProdutos 
            Caption         =   "Produtos"
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   660
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         TabIndex        =   4
         Top             =   4980
         Width           =   1845
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
         Height          =   495
         Left            =   -71460
         TabIndex        =   5
         Top             =   4980
         Width           =   1845
      End
      Begin VB.Frame Frame4 
         Caption         =   "Processamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   -74880
         TabIndex        =   8
         Top             =   3720
         Width           =   5265
         Begin VB.Label lblVendedor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rótulo"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   330
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblTermometro_aux 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999999-Indicador"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   150
            TabIndex        =   9
            Top             =   660
            Visible         =   0   'False
            Width           =   1515
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSecao_Tabela 
         Height          =   3705
         Left            =   -74880
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Selecione um produto"
         Top             =   2280
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   6535
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         Enabled         =   0   'False
         FocusRect       =   2
         Appearance      =   0
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo dtcEmpresa 
         Height          =   360
         Left            =   -74880
         TabIndex        =   0
         Top             =   780
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path dos arquivos de retorno de dispositivos móveis"
         Height          =   240
         Left            =   -74880
         TabIndex        =   26
         Top             =   1890
         Width           =   4440
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo da Tabela"
         Height          =   240
         Left            =   -74880
         TabIndex        =   25
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa [F2 ]"
         Height          =   240
         Left            =   -74880
         TabIndex        =   24
         Top             =   540
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmMovimentacoes_exportacao_balancas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Faturamento                                                    '
' Objetivo...............: Exportação para balanças                                       '
' Data de Criação........: 07/05/2006                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:                                                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Public strSql As String
Dim log As New DLLSystemManager.log
Dim strCaminho_Enviar As String
Dim strCaminho_Enviado As String
            

Private Sub cmdConsulta_Empresa_Rx_Click()
    frmCaminho_balanca.Show
End Sub

Private Sub cmdExportar_Click()

    Dim NumArq As Integer
    Dim rstSecao As New ADODB.Recordset
    Dim rstProdutos As New ADODB.Recordset
    
    If Me.txtCaminho.Text = "" Then
       MsgBox "Caminho para geração dos arquivos não informados!Verifique.", vbCritical, "Onlytech"
       Me.txtCaminho.SetFocus
       Exit Sub
    End If
    
    lblTermometro_aux.Visible = True
    lblTermometro_aux.Caption = ""
    
    frmAguarde.Show
'''''''''
'''''''''    If Me.chkSecao.Value = 1 Then
'''''''''        Do While Not rstSecao.EOF
'''''''''           NumArq = FreeFile
'''''''''           Open strCaminho_representante & "\FAMILIA.TXT" For Append As #NumArq
'''''''''
'''''''''           Print #NumArq, "" & rstSecao!PKCodigo_TBSecao & "|" & rstSecao!DFDescricao_TBsecao & "||"
'''''''''
'''''''''           Close #NumArq
'''''''''
'''''''''           rstSecao.MoveNext
'''''''''
'''''''''           lblTermometro_aux.Caption = "Seções - " & rstSecao.AbsolutePosition & " --> " & rstSecao.RecordCount
'''''''''           Me.Refresh
'''''''''        Loop
'''''''''
'''''''''        Set rstSecao = Nothing
'''''''''
'''''''''    End If
    
    If Me.chkProdutos.Value = 1 Then
        
        'Movendo os arquivos antigos para a pasta BKP da raiz da respectiva
        strCaminho_Enviar = Me.txtCaminho.Text & "\CADTXT.TXT"
        strCaminho_Enviado = Me.txtCaminho.Text & "\BKP" & "\CADTXT" & Format(Now, "YYYYMMDD") & "_" & Format(Now, "HHMMSS") & ".TXT"
        
        Call CopyFile(strCaminho_Enviar, strCaminho_Enviado)
        
        If Dir(strCaminho_Enviar) <> "" Then
           Kill strCaminho_Enviar
        End If
        
        strSql = Empty
        'Geração dos arquivos texto para produtos
        strSql = "SELECT IXCodigo_TBProduto, " & _
                 "DFDescricao_resumida_TBProduto, " & _
                 "TBItens_tabela_preco.DFPreco_varejo_TBItens_tabela_preco " & _
                 "FROM TBPRODUTO " & _
                 "LEFT JOIN TBItens_tabela_preco " & _
                 "ON TBPRODUTO.PKId_TBProduto = TBItens_tabela_preco.FKId_TBProduto " & _
                 "LEFT JOIN TBFaixa_comissao_vendedor " & _
                 "ON TBPRODUTO.FKCodigo_TBFaixa_comissao_vendedor = TBFaixa_comissao_vendedor.PKCodigo_TBFaixa_comissao_vendedor " & _
                 "WHERE TBItens_tabela_preco.FKCodigo_TBTabela_preco = (SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBPARAMETROS_VENDA WHERE IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "') AND DFInativo_TBProduto = 0 " & _
                 "AND TBProduto.DFPeso_variavel_TBProduto  = 1 " & _
                 "AND TBProduto.IXCodigo_TBEmpresa = '" & dtcEmpresa.BoundText & "' " & _
                 "ORDER BY TBPRODUTO.IXCodigo_TBProduto"
                 
        Movimentacoes.Select_geral strSql, "BDRetaguarda", rstProdutos, "Otica", Me
        
        rstProdutos.MoveFirst
            
        'CADTXT.TXT
        Do While Not rstProdutos.EOF
            Dim strProdutos As String
            Dim strSituacao_Produto As String
            
            'Iniciando o processamento do arquivo
            Me.Refresh
            NumArq = FreeFile
            
            Open strCaminho_Enviar For Append As #NumArq
            
            Dim strCodigo_Produto As String * 6
            Dim strTipo_Produto As String * 1
            Dim strDescricao_Produto As String * 22
            Dim strPreco_Produto As String * 7
            
            strCodigo_Produto = Format(rstProdutos!IXCodigo_TBProduto, "000000")
            strTipo_Produto = "P"
            strDescricao_Produto = rstProdutos!DFDescricao_resumida_TBProduto
            strPreco_Produto = Format(Replace(Format(rstProdutos!DFPreco_varejo_TBItens_tabela_preco, "#,###0.00"), ",", ""), "0000000")
            
            strProdutos = strCodigo_Produto & strTipo_Produto & strDescricao_Produto & strPreco_Produto
            
            Print #NumArq, strProdutos
            
            Close #NumArq
            
            lblTermometro_aux.Caption = "Produtos - " & rstProdutos.AbsolutePosition & " --> " & rstProdutos.RecordCount
            
            rstProdutos.MoveNext
        Loop
        
        Set rstProdutos = Nothing
    End If
    
    MsgBox "Arquivos processados com sucesso!", vbInformation, "Only Tech"
    
    Unload frmAguarde
    
End Sub
Private Sub cmdCancelar_Click()
    Me.chkProdutos.Value = 0
    Me.chkSecao.Value = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = "113" Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo erro
    
    'Informações Constantes para o log
    log.Usuario = MDIPrincipal.ocxUsuario.Nome
    log.Programa = "Exportação para dispositivos móveis"
    log.Estacao = MDIPrincipal.ocxUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.ocxUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando Exportação para balanças"
    'Gravando o log
    log.Gravar_log "Otica", Me

    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSql, "BDRetaguarda", "Otica", Me
    
    dtcEmpresa.BoundText = MDIPrincipal.ocxUsuario.Empresa
    
    Exit Sub
    
erro:
    Call erro.erro(Me, "Otica", "Load")
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro
    
    log.Evento = "Unload"
    
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "Otica", Me
    
    Exit Sub
erro:
    Call erro.erro(Me, "Otica", "Unload")
    Exit Sub
    
End Sub

Private Sub chkTodos_Click()
    If Me.chkTodos.Value = 1 Then
        Me.chkProdutos.Value = 1
        Me.chkSecao.Value = 1
    Else
        Me.chkProdutos.Value = 0
        Me.chkSecao.Value = 0
    End If
End Sub

Private Function CopyFile(strOrigem As String, strDestino As String) As Single
    Static Buf$
    Dim BTest!, FSize!
    Dim Chunk%, F1%, F2%

    Const BUFSIZE = 1024
           
    If Dir(strOrigem) = "" Then
       Exit Function
    End If
    If Len(Dir(strDestino)) Then
       Kill strDestino
    End If

    On Error GoTo FileCopyError
    F1 = FreeFile
    Open strOrigem For Binary As F1
    F2 = FreeFile
    Open strDestino For Binary As F2

    FSize = LOF(F1)
    BTest = FSize - LOF(F2)
    Do
       If BTest < BUFSIZE Then
          Chunk = BTest
       Else
          Chunk = BUFSIZE
       End If
       Buf = String(Chunk, " ")
       Get F1, , Buf
       Put F2, , Buf
       BTest = FSize - LOF(F2)
       ' __Call percent display here__
       'PercentDone (100 - Int(100 * BTest / FSize))
    Loop Until BTest = 0
    Close F1
    Close F2
    CopyFile = FSize
    
    'Kill strOrigem
    
    Exit Function

FileCopyError:
   MsgBox "Erro ao copiar, verifique!"
   Close F1
   Close F2
   Exit Function
   
End Function
