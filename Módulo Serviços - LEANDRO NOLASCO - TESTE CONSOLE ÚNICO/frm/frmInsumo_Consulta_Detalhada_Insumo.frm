VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInsumo_Consulta_Detalhada_Insumo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Detalhada Insumo"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsumo_Consulta_Detalhada_Insumo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7470
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar"
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
      Height          =   585
      Left            =   6300
      Picture         =   "frmInsumo_Consulta_Detalhada_Insumo.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3990
      Width           =   1065
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00800000&
      Height          =   1725
      Left            =   3480
      TabIndex        =   10
      Top             =   30
      Width           =   3885
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
         Left            =   2700
         Picture         =   "frmInsumo_Consulta_Detalhada_Insumo.frx":1B0C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1020
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
         Left            =   1560
         Picture         =   "frmInsumo_Consulta_Detalhada_Insumo.frx":1E96
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox txtConsulta 
         Height          =   360
         Left            =   90
         MaxLength       =   40
         TabIndex        =   6
         Top             =   570
         Width           =   3675
      End
      Begin VB.Label lblVariavel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Width           =   60
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opções de Filtro"
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
      Height          =   1725
      Left            =   90
      TabIndex        =   9
      Top             =   30
      Width           =   3375
      Begin VB.OptionButton optEspecificacao 
         Caption         =   "Especificação"
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
         Left            =   1860
         TabIndex        =   5
         Top             =   660
         Width           =   1455
      End
      Begin VB.OptionButton optAnalise 
         Caption         =   "Análise"
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
         Left            =   1860
         TabIndex        =   4
         Top             =   330
         Width           =   1455
      End
      Begin VB.OptionButton optNome_Cientifico 
         Caption         =   "Nome Científico"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton optCodigo 
         Caption         =   "Código"
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
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   1575
      End
      Begin VB.OptionButton optDescricao 
         Caption         =   "Descrição"
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
         Left            =   120
         TabIndex        =   2
         Top             =   990
         Width           =   1755
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
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
         Left            =   120
         TabIndex        =   0
         Top             =   330
         Width           =   885
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgConsulta_Insumo 
      Height          =   2145
      Left            =   90
      TabIndex        =   13
      Top             =   1770
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   3784
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
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   150
      TabIndex        =   15
      Top             =   4140
      Width           =   60
   End
   Begin VB.Label lblDescricao 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   14
      Top             =   4140
      Width           =   4995
   End
End
Attribute VB_Name = "frmInsumo_Consulta_Detalhada_Insumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Informações sobre o Insumo                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Sá                                                       '
' Data de Criação........: 04/05/2006                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim strsql As String
Dim intContador As Integer

Private Sub cmdAplicar_Click()
    'VERIFICA SE JÁ EXISTEM ANALISES E ESPECIFICAÇÕES LANÇADAS
    frmInsumo.hfgAnalise.Row = 1
    frmInsumo.hfgAnalise.Col = 0
    
    If frmInsumo.hfgAnalise.Text = Empty Then
        
       Dim rstAplicacao As New ADODB.Recordset
       
       hfgConsulta_Insumo.Col = 1
       
       strsql = "SELECT DFDescricao_TBAnalise_insumo," & _
                "DFDescricao_TBEspecificacao_analise_insumo " & _
                "FROM TBEspecificacao_analise_insumo,TBAnalise_insumo " & _
                "WHERE TBEspecificacao_analise_insumo.FKId_TBAnalise_Insumo = TBAnalise_insumo.PKId_TBAnalise_Insumo " & _
                "AND TBAnalise_insumo.FKCodigo_TBInsumo = " & hfgConsulta_Insumo.Text & " " & _
                "ORDER BY TBAnalise_insumo.DFDescricao_TBAnalise_insumo"
                
       Movimenta_HFlex_Grid strsql, frmInsumo.hfgAnalise, "2050,5500", "Análises,Especificações", "BDRetaguarda", "Otica", Me
       
       
       'centralizando o texto do cabeçalho
       frmInsumo.hfgAnalise.ColAlignmentFixed(1) = 4
       frmInsumo.hfgAnalise.ColAlignmentFixed(2) = 4
           
       frmInsumo.hfgAnalise.Col = 1
       frmInsumo.hfgAnalise.Row = 1
        
       If frmInsumo.hfgAnalise.Text <> Empty Then
          Dim strAnalise As String
          Dim intContador As Integer
          'Acertando o tamanho das linhas - vide botao incluir e alterar
          intContador = 1
          Do While intContador <= frmInsumo.hfgAnalise.Rows - 1
             frmInsumo.hfgAnalise.Row = intContador
             
             frmInsumo.hfgAnalise.Col = 1
             strAnalise = frmInsumo.hfgAnalise.Text
             
             frmInsumo.hfgAnalise.Col = 2
             'verificacao para montagem da proporcionalidade de acordo com o maior dos campos
             If Len(frmInsumo.hfgAnalise.Text) > Len(strAnalise) Then
                If Len(frmInsumo.hfgAnalise.Text) > 60 Then
                   frmInsumo.hfgAnalise.RowHeight(intContador) = 285 * CDbl((Len(frmInsumo.hfgAnalise.Text)) / 49)
                   frmInsumo.hfgAnalise.WordWrap = True
                End If
             Else
                If Len(strAnalise) > 20 Then
                   frmInsumo.hfgAnalise.RowHeight(intContador) = 285 * CDbl((Len(strAnalise)) / 10)
                   frmInsumo.hfgAnalise.WordWrap = True
                End If
             End If
             intContador = intContador + 1
          Loop
          'rotina para ajuste da mesclagem
          Call frmInsumo.Ajusta_Analise
          'habilitando a mesclagem
          frmInsumo.hfgAnalise.MergeCol(1) = True
          frmInsumo.hfgAnalise.MergeCells = flexMergeRestrictColumns
          frmInsumo.hfgAnalise.ColAlignment(0) = 7
          frmInsumo.hfgAnalise.ColAlignment(1) = 4
       Else
          frmInsumo.hfgAnalise.Rows = 2
          Movimentacoes.Monta_HFlex_Grid frmInsumo.hfgAnalise, "2050,5500", "Análises,Especificações", 2, "Otica", Me
          frmInsumo.hfgAnalise.ColAlignmentFixed(1) = 4
          frmInsumo.hfgAnalise.ColAlignmentFixed(2) = 4
       End If
    End If
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
   'Limpando o grid de itens
    
    hfgConsulta_Insumo.ClearStructure
    
    hfgConsulta_Insumo.Rows = 2
    Movimentacoes.Monta_HFlex_Grid hfgConsulta_Insumo, "1000,4000,1600", "Código,Descrição,Nome Científico", 3, "OTICA", Me
    
    lblCodigo.Caption = Empty
    lblDescricao.Caption = Empty
    cmdAplicar.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Habilita a troca de campos pelo ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdConsultar_Click()
    Dim booRetorno As Boolean
    
    lblCodigo.Caption = Empty
    lblDescricao.Caption = Empty
    
    If txtConsulta.Text = Empty And optTodos.Value = False Then
       MsgBox "Selecione um campo e digite os dados para consulta.", vbCritical, "Only Tech"
       txtConsulta.SetFocus
       Exit Sub
    End If
    
    strsql = "SELECT PKCodigo_TBInsumo,DFDescricao_TBInsumo,DFNome_cientifico_TBInsumo " & _
             "FROM TBInsumo " & _
             "INNER JOIN TBAnalise_Insumo " & _
             "ON TBInsumo.PKCodigo_TBInsumo = TBAnalise_Insumo.FKCodigo_TBInsumo " & _
             "INNER JOIN TBEspecificacao_analise_insumo " & _
             "ON TBAnalise_Insumo.PKId_TBAnalise_Insumo = TBEspecificacao_analise_insumo.FKId_TBAnalise_Insumo "
             
    If optCodigo.Value = True Then
       If IsNumeric(txtConsulta.Text) = False Then txtConsulta.Text = Empty
       strsql = strsql & "WHERE PKCodigo_TBInsumo = '" & Me.txtConsulta.Text & "' "
    ElseIf optDescricao.Value = True Then
        strsql = strsql & "WHERE DFDescricao_TBInsumo LIKE '%" & Me.txtConsulta.Text & "%' "
    ElseIf optNome_Cientifico.Value = True = True Then
       strsql = strsql & "WHERE DFNome_cientifico_TBInsumo LIKE '%" & Me.txtConsulta.Text & "%' "
    ElseIf optAnalise.Value = True Then
        strsql = strsql & "WHERE DFDescricao_TBAnalise_insumo LIKE '%" & Me.txtConsulta.Text & "%' "
    ElseIf optEspecificacao.Value = True Then
        strsql = strsql & "WHERE DFDescricao_TBEspecificacao_analise_insumo LIKE '%" & Me.txtConsulta.Text & "%' "
    End If
    
    strsql = strsql & " GROUP BY PKCodigo_TBInsumo,DFDescricao_TBInsumo,DFNome_cientifico_TBInsumo "
    
    frmAguarde.Show
    
    Call Movimentacoes.Movimenta_HFlex_Grid(strsql, hfgConsulta_Insumo, "1000,4000,1600", "Código,Descrição,Nome Científico", "BDRetaguarda", "Otica", Me)
      
    hfgConsulta_Insumo.Row = 1
    hfgConsulta_Insumo.Col = 0
    If hfgConsulta_Insumo.Text = Empty Then
       hfgConsulta_Insumo.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgConsulta_Insumo, "1000,4000,1600", "Código,Descrição,Nome Científico", 3, "Otica", Me
    End If
    
    cmdAplicar.Enabled = False
    Unload frmAguarde
    
End Sub

Private Sub Form_Load()
    Movimentacoes.Monta_HFlex_Grid Me.hfgConsulta_Insumo, "1000,4000,1600", "Código,Descrição,Nome Científico", 3, "Otica", Me
End Sub

Private Sub optAnalise_Click()
    txtConsulta.Visible = True
    lblVariavel.Caption = "Análise"
    txtConsulta.SetFocus
End Sub

Private Sub optCodigo_Click()
    txtConsulta.Visible = True
    lblVariavel.Caption = "Código"
    txtConsulta.SetFocus
End Sub

Private Sub optDescricao_Click()
    txtConsulta.Visible = True
    lblVariavel.Caption = "Descrição"
    txtConsulta.SetFocus
End Sub

Private Sub optEspecificacao_Click()
    txtConsulta.Visible = True
    lblVariavel.Caption = "Especificação"
    txtConsulta.SetFocus
End Sub

Private Sub optNome_Cientifico_Click()
    txtConsulta.Visible = True
    lblVariavel.Caption = "Nome Científico"
    txtConsulta.SetFocus
End Sub

Private Sub optTodos_Click()
    txtConsulta.Visible = False
    lblVariavel.Caption = ""
    cmdConsultar.SetFocus
End Sub

Private Sub hfgConsulta_insumo_Click()
    If hfgConsulta_Insumo.Text <> Empty Then
       hfgConsulta_Insumo.Col = 0
       lblCodigo.Caption = hfgConsulta_Insumo.TextArray((hfgConsulta_Insumo.Row * hfgConsulta_Insumo.Cols + hfgConsulta_Insumo.Col + 1))
       lblDescricao.Caption = hfgConsulta_Insumo.TextArray((hfgConsulta_Insumo.Row * hfgConsulta_Insumo.Cols + hfgConsulta_Insumo.Col + 2))
       cmdAplicar.Enabled = True
    End If
End Sub

Private Sub txtConsulta_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtConsulta_LostFocus()
   txtConsulta.Text = UCase(txtConsulta.Text)
End Sub
