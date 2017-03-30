VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMarcha_Informacoes_Insumo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informações Adicionais de Insumo"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMarcha_Informacoes_Insumo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   8280
   Begin VB.Frame Frame2 
      Caption         =   "Análises"
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
      Height          =   3945
      Left            =   90
      TabIndex        =   4
      Top             =   600
      Width           =   8085
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgAnalise 
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   330
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   6165
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
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
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.TextBox txtDescricao_Insumo 
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
      Left            =   1665
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   "Nº da Nota Fiscal"
      Top             =   240
      Width           =   6525
   End
   Begin VB.TextBox txtCodigo_Insumo 
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
      Left            =   90
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Peso"
      Top             =   240
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Referências Bibliográficas"
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
      Height          =   3135
      Left            =   90
      TabIndex        =   0
      Top             =   4590
      Width           =   8085
      Begin VB.TextBox txtTecnica_Aplicada 
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
         Height          =   720
         Left            =   120
         MaxLength       =   800
         MultiLine       =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Nº da Nota Fiscal"
         Top             =   2280
         Width           =   7825
      End
      Begin VB.TextBox txtReferencia5 
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
         MaxLength       =   100
         TabIndex        =   14
         ToolTipText     =   "Nº da Nota Fiscal"
         Top             =   1710
         Width           =   7825
      End
      Begin VB.TextBox txtReferencia4 
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
         Left            =   4065
         MaxLength       =   100
         TabIndex        =   11
         ToolTipText     =   "Nº da Nota Fiscal"
         Top             =   1140
         Width           =   3885
      End
      Begin VB.TextBox txtReferencia3 
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
         MaxLength       =   100
         TabIndex        =   10
         ToolTipText     =   "Peso"
         Top             =   1140
         Width           =   3885
      End
      Begin VB.TextBox txtReferencia2 
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
         Left            =   4065
         MaxLength       =   100
         TabIndex        =   7
         ToolTipText     =   "Nº da Nota Fiscal"
         Top             =   540
         Width           =   3885
      End
      Begin VB.TextBox txtReferencia1 
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
         MaxLength       =   100
         TabIndex        =   6
         ToolTipText     =   "Peso"
         Top             =   540
         Width           =   3885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Técnica Aplicada"
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
         TabIndex        =   17
         Top             =   2070
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Referência Bibliográfica V"
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
         TabIndex        =   15
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Referência Bibliográfica IV"
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
         Left            =   4065
         TabIndex        =   13
         Top             =   930
         Width           =   1875
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Referência Bibliográfica III"
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
         TabIndex        =   12
         Top             =   930
         Width           =   1905
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Referência Bibliográfica II"
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
         Left            =   4065
         TabIndex        =   9
         Top             =   330
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Referência Bibliográfica I"
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
         TabIndex        =   8
         Top             =   330
         Width           =   1785
      End
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "Insumo"
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
      TabIndex        =   3
      Top             =   30
      Width           =   525
   End
End
Attribute VB_Name = "frmMarcha_Informacoes_Insumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Serviços                                                       '
' Objetivo...............: Informaçoes Adicionais de Insumo                               '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Jones Peixoto                                                  '
' Data de Criação........: 04/03/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSql As String
Dim rstAplicacao As New ADODB.Recordset

Private Sub Form_Load()
    'ABASTECENDO ANÁLISES
    strSql = "SELECT DFDescricao_TBAnalise_insumo, " & _
             "DFDescricao_TBEspecificacao_analise_insumo " & _
             "FROM TBAnalise_insumo " & _
             "INNER JOIN TBEspecificacao_analise_insumo " & _
             "ON TBAnalise_insumo.PKId_TBAnalise_Insumo = TBEspecificacao_analise_insumo.FKId_TBAnalise_Insumo " & _
             "WHERE FKCodigo_TBInsumo = " & frmMarcha_Analitica.txtInsumo.Text & " " & _
             "ORDER BY PKId_TBAnalise_Insumo,PKId_TBEspecificacao_analise_insumo"
    
    Movimenta_HFlex_Grid strSql, hfgAnalise, "2050,5000", "Análises,Especificações", "BDRetaguarda", "Otica", Me, "N"
    
    'centralizando o texto do cabeçalho
    hfgAnalise.ColAlignmentFixed(1) = 4
    hfgAnalise.ColAlignmentFixed(2) = 4
       
    hfgAnalise.Col = 1
    hfgAnalise.Row = 1
    
    If hfgAnalise.Text <> Empty Then
       Dim strAnalise As String
       'Acertando o tamanho das linhas - vide botao incluir e alterar
       intContador = 1
       Do While intContador <= hfgAnalise.Rows - 1
          hfgAnalise.Row = intContador
          
          hfgAnalise.Col = 1
          strAnalise = hfgAnalise.Text
          
          hfgAnalise.Col = 2
          'verificacao para montagem da proporcionalidade de acordo com o maior dos campos
          If Len(hfgAnalise.Text) > Len(strAnalise) Then
             If Len(hfgAnalise.Text) > 60 Then
                hfgAnalise.RowHeight(intContador) = 285 * CDbl((Len(hfgAnalise.Text)) / 57)
                hfgAnalise.WordWrap = True
             End If
          Else
             If Len(strAnalise) > 20 Then
                hfgAnalise.RowHeight(intContador) = 285 * CDbl((Len(strAnalise)) / 12)
                hfgAnalise.WordWrap = True
             End If
          End If
          intContador = intContador + 1
       Loop
       'rotina para ajuste da mesclagem
       'Call Ajusta_Analise
       'habilitando a mesclagem
       hfgAnalise.MergeCol(1) = True
       hfgAnalise.MergeCells = flexMergeRestrictColumns
       hfgAnalise.ColAlignment(0) = 7
       hfgAnalise.ColAlignment(1) = 4
    Else
       hfgAnalise.Rows = 2
       Movimentacoes.Monta_HFlex_Grid hfgAnalise, "2050,5000", "Análises,Especificações", 2, "Otica", Me
       hfgAnalise.ColAlignmentFixed(1) = 4
       hfgAnalise.ColAlignmentFixed(2) = 4
    End If
    hfgAnalise.Col = 0
    hfgAnalise.Row = 1
    
    txtCodigo_Insumo.Text = frmMarcha_Analitica.txtInsumo.Text
    txtDescricao_Insumo.Text = frmMarcha_Analitica.dtcInsumo.Text
    
    'buscando as referencias e última análise
    strSql = "SELECT DFReferencia_biografica1_TBInsumo,DFReferencia_biografica2_TBInsumo," & _
             "DFReferencia_biografica3_TBInsumo,DFReferencia_biografica4_TBInsumo," & _
             "DFReferencia_biografica5_TBInsumo,DFTecnica_aplicada_TBInsumo " & _
             "FROM TBInsumo " & _
             "WHERE PKCodigo_TBInsumo = " & frmMarcha_Analitica.txtInsumo.Text & ""
          
    Select_geral strSql, "BDRetaguarda", rstAplicacao, "Otica", Me
    
    If rstAplicacao.RecordCount <> 0 Then
       If IsNull(rstAplicacao.Fields("DFReferencia_biografica1_TBInsumo")) = False Then txtReferencia1.Text = rstAplicacao.Fields("DFReferencia_biografica1_TBInsumo")
       If IsNull(rstAplicacao.Fields("DFReferencia_biografica2_TBInsumo")) = False Then txtReferencia2.Text = rstAplicacao.Fields("DFReferencia_biografica2_TBInsumo")
       If IsNull(rstAplicacao.Fields("DFReferencia_biografica3_TBInsumo")) = False Then txtReferencia3.Text = rstAplicacao.Fields("DFReferencia_biografica3_TBInsumo")
       If IsNull(rstAplicacao.Fields("DFReferencia_biografica4_TBInsumo")) = False Then txtReferencia4.Text = rstAplicacao.Fields("DFReferencia_biografica4_TBInsumo")
       If IsNull(rstAplicacao.Fields("DFReferencia_biografica5_TBInsumo")) = False Then txtReferencia5.Text = rstAplicacao.Fields("DFReferencia_biografica5_TBInsumo")
       If IsNull(rstAplicacao.Fields("DFTecnica_aplicada_TBInsumo")) = False Then txtTecnica_Aplicada.Text = rstAplicacao.Fields("DFTecnica_aplicada_TBInsumo")
    End If
    Set rstAplicacao = Nothing

End Sub
