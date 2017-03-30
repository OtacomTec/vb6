VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{40BD39E3-6F1E-11D1-B2DF-444553540000}#1.0#0"; "OCXShape.ocx"
Begin VB.Form frmInformacoes_Adicionais_Plano_Pagamento 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmInformacoes_Adicionais_Plano_Pagamento.frx":0000
   ScaleHeight     =   3900
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgNumero_Dias 
      Height          =   855
      Left            =   180
      TabIndex        =   9
      Top             =   2580
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   1508
      _Version        =   393216
      BackColor       =   15335166
      ForeColor       =   8388608
      FixedCols       =   0
      BackColorFixed  =   13098744
      BackColorBkg    =   -2147483624
      FocusRect       =   2
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin FormShape.FormShape FormShape1 
      Left            =   5430
      Top             =   3450
      ShapeType       =   1
      MaskColor       =   16777215
      AutoScale       =   -1  'True
      ScaleX          =   1
      ScaleY          =   1
      ShapeString     =   ""
   End
   Begin VB.Label lblDescricao 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   750
      Width           =   5745
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
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
      Left            =   180
      TabIndex        =   7
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label lblSupervisor 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   2130
      Width           =   5745
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão Supervisor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   1920
      Width           =   1605
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão Vendedor"
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
      Left            =   180
      TabIndex        =   4
      Top             =   1470
      Width           =   1485
   End
   Begin VB.Label lblVendedor 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1680
      Width           =   5745
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Modalidade"
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
      Left            =   180
      TabIndex        =   2
      Top             =   990
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Informações Adicionais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   5535
   End
   Begin VB.Label lblModalidade 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   1200
      Width           =   5745
   End
End
Attribute VB_Name = "frmInformacoes_Adicionais_Plano_Pagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Faturamento                                                    '
' Objetivo...............: Informações do Produto                                         '
' Data de Criação........: 26/06/2006                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......:                                                                '
' Desenvolvedor..........: Rodrigo Santos                                                 '
' Data última manutenção.:                                                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Info_Plano(Codigo_plano As Integer, Empresa As Integer, Aplicacao As String, Banco As String, Optional Top As Integer, Optional Left As Integer, Optional width As Integer, Optional Height As Integer)
        
    Dim strSql As String
    Dim strID As String
    Dim rstPlano As New ADODB.Recordset
    Dim rstDias As New ADODB.Recordset
    Dim conexao_Integracao As New DLLConexao_Sistema.Conexao
           
    'INDICANDO O BANCO A CONECTAR-SE
    conexao_Integracao.Initial_Catalog = Banco
    
    'ABRINDO CONEXAO COM BANCO
    conexao_Integracao.Abrir_conexao (Aplicacao)
    
    DoEvents
    
    rstPlano.CursorLocation = adUseClient
    
    'STRING QUE COLETA DADOS RELATIVOS AO PLANO
    strSql = " SELECT PKId_TBPlano_pagamento, IXCodigo_TBEmpresa," & _
             " IXCodigo_TBPlano_pagamento, DFDescricao_TBPlano_pagamento," & _
             " DFModalidade_TBPlano_pagamento, DFAcrescimo_desconto_TBPlano_pagamento," & _
             " DFPercentual_TBPlano_pagamento, DFAtivo_inativo_TBPlano_pagamento," & _
             " DFValor_minimo_pedido_TBPlano_pagamento," & _
             " DFBaixa_Titulo_TBPlano_pagamento," & _
             " DFImprime_vencimento_TBPlano_pagamento," & _
             " DFDigita_vencimento_TBPlano_pagamento," & _
             " FKCodigo_TBFaixa_comissao_vendedor," & _
             " TBFaixa_comissao_vendedor.DFPercentual_TBFaixa_comissao_vendedor," & _
             " FKCodigo_TBFaixa_comissao_supervisor," & _
             " TBFaixa_comissao_supervisor.DFPercentual_TBFaixa_comissao_supervisor," & _
             " DFIntegrado_filiais_TBPlano_pagamento," & _
             " DFIntegrado_portal_TBPlano_pagamento," & _
             " DFData_alteracao_TBPlano_pagamento," & _
             " DFCodigo_Identificador_TBPlano_pagamento" & _
             " FROM TBPlano_pagamento " & _
             " LEFT JOIN TBFaixa_comissao_vendedor " & _
             " ON TBPlano_pagamento.FKCodigo_TBFaixa_comissao_vendedor = TBFaixa_comissao_vendedor.PKCodigo_TBFaixa_comissao_vendedor " & _
             " LEFT JOIN TBFaixa_comissao_supervisor " & _
             " ON TBPlano_pagamento.FKCodigo_TBFaixa_comissao_supervisor = TBFaixa_comissao_supervisor.PKCodigo_TBFaixa_comissao_supervisor " & _
             " WHERE IXCodigo_TBPlano_pagamento = '" & Codigo_plano & "' AND IXCodigo_TBEmpresa = '" & Empresa & "'"
             
    rstPlano.Open strSql, conexao_Integracao.CNconexao, adOpenStatic, adLockReadOnly
    
    rstPlano.MoveFirst
    Me.Show
    
    'PREENCHENDO LABELS
    If rstPlano.BOF = False Then
       If IsNull(rstPlano!PKId_TBPlano_pagamento) = False Then
          strID = rstPlano!PKId_TBPlano_pagamento
       End If
       If IsNull(rstPlano!DFDescricao_TBPlano_pagamento) = False Then
          lblDescricao.Caption = rstPlano!DFDescricao_TBPlano_pagamento
       End If
       If IsNull(rstPlano!DFModalidade_TBPlano_pagamento) = False Then
          lblModalidade.Caption = rstPlano!DFModalidade_TBPlano_pagamento
       End If
       If IsNull(rstPlano!FKCodigo_TBFaixa_comissao_vendedor) = False Then
          lblVendedor.Caption = rstPlano!FKCodigo_TBFaixa_comissao_vendedor & " - " & rstPlano!DFPercentual_TBFaixa_comissao_vendedor
       End If
       If IsNull(rstPlano!FKCodigo_TBFaixa_comissao_supervisor) = False Then
          lblSupervisor.Caption = rstPlano!FKCodigo_TBFaixa_comissao_supervisor & " - " & rstPlano!DFPercentual_TBFaixa_comissao_supervisor
       End If
    End If
      
    Set rstPlano = Nothing
    
    conexao_Integracao.Fechar_conexao
    
    'MONTANDO GRID
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim matDias(20) As String
    Dim I As Integer
    
    
    Call Movimentacoes.Monta_HFlex_Grid(hfgNumero_Dias, "1000", " ", 1, "Otica", Me)
        
        hfgNumero_Dias.Col = 0
        hfgNumero_Dias.Row = 1
        hfgNumero_Dias.Font.Name = "Tahoma"
        hfgNumero_Dias.CellFontSize = 7
        hfgNumero_Dias.CellBackColor = &H80FFFF
        hfgNumero_Dias.Text = 1
        
        I = 1
        
        strSql = "SELECT * FROM TBDias_Pagamento WHERE FKId_TBPlano_pagamento = " & strID & ""
        Call Movimentacoes.Select_geral(strSql, "BDRetaguarda", rstDias, "Otica", Me)
        
        If rstDias.RecordCount <> 0 Then
           Do While rstDias.EOF = False
              On Error Resume Next
              matDias(I) = rstDias.Fields("DFNumero_dias_TBDias_pagamento")
              hfgNumero_Dias.Row = 1
              hfgNumero_Dias.Col = I
              If Err.Number = 30010 Then
                 hfgNumero_Dias.Cols = hfgNumero_Dias.Cols + 1
                 hfgNumero_Dias.Col = I
              End If
              hfgNumero_Dias.CellFontSize = 10
              hfgNumero_Dias.Text = rstDias.Fields("DFNumero_dias_TBDias_pagamento")
              hfgNumero_Dias.Row = 0
              hfgNumero_Dias.Col = I
              hfgNumero_Dias.Text = I & "ª"
              I = I + 1
              rstDias.MoveNext
           Loop
        Else
           Call Movimentacoes.Monta_HFlex_Grid(hfgNumero_Dias, "1000", " ", 1, "Otica", Me)
        End If
        
        'Não Retirar
        rstDias.Close
        I = 0
        
        'PAREI
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    DoEvents
    
    'POSICIONA FORM
    Me.Left = Left + width + 60
    Me.Top = Top
    
End Function

Private Sub Form_Click()
    Unload frmInformacoes_Adicionais_Plano_Pagamento
End Sub

Private Sub Form_Load()
    FormShape1.hWnd = frmInformacoes_Adicionais_Plano_Pagamento.hWnd
    FormShape1.ShapePicture = frmInformacoes_Adicionais_Plano_Pagamento.Picture
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub Label10_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Label4_Click()
    Unload Me
End Sub

Private Sub Label6_Click()
    Unload Me
End Sub

Private Sub lblDescricao_Click()
    Unload Me
End Sub

Private Sub lblModalidade_Click()
    Unload Me
End Sub

Private Sub lblSupervisor_Click()
    Unload Me
End Sub

Private Sub lblVendedor_Click()
    Unload Me
End Sub
