VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTira_Teima_Fechamento_Caixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tira Teima Fechamento Caixa"
   ClientHeight    =   5745
   ClientLeft      =   3780
   ClientTop       =   6690
   ClientWidth     =   6255
   Icon            =   "frmTira_Teima_Fechamento_Caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6255
   Begin VB.Frame fraOpcoes 
      Caption         =   "Opções"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   19
      Top             =   2460
      Width           =   6015
      Begin VB.OptionButton optCupom 
         Caption         =   "Cupom"
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
         Left            =   2190
         TabIndex        =   6
         Top             =   330
         Width           =   1965
      End
      Begin VB.OptionButton optOperacao_Caixa 
         Caption         =   "Operação Caixa"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1965
      End
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
      Height          =   405
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Limpa os Filtros"
      Top             =   5190
      Width           =   1185
   End
   Begin VB.CommandButton cmdImprimir 
      Cancel          =   -1  'True
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Visualiza Impressão"
      Top             =   5190
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Caption         =   "Canceladas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   27
      Top             =   4230
      Width           =   6015
      Begin VB.OptionButton optCancelada_Sim 
         Caption         =   "Sim"
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
         TabIndex        =   12
         Top             =   330
         Width           =   1095
      End
      Begin VB.OptionButton optCancelada_Nao 
         Caption         =   "Não"
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
         Left            =   2160
         TabIndex        =   13
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optCancelada_Todos 
         Caption         =   "Todos"
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
         Left            =   4200
         TabIndex        =   14
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.Frame fraOrdenar 
      Caption         =   "Ordenar"
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
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "Ordenar Impressão"
      Top             =   3180
      Width           =   6015
      Begin VB.OptionButton optData 
         Caption         =   "Data"
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
         Left            =   2190
         TabIndex        =   8
         Top             =   330
         Width           =   735
      End
      Begin VB.OptionButton optFaixa_Horaria 
         Caption         =   "Faixa Hora"
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
         Left            =   4230
         TabIndex        =   9
         Top             =   330
         Width           =   1665
      End
      Begin VB.OptionButton optPDV 
         Caption         =   "PDV"
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
         TabIndex        =   10
         Top             =   660
         Width           =   675
      End
      Begin VB.OptionButton optFinalizadora 
         Caption         =   "Finalizadora"
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
         Left            =   2190
         TabIndex        =   11
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton optOperador 
         Caption         =   "Operador"
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
         TabIndex        =   7
         Top             =   330
         Width           =   1125
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
      Height          =   1725
      Left            =   120
      TabIndex        =   22
      Top             =   690
      Width           =   6015
      Begin VB.TextBox txtFinalizadora 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Código da Finalizadora"
         Top             =   1215
         Width           =   1155
      End
      Begin VB.TextBox txtPdv 
         Height          =   360
         Left            =   4680
         TabIndex        =   2
         ToolTipText     =   "Código do Ponto de Venda"
         Top             =   570
         Width           =   1185
      End
      Begin VB.TextBox txtOperador 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Código do Operador"
         Top             =   570
         Width           =   1155
      End
      Begin MSDataListLib.DataCombo dtcFinalizadora 
         Height          =   360
         Left            =   1320
         TabIndex        =   4
         Top             =   1215
         Width           =   4575
         _ExtentX        =   8070
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
      Begin MSDataListLib.DataCombo dtcOperador 
         Height          =   360
         Left            =   1320
         TabIndex        =   1
         Top             =   570
         Width           =   3315
         _ExtentX        =   5847
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Finalizadora"
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
         TabIndex        =   25
         Top             =   975
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PDV"
         Height          =   240
         Left            =   4710
         TabIndex        =   24
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Operador"
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
         TabIndex        =   23
         Top             =   330
         Width           =   810
      End
   End
   Begin MSDataListLib.DataCombo dtcEmpresa 
      Height          =   360
      Left            =   120
      TabIndex        =   20
      Top             =   300
      Width           =   6015
      _ExtentX        =   10610
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
   Begin MSComCtl2.DTPicker dtpInicial 
      Height          =   360
      Left            =   120
      TabIndex        =   15
      Top             =   5220
      Width           =   1515
      _ExtentX        =   2672
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
      Format          =   20643841
      CurrentDate     =   37881
   End
   Begin MSComCtl2.DTPicker dtpFinal 
      Height          =   360
      Left            =   2070
      TabIndex        =   16
      Top             =   5220
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   20643841
      CurrentDate     =   37881
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Período"
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
      TabIndex        =   29
      Top             =   4980
      Width           =   645
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Left            =   1695
      TabIndex        =   28
      Top             =   5340
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
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
      TabIndex        =   21
      Top             =   60
      Width           =   1290
   End
End
Attribute VB_Name = "frmTira_Teima_Fechamento_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador Vendas                                            '
' Objetivo...............: Tira TeimaFEchamento de Caixa                                  '
' Data de Criação........: 22/06/2005                                                     '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Última Manutenção......: Criação do Formulário e dos Relatórios Tira-Teima              '
' Desenvolvedor..........: Leandro Lawall Guedes                                          '
' Data última manutenção.: 05/08/2006                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public strSQL As String
Dim log As New DLLSystemManager.log
'Declaração da variavel do intercomunicador de mensagens
Private Cliente_mensagem_exe As VetorDeMensagens.ClienteDeMensagens

Private Sub Cancelar()

    On Error GoTo Erro
    
    Call Objetos.Limpa_TXT(Me)
            
    'Inserir log
    log.Evento = "Cancelar"
    log.Descricao = "Cancelamento de Relatorio de Operação de Caixa"
    log.Tipo = 1
    log.Hora = Format(Now, "hh:mm:ss")
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    dtpInicial.Value = Date
    dtpInicial.Day = 1
    dtpFinal = Date
    txtOperador.SetFocus
    optOperador.Value = True
    optAnalitico.Value = True
    optCancelada_Todos.Value = True
    
    Exit Sub
    
Erro:

    Call Erro.Erro(Me, "OTICA", "Cancelar")
    Exit Sub
    
End Sub

Private Sub Impressao()
    
    Dim frmAux_Imp As Form
    Dim strFormulas As String
    Dim strValores As String
    
        If optOperacao_Caixa.Value Then
        strSQL = "SELECT distinct FKCodigo_TBPdv, " & _
                 "TBFinalizadora.IXCodigo_TBFinalizadora, " & _
                 "TBFinalizadora.DFDescricao_TBFinalizadora, " & _
                 "TBOperacao_Caixa.FKCodigo_TBOperadores_ecf, " & _
                 "DFNome_TBOperadores_ecf, " & _
                 "DFData_TBOperacao_caixa, " & _
                 "DFHora_TBOperacao_caixa, " & _
                 "TBFinalizadora.DFDebito_credito_TBFinalizadora, " & _
                 "DFValor_TBOperacao_caixa, " & _
                 "DFTipo_operacao_TBOperacao_caixa, " & _
                 "DFStatus_aberto_fechado_TBOperacao_caixa, " & _
                 "DFCodigo_cupom_impressora_TBCupom, " & _
                 "DFCodigo_cupom_impressora_TBOperacao_caixa AS DFNumero_Cupom_TBOperacao_caixa, " & _
                 "DFGrant_total_impressora_TBOperacao_caixa, " & _
                 "DFObservacao_TBOperacao_caixa " & _
                 "FROM TBOperacao_caixa " & _
                 "INNER JOIN TBFinalizadora ON TBOPERACAO_CAIXA.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                 "INNER JOIN TBOperadores_ecf ON TBOPERACAO_CAIXA.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
                 "LEFT JOIN TBCUPOM ON TBOPERACAO_CAIXA.DFNumero_Cupom_TBOperacao_caixa = TBCupom.DFNumero_TBCupom " & _
                 "WHERE TBOPERACAO_CAIXA.FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
                 "AND (TBOPERACAO_CAIXA.DFData_TBOperacao_caixa BETWEEN '" & Format(Me.dtpInicial.Value, "YYYYMMDD") & "' " & _
                 "AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "' ) "
        Else
            strSQL = "SELECT distinct FKCodigo_TBPdv AS PDV, " & _
                     "TBFinalizadora.IXCodigo_TBFinalizadora, " & _
                     "TBFinalizadora.DFDescricao_TBFinalizadora, " & _
                     "TBOperacao_Caixa.FKCodigo_TBOperadores_ecf, " & _
                     "TBOperadores_ecf.DFNome_TBOperadores_ecf, " & _
                     "DFData_TBOperacao_caixa, " & _
                     "DFHora_TBOperacao_caixa, " & _
                     "TBFinalizadora.DFDebito_credito_TBFinalizadora, " & _
                     "DFValor_TBOperacao_caixa, " & _
                     "DFTipo_operacao_TBOperacao_caixa, " & _
                     "DFStatus_aberto_fechado_TBOperacao_caixa, " & _
                     "DFCodigo_cupom_impressora_TBCupom, " & _
                     "DFCodigo_cupom_impressora_TBOperacao_caixa AS DFNumero_Cupom_TBOperacao_caixa, " & _
                     "DFGrant_total_impressora_TBOperacao_caixa, " & _
                     "DFObservacao_TBOperacao_caixa " & _
                     "FROM TBOperacao_caixa " & _
                     "INNER JOIN TBFinalizadora ON TBOPERACAO_CAIXA.FKId_TBFinalizadora = TBFinalizadora.PKId_TBFinalizadora " & _
                     "INNER JOIN TBOperadores_ecf ON TBOPERACAO_CAIXA.FKCodigo_TBOperadores_ecf = TBOperadores_ecf.PKCodigo_TBOperadores_ecf " & _
                     "LEFT JOIN TBCUPOM ON TBOPERACAO_CAIXA.DFNumero_Cupom_TBOperacao_caixa = TBCupom.DFNumero_TBCupom " & _
                     "WHERE TBOPERACAO_CAIXA.FKCodigo_TBEmpresa = " & dtcEmpresa.BoundText & " " & _
                     "AND (TBOPERACAO_CAIXA.DFData_TBOperacao_caixa BETWEEN '" & Format(Me.dtpInicial.Value, "YYYYMMDD") & "' " & _
                     "AND '" & Format(dtpFinal.Value, "YYYYMMDD") & "' ) "
        End If
    ' Finalizadora
    If dtcFinalizadora.Text <> "" Then
       strSQL = strSQL & " AND TBFinalizadora.IXCodigo_TBFinalizadora = " & dtcFinalizadora.BoundText
    End If
    
    ' PDV
    If txtPdv.Text <> "" Then
       strSQL = strSQL & " AND FKCodigo_TBPdv = " & txtPdv.Text
    End If
    
    ' Operador
    If dtcOperador.BoundText <> "" Then
       strSQL = strSQL & " AND TBOperadores_ecf.PKCodigo_TBOperadores_ecf = " & dtcOperador.BoundText
    End If
    
    If optCancelada_Sim.Value = True Then
       strSQL = strSQL & " AND TBCupom.DFCancelado_TBCupom = '1' "
    End If
    
    If optCancelada_Nao.Value = True Then
       strSQL = strSQL & " AND TBCupom.DFCancelado_TBCupom = '0' "
    End If
      
    
    ' Ordenacao do relatorio
    'If optFaixa_Horaria.Value = True Then
    '    strSQL = strSQL & " ORDER BY DFHora_TBOperacao_caixa "
    'ElseIf optData.Value = True Then
    '    strSQL = strSQL & " ORDER BY DFData_TBOperacao_caixa "
    'ElseIf optPDV.Value = True Then
    '    strSQL = strSQL & " ORDER BY FKCodigo_TBPdv "
    'ElseIf optOperador.Value = True Then
    '    strSQL = strSQL & " ORDER BY TBOperadores_ecf.PKCodigo_TBOperadores_ecf"
    'End If
    
    'Inicio da Chamada do Formulário
    'Tratamento de Erro
    If strSQL = "" Then
       MsgBox "Não existem informações suficientes para a geração deste relatório.Verifique.", vbInformation, "Only Tech"
       Exit Sub
    End If

    frmAguarde.Show
    DoEvents

    Set frmAux_Imp = New frmConsole_Geral
    
    strFormulas = "Cliente;Tipo_relatorio;Periodo"
    strValores = Funcoes_Gerais.Abrir_nome_cliente_registro("Otica", Me) & ";Listagem Geral;Período: De " & dtpInicial.Value & " Até " & dtpFinal.Value
    
    If optOperacao_Caixa.Value Then
        If optOperador.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_Operador.rpt", strFormulas, strValores
        ElseIf optFinalizadora.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_Finalizadora.rpt", strFormulas, strValores
        ElseIf optPDV.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_PDV.rpt", strFormulas, strValores
        ElseIf optFaixa_Horaria.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_Faixa_Hora.rpt", strFormulas, strValores
        ElseIf optData.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_Data.rpt", strFormulas, strValores
        End If
    ElseIf optCupom.Value Then
        If optOperador.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_Cupom_Operador.rpt", strFormulas, strValores
        ElseIf optFinalizadora.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_Cupom_Finalizadora.rpt", strFormulas, strValores
        ElseIf optPDV.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_Cupom_PDV.rpt", strFormulas, strValores
        ElseIf optFaixa_Horaria.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_Cupom_Faixa_Hora.rpt", strFormulas, strValores
        ElseIf optData.Value Then
            frmAux_Imp.setParametros strSQL, "rptTira_Teima_Fechamento_Caixa_Cupom_Data.rpt", strFormulas, strValores
        End If
    End If
    
    
    frmAux_Imp.Show
    
    Unload frmAguarde
    
    Set frmAux_Imp = Nothing

End Sub

Private Sub cmdCancelar_Click()
    Call Cancelar
End Sub

Private Sub cmdImprimir_Click()
    If dtpInicial.Value > dtpFinal.Value Then
       MsgBox "Data Final menor que Data Inicial. Verifique!", vbInformation, "Only Tech"
       Exit Sub
    End If

    frmAguarde.Show
    DoEvents
    Call Impressao
    Unload frmAguarde
    
End Sub

Private Sub dtcFinalizadora_LostFocus()
    txtFinalizadora.Text = dtcFinalizadora.BoundText
    If IsNumeric(txtFinalizadora.Text) = False Or dtcFinalizadora.Text = Empty Then txtFinalizadora.Text = Empty: Exit Sub
End Sub

Private Sub dtcOperador_LostFocus()
    If txtOperador.Text = Empty Then
       Call Movimentacoes.Verifica_DataCombo(dtcOperador.Text)
    End If
    txtOperador.Text = dtcOperador.BoundText
    If IsNumeric(txtOperador.Text) = False Or dtcOperador.Text = Empty Then txtOperador.Text = Empty: Exit Sub
End Sub

Private Sub dtpFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

Private Sub dtpInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then KeyCode = vbKeyTab
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyF2 And booAlterar = False Then Movimentacoes.Verifica_Acesso_Usuario dtcEmpresa, "Otica", "BDRetaguarda", Me
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
    log.Programa = "Tira Teima Fechamento Caixa"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
 '   If MDIPrincipal.booDesign_time = False Then
        Call Movimentacoes.Acessibilidade_inicio_relatorios("Tira Teima Fechamento Caixa", MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
 '   End If
    
    log.Descricao = "Inicializando Relatorio de Fechamento Diário de Vendas"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    'Montando os datacombo de tela
    strSQL = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcEmpresa, strSQL, "BDRetaguarda", "Otica", Me

    strSQL = "SELECT IXCodigo_TBFinalizadora, DFDescricao_TBFinalizadora FROM TBFinalizadora "
    Movimentacoes.Movimenta_DataCombo "IXCodigo_TBFinalizadora", "DFDescricao_TBFinalizadora", dtcFinalizadora, strSQL, "BDRetaguarda", "Otica", Me
    
    strSQL = "SELECT PKCodigo_TBOperadores_ecf, DFNome_TBOperadores_ecf FROM TBOperadores_ecf "
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBOperadores_ecf", "DFNome_TBOperadores_ecf", dtcOperador, strSQL, "BDRetaguarda", "Otica", Me
     
    'dtcCodigo_empresa.boundtext = ---- Inserir aqui informações da DLLIntercomunicador de EXE's
    dtcEmpresa.BoundText = MDIPrincipal.OCXUsuario.Empresa
 
    dtpInicial.Value = Date
    dtpFinal.Value = Date
    
    optOperacao_Caixa.Value = True
    optOperador.Value = True
    optCancelada_Todos.Value = True
    Exit Sub
    
Erro:

    Call Erro.Erro(Me, "OTICA", "Load")
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Erro
    
    log.Evento = "Unload"
    log.Hora = Format(Now, "hh:mm:ss")
    log.Descricao = "Saindo de Tira Teima Fechamento Caixa"
    
    'Gravando Log
    log.Gravar_log "OTICA", Me
    
    Exit Sub
    
Erro:

    Call Erro.Erro(Me, "OTICA", "Unload")
    Exit Sub
  
End Sub

Private Sub txtFinalizadora_Change()
    dtcFinalizadora.BoundText = txtFinalizadora.Text
    If IsNumeric(txtFinalizadora.Text) = False Then txtFinalizadora.Text = Empty: Exit Sub
End Sub

Private Sub txtOperador_Change()
    dtcOperador.BoundText = txtOperador.Text
    If IsNumeric(txtOperador.Text) = False Then txtOperador.Text = Empty: Exit Sub
End Sub
