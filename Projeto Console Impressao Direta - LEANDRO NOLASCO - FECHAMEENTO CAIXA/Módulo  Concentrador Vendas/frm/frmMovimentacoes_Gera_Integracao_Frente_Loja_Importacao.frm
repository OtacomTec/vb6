VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{123406F5-5DCA-4A4C-92CB-A113A0C83143}#1.0#0"; "AUTOCOMPLETAR.OCX"
Begin VB.Form frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gera Integração Frente Loja - Importação"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.frx":1782
   ScaleHeight     =   4470
   ScaleWidth      =   8205
   Begin AutoCompletar.CbCompleta cbbEmpresa_Exportadora 
      Height          =   360
      Left            =   5970
      TabIndex        =   1
      Top             =   600
      Width           =   2145
      _ExtentX        =   3784
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
   Begin VB.TextBox txtCaminho 
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   1290
      Width           =   7515
   End
   Begin VB.CommandButton cmdCaminho 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   365
      Left            =   7690
      Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.frx":1AC4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Localize o caminho onde o arquivo será salvo"
      Top             =   1305
      Width           =   405
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   10350
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.frx":1E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.frx":2168
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.frx":2482
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.frx":281C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.frx":2BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.frx":2ED0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotoes 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1(0)"
      HotImageList    =   "ImageList1(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Confirmar"
            Object.ToolTipText     =   "Gravar registro - CTRL+G"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar registro - CTRL+C"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSDataListLib.DataCombo dtcCodigo_empresa 
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      BackColor       =   -2147483639
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgOpcoes_Importacao 
      Height          =   2655
      Left            =   90
      TabIndex        =   3
      Top             =   1740
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4683
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa Exportadora"
      Height          =   240
      Left            =   5970
      TabIndex        =   8
      Top             =   360
      Width           =   1845
   End
   Begin VB.Label lblCaminho 
      AutoSize        =   -1  'True
      Caption         =   "Caminho do Arquivo"
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   1050
      Width           =   1725
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Empresa [ F2 ]"
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   360
      Width           =   1290
   End
End
Attribute VB_Name = "frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' Módulo.................: Concentrador de Vendas                                         '
' Objetivo...............: Movimentação Gera Integração Frente de Loja - Importação       '
' Equipe Responsável.....: Only Tech Solutions                                            '
' Desenvolvedor..........: Rafael de Oliveira Gomes                                       '
' Data de Criação........: 23/12/2005                                                     '
' Desenvolvedor..........:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strID_Vendedor As String
Dim intConta_Nota_Entrada As Integer
Dim intConta_Item_Nota_Entrada As Integer
Public strSql As String
Dim strLinha_Arquivo As String
Dim strDescricao_Campo_Arquivo As String
Dim strCampo_Arquivo As String
Dim intConta_Linha_Grid As Integer
Dim intConta_Coluna_Grid As Integer
Dim intConta_Letras As Integer
Dim I As Integer
Dim strData_Emissao As String
Dim strHora_Emissao As String
Dim strCaminho_Cupom As String
Dim strCaminho_Cupom_FP As String
Dim strCaminho_Cupom_IT As String
Dim strCaminho_Backup As String
Dim log As New DLLSystemManager.log
Dim Conexao As New DLLConexao_Sistema.Conexao
Option Explicit

Private Sub cmdCaminho_Click()
    Unload frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao_Caminho
    frmAguarde.Show
    DoEvents
    frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao_Caminho.Show
    Unload frmAguarde
End Sub

Private Function Grava_Corpo_Nota_Saida_1()
    Dim strObservacao As String
    Dim strNumero_Pedido_ID As String
    Dim strID_Plano_pagamento  As String

    intConta_Coluna_Grid = 1
    intConta_Linha_Grid = 1

    Open strCaminho_Cupom For Input As #FreeFile
    
    strID_Plano_pagamento = Funcoes_Gerais.Localiza_ID("PKId_TBPlano_pagamento", "IXCodigo_TBPlano_pagamento", 999999, "TBPlano_pagamento", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcCodigo_empresa.BoundText)

    'INDICANDO O BANCO A CONECTAR-SE
    Conexao.Initial_Catalog = "BDRetaguarda"

    'ESTABELECENDO CONEXÃO COM O BANCO
    Conexao.Abrir_conexao ("Otica")

    'INDICA O INICIO DA TRANSAÇÃO JUNTO O BANCO
    Conexao.CNconexao.BeginTrans
        
    Do While Not EOF(1)
       Line Input #1, strLinha_Arquivo

       On Error GoTo Erro_transacao

       strSql = "INSERT INTO TBNota_saida (DFNumero_TBNota_saida," & _
                "DFData_Emissao_TBNota_saida," & _
                "DFHora_Emissao_TBNota_saida," & _
                "DFSerie_TBNota_saida," & _
                "DFTipo_emitente_TBNota_saida," & _
                "DFEmitente_TBNota_saida," & _
                "FKId_TBVendedor," & _
                "DFTotal_nota_TBNota_saida," & _
                "DFDesconto_especial_TBNota_saida," & _
                "DFDespesas_acessorias_TBNota_saida," & _
                "DFCancelado_TBNota_saida," & _
                "DFMotivo_cancelamento_TBNota_saida," & _
                "DFUsuario_cancelamento_TBNota_saida," & _
                "DFFaturista_TBNota_saida," & _
                "FKCodigo_TBEmpresa," & _
                "DFObservacao_TBNota_saida,"
                
       strSql = strSql & "FKId_TBPlano_pagamento," & _
                         "FKCodigo_TBTabela_preco," & _
                         "FKCodigo_TBTransportadora," & _
                         "DFTipo_operacao_TBNota_saida," & _
                         "DFTotal_itens_TBNota_saida," & _
                         "DFTotal_nota_tabela_TBNota_saida," & _
                         "DFData_Saida_TBNota_saida," & _
                         "DFDigitador_TBNota_saida," & _
                         "DFPrevisao_TBNota_saida," & _
                         "DFTotal_peso_liquido_TBNota_saida," & _
                         "DFTotal_peso_bruto_TBNota_saida," & _
                         "DFTotal_custo_medio_TBNota_saida," & _
                         "DFTotal_custo_real_TBNota_saida," & _
                         "DFTotal_custo_contabil_TBNota_saida," & _
                         "DFNumero_pedido_TBNota_saida) " & _
                         "VALUES ( "

       strLinha_Arquivo = Replace(Replace(strLinha_Arquivo, "'", ""), Chr$(9), "####")

       intConta_Letras = Len(strLinha_Arquivo)
              
       If strLinha_Arquivo <> "*" Then
          For I = 1 To intConta_Letras + 2
             If I = 1 Then
                strCampo_Arquivo = Empty
             Else
                strCampo_Arquivo = Mid(strLinha_Arquivo, I - 1, 1)
             
                If strCampo_Arquivo <> "#" And strCampo_Arquivo <> "" Then
                   strDescricao_Campo_Arquivo = strDescricao_Campo_Arquivo & Mid(strLinha_Arquivo, I - 1, 1)
                Else
                   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                   '        INICIO DA MONTAGEM DA STRING DE INSERÇÃO DOS VALORES DO CORPO DA NOTA SAIDA         '
                   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                   
                   'CONCATENANDO ID DA NOTA NA STRNUMERO_PEDIDO_ID
                   If intConta_Coluna_Grid = 1 And intConta_Linha_Grid <> 1 Then
                      strNumero_Pedido_ID = strDescricao_Campo_Arquivo
                      
                   'CONCATENANDO NUMERO DA NOTA (CUPOM) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 2 And intConta_Linha_Grid <> 1 Then
                      'VERIFICANDO SE O CAMPO NUMERO DA NOTA É INTEIRO
                      If IsNumeric(strDescricao_Campo_Arquivo) = True Then
                         strSql = strSql & "'" & CDbl(Right(strDescricao_Campo_Arquivo, 6)) & "',"
                         
                         strObservacao = strDescricao_Campo_Arquivo
                      Else
                         strSql = strSql & "'" & 999999 & "',"
                      End If
                      
                   'CONCATENANDO DATA EMISSAO DA NOTA (DATA) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 3 And intConta_Linha_Grid <> 1 Then
                      'BUSCANDO DATA DE EMISÃO NA STRING MONTADADA
                      If strDescricao_Campo_Arquivo = "NULL" Or strDescricao_Campo_Arquivo = "Null" Then
                         strData_Emissao = "19000101"
                      Else
                         strData_Emissao = Mid(strDescricao_Campo_Arquivo, 7, 4) & Mid(strDescricao_Campo_Arquivo, 1, 2) & Mid(strDescricao_Campo_Arquivo, 4, 2)
                         strHora_Emissao = Right(strDescricao_Campo_Arquivo, 8)
                      End If
                       
                      'VERIFICANDO SE A DATA FOI PEGA CORRETAMENTE
                      If IsNumeric(strData_Emissao) = True Then
                         strSql = strSql & "'" & strData_Emissao & "','" & strHora_Emissao & "',"
                      Else
                         strSql = strSql & "'" & 19000101 & "','" & strHora_Emissao & "',"
                      End If
                      
                   'CONCATENANDO SERIE DA NOTA (CAIXA) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 4 And intConta_Linha_Grid <> 1 Then
                      If strDescricao_Campo_Arquivo = "NULL" Or strDescricao_Campo_Arquivo = "Null" Then
                         strSql = strSql & "'" & 99 & "','" & 0 & "',"
                      Else
                         strSql = strSql & "'" & Left(strDescricao_Campo_Arquivo, 3) & "','" & 0 & "',"
                      End If
                   
                   'CONCATENANDO CLIENTE DA NOTA (CLIENTE) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 5 And intConta_Linha_Grid <> 1 Then
                      If strDescricao_Campo_Arquivo = "NULL" Or strDescricao_Campo_Arquivo = "Null" Or IsNumeric(strDescricao_Campo_Arquivo) = False Then
                         strSql = strSql & "'" & 999999 & "',"
                      Else
                         strSql = strSql & "'0','" & strDescricao_Campo_Arquivo & "',"
                      End If
                   
                   'CONCATENANDO VENDEDOR DA NOTA (VENDEDOR) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 6 And intConta_Linha_Grid <> 1 Then
                      strSql = strSql & "'" & strID_Vendedor & "',"
                   
                   'CONCATENANDO LOCAL DA NOTA NA STROBSERVACAO
                   ElseIf intConta_Coluna_Grid = 7 And intConta_Linha_Grid <> 1 Then
                      strObservacao = strObservacao & "#" & strDescricao_Campo_Arquivo
                   
                   'CONCATENANDO TOTAL DA NOTA (TOTAL) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 8 And intConta_Linha_Grid <> 1 Then
                      strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda(strDescricao_Campo_Arquivo) & ","
                   
                   'CONCATENANDO DESCONTO ESPECIAL DA NOTA (DESCONTO) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 9 And intConta_Linha_Grid <> 1 Then
                      strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda(strDescricao_Campo_Arquivo) & ","
                   
                   'CONCATENANDO DESPESA ACESSORIO DA NOTA (ACRESCIMO) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 10 And intConta_Linha_Grid <> 1 Then
                      strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda(strDescricao_Campo_Arquivo) & ","
                   
                   'CONCATENANDO CANCELAMENTO DA NOTA (CANCELOU) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 11 And intConta_Linha_Grid <> 1 Then
                      strSql = strSql & "'" & strDescricao_Campo_Arquivo & "',"
                      strSql = strSql & "'IMPORTAÇÃO FANTASTSOFT',"
                      strSql = strSql & "'FANTASTSOFT',"
                   
                   'CONCATENANDO FATURISTA DA NOTA (OPERADOR) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 12 And intConta_Linha_Grid <> 1 Then
                      strSql = strSql & "'" & strDescricao_Campo_Arquivo & "',"
                   
                   'CONCATENANDO EMPRESA DA NOTA (EMPRESA) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 13 And intConta_Linha_Grid <> 1 Then
                      strSql = strSql & "'" & dtcCodigo_empresa.BoundText & "',"
                   
                   'CONCATENANDO NUMEROSERIE DA NOTA NA STROBSERVACAO
                   ElseIf intConta_Coluna_Grid = 14 And intConta_Linha_Grid <> 1 Then
                      strObservacao = strObservacao & "#" & strDescricao_Campo_Arquivo
                   
                   'CONCATENANDO COO DA NOTA NA STROBSERVACAO
                   ElseIf intConta_Coluna_Grid = 15 And intConta_Linha_Grid <> 1 Then
                      strObservacao = strObservacao & "#" & strDescricao_Campo_Arquivo
                   
                   'CONCATENANDO CRZ DA NOTA NA STROBSERVACAO
                   ElseIf intConta_Coluna_Grid = 16 And intConta_Linha_Grid <> 1 Then
                      strObservacao = strObservacao & "#" & strDescricao_Campo_Arquivo
                   
                   'CONCATENANDO CRO DA NOTA NA STROBSERVACAO
                   ElseIf intConta_Coluna_Grid = 17 And intConta_Linha_Grid <> 1 Then
                      strObservacao = strObservacao & "#" & strDescricao_Campo_Arquivo
                   
                   'CONCATENANDO VENDABRUTA DA NOTA NA STROBSERVACAO
                   ElseIf intConta_Coluna_Grid = 18 And intConta_Linha_Grid <> 1 Then
                      strObservacao = strObservacao & "#" & strDescricao_Campo_Arquivo
                   
                   'CONCATENANDO GT DA NOTA NA STROBSERVACAO
                   ElseIf intConta_Coluna_Grid = 19 And intConta_Linha_Grid <> 1 Then
                      strObservacao = strObservacao & "#" & strDescricao_Campo_Arquivo
                      
                      strSql = strSql & "'" & strObservacao & "',"
                      
                      strSql = strSql & "'" & strID_Plano_pagamento & "'," & _
                                        "'999999'," & _
                                        "'999999'," & _
                                        "'1'," & _
                                        "0," & _
                                        "0," & _
                                        "'00:00:00'," & _
                                        "'FANTASTSOFT'," & _
                                        "'0'," & _
                                        "0," & _
                                        "0," & _
                                        "0," & _
                                        "0," & _
                                        "0," & _
                                        "'" & strNumero_Pedido_ID & "') "
                      
                      'GRAVANDO INCLUSAO NA TBNOTA_SAIDA
                      Conexao.CNconexao.Execute strSql
                   End If
                   
                   intConta_Coluna_Grid = intConta_Coluna_Grid + 1
                   
                   I = I + 3
                   
                   strDescricao_Campo_Arquivo = Empty
                End If
             End If
          Next I
          
          intConta_Linha_Grid = intConta_Linha_Grid + 1
          intConta_Coluna_Grid = 1
          
          strObservacao = Empty
          strDescricao_Campo_Arquivo = Empty
          strSql = Empty
       End If
    Loop
    
    Close #1
    
    'COMITANDO TRANSAÇÃO
    Conexao.CNconexao.CommitTrans
    
    'FECHANDO A CONEXÃO
    Conexao.CNconexao.Close
    
    hfgOpcoes_Importacao.Row = 1: hfgOpcoes_Importacao.Col = 2: hfgOpcoes_Importacao.Text = (intConta_Linha_Grid - 2) & " registro(s) importado(s)."

    Exit Function

Erro_transacao:
    'ROOLBACK NA TRANSAÇÃO
    Conexao.CNconexao.RollbackTrans
    
    'FECHANDO A CONEXÃO
    Conexao.CNconexao.Close

    Call erro.erro(Me, "Otica", "Gravar")
    Exit Function
End Function

Private Function Grava_Corpo_Nota_Saida_2()
    Dim strNumero_Pedido As String
    Dim strID_Nota_Saida As String
    Dim strCodigo_Plano_Pagamento As String

    intConta_Coluna_Grid = 1
    intConta_Linha_Grid = 1
    
    Open strCaminho_Cupom_FP For Input As #FreeFile

    'INDICANDO O BANCO A CONECTAR-SE
    Conexao.Initial_Catalog = "BDRetaguarda"

    'ESTABELECENDO CONEXÃO COM O BANCO
    Conexao.Abrir_conexao ("Otica")
        
    Do While Not EOF(1)
       Line Input #1, strLinha_Arquivo

       On Error GoTo Erro_transacao

       strSql = "UPDATE TBNota_saida " & _
                "SET FKId_TBPlano_pagamento = "

       strLinha_Arquivo = Replace(Replace(strLinha_Arquivo, "'", ""), Chr$(9), "####")
       
       intConta_Letras = Len(strLinha_Arquivo)
       
       If strLinha_Arquivo <> "*" Then
          For I = 1 To intConta_Letras + 2
             If I = 1 Then
                strCampo_Arquivo = Empty
             Else
                strCampo_Arquivo = Mid(strLinha_Arquivo, I - 1, 1)
             
                If strCampo_Arquivo <> "#" And strCampo_Arquivo <> "" Then
                   strDescricao_Campo_Arquivo = strDescricao_Campo_Arquivo & Mid(strLinha_Arquivo, I - 1, 1)
                Else
                   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                   '        INICIO DA MONTAGEM DA STRING DE ALTERAÇÃO DO PLANO PAGAMENTO DA  NOTA SAIDA         '
                   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                   
                   'DESCARTANDO ID CUPOMFP
                   If intConta_Coluna_Grid = 1 And intConta_Linha_Grid <> 1 Then

                   'CONCATENANDO ID NOTA SAIDA GERADO NO NUMERO PEDIDO
                   ElseIf intConta_Coluna_Grid = 2 And intConta_Linha_Grid <> 1 Then
                      'VERIFICANDO SE O CAMPO ID DA NOTA É INTEIRO
                      If IsNumeric(strDescricao_Campo_Arquivo) = True Then
                         strNumero_Pedido = strDescricao_Campo_Arquivo
                      Else
                         strNumero_Pedido = "999999"
                      End If
                      
                   'DESCARTANDO NUMERO CUPOMFP
                   ElseIf intConta_Coluna_Grid = 3 And intConta_Linha_Grid <> 1 Then
                   
                   'DESCARTANDO DATA CUPOMFP
                   ElseIf intConta_Coluna_Grid = 4 And intConta_Linha_Grid <> 1 Then
                   
                   'DESCARTANDO CAIXA CUPOMFP
                   ElseIf intConta_Coluna_Grid = 5 And intConta_Linha_Grid <> 1 Then
                   
                   'CONCATENANDO MODALIDADE NOTA SAIDA GERADO NO PLANO PAGAMENTO (FINALIZADORA)
                   ElseIf intConta_Coluna_Grid = 6 And intConta_Linha_Grid <> 1 Then
                      strCodigo_Plano_Pagamento = Funcoes_Gerais.Localiza_ID("PKID_TBPlano_pagamento", "IXCodigo_TBPlano_pagamento", strDescricao_Campo_Arquivo, "TBPlano_pagamento", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcCodigo_empresa.BoundText)
                      strID_Nota_Saida = Funcoes_Gerais.Localiza_ID("PKId_TBNota_saida", "DFNumero_pedido_TBNota_saida", strNumero_Pedido, "TBNota_Saida", "Otica", Me, "BDRetaguarda", "FKCodigo_TBEmpresa", dtcCodigo_empresa.BoundText)
                      
                      strSql = strSql & "'" & strCodigo_Plano_Pagamento & "' "
                      
                      strSql = strSql & "WHERE FKCodigo_TBEmpresa = '" & dtcCodigo_empresa.BoundText & "' " & _
                                        "AND PKId_TBNota_saida = '" & strID_Nota_Saida & "' "
                      
                      'INDICA O INICIO DA TRANSAÇÃO JUNTO O BANCO
                      Conexao.CNconexao.BeginTrans

                      'GRAVANDO ALTERAÇÃO NA TBNOTA_SAIDA
                      Conexao.CNconexao.Execute strSql
                      
                      'COMITANDO TRANSAÇÃO
                      Conexao.CNconexao.CommitTrans
                      
                   'DESCARTANDO VALOR CUPOMFP
                   ElseIf intConta_Coluna_Grid = 7 And intConta_Linha_Grid <> 1 Then

                   'DESCARTANDO EMISSAODOC CUPOMFP
                   ElseIf intConta_Coluna_Grid = 8 And intConta_Linha_Grid <> 1 Then

                   'DESCARTANDO PARCELAS CUPOMFP
                   ElseIf intConta_Coluna_Grid = 9 And intConta_Linha_Grid <> 1 Then

                   'DESCARTANDO TROCO CUPOMFP
                   ElseIf intConta_Coluna_Grid = 10 And intConta_Linha_Grid <> 1 Then

                   'DESCARTANDO EMPRESA CUPOMFP
                   ElseIf intConta_Coluna_Grid = 11 And intConta_Linha_Grid <> 1 Then

                   End If
                   
                   intConta_Coluna_Grid = intConta_Coluna_Grid + 1
                   
                   I = I + 3
                   
                   strDescricao_Campo_Arquivo = Empty
                End If
             End If
          Next I
                   
          intConta_Linha_Grid = intConta_Linha_Grid + 1
          intConta_Coluna_Grid = 1
          
          strNumero_Pedido = Empty
          strID_Nota_Saida = Empty
          strDescricao_Campo_Arquivo = Empty
          strSql = Empty
       End If
    Loop
        
    Close #1
    
    'FECHANDO A CONEXÃO
    Conexao.CNconexao.Close
    
    Exit Function
    
Erro_transacao:
    'ROOLBACK NA TRANSAÇÃO
    Conexao.CNconexao.RollbackTrans
    
    'DELETANDO CORPO DA NOTA SAIDA
    Conexao.CNconexao.Execute "DELETE FROM TBNota_saida WHERE DFData_Emissao_TBNota_saida = '" & strData_Emissao & "' AND DFHora_Emissao_TBNota_saida = '" & strHora_Emissao & "' "

    'FECHANDO A CONEXÃO
    Conexao.CNconexao.Close

    Call erro.erro(Me, "Otica", "Gravar")
    Exit Function
End Function

Private Function Grava_Itens_Nota_Saida()
    Dim rstCST As New ADODB.Recordset
    Dim strCst As String
    Dim strNumero_Pedido_Itens As String
    Dim strID_Nota_Saida_Itens As String
    Dim strID_Produto As String
    Dim strID_Cfop As String
    Dim dblPreco As Double
    Dim dblTotal As Double
    Dim dblTotal_Icms As Double
    Dim dblTotal_Praticado As Double
    Dim dblPercentual As Double
    Dim dblPercentual_Icms As Double
    Dim strCst1 As String
    Dim strCst2 As String
    Dim strUnidade As String
    Dim dblCusto_Real As Double
    Dim dblCusto_Medio As Double
    Dim dblCusto_Contabil As Double
    Dim dblPeso_Liquido As Double
    Dim dblPeso_Bruto As Double
    Dim intQtde As Integer
    Dim dblEstoque_Atual As Double
    Dim dblEstoque_Anterior As Double
    
    intConta_Coluna_Grid = 1
    intConta_Linha_Grid = 1

    Open strCaminho_Cupom_IT For Input As #FreeFile

    'INDICANDO O BANCO A CONECTAR-SE
    Conexao.Initial_Catalog = "BDRetaguarda"

    'ESTABELECENDO CONEXÃO COM O BANCO
    Conexao.Abrir_conexao ("Otica")
        
    Do While Not EOF(1)
       Line Input #1, strLinha_Arquivo

       On Error GoTo Erro_transacao
       
       strSql = Empty
       strSql = "INSERT INTO TBItens_nota_saida (FKId_TBNota_saida," & _
                "FKId_TBProduto," & _
                "DFQuantidade_TBItens_nota_saida," & _
                "DFPreco_praticado_TBItens_nota_saida," & _
                "DFPercentual_desconto_TBItens_nota_saida," & _
                "DFValor_total_item_TBItens_nota_saida," & _
                "DFPercentual_icms_TBItens_nota_saida," & _
                "FKId_TBVendedor,"
               
       strSql = strSql & "FKId_TBCfop," & _
                         "DFCst1_TBItens_nota_saida," & _
                         "DFCst2_TBItens_nota_saida," & _
                         "DFTipo_preco_TBItens_nota_saida," & _
                         "DFPreco_tabela_TBItens_nota_saida," & _
                         "DFValor_total_tabela_TBItens_nota_saida," & _
                         "DFValor_total_praticado_TBItens_nota_saida," & _
                         "DFValor_total_icms_TBItens_nota_saida," & _
                         "DFUnidade_TBItens_nota_saida," & _
                         "DFCusto_real_TBItens_nota_saida," & _
                         "DFCusto_contabil_TBItens_nota_saida," & _
                         "DFCusto_medio_TBItens_nota_saida," & _
                         "DFPeso_liquido_TBItens_nota_saida," & _
                         "DFPeso_bruto_TBItens_nota_saida," & _
                         "DFQuantidade_baixa_estoque_TBItens_nota_saida) " & _
                         "VALUES ( "

       strLinha_Arquivo = Replace(Replace(strLinha_Arquivo, "'", ""), Chr$(9), "####")
       
       intConta_Letras = Len(strLinha_Arquivo)
       
       If strLinha_Arquivo <> "*" Then
          For I = 1 To intConta_Letras + 2
             If I = 1 Then
                strCampo_Arquivo = Empty
             Else
                strCampo_Arquivo = Mid(strLinha_Arquivo, I - 1, 1)
             
                If strCampo_Arquivo <> "#" And strCampo_Arquivo <> "" Then
                   strDescricao_Campo_Arquivo = strDescricao_Campo_Arquivo & Mid(strLinha_Arquivo, I - 1, 1)
                Else
                   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                   '        INICIO DA MONTAGEM DA STRING DE INSERÇÃO DOS VALORES DOS ITENS NOTA SAIDA           '
                   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                   
                   'DESCARTANDO ID, NUMERO, DATA, CAIXA CUPOMIT CUPOMIT
                   If intConta_Coluna_Grid = 1 Or intConta_Coluna_Grid = 3 Or intConta_Coluna_Grid = 4 Or intConta_Coluna_Grid = 5 And intConta_Linha_Grid <> 1 Then

                   'CONCATENANDO ID NOTA SAIDA GERADO NO NUMERO PEDIDO
                   ElseIf intConta_Coluna_Grid = 2 And intConta_Linha_Grid <> 1 Then
                      'VERIFICANDO SE O CAMPO ID DA NOTA É INTEIRO
                      If IsNumeric(strDescricao_Campo_Arquivo) = True Then
                         strNumero_Pedido_Itens = strDescricao_Campo_Arquivo
                      Else
                         strNumero_Pedido_Itens = "999999"
                      End If
                      
                      strID_Nota_Saida_Itens = Funcoes_Gerais.Localiza_ID("PKId_TBNota_saida", "DFNumero_pedido_TBNota_saida", strNumero_Pedido_Itens, "TBNota_Saida", "Otica", Me, "BDRetaguarda", "FKCodigo_TBEmpresa", dtcCodigo_empresa.BoundText)

                      strSql = strSql & "'" & strID_Nota_Saida_Itens & "',"
                      
                   'CONCATENANDO PRODUTO ITEM (PRODUTO) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 6 And intConta_Linha_Grid <> 1 Then
                      strID_Produto = Funcoes_Gerais.Localiza_ID("PKID_TBProduto", "IXCodigo_TBProduto", strDescricao_Campo_Arquivo, "TBProduto", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcCodigo_empresa.BoundText)
                      
                      'BUSCANDO CST DO PRODUTO
                      strCst = "SELECT DFCusto_real_TBProduto," & _
                               "DFCusto_contabil_TBProduto," & _
                               "DFCusto_medio_TBProduto," & _
                               "DFPeso_liquido_TBProduto," & _
                               "DFPeso_bruto_TBProduto," & _
                               "DFUnidade_venda_TBProduto," & _
                               "DFCst1_TBProduto," & _
                               "DFCst2_TBProduto," & _
                               "DFEstoque_Atual_TBProduto " & _
                               "FROM TBProduto " & _
                               "WHERE PKID_TBProduto = '" & strID_Produto & "' "
                      
                      Movimentacoes.Select_geral strCst, "BDRetaguarda", rstCST, "Otica", Me
                      
                      If rstCST.RecordCount <> 0 Then
                         If Not IsNull(rstCST!DFCst1_TBProduto) Then
                            strCst1 = rstCST!DFCst1_TBProduto
                         Else
                            strCst1 = "0"
                         End If
                         
                         If Not IsNull(rstCST!DFCst2_TBProduto) Then
                            strCst2 = rstCST!DFCst2_TBProduto
                         Else
                            strCst2 = "00"
                         End If
                         
                         If Not IsNull(rstCST!DFUnidade_venda_TBProduto) Then
                            strUnidade = rstCST!DFUnidade_venda_TBProduto
                         Else
                            strUnidade = "UN"
                         End If
                      
                         If Not IsNull(rstCST!DFCusto_real_TBProduto) Then
                            dblCusto_Real = rstCST!DFCusto_real_TBProduto
                         Else
                            dblCusto_Real = "0"
                         End If
                         
                         If Not IsNull(rstCST!DFCusto_contabil_TBProduto) Then
                            dblCusto_Contabil = rstCST!DFCusto_contabil_TBProduto
                         Else
                            dblCusto_Contabil = "0"
                         End If
                         
                         If Not IsNull(rstCST!DFCusto_medio_TBProduto) Then
                            dblCusto_Medio = rstCST!DFCusto_medio_TBProduto
                         Else
                            dblCusto_Medio = "0"
                         End If
                         
                         If Not IsNull(rstCST!DFPeso_liquido_TBProduto) Then
                            dblPeso_Liquido = rstCST!DFPeso_liquido_TBProduto
                         Else
                            dblPeso_Liquido = "0"
                         End If
                         
                         If Not IsNull(rstCST!DFPeso_bruto_TBProduto) Then
                            dblPeso_Bruto = rstCST!DFPeso_bruto_TBProduto
                         Else
                            dblPeso_Bruto = "0"
                         End If
                         
                         If Not IsNull(rstCST!DFEstoque_Atual_TBProduto) Then
                            dblEstoque_Anterior = rstCST!DFEstoque_Atual_TBProduto
                         Else
                            dblEstoque_Anterior = "0"
                         End If
                         
                         If IsNumeric(dblEstoque_Anterior) = True Then
                            dblEstoque_Atual = CDbl(dblEstoque_Anterior) - CDbl(intQtde)
                         Else
                            dblEstoque_Atual = 0
                         End If
                      End If
                      
                      Set rstCST = Nothing
                      
                      strSql = strSql & "'" & strID_Produto & "',"
                      
                   'DESCARTANDO CODIGOEAN CUPOMIT
                   ElseIf intConta_Coluna_Grid = 7 And intConta_Linha_Grid <> 1 Then
                   
                   'CONCATENANDO QTDE ITEM (QUANTIDADE) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 8 And intConta_Linha_Grid <> 1 Then
                      intQtde = strDescricao_Campo_Arquivo
                      
                      strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda(intQtde) & ","
                      
                   'CONCATENANDO PREÇO PRATICADO ITEM (PRECO) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 9 And intConta_Linha_Grid <> 1 Then
                      dblPreco = strDescricao_Campo_Arquivo
                      
                      strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda(dblPreco) & ","
                      
                   'CONCATENANDO PERCENTUAL ITEM (DESCONTO) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 10 And intConta_Linha_Grid <> 1 Then
                      dblPercentual = (CDbl(strDescricao_Campo_Arquivo) * 100) / dblPreco
                      
                      strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda(dblPercentual) & ","
                      
                   'CONCATENANDO TOTAL ITEM (TOTAL) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 11 And intConta_Linha_Grid <> 1 Then
                      dblTotal = strDescricao_Campo_Arquivo
                   
                      strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda(dblTotal) & ","
                   
                   'DESCARTANDO CANCELOU, EMPRESA, TIPOTRIBUTACAO CUPOMIT
                   ElseIf intConta_Coluna_Grid = 12 Or intConta_Coluna_Grid = 13 Or intConta_Coluna_Grid = 14 And intConta_Linha_Grid <> 1 Then

                   'CONCATENANDO PERCENTUAL ICMS ITEM NA STRSQL
                   ElseIf intConta_Coluna_Grid = 15 And intConta_Linha_Grid <> 1 Then
                      If IsNumeric(strDescricao_Campo_Arquivo) = True Then
                         strSql = strSql & "" & Funcoes_Gerais.Grava_Moeda(strDescricao_Campo_Arquivo) & ","
                      Else
                         strSql = strSql & "" & 0 & ","
                      End If
                      
                   'DESCARTANDO ITEM, DESCADICIONAL1, DESCADICIONAL2, DESCADICIONAL3, DESCADICIONAL4, DESCADICIONAL5 CUPOMIT CUPOMIT
                   ElseIf intConta_Coluna_Grid = 16 Or intConta_Coluna_Grid = 17 Or intConta_Coluna_Grid = 18 Or intConta_Coluna_Grid = 19 Or intConta_Coluna_Grid = 20 Or intConta_Coluna_Grid = 21 And intConta_Linha_Grid <> 1 Then
                   
                   'CONCATENANDO VENDEDOR DA NOTA (VENDEDOR) NA STRSQL
                   ElseIf intConta_Coluna_Grid = 22 And intConta_Linha_Grid <> 1 Then
                      strSql = strSql & "'" & strID_Vendedor & "',"
                      
                      If IsNumeric(strDescricao_Campo_Arquivo) = True Then
                         strID_Cfop = Funcoes_Gerais.Localiza_ID("PKId_TBCfop", "DFCodigo_TBCfop", strDescricao_Campo_Arquivo, "TBCfop", "Otica", Me, "BDRetaguarda")
                      Else
                         strID_Cfop = Funcoes_Gerais.Localiza_ID("PKId_TBCfop", "DFCodigo_TBCfop", 999999, "TBCfop", "Otica", Me, "BDRetaguarda")
                      End If
                      
                      strData_Emissao = Format(Date, "YYYYMMDD")
                      strHora_Emissao = Format(Now, "HH:MM:SS")
                      
                      dblTotal_Icms = dblTotal * dblPercentual_Icms
                      dblTotal_Praticado = intQtde * dblPreco
                      
                      dblCusto_Real = intQtde * dblCusto_Real
                      dblCusto_Contabil = intQtde * dblCusto_Contabil
                      dblCusto_Medio = intQtde * dblCusto_Medio
                      
                      dblPeso_Liquido = intQtde * dblPeso_Liquido
                      dblPeso_Bruto = intQtde * dblPeso_Bruto
                      
                      strSql = strSql & "'" & strID_Cfop & "'," & _
                                        "'" & strCst1 & "'," & _
                                        "'" & strCst2 & "'," & _
                                        "'1'," & _
                                        "0," & _
                                        "0," & _
                                        "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Praticado) & "," & _
                                        "" & Funcoes_Gerais.Grava_Moeda(dblTotal_Icms) & "," & _
                                        "'" & Trim(strUnidade) & "'," & _
                                        "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Real) & "," & _
                                        "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Contabil) & "," & _
                                        "" & Funcoes_Gerais.Grava_Moeda(dblCusto_Medio) & "," & _
                                        "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Liquido) & "," & _
                                        "" & Funcoes_Gerais.Grava_Moeda(dblPeso_Bruto) & "," & _
                                        "0) "
                      
                      'INDICA O INICIO DA TRANSAÇÃO JUNTO O BANCO
                      Conexao.CNconexao.BeginTrans
                      
                      'GRAVANDO INCLUSAO NA TBITENS_NOTA_SAIDA
                      Conexao.CNconexao.Execute strSql
                      
                      'GRAVANDO INCLUSAO NA TBCFOP_NOTA_SAIDA
                      Conexao.CNconexao.Execute "INSERT INTO TBCfop_nota_saida(FKId_TBCfop,FKId_TBnota_saida) VALUES ( '" & strID_Cfop & "','" & strID_Nota_Saida_Itens & "')"
                      
                      strSql = Empty
                      strSql = "INSERT INTO TBOcorrencias_produto " & _
                               "(FKId_TBProduto," & _
                               "DFQuantidade_movimento_TBOcorrencia_produto," & _
                               "DFEstoque_anterior_TBOcorrencia_produto," & _
                               "DFEstoque_atual_TBOcorrencia_produto," & _
                               "DFData_movimento_TBOcorrencia_produto," & _
                               "DFHora_movimento_TBOcorrencia_produto," & _
                               "DFUsuario_TBOcorrencia_produto," & _
                               "DFPrograma_TBOcorrencia_produto," & _
                               "DFObservacao_TBOcorrencia_produto) " & _
                               "VALUES ('" & strID_Produto & "'," & _
                               "" & CDbl(intQtde) & "," & _
                               "" & dblEstoque_Anterior & "," & _
                               "" & dblEstoque_Atual & "," & _
                               "'" & strData_Emissao & "'," & _
                               "'" & strHora_Emissao & "'," & _
                               "'" & MDIPrincipal.OCXUsuario.Nome & "'," & _
                               "'MOV.GINTEGRAÇÃO FLOJA - IMP'," & _
                               "'GRAVADO APARTIR DO ARQUIVO GERADO PELA FANTASTSOFT')"
                      
                      'GRAVANDO INCLUSAO NA TBOCORRENCIAS_PRODUTO
                      Conexao.CNconexao.Execute strSql
                      
                      'GRAVANDO INCLUSAO NA TBOCORRENCIAS_PRODUTO
                      Conexao.CNconexao.Execute "UPDATE TBProduto SET DFEstoque_Atual_TBProduto = " & Funcoes_Gerais.Grava_Moeda(dblEstoque_Atual) & " WHERE PKId_TBProduto = '" & strID_Produto & "' "
                      
                      'COMITANDO TRANSAÇÃO
                      Conexao.CNconexao.CommitTrans
                      
                      dblCusto_Real = Empty
                      dblCusto_Contabil = Empty
                      dblCusto_Medio = Empty
                      dblPeso_Liquido = Empty
                      dblPeso_Bruto = Empty
                   End If
                   
                   intConta_Coluna_Grid = intConta_Coluna_Grid + 1
                   
                   I = I + 3

                   strDescricao_Campo_Arquivo = Empty
                End If
             End If
          Next I
          
          intConta_Linha_Grid = intConta_Linha_Grid + 1
          intConta_Coluna_Grid = 1
          
          strDescricao_Campo_Arquivo = Empty
          strSql = Empty
       End If
    Loop
    
    Close #1
    
    'FECHANDO A CONEXÃO
    Conexao.CNconexao.Close
        
    hfgOpcoes_Importacao.Row = 2: hfgOpcoes_Importacao.Col = 2: hfgOpcoes_Importacao.Text = (intConta_Linha_Grid - 2) & " registro(s) importado(s)."
    
    Exit Function
    
Erro_transacao:
    'ROOLBACK NA TRANSAÇÃO
    Conexao.CNconexao.RollbackTrans
    
    'DELETANDO CORPO DA NOTA SAIDA
    Conexao.CNconexao.Execute "DELETE FROM TBNota_saida WHERE DFData_Emissao_TBNota_saida = '" & strData_Emissao & "' AND DFHora_Emissao_TBNota_saida = '" & strHora_Emissao & "' "
    
    'FECHANDO A CONEXÃO
    Conexao.CNconexao.Close

    Call erro.erro(Me, "Otica", "Gravar")
    Exit Function
End Function

Private Sub dtcCodigo_empresa_LostFocus()
    dtcCodigo_empresa.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'HABILITA A TROCA DE CAMPOS PELO ENTER
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'TECLAS DE ATALHO DA TOOLBAR
    Select Case Shift
           Case 2
                Select Case KeyCode
                       Case 71: Call Gravar   'CTRL+G
                       Case 67: Call Cancelar 'CTRL+C
                       Case 83: Unload Me     'CTRL+S
                End Select
    End Select

    If KeyCode = "113" Then Movimentacoes.Verifica_Acesso_Usuario dtcCodigo_empresa, "Otica", "BDRetaguarda", Me
End Sub

Private Sub tlbBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
           Case 1: Call Gravar
           Case 2: Call Cancelar
           Case 4: Unload Me
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo erro
    
    'INFORMAÇÕES CONSTANTES PARA O LOG
    log.Usuario = MDIPrincipal.OCXUsuario.Nome
    log.Programa = "Movimentacao Gera Integração Frente de Loja"
    log.Estacao = MDIPrincipal.OCXUsuario.Estacao
    
    'INFORMAÇÕES VARIAVEIS PARA O LOG
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    If MDIPrincipal.booDesign_time = False Then
       Call Movimentacoes.Acessibilidade_inicio_relatorios(Me.Caption, MDIPrincipal.OCXUsuario, Me, "Otica", "BDRetaguarda")
    End If
    
    log.Descricao = "Inicializando a Movimentacao Gera Integração Frente de Loja"
    'GRAVANDO O LOG
    log.Gravar_log "Otica", Me

    'MONTANDO DATA COMBO DA EMPRESA
    strSql = "SELECT TBEmpresa.PKCodigo_TBEmpresa,DFRazao_Social_TBEmpresa FROM TBEmpresa"
    Movimentacoes.Movimenta_DataCombo "PKCodigo_TBEmpresa", "DFRazao_Social_TBEmpresa", dtcCodigo_empresa, strSql, "BDRetaguarda", "Otica", Me

    dtcCodigo_empresa.BoundText = MDIPrincipal.OCXUsuario.Empresa

    'VERIFICANDO EXISTENCIA DOS DIRETÓRIOS DE DESTINO
    Dim strVerifica_Diretorio As String

    strVerifica_Diretorio = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
    strVerifica_Diretorio = Left(strVerifica_Diretorio, CDbl(Len(strVerifica_Diretorio) - 3)) & "INTEGRAÇÃO"

    If Dir(strVerifica_Diretorio, vbDirectory) = "" Then
       MkDir strVerifica_Diretorio
    End If

    strVerifica_Diretorio = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
    strVerifica_Diretorio = Left(strVerifica_Diretorio, CDbl(Len(strVerifica_Diretorio) - 3)) & "INTEGRAÇÃO\INTEGRADOS"

    If Dir$(strVerifica_Diretorio, vbDirectory) = Empty Then
       MkDir strVerifica_Diretorio
    End If

    cbbEmpresa_Exportadora.Clear
    cbbEmpresa_Exportadora.AddItem ("Only Tech")
    cbbEmpresa_Exportadora.AddItem ("Fantastsoft")
    
    Call Monta_Opcoes_Importacao
    
    Exit Sub
erro:
    Call erro.erro(Me, "Otica", "Load")
    Exit Sub
End Sub

    Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro
    
    log.Evento = "Unload"
    log.Descricao = "Finalizando a Movimentação Gera Integração Frente de Loja - Importação"
    log.Hora = Format(Now, "hh:mm:ss")
    
    'GRAVANDO O LOG
    log.Gravar_log "Otica", Me
    
    Exit Sub
erro:
    Call erro.erro(Me, "Otica", "Unload")
    Exit Sub
End Sub

Function Gravar()
    On Error GoTo erro

    hfgOpcoes_Importacao.Row = 1: hfgOpcoes_Importacao.Col = 2: hfgOpcoes_Importacao.Text = Empty
    hfgOpcoes_Importacao.Row = 2: hfgOpcoes_Importacao.Col = 2: hfgOpcoes_Importacao.Text = Empty

    'VERIFICANDO SE O CAMINHO DO ARQUIVO FOI CARREGADO
    If txtCaminho.Text = Empty Then
       MsgBox "Caminho para importação do arquivo não informado. Verifique!", vbInformation, "Only Tech"
       txtCaminho.SetFocus
       Exit Function
    End If
    
    'VERIFICANDO SE A EMPRESA EXPORTADORA DO ARQUIVO FOI CARREGADA
    If cbbEmpresa_Exportadora.Text = Empty Then
       MsgBox "Empresa exportadora do arquivo não informado. Verifique!", vbInformation, "Only Tech"
       cbbEmpresa_Exportadora.SetFocus
       Exit Function
    End If
       
    frmAguarde.Show
    DoEvents
    
    strCaminho_Cupom = Empty
    strCaminho_Cupom_FP = Empty
    strCaminho_Cupom_IT = Empty
    strCaminho_Backup = Empty
        
    strCaminho_Backup = Funcoes_Gerais.Abrir_figura_registro("Otica", Me)
    strCaminho_Backup = Left(strCaminho_Backup, CDbl(Len(strCaminho_Backup) - 3)) & "INTEGRAÇÃO\INTEGRADOS"
    
    strCaminho_Cupom = txtCaminho.Text & "\CUPOM.DAT"
    strCaminho_Cupom_FP = txtCaminho.Text & "\CUPOMFP.DAT"
    strCaminho_Cupom_IT = txtCaminho.Text & "\CUPOMIT.DAT"
    
    'VERIFICANDO A EXISTENCIA DO ARQUIVO COM INFORMAÇÕES DO CORPO DA NOTA SAIDA
    If Dir$(strCaminho_Cupom) = Empty Then
       MsgBox "Arquivo CUPOM.DAT inexistente na pasta selecionada. Verifique!", vbInformation, "Only Tech"
       Unload frmAguarde
       Exit Function
    End If

    'VERIFICANDO A EXISTENCIA DO ARQUIVO COM INFORMAÇÕES RESTANTES DO CORPO DA NOTA SAIDA
    If Dir$(strCaminho_Cupom_FP) = Empty Then
       MsgBox "Arquivo CUPOMFP.DAT inexistente na pasta selecionada. Verifique!", vbInformation, "Only Tech"
       Unload frmAguarde
       Exit Function
    End If

    'VERIFICANDO A EXISTENCIA DO ARQUIVO COM INFORMAÇÕES DOS ITENS DA NOTA SAIDA
    If Dir$(strCaminho_Cupom_IT) = Empty Then
       MsgBox "Arquivo CUPOMIT.DAT inexistente na pasta selecionada. Verifique!", vbInformation, "Only Tech"
       Unload frmAguarde
       Exit Function
    End If

    'COPIANDO ARQUIVOS PARA PASTA ONLYTECH INTEGRADOS
    FileCopy strCaminho_Cupom, (strCaminho_Backup & "\CUPOM.DAT")
    FileCopy strCaminho_Cupom_FP, (strCaminho_Backup & "\CUPOMFP.DAT")
    FileCopy strCaminho_Cupom_IT, (strCaminho_Backup & "\CUPOMIT.DAT")
    
    strID_Vendedor = Funcoes_Gerais.Localiza_ID("PKId_TBVendedor", "IXCodigo_TBVendedor", 999999, "TBVendedor", "Otica", Me, "BDRetaguarda", "IXCodigo_TBEmpresa", dtcCodigo_empresa.BoundText)
    
    Call Grava_Corpo_Nota_Saida_1
    Call Grava_Corpo_Nota_Saida_2
    Call Grava_Itens_Nota_Saida
    
    Unload frmAguarde
    
    txtCaminho.Text = Empty
    cbbEmpresa_Exportadora.Text = Empty
    cbbEmpresa_Exportadora.SetFocus
    
    Exit Function
erro:
    Call erro.erro(Me, "Otica", "Gravar")
    
    Exit Function
End Function

Private Function Cancelar()
    On Error GoTo erro
    
    txtCaminho.Text = Empty
    cbbEmpresa_Exportadora.Text = Empty
    
    hfgOpcoes_Importacao.Row = 1: hfgOpcoes_Importacao.Col = 2: hfgOpcoes_Importacao.Text = Empty
    hfgOpcoes_Importacao.Row = 2: hfgOpcoes_Importacao.Col = 2: hfgOpcoes_Importacao.Text = Empty
   
    txtCaminho.SetFocus
    
    Exit Function
erro:
    Call erro.erro(Me, "Otica", "Cancelar")
    Exit Function
End Function

Private Sub txtCaminho_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Function Monta_Opcoes_Importacao()
    Dim I As Integer
    
    'DEFININDO NÚMERO DE LINHAS E COLUNAS
    hfgOpcoes_Importacao.Col = 0
    hfgOpcoes_Importacao.Cols = 3
    hfgOpcoes_Importacao.Rows = 3
    
    'DEFININDO TAMANHO DAS COLUNAS
    hfgOpcoes_Importacao.ColWidth(0) = 500
    hfgOpcoes_Importacao.ColWidth(1) = 4000
    hfgOpcoes_Importacao.ColWidth(2) = 3500
    
    'DEFININDO ALINHAMENTO DO CABEÇALHO
    hfgOpcoes_Importacao.ColAlignmentFixed(1) = 5
    hfgOpcoes_Importacao.ColAlignmentFixed(2) = 5
    
    'DEFININDO ALINHAMENTO DA COLUNA SIM DO GRID
    hfgOpcoes_Importacao.ColAlignment(2) = 5
    
    'COLORINDO COLUNA DE INDICE
    I = hfgOpcoes_Importacao.Rows - 1
    
    Do While I <> 0
       
       hfgOpcoes_Importacao.Col = 0: hfgOpcoes_Importacao.Row = I
       
       hfgOpcoes_Importacao.CellBackColor = &H80FFFF
       
       I = I - 1
    Loop
    
    'DEFININDO CABEÇALHO PADRÃO
    hfgOpcoes_Importacao.Row = 0
    hfgOpcoes_Importacao.Col = 2: hfgOpcoes_Importacao.CellBackColor = &H8000000F: hfgOpcoes_Importacao.CellFontBold = True: hfgOpcoes_Importacao.CellFontSize = 10
    hfgOpcoes_Importacao.Col = 1: hfgOpcoes_Importacao.CellBackColor = &H8000000F: hfgOpcoes_Importacao.CellFontBold = True: hfgOpcoes_Importacao.CellFontSize = 10
    hfgOpcoes_Importacao.Col = 0: hfgOpcoes_Importacao.CellBackColor = &H8000000F: hfgOpcoes_Importacao.CellFontBold = True: hfgOpcoes_Importacao.CellFontSize = 10
    hfgOpcoes_Importacao.TextArray((hfgOpcoes_Importacao.Row * hfgOpcoes_Importacao.Cols + hfgOpcoes_Importacao.Col + 1)) = "PROGRAMAS PARA IMPORTAÇÃO"
    hfgOpcoes_Importacao.TextArray((hfgOpcoes_Importacao.Row * hfgOpcoes_Importacao.Cols + hfgOpcoes_Importacao.Col + 2)) = "RESUMO INTEGRAÇÃO"
    
    hfgOpcoes_Importacao.Row = 1
    hfgOpcoes_Importacao.TextArray((hfgOpcoes_Importacao.Row * hfgOpcoes_Importacao.Cols + hfgOpcoes_Importacao.Col + 0)) = "1"
    hfgOpcoes_Importacao.TextArray((hfgOpcoes_Importacao.Row * hfgOpcoes_Importacao.Cols + hfgOpcoes_Importacao.Col + 1)) = "NOTA SAÍDA"
    
    hfgOpcoes_Importacao.Row = 2
    hfgOpcoes_Importacao.TextArray((hfgOpcoes_Importacao.Row * hfgOpcoes_Importacao.Cols + hfgOpcoes_Importacao.Col + 0)) = "2"
    hfgOpcoes_Importacao.TextArray((hfgOpcoes_Importacao.Row * hfgOpcoes_Importacao.Cols + hfgOpcoes_Importacao.Col + 1)) = "ITENS NOTA SAÍDA"
End Function
