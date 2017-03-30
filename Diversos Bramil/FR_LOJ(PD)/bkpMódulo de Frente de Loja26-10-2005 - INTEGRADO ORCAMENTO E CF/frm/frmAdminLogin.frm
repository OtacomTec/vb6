VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAdminLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   5130
   ClientTop       =   3930
   ClientWidth     =   5205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAdminLogin.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wskLan 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2610
      Width           =   1875
   End
   Begin VB.TextBox txtOperador 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   " "
      Top             =   2040
      Width           =   1875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "Versão: 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4140
      TabIndex        =   4
      Top             =   2700
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Operador:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Dim rstOperador As New ADODB.Recordset
Dim rstDia As New ADODB.Recordset
Dim datValidade_usuario As Date
Public intNumero_pdv As Integer
Public intImpressoes_suportadas As Integer
Dim booIntegracao_online As Boolean
Dim lngCodigo_operador As Long
Dim booDia_anterior  As Boolean
Public dtpDia_operacao As Date
Dim booLeitor_serial As Boolean
Dim strPorta_leitor As String
Dim intTipo_imp_orcamento As Integer
Public booGaveta_integrada As Boolean
Dim rstParametro_ecf As New ADODB.Recordset
Dim intEmpresa As Integer
Dim strCaminho_impComum As String
Dim booFinaliza_direto As Boolean
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
   End If
   If KeyAscii = 27 Then
        End
   End If
End Sub

Private Sub Form_Load()
    
    Dim rstPDV As New ADODB.Recordset
    Dim strIP As String
    Dim rstIntegracao_online As New ADODB.Recordset
    
    'Valida Travamento
    If Funcoes_Gerais.Valida_Trava_Sistema(Me) = True Then
       MsgBox "Contacte Only Tech - Result : TFPG", vbCritical, "Only Tech"
       End
    End If
     
    'Pegando o endereço da maquina e verificando à qual pdv se refere
    strIP = Me.wskLan.LocalIP
    
    'Carregando a combo de pdv
    strSql = Empty
    strSql = "SELECT * FROM TBPDV WHERE DFEndereco_ip_TBPdv = '" & strIP & "'"
    
    Movimentacoes.Select_geral strSql, "BDPDV", rstPDV, "PDV", Me
    
    If rstPDV.BOF = True And rstPDV.EOF = True Then
       MsgBox "Ponto de venda não cadastrado!Verifique.", vbCritical, "Only Tech"
       Set rstPDV = Nothing
       End
    End If
    
    If rstPDV!DFImpressoes_suportadas_TBPdv = "" Or IsNull(rstPDV!DFImpressoes_suportadas_TBPdv) Then
       MsgBox "Tipo de impressoras suportado não informado!Verifique.", vbCritical, "Only Tech"
       Set rstPDV = Nothing
       End
    End If
    
    If rstPDV!DFTipo_impressora_orcamento_balcao_TBpdv = "" Or IsNull(rstPDV!DFTipo_impressora_orcamento_balcao_TBpdv) Then
       MsgBox "Tipo de impressoras orçamento não informado!Verifique.", vbCritical, "Only Tech"
       Set rstPDV = Nothing
       End
    End If
    
    If rstPDV!DFGaveta_integrada_TBPdv = "" Or IsNull(rstPDV!DFGaveta_integrada_TBPdv) Then
       MsgBox "Integração com a gaveta não informada!Verifique.", vbCritical, "Only Tech"
       Set rstPDV = Nothing
       End
    End If
    
    If rstPDV!IXCodigo_TBEmpresa = "" Or IsNull(rstPDV!IXCodigo_TBEmpresa) Then
       MsgBox "Empresa deste PDV não informada!Verifique.", vbCritical, "Only Tech"
       Set rstPDV = Nothing
       End
    End If
    
    If rstPDV!DFTipo_impressora_orcamento_balcao_TBpdv = 1 Then
       If rstPDV!DFCaminho_Impressora_Comum = "" Or IsNull(rstPDV!DFCaminho_Impressora_Comum) Then
          MsgBox "Caminho de imp. comum não informado!Verifique.", vbCritical, "Only Tech"
          Set rstPDV = Nothing
          End
       Else
          strCaminho_impComum = rstPDV!DFCaminho_Impressora_Comum
       End If
    End If
    
    If IsNull(rstPDV!DFFinaliza_venda_direto) = True Then
       booFinaliza_direto = False
    Else
       booFinaliza_direto = rstPDV!DFFinaliza_venda_direto
    End If
    
    intNumero_pdv = rstPDV!PKCodigo_TBPdv
    intImpressoes_suportadas = rstPDV!DFImpressoes_suportadas_TBPdv
    intTipo_imp_orcamento = rstPDV!DFTipo_impressora_orcamento_balcao_TBpdv
    booGaveta_integrada = rstPDV!DFGaveta_integrada_TBPdv
    intEmpresa = rstPDV!IXCodigo_TBEmpresa
    
    'Leitor
    'Verifica se no parametro esta marcado o uso de leitor serial
    If IsNull(rstPDV!DFLeitor_Serial_integrado) = False Or rstPDV!DFLeitor_Serial_integrado = "" Then
       booLeitor_serial = False
    Else
       booLeitor_serial = rstPDV!DFLeitor_Serial_integrado
       'Verifica se há porta inf. do coletor.
       If IsNull(rstPDV!DFPorta_com_leitor_serial) Or rstPDV!DFPorta_com_leitor_serial = "" Then
          MsgBox "Leitor serial suportado,mas porta com não informada!Verifique.", vbCritical, "Only Tech"
          Set rstPDV = Nothing
          End
       Else
          strPorta_leitor = rstPDV!DFPorta_com_leitor_serial
       End If
    End If
    
    strSql = Empty
    strSql = "SELECT DFIntegracao_online_pdv_retaguarda_TBParametros_ecf FROM TBParametros_ecf WHERE FKCodigo_TBEmpresa = " & intEmpresa & ""
    Movimentacoes.Select_geral strSql, "BDPDV", rstIntegracao_online, "PDV", Me
    
    If rstIntegracao_online.BOF = True And rstIntegracao_online.EOF = True Then
       MsgBox "Integração com retaguarda não definida.Verifique.", vbCritical, "Only Tech"
       Set rstIntegracao_online = Nothing
       Set rstPDV = Nothing
       End
    End If
    
    booIntegracao_online = rstIntegracao_online!DFIntegracao_online_pdv_retaguarda_TBParametros_ecf
    
    'Parametros do ECF
    strSql = Empty
    strSql = "SELECT * FROM TBPARAMETROS_ECF WHERE FKCodigo_TBEmpresa = " & intEmpresa & ""
    
    If booIntegracao_online = True Then
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstParametro_ecf, "Otica", Me
    Else
       Movimentacoes.Select_geral strSql, "BDPDV", rstParametro_ecf, "PDV", Me
    End If
        
    'Verificação do checklist da base do pdv
    Call CheckList_pdv_inicio
    
    Set rstIntegracao_online = Nothing
    Set rstPDV = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstOperador = Nothing
End Sub

Private Sub txtOperador_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_LostFocus()

    On Error GoTo Erro
    
    Me.txtOperador.Text = UCase(Me.txtOperador)
    
    strSql = Empty
    strSql = "SELECT * FROM TBOperadores_ecf WHERE DFNome_TBOperadores_ecf = '" & Me.txtOperador.Text & "' AND FKCodigo_TBEmpresa = " & intEmpresa & ""
    
    If booIntegracao_online = True Then
        Movimentacoes.Select_geral strSql, "BDRetaguarda", rstOperador, "Otica", Me
    Else
        Movimentacoes.Select_geral strSql, "BDPDV", rstOperador, "PDV", Me
    End If
       
    If rstOperador.BOF = True And rstOperador.EOF = True Then
       MsgBox "Operador não cadastrado.Verifique!", vbCritical, "Only Tech"
       txtOperador.Text = ""
       txtOperador.SetFocus
       Set rstOperador = Nothing
       Exit Sub
    End If
    
    lngCodigo_operador = rstOperador!PKCodigo_TBOperadores_ecf
    
    Set rstOperador = Nothing
    
    Me.txtSenha.SetFocus
    
    Exit Sub
    
Erro:
    If Err.Number = -2147220503 Then
       MsgBox "Fonte de dados não encontrada!", vbCritical, "Only Tech"
       Shell App.Path & "Configurador de Sistemas.exe", vbNormalFocus
       End
       Exit Sub
    Else
       If booIntegracao_online = True Then
          Call Erro.Erro(Me, "Otica", "Load", Err.Number)
       Else
          Call Erro.Erro(Me, "PDV", "Load", Err.Number)
       End If
       
       Exit Sub
    End If
    
End Sub

Private Sub txtSenha_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtSenha_LostFocus()

    On Error GoTo Erro
    
    If Me.txtSenha.Text <> "" Then
          Dim intRetorno As Integer
          
          'Verificação de dia de Operação
          strSql = Empty
          strSql = "SELECT MAX(DFDia_TBDia_operacao_pdv) as Ultimo_dia_operacao FROM TBDia_operacao_pdv WHERE DFStatus_dia_TBDia_Operacao_pdv = 'A' AND DFNumero_pdv_TBDia_operacao_pdv = " & intNumero_pdv & " AND IXCodigo_TBEmpresa = " & intEmpresa & ""
          
          If booIntegracao_online = True Then
             Movimentacoes.Select_geral strSql, "BDRetaguarda", rstDia, "Otica", Me
          Else
             Movimentacoes.Select_geral strSql, "BDPDV", rstDia, "PDV", Me
          End If
          
          strSql = Empty
          strSql = "SELECT * FROM TBOperadores_ecf WHERE PKCodigo_TBOperadores_ecf = " & lngCodigo_operador & " AND DFSenha_TBOperadores_ecf = '" & Me.txtSenha.Text & "' AND FKCodigo_TBEmpresa = " & intEmpresa & ""
          
          If booIntegracao_online = True Then
             Movimentacoes.Select_geral strSql, "BDRetaguarda", rstOperador, "Otica", Me
          Else
             Movimentacoes.Select_geral strSql, "BDPDV", rstOperador, "PDV", Me
          End If
         
          If rstOperador.BOF = True And rstOperador.EOF = True Then
             MsgBox "Senha não cadastrada.Verifique!", vbCritical, "Only Tech"
             Set rstOperador = Nothing
             Set rstDia = Nothing
             txtSenha.Text = Empty
             txtSenha.SetFocus
             Exit Sub
          End If
          
          If IsNull(rstDia!Ultimo_dia_operacao) Then
             Dim strCampos As String
             Dim strValores As String
             
             dtpDia_operacao = Now
             
             strCampos = "DFDia_TBDia_operacao_pdv,DFStatus_dia_TBDia_Operacao_pdv,DFNumero_pdv_TBDia_operacao_pdv,IXCodigo_TBEmpresa"
             strValores = "'" & Format(Now, "YYYYMMDD") & "','A'," & intNumero_pdv & "," & intEmpresa & ""
             
             If booIntegracao_online = True Then
                funcoes_banco.Gravar "TBDia_Operacao_pdv", strCampos, strValores, "Otica", Me, "BDRetaguarda"
             Else
                funcoes_banco.Gravar "TBDia_Operacao_pdv", strCampos, strValores, "Otica", Me, "BDPDV"
             End If
          Else
             'Verifica se o dia da operação é hoje, se não perguntará ao operador o q fazer.
             If CDate(rstDia!Ultimo_dia_operacao) < Format(Now, "DD/MM/YYYY") Then
                intRetorno = MsgBox("Consta que último dia de operação aberto( " & Format(rstDia!Ultimo_dia_operacao, "DD/MM/YYYY") & " ),não foi fechado.Deseja proseguir operando com a data desatualizada?", vbYesNo, "Only Tech")
                If intRetorno = 6 Then
                   dtpDia_operacao = rstDia!Ultimo_dia_operacao
                   booDia_anterior = True
                Else
                   MsgBox "Para ajustar o dia de operação do PDV, confirme que deseja proseguir com o último dia de operação e efetue um fechamento de dia na tela de venda (Prescione F11).", vbInformation, "Only Tech"
                   End
                End If
             Else
                dtpDia_operacao = rstDia!Ultimo_dia_operacao
             End If
          End If
          
          '-------------------------------------------------------------------------------------------
          'Verifcando se o operador já fechou suas operações neste PDV no dia de hj e verifica se há necessidade de se abrir o caixa
          Dim rstCaixa_Aberto As New ADODB.Recordset
          
          strSql = Empty
          strSql = "SELECT * FROM TBOperacao_caixa " & _
                   "WHERE FKCodigo_TBPdv = " & intNumero_pdv & " " & _
                   "AND FKCodigo_TBOperadores_ecf = " & lngCodigo_operador & " " & _
                   "AND DFData_TBOperacao_caixa = '" & Format(Now, "YYYYMMDD") & "' " & _
                   "AND DFStatus_aberto_fechado_TBOperacao_caixa = 0 " & _
                   "AND FKCodigo_TBEmpresa = " & intEmpresa & ""
                   
          If booIntegracao_online = True Then
             Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCaixa_Aberto, "Otica", Me
          Else
             Movimentacoes.Select_geral strSql, "BDPDV", rstCaixa_Aberto, "PDV", Me
          End If
          
          'Verifica se existe a necessidade de se abrir o caixa
          If rstCaixa_Aberto.BOF = True And rstCaixa_Aberto.EOF = True And booDia_anterior = False Then
             
             Dim rstCaixa_Fechado As New ADODB.Recordset
             
             'Verifica se o caixa já foi fechado por este operador
             strSql = Empty
             strSql = "SELECT * FROM TBOperacao_caixa " & _
                      "WHERE FKCodigo_TBPdv = " & intNumero_pdv & " " & _
                      "AND FKCodigo_TBOperadores_ecf = " & lngCodigo_operador & " " & _
                      "AND DFData_TBOperacao_caixa = '" & Format(Now, "YYYYMMDD") & "' " & _
                      "AND DFStatus_aberto_fechado_TBOperacao_caixa = 1 " & _
                      "AND FKCodigo_TBEmpresa = " & intEmpresa & ""
                      
             If booIntegracao_online = True Then
                Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCaixa_Fechado, "Otica", Me
             Else
                Movimentacoes.Select_geral strSql, "BDPDV", rstCaixa_Fechado, "PDV", Me
             End If
             
             If rstCaixa_Fechado.BOF = True And rstCaixa_Fechado.EOF = True Then
                'Enviando inf. para cadastro de abertura de caixa
                frmAbertura_Caixa.strNumero_PDV = intNumero_pdv
                frmAbertura_Caixa.strEmpresa_Operador = rstOperador!FKCodigo_TBEmpresa
                frmAbertura_Caixa.strCodigo_Operador = lngCodigo_operador
                frmAbertura_Caixa.strOperador = rstOperador!PKCodigo_TBOperadores_ecf & "-" & rstOperador!DFNome_TBOperadores_ecf
                frmAbertura_Caixa.booIntegracao_Retaguarda = booIntegracao_online
                frmAbertura_Caixa.intImpressoes_suportadas = intImpressoes_suportadas
                frmAbertura_Caixa.dtpData_operacao = dtpDia_operacao
                frmAbertura_Caixa.booLeitor_serial = booLeitor_serial
                frmAbertura_Caixa.strCom_leitor_serial = strPorta_leitor
                frmAbertura_Caixa.intTipo_imp_orcamento = intTipo_imp_orcamento
                frmAbertura_Caixa.booGaveta_integrada = booGaveta_integrada
                
                'Informações pertinentes à lei
                frmAbertura_Caixa.intIP_Concentrador = rstParametro_ecf!DFEndereco_ip_concentrador_TBParametros_ecf
                frmAbertura_Caixa.booPreco_online = rstParametro_ecf!DFAtualizacao_preco_online_retaguarda_pdv_TBParametros_ecf
                frmAbertura_Caixa.booComissao_vendedor = rstParametro_ecf!DFComissao_vendedor_TBParametros_ecf
                frmAbertura_Caixa.strNumero_check_out = intNumero_pdv
                frmAbertura_Caixa.strNumero_Nome_Operadora = "Operador: " & lngCodigo_operador & " - " & rstOperador!DFNome_TBOperadores_ecf
                frmAbertura_Caixa.strVersao_software = "Versão 1.0"
                frmAbertura_Caixa.strNumero_loja = "Loja: " & intEmpresa
                frmAbertura_Caixa.intFinalizadora_sangria = rstParametro_ecf!DFFinalizadora_sangria_TBParametros_ecf
                frmAbertura_Caixa.strTipo_quantidade = rstParametro_ecf!DFTipo_quantidade_TBParametros_ecf
                frmAbertura_Caixa.strCasas_Decimais = rstParametro_ecf!DFNumero_decimais_TBParametros_ecf
                frmAbertura_Caixa.strTipo_desconto = rstParametro_ecf!DFTipo_desconto_TBParametros_ecf
                frmAbertura_Caixa.strTipo_quantidade = rstParametro_ecf!DFTipo_quantidade_TBParametros_ecf
                frmAbertura_Caixa.strDigito_Peso_Variavel = rstParametro_ecf.Fields!DFCodigo_inicial_peso_variavel_TBParametros_ecf
                frmAbertura_Caixa.booPreco_peso_balanca_TBParametros_ecf = rstParametro_ecf!DFPreco_peso_balanca_TBParametros_ecf
                frmAbertura_Caixa.strCaminho_impComum = strCaminho_impComum
                frmAbertura_Caixa.booFinaliza_direto = booFinaliza_direto
                
                If IsNull(rstParametro_ecf!DFPerfil_varejo_TBParametros_ecf) = True Then
                   frmAbertura_Caixa.intPerfil_ECF = 1
                Else
                   frmAbertura_Caixa.intPerfil_ECF = rstParametro_ecf!DFPerfil_varejo_TBParametros_ecf
                End If
                frmAbertura_Caixa.Show
                
                Set rstParametro_ecf = Nothing
                Set rstCaixa_Aberto = Nothing
                Set rstCaixa_Fechado = Nothing
                Set rstOperador = Nothing
                Set rstDia = Nothing
                Unload Me
             Else
                MsgBox "Este PDV já foi aberto e encerrado na data de hoje por este operador!Verifique.", vbCritical, "Only Tech"
                Set rstCaixa_Fechado = Nothing
                Set rstCaixa_Aberto = Nothing
                Set rstOperador = Nothing
                Set rstDia = Nothing
                Me.txtOperador.Text = Empty
                Me.txtSenha.Text = Empty
                Me.txtOperador.SetFocus
                Exit Sub
             End If
             
             Set rstCaixa_Fechado = Nothing
             
          Else
            'Verifica se o caixa já foi fechado por este operador
            strSql = Empty
            strSql = "SELECT * FROM TBOperacao_caixa " & _
                     "WHERE FKCodigo_TBPdv = " & intNumero_pdv & " " & _
                     "AND FKCodigo_TBOperadores_ecf = " & rstOperador!PKCodigo_TBOperadores_ecf & " " & _
                     "AND DFData_TBOperacao_caixa = '" & Format(Now, "YYYYMMDD") & "' " & _
                     "AND DFStatus_aberto_fechado_TBOperacao_caixa = 1 " & _
                     "AND FKCodigo_TBEmpresa = " & intEmpresa & ""
                     
            If booIntegracao_online = True Then
               Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCaixa_Fechado, "Otica", Me
            Else
               Movimentacoes.Select_geral strSql, "BDPDV", rstCaixa_Fechado, "PDV", Me
            End If
                
            If Not rstCaixa_Fechado.BOF = True And Not rstCaixa_Fechado.EOF = True Then
                MsgBox "Este PDV já foi aberto e encerrado na data de hoje por este operador!Verifique.", vbCritical, "Only Tech"
                Set rstCaixa_Fechado = Nothing
                Set rstCaixa_Aberto = Nothing
                Set rstOperador = Nothing
                Set rstDia = Nothing
                Me.txtOperador.Text = Empty
                Me.txtSenha.Text = Empty
                Me.txtOperador.SetFocus
                Exit Sub
            End If
          
            'Enviando inf. para Tela de Vendas
            frmTela_Venda.strEmpresa_Operador = rstOperador!FKCodigo_TBEmpresa
            frmTela_Venda.strOperador = rstOperador!PKCodigo_TBOperadores_ecf & "-" & rstOperador!DFNome_TBOperadores_ecf
            frmTela_Venda.strCodigo_Operador = rstOperador!PKCodigo_TBOperadores_ecf
            frmTela_Venda.booIntegracao_Retaguarda = booIntegracao_online
            frmTela_Venda.dtpData_operacao = dtpDia_operacao
            frmTela_Venda.intImpressoes_suportadas = intImpressoes_suportadas
            frmTela_Venda.booLeitor_serial = booLeitor_serial
            frmTela_Venda.strCom_leitor_serial = strPorta_leitor
            frmTela_Venda.intTipo_imp_orcamento = intTipo_imp_orcamento
            frmTela_Venda.booGaveta_integrada = booGaveta_integrada
            frmTela_Venda.intIP_Concentrador = rstParametro_ecf!DFEndereco_ip_concentrador_TBParametros_ecf
            frmTela_Venda.booPreco_online = rstParametro_ecf!DFAtualizacao_preco_online_retaguarda_pdv_TBParametros_ecf
            frmTela_Venda.booComissao_vendedor = rstParametro_ecf!DFComissao_vendedor_TBParametros_ecf
            frmTela_Venda.txtNumero_check_out.Text = intNumero_pdv
            frmTela_Venda.txtNumero_Nome_Operadora.Text = "Operador:" & lngCodigo_operador & " - " & rstOperador!DFNome_TBOperadores_ecf
            frmTela_Venda.txtVersao_software.Text = "Versão 1.0"
            frmTela_Venda.txtNumero_loja.Text = "Loja: " & intEmpresa
            frmTela_Venda.intFinalizadora_sangria = rstParametro_ecf!DFFinalizadora_sangria_TBParametros_ecf
            frmTela_Venda.strTipo_quantidade = rstParametro_ecf!DFTipo_quantidade_TBParametros_ecf
            frmTela_Venda.strCasas_Decimais = rstParametro_ecf!DFNumero_decimais_TBParametros_ecf
            frmTela_Venda.strTipo_desconto = rstParametro_ecf!DFTipo_desconto_TBParametros_ecf
            frmTela_Venda.strTipo_quantidade = rstParametro_ecf!DFTipo_quantidade_TBParametros_ecf
            frmTela_Venda.strDigito_Peso_Variavel = rstParametro_ecf.Fields!DFCodigo_inicial_peso_variavel_TBParametros_ecf
            frmTela_Venda.booPreco_peso_balanca_TBParametros_ecf = rstParametro_ecf!DFPreco_peso_balanca_TBParametros_ecf
            frmTela_Venda.strCaminho_impComum = strCaminho_impComum
            frmTela_Venda.intPerfil_ECF = rstParametro_ecf!DFPerfil_varejo_TBParametros_ecf
            frmTela_Venda.booFinaliza_direto = booFinaliza_direto
            
            frmTela_Venda.Show
            
            Set rstParametro_ecf = Nothing
            Set rstCaixa_Aberto = Nothing
            Set rstOperador = Nothing
            
            Unload Me
            
          End If
          
          Set rstParametro_ecf = Nothing
          Set rstCaixa_Aberto = Nothing
          Set rstOperador = Nothing
          Set rstDia = Nothing
          
          Unload Me
    End If
    
    Exit Sub
    
Erro:

    If Err.Number = -2147220503 Then
       MsgBox "Fonte de dados não encontrada!", vbCritical, "Only Tech"
       Shell App.Path & "Configurador de Sistemas.exe", vbNormalFocus
       End
       Exit Sub
    Else
       If booIntegracao_online = True Then
          Call Erro.Erro(Me, "Otica", "Load", Err.Number)
       Else
          Call Erro.Erro(Me, "PDV", "Load", Err.Number)
       End If
        
       Exit Sub
    End If
    
End Sub

Private Function CheckList_pdv_inicio()
    Dim booInterromper As Boolean
    
    booInterromper = False
    
    Dim rstCheck_Vendedor As New ADODB.Recordset
    Dim rstCheck_Plano As New ADODB.Recordset
    Dim rstCheck_Transportador As New ADODB.Recordset
    Dim rstCheck_Cliente As New ADODB.Recordset
    Dim rstCheck_Finalizadora As New ADODB.Recordset
    Dim rstCheck_Tabela As New ADODB.Recordset
    Dim rstCheck_Tributação As New ADODB.Recordset
    Dim rstCheck_Operador As New ADODB.Recordset
    Dim rstCheck_Empresa As New ADODB.Recordset
    Dim rstCheck_PDV As New ADODB.Recordset
    Dim rstCheck_Parametros_ecf As New ADODB.Recordset
    
    'Vendedor
    strSql = Empty
    strSql = "SELECT IXCodigo_TBVendedor FROM TBVendedor WHERE IXCodigo_TBVendedor = 9999 " & _
             "AND IXCodigo_TBEmpresa = " & intEmpresa & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Vendedor, "Otica", Me
    
    If rstCheck_Vendedor.EOF = True And rstCheck_Vendedor.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhum vendedor com o código 9999 e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If
    
    'Plano de pagamento
    strSql = Empty
    strSql = "SELECT IXCodigo_TBPlano_pagamento FROM TBPlano_pagamento WHERE IXCodigo_TBPlano_pagamento = 9999 " & _
             "AND IXCodigo_TBEmpresa = " & intEmpresa & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Plano, "Otica", Me
    
    If rstCheck_Plano.EOF = True And rstCheck_Plano.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhum plano de pagamento com o código 9999 e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If

    'Transportador
    strSql = Empty
    strSql = "SELECT PKCodigo_TBTransportadora FROM TBTransportadora WHERE PKCodigo_TBTransportadora = 9999"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Transportador, "Otica", Me
    
    If rstCheck_Transportador.EOF = True And rstCheck_Transportador.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhuma transportadora com o código 9999 e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If

    'Cliente
    strSql = Empty
    strSql = "SELECT IXCodigo_TBCliente FROM TBCliente WHERE IXCodigo_TBCliente = 9999 " & _
             "AND IXCodigo_TBEmpresa = " & intEmpresa & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Cliente, "Otica", Me
    
    If rstCheck_Cliente.EOF = True And rstCheck_Cliente.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhum cliente com o código 9999 e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If

    'Finalizadora
    strSql = Empty
    strSql = "SELECT * FROM TBFinalizadora"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Finalizadora, "Otica", Me
            
    If rstCheck_Finalizadora.EOF = True And rstCheck_Finalizadora.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhuma finalizadora cadastrada e o mesma é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If
    
    'Parametros Venda - Tabela Vigente
    strSql = Empty
    strSql = "SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBParametros_venda " & _
             "WHERE IXCodigo_TBEmpresa = " & intEmpresa & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Tabela, "Otica", Me
    
    If rstCheck_Tabela.EOF = True And rstCheck_Tabela.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhuma tabela vigente cadastrada e o mesma é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If
    
    'Parametros Fiscais - Tributação
    strSql = Empty
    strSql = "SELECT DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais,DFProximo_cfop_venda_dentro_estado_TBParametros_fiscais FROM TBParametros_fiscais " & _
             "WHERE  FKCodigo_TBEmpresa = " & intEmpresa & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Tributação, "Otica", Me
    
    If rstCheck_Tributação.EOF = True And rstCheck_Tributação.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhum parâmetro fiscal cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    Else
       If IsNull(rstCheck_Tributação!DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais) Or rstCheck_Tributação!DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais = "" Then
           MsgBox "RETAGUARDA - Não consta em sua base nenhum cfop de venda dentro do estado com substituição no parâmetro fiscal cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
           booInterromper = True
       End If
       If IsNull(rstCheck_Tributação!DFProximo_cfop_venda_dentro_estado_TBParametros_fiscais) Or rstCheck_Tributação!DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais = "" Then
           MsgBox "RETAGUARDA - Não consta em sua base nenhum cfop de venda dentro do estado no parâmetro fiscal cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
           booInterromper = True
       End If
    End If
    
    'Operadores
    strSql = Empty
    strSql = "SELECT * FROM TBOperadores_ecf " & _
             "WHERE  FKCodigo_TBEmpresa = " & intEmpresa & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Operador, "Otica", Me
    
    If rstCheck_Operador.EOF = True And rstCheck_Operador.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhum operador cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If
    
    'Empresa
    strSql = Empty
    strSql = "SELECT * FROM TBEmpresa"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Empresa, "Otica", Me
    
    If rstCheck_Empresa.EOF = True And rstCheck_Empresa.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhuma empresa cadastrada e o mesma é necessária para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If
    
    'PDV
    strSql = Empty
    strSql = "SELECT * FROM TBPDV " & _
             "WHERE IXCodigo_TBEmpresa = " & intEmpresa & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_PDV, "Otica", Me
    
    If rstCheck_PDV.EOF = True And rstCheck_PDV.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhum PDV cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If
    
    'Parâmetros ECF
    strSql = Empty
    strSql = "SELECT * FROM TBParametros_ecf WHERE FKCodigo_TBEmpresa = " & intEmpresa & ""
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Parametros_ecf, "Otica", Me
    
    If rstCheck_Parametros_ecf.EOF = True And rstCheck_Parametros_ecf.BOF = True Then
       MsgBox "RETAGUARDA - Não consta em sua base nenhum registro de Parametros do ecf cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
       booInterromper = True
    End If
    
    Set rstCheck_Vendedor = Nothing
    Set rstCheck_Plano = Nothing
    Set rstCheck_Transportador = Nothing
    Set rstCheck_Cliente = Nothing
    Set rstCheck_Finalizadora = Nothing
    Set rstCheck_Tabela = Nothing
    Set rstCheck_Tributação = Nothing
    Set rstCheck_Operador = Nothing
    Set rstCheck_Empresa = Nothing
    Set rstCheck_PDV = Nothing
    Set rstCheck_Parametros_ecf = Nothing
    
    '--------------------------------------------------------------------------------------------------
    'Local
    '--------------------------------------------------------------------------------------------------
''
''    'Vendedor
''    strSql = Empty
''    strSql = "SELECT IXCodigo_TBVendedor FROM TBVendedor WHERE IXCodigo_TBVendedor = 9999"
''    Movimentacoes.Select_geral strSql, "BDPDV", rstCheck_Vendedor, "PDV", Me
''
''    If rstCheck_Vendedor.EOF = True And rstCheck_Vendedor.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhum vendedor com o código 9999 e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
''
''    'Plano de pagamento
''    strSql = Empty
''    strSql = "SELECT IXCodigo_TBPlano_pagamento FROM TBPlano_pagamento WHERE IXCodigo_TBPlano_pagamento = 9999"
''    Movimentacoes.Select_geral strSql, "BDPDV", rstCheck_Plano, "PDV", Me
''
''    If rstCheck_Plano.EOF = True And rstCheck_Plano.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhum plano de pagamento com o código 9999 e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
''
''    'Transportador
''    strSql = Empty
''    strSql = "SELECT PKCodigo_TBTransportadora FROM TBTransportadora WHERE PKCodigo_TBTransportadora = 9999"
''    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Transportador, "PDV", Me
''
''    If rstCheck_Transportador.EOF = True And rstCheck_Transportador.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhuma transportadora com o código 9999 e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
''
''    'Cliente
''    strSql = Empty
''    strSql = "SELECT IXCodigo_TBCliente FROM TBCliente WHERE IXCodigo_TBCliente = 9999"
''    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Cliente, "PDV", Me
''
''    If rstCheck_Cliente.EOF = True And rstCheck_Cliente.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhum cliente com o código 9999 e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
''
''    'Finalizadora
''    strSql = Empty
''    strSql = "SELECT * FROM TBFinalizadora"
''    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Finalizadora, "PDV", Me
''
''    If rstCheck_Finalizadora.EOF = True And rstCheck_Finalizadora.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhuma finalizadora cadastrada e o mesma é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
''
''    'Parametros Venda - Tabela Vigente
''    strSql = Empty
''    strSql = "SELECT DFNumero_tabela_vigente_TBParametros_venda FROM TBParametros_venda"
''    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Tabela, "PDV", Me
''
''    If rstCheck_Tabela.EOF = True And rstCheck_Tabela.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhuma tabela vigente cadastrada e o mesma é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
''
''    'Parametros Fiscais - Tributação
''    strSql = Empty
''    strSql = "SELECT DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais,DFProximo_cfop_venda_dentro_estado_TBParametros_fiscais FROM TBParametros_fiscais"
''    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Tributação, "PDV", Me
''
''    If rstCheck_Tributação.EOF = True And rstCheck_Tributação.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhum parâmetro fiscal cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    Else
''       If IsNull(rstCheck_Tributação!DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais) Or rstCheck_Tributação!DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais = "" Then
''           MsgBox "LOCAL - Não consta em sua base nenhum cfop de venda dentro do estado com substituição no parâmetro fiscal cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''           booInterromper = True
''       End If
''       If IsNull(rstCheck_Tributação!DFProximo_cfop_venda_dentro_estado_TBParametros_fiscais) Or rstCheck_Tributação!DFProximo_cfop_venda_dentro_substituicao_estado_TBParametros_fiscais = "" Then
''           MsgBox "LOCAL - Não consta em sua base nenhum cfop de venda dentro do estado no parâmetro fiscal cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''           booInterromper = True
''       End If
''    End If
''
''    'Operadores
''    strSql = Empty
''    strSql = "SELECT * FROM TBOperadores_ecf"
''    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Operador, "PDV", Me
''
''    If rstCheck_Operador.EOF = True And rstCheck_Operador.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhum operador cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
''
''    'Empresa
''    strSql = Empty
''    strSql = "SELECT * FROM TBEmpresa"
''    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Empresa, "PDV", Me
''
''    If rstCheck_Empresa.EOF = True And rstCheck_Empresa.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhuma empresa cadastrada e o mesma é necessária para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
''
''    'PDV
''    strSql = Empty
''    strSql = "SELECT * FROM TBPDV"
''    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_PDV, "PDV", Me
''
''    If rstCheck_PDV.EOF = True And rstCheck_PDV.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhum PDV cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
''
''    'Parâmetros ECF
''    strSql = Empty
''    strSql = "SELECT * FROM TBParametros_ecf"
''    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCheck_Parametros_ecf, "PDV", Me
''
''    If rstCheck_Parametros_ecf.EOF = True And rstCheck_Parametros_ecf.BOF = True Then
''       MsgBox "LOCAL - Não consta em sua base nenhum registro de Parametros do ecf cadastrado e o mesmo é necessário para prosseguir!Verifique.", vbCritical, "Only Tech"
''       booInterromper = True
''    End If
    
    Set rstCheck_Vendedor = Nothing
    Set rstCheck_Plano = Nothing
    Set rstCheck_Transportador = Nothing
    Set rstCheck_Cliente = Nothing
    Set rstCheck_Finalizadora = Nothing
    Set rstCheck_Tabela = Nothing
    Set rstCheck_Tributação = Nothing
    Set rstCheck_Operador = Nothing
    Set rstCheck_Empresa = Nothing
    Set rstCheck_PDV = Nothing
    Set rstCheck_Parametros_ecf = Nothing
    
    'Interrompe a execução por faltar inf. que serão necessárias para validar o aplicativo
    If booInterromper = True Then
       End
    End If
    
End Function
