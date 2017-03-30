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
    
    strSql = Empty
    strSql = "SELECT DFIntegracao_online_pdv_retaguarda_TBParametros_ecf FROM TBParametros_ecf "
    Movimentacoes.Select_geral strSql, "BDPDV", rstIntegracao_online, "PDV", Me
    
    If rstIntegracao_online.BOF = True And rstIntegracao_online.EOF = True Then
       MsgBox "Integração com retaguarda não definida.Verifique.", vbCritical, "Only Tech"
       Set rstPDV = Nothing
       End
    End If
    
    'Carregando a combo de pdv
    strSql = Empty
    strSql = "SELECT PKCodigo_TBPdv,DFImpressoes_suportadas_TBPdv,DFLeitor_Serial_integrado,DFPorta_com_leitor_serial,DFTipo_impressora_orcamento_balcao_TBpdv,DFGaveta_integrada_TBPdv FROM TBPDV WHERE DFEndereco_ip_TBPdv = '" & strIP & "'"
    
    If rstIntegracao_online!DFIntegracao_online_pdv_retaguarda_TBParametros_ecf = True Then
       Movimentacoes.Select_geral strSql, "BDRetaguarda", rstPDV, "Otica", Me
    Else
       Movimentacoes.Select_geral strSql, "BDPDV", rstPDV, "PDV", Me
    End If
    
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
    
    intNumero_pdv = rstPDV!PKCodigo_TBPdv
    intImpressoes_suportadas = rstPDV!DFImpressoes_suportadas_TBPdv
    intTipo_imp_orcamento = rstPDV!DFTipo_impressora_orcamento_balcao_TBpdv
    booGaveta_integrada = rstPDV!DFGaveta_integrada_TBPdv
    
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
       
    booIntegracao_online = rstIntegracao_online!DFIntegracao_online_pdv_retaguarda_TBParametros_ecf
    
    Set rstIntegracao_online = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstOperador = Nothing
End Sub

Private Sub txtOperador_GotFocus()
'    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtOperador_LostFocus()

    On Error GoTo erro
    
    Me.txtOperador.Text = UCase(Me.txtOperador)
    strSql = Empty
    strSql = "SELECT * FROM TBOperadores_ecf WHERE DFNome_TBOperadores_ecf = '" & Me.txtOperador.Text & "'"
    
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
    
erro:
    If Err.Number = -2147220503 Then
       MsgBox "Fonte de dados não encontrada!", vbCritical, "Only Tech"
       Shell App.Path & "Configurador de Sistemas.exe", vbNormalFocus
       End
       Exit Sub
    Else
       If booIntegracao_online = True Then
          Call erro.erro(Me, "Otica", "Load", Err.Number)
       Else
          Call erro.erro(Me, "PDV", "Load", Err.Number)
       End If
       
       Exit Sub
    End If
    
End Sub

Private Sub txtSenha_GotFocus()
    On Error Resume Next: Me.ActiveControl.SelStart = 0: Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtSenha_LostFocus()

    On Error GoTo erro
    
    If Me.txtSenha.Text <> "" Then
          Dim intRetorno As Integer
          
          'Verificação de dia de Operação
          strSql = Empty
          strSql = "SELECT MAX(DFDia_TBDia_operacao_pdv) as Ultimo_dia_operacao FROM TBDia_operacao_pdv WHERE DFStatus_dia_TBDia_Operacao_pdv = 'A' AND DFNumero_pdv_TBDia_operacao_pdv = " & intNumero_pdv & ""
          
          If booIntegracao_online = True Then
             Movimentacoes.Select_geral strSql, "BDRetaguarda", rstDia, "Otica", Me
          Else
             Movimentacoes.Select_geral strSql, "BDPDV", rstDia, "PDV", Me
          End If
          
          strSql = Empty
          strSql = "SELECT * FROM TBOperadores_ecf WHERE PKCodigo_TBOperadores_ecf = " & lngCodigo_operador & " AND DFSenha_TBOperadores_ecf = '" & Me.txtSenha.Text & "'"
          
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
             
             strCampos = "DFDia_TBDia_operacao_pdv,DFStatus_dia_TBDia_Operacao_pdv,DFNumero_pdv_TBDia_operacao_pdv"
             strValores = "'" & Format(dtpDia_operacao, "YYYYMMDD") & "','A'," & intNumero_pdv & ""
             
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

          'Verifcando se o operador já fechou susas operações neste PDV no dia de hj e verifica se há necessidade de se abrir o caixa
          Dim rstCaixa_Aberto As New ADODB.Recordset
          
          strSql = Empty
          strSql = "SELECT * FROM TBOperacao_caixa " & _
                   "WHERE FKCodigo_TBPdv = " & intNumero_pdv & " " & _
                   "AND FKCodigo_TBOperadores_ecf = " & lngCodigo_operador & " " & _
                   "AND DFData_TBOperacao_caixa = '" & Format(Now, "YYYYMMDD") & "' " & _
                   "AND DFStatus_aberto_fechado_TBOperacao_caixa = 0"
                   
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
                      "AND DFStatus_aberto_fechado_TBOperacao_caixa = 1"
                      
             If booIntegracao_online = True Then
                Movimentacoes.Select_geral strSql, "BDRetaguarda", rstCaixa_Fechado, "Otica", Me
             Else
                Movimentacoes.Select_geral strSql, "BDPDV", rstCaixa_Fechado, "PDV", Me
             End If
             
             If rstCaixa_Fechado.BOF = True And rstCaixa_Fechado.EOF = True Then
                'Enviando inf. para cadastro de abertura de caixa
                frmAbertura_Caixa.strPDV = intNumero_pdv
                frmAbertura_Caixa.strNumero_PDV = intNumero_pdv
                frmAbertura_Caixa.strEmpresa_Operador = rstOperador!FKCodigo_TBEmpresa
                frmAbertura_Caixa.strCodigo_Operador = lngCodigo_operador
                frmAbertura_Caixa.strOperador = lngCodigo_operador & "-" & rstOperador!DFNome_TBOperadores_ecf
                frmAbertura_Caixa.booIntegracao_Retaguarda = booIntegracao_online
                frmAbertura_Caixa.intImpressoes_suportadas = intImpressoes_suportadas
                frmAbertura_Caixa.dtpData_operacao = dtpDia_operacao
                frmAbertura_Caixa.booLeitor_serial = booLeitor_serial
                frmAbertura_Caixa.strCom_leitor_serial = strPorta_leitor
                frmAbertura_Caixa.intTipo_imp_orcamento = intTipo_imp_orcamento
                frmAbertura_Caixa.booGaveta_integrada = booGaveta_integrada
                frmAbertura_Caixa.Show
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
                     "AND DFStatus_aberto_fechado_TBOperacao_caixa = 1"
            
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
            frmTela_Venda.strPDV = intNumero_pdv
            frmTela_Venda.dtpData_operacao = dtpDia_operacao
            frmTela_Venda.intImpressoes_suportadas = intImpressoes_suportadas
            frmTela_Venda.booLeitor_serial = booLeitor_serial
            frmTela_Venda.strCom_leitor_serial = strPorta_leitor
            frmTela_Venda.intTipo_imp_orcamento = intTipo_imp_orcamento
            frmTela_Venda.booGaveta_integrada = booGaveta_integrada
            frmTela_Venda.Show
            
            Set rstCaixa_Aberto = Nothing
            Set rstOperador = Nothing
            Unload Me
            
          End If
          
          Set rstCaixa_Aberto = Nothing
          Set rstOperador = Nothing
          Set rstDia = Nothing
          
          Unload Me
    End If
    
    Exit Sub
    
erro:
    If Err.Number = -2147220503 Then
       MsgBox "Fonte de dados não encontrada!", vbCritical, "Only Tech"
       Shell App.Path & "Configurador de Sistemas.exe", vbNormalFocus
       End
       Exit Sub
    Else
       If booIntegracao_online = True Then
          Call erro.erro(Me, "Otica", "Load", Err.Number)
       Else
          Call erro.erro(Me, "PDV", "Load", Err.Number)
       End If
        
       Exit Sub
    End If
    
End Sub
