VERSION 5.00
Object = "{D0159C1D-A983-4698-8940-3BE45A260C35}#1.0#0"; "SegundoPlanoMDI.ocx"
Object = "{C5014412-BD55-402F-8335-07C273732964}#1.1#0"; "AplicativoUsuário.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Concentrador - Cadastros Base"
   ClientHeight    =   8100
   ClientLeft      =   1740
   ClientTop       =   1200
   ClientWidth     =   13545
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrincipal.frx":1782
   Begin AplicativoUsuárioOCX.AplicativoUsuário OCXUsuario 
      Left            =   6480
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin OCXSegundoPlano.SegundoPlanoMDI SegundoPlanoMDI 
      Left            =   5520
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastros"
      Begin VB.Menu smnEmpresa 
         Caption         =   "&Empresa"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnParametros_ecf 
         Caption         =   "&Parâmetros ECF"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnOperador_Ecf 
         Caption         =   "&Operador ECF"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnFinalizadora 
         Caption         =   "&Finalizadora"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnImpressora_ecf 
         Caption         =   "Impressora ECF"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnPdv 
         Caption         =   "&Pdv"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnOperacao_Caixa 
         Caption         =   "Opera&ção Caixa"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnComanda 
         Caption         =   "&Comanda"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu smnrelTira_Teima 
         Caption         =   "Tira Teima"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuMovimentacoes 
      Caption         =   "&Movimentações"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuEstatisticas 
      Caption         =   "&Estatisticas"
      Enabled         =   0   'False
      Begin VB.Menu smnMovimeneto_Caixa 
         Caption         =   "Mov&imento Caixa"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnVenda_Diaria 
         Caption         =   "Ve&nda Diária"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnComanda_Nao_Finalizada 
         Caption         =   "Coman&da Não Finalizada"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "&?"
      Begin VB.Menu smnAjuda 
         Caption         =   "&Ajuda"
      End
      Begin VB.Menu smnSobre 
         Caption         =   "&Sobre"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                         '
' Módulo.................: Cadastro Base                                                  '
' Objetivo...............: MDI Principal                                                  '
' Data de Criação........: 14/01/2005                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex Baião,Rafael Gomes, Sérgio   '
' Última Manutenção......:                                                                '
' Data última manutenção.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim acesso As New DLLSystemManager.Acessibilidade
'------------------------------------------------------------
'Declaração da variavel do intercomunicador de mensagens
Private Cliente_mensagem_exe As VetorDeMensagens.ClienteDeMensagens
Dim log As New DLLSystemManager.log
Public booDesign_time As Boolean
Option Explicit

Private Sub MDIForm_Activate()
    Dim strMensagem_cliente() As String
    Dim mensagem_design As String
    
    On Error GoTo Erro
    
    OCXUsuario.Nome = "Acessibilidade"
    OCXUsuario.Estacao = "Acessibilidade"
    
    'Informações Constantes para o log
    log.Usuario = OCXUsuario.Nome
    log.Programa = "Módulo Cadastros Base"
    log.Estacao = OCXUsuario.Estacao

    'Informações Variaveis para o log
    log.Evento = "Acessibilidade"
    log.Tipo = 4
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")

    log.Descricao = "Inicializando a acessibilidade do Módulo Cadastros Base"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    Call Access
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Activate")
    Exit Sub
    
End Sub

Private Sub MDIForm_Load()
   'Setando e recebendo as mensagens do admin
    Set Cliente_mensagem_exe = New VetorDeMensagens.ClienteDeMensagens
    Cliente_mensagem_exe.ID_Aplicativo = Me.hwnd
    Cliente_mensagem_exe.Interceptar
End Sub

Private Sub smnComanda_Click()
    frmComanda.Show
End Sub

Private Sub smnEmpresa_Click()
    frmEmpresa.Show
End Sub

Private Sub smnFinalizadora_Click()
    frmFinalizadora.Show
End Sub

Private Sub smnImpressora_ecf_Click()
    frmImpressora_Ecf.Show
End Sub

Private Sub smnOperacao_Caixa_Click()
    frmOperacao_Caixa.Show
End Sub

Private Sub smnOperador_Ecf_Click()
    frmOperador_Ecf.Show
End Sub

Private Sub smnParametros_ecf_Click()
    frmParametros_Ecf.Show
End Sub

Private Sub smnPdv_Click()
    frmPdv.Show
End Sub

Private Sub smnrelTira_Teima_Click()
    FrmTira_Teima.Show
End Sub

Private Sub smnSobre_Click()
    frmSobre.Show
End Sub

Private Sub Access()

    Dim strMensagem_cliente() As String
    Dim mensagem_design As String
    
    On Error GoTo Erro
   
    'Marcar sempre essa variavel com false quando for compilar e testar no admin
    booDesign_time = True
    
    If booDesign_time = True Then
       mensagem_design = "ESTAÇÃOTESTE¤MARCOS¤SENHA_TESTE¤marcos¤1¤AREA_TRABALHO_TESTE¤100"
       strMensagem_cliente = Split(mensagem_design, "¤")
       MDIPrincipal.OCXUsuario.Empresa = strMensagem_cliente(6)
    Else
       If Cliente_mensagem_exe.MensagemRecebida = "" Then
          'Tentativa de acessar inf dos usuários
          '-------------------------------------------------------------------------------------------
          'Log
          log.Tipo = 4    ' Tipo de Log de uso da Only Tech
          log.Data = Date
          log.Hora = Format(Now, "hh:mm:ss")
          log.Descricao = "Sistema acessando registro da máquina na tentativa de obter inf. do usuário."
          'Gravando o log
          log.Gravar_log "Otica", Me
          '--------------------------------------------------------------------------------------------
          Dim strMensagem_Registro As String
        
          strMensagem_Registro = Movimentacoes.Consulta_Contingencia_Acessibilidade("Otica")
         
          If strMensagem_Registro = "" Then
             'Falha nas 2 primeiras tentativas, sistema impossibilitado de acessar inf. do usuário
             '-------------------------------------------------------------------------------------------
             'Log
             log.Tipo = 4    ' Tipo de Log de uso da Only Tech
             log.Data = Date
             log.Hora = Format(Now, "hh:mm:ss")
             log.Descricao = "Falha no acesso a memória e ao registro da máquina, sistema impossibilitado de acessar inf. do usuário."
             'Gravando o log
             log.Gravar_log "Otica", Me
             '--------------------------------------------------------------------------------------------
              MsgBox "Acessibilidade - Ocorreu uma falha de execução interna do aplicativo,reexecute o mesmo,se o problema persistir contacte Only Tech Solutions", vbInformation, "Only Tech"
              Exit Sub
          End If
          strMensagem_cliente = Split(strMensagem_Registro, "¤")
          MDIPrincipal.OCXUsuario.Empresa = strMensagem_cliente(11)
       Else
          strMensagem_cliente = Split(Cliente_mensagem_exe.MensagemRecebida, "¤")
          MDIPrincipal.OCXUsuario.Empresa = strMensagem_cliente(11)
       End If
    End If
    
    OCXUsuario.Nome = strMensagem_cliente(3)
    OCXUsuario.Estacao = strMensagem_cliente(0)
    OCXUsuario.Codigo = strMensagem_cliente(4)
    
    'Informações Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    log.Descricao = "Inicializando o Módulo de Cadastros Base"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    Call Acessibilidade
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Access")
    Exit Sub
    
End Sub
Private Function Acessibilidade()

    'Parâmetros ECF
    Movimentacoes.Acessibilidade_Item_Menu "Parâmetros ECF", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnParametros_ecf
    'Empresa
    Movimentacoes.Acessibilidade_Item_Menu "Empresa", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnEmpresa
    'Tira Teima
    Movimentacoes.Acessibilidade_Item_Menu "Tira Teima", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnrelTira_Teima
    'Comanda
    Movimentacoes.Acessibilidade_Item_Menu "Comanda", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnComanda
    'Operador ECF
    Movimentacoes.Acessibilidade_Item_Menu "Operador ECF", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnOperador_Ecf
    'Impressora ECF
    Movimentacoes.Acessibilidade_Item_Menu "Impressora ECF", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnImpressora_ecf
    'PDV
    Movimentacoes.Acessibilidade_Item_Menu "PDV", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnPdv
    'Operação Caixa
    Movimentacoes.Acessibilidade_Item_Menu "Operação Caixa", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnOperacao_Caixa
    'Finalizadora
    Movimentacoes.Acessibilidade_Item_Menu "Finalizadora", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnFinalizadora
    
End Function

