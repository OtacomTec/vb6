VERSION 5.00
Object = "{D0159C1D-A983-4698-8940-3BE45A260C35}#1.0#0"; "SegundoPlanoMDI.ocx"
Object = "{C5014412-BD55-402F-8335-07C273732964}#1.1#0"; "AplicativoUsu�rio.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Servi�os"
   ClientHeight    =   7980
   ClientLeft      =   2325
   ClientTop       =   1365
   ClientWidth     =   13545
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrincipal.frx":1782
   Begin AplicativoUsu�rioOCX.AplicativoUsu�rio OCXUsuario 
      Left            =   5220
      Top             =   630
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin OCXSegundoPlano.SegundoPlanoMDI SegundoPlanoMDI 
      Left            =   5910
      Top             =   630
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastros"
      Begin VB.Menu smnParametros_Servicos 
         Caption         =   "Par�metros de Servi�os"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnStatus_Pendencia 
         Caption         =   "Status Pen&d�ncia"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnServicos 
         Caption         =   "Servi�os"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnPlano_servicos 
         Caption         =   "Plano de Servi�os"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnPrioridade_Pendencia 
         Caption         =   "P&rioridade Pend�ncia"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnContrato_servicos 
         Caption         =   "Contrato de Servi�os"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnFuncao_Insumo 
         Caption         =   "&Fun��o Insumo"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnInsumo 
         Caption         =   "&Insumo"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnEquipamento_Laboratorio 
         Caption         =   "&Equipamento Laborat�rio"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnTipo_Marcha 
         Caption         =   "&Tipo Marcha"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnTipo_Servico 
         Caption         =   "Tipo S&ervi�o"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnMarcha 
         Caption         =   "&Marcha Anal�tica"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnTriagem_Laboratorio 
         Caption         =   "Tria&gem Laborat�rio"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnSoftware 
         Caption         =   "&Software"
         WindowList      =   -1  'True
         Begin VB.Menu smnSoftware_agenda 
            Caption         =   "&Agenda"
            Enabled         =   0   'False
         End
         Begin VB.Menu smnSoftware_pendencias 
            Caption         =   "&Pend�ncias"
            Enabled         =   0   'False
         End
         Begin VB.Menu smnSoftware_triagem 
            Caption         =   "&Triagem"
            Enabled         =   0   'False
         End
         Begin VB.Menu smnSoftware_visitas 
            Caption         =   "&Visitas"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu smncadTipo_Atendimento_Servico 
         Caption         =   "Tipo Atendimento Servi�o"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relat�rios"
      Begin VB.Menu smnRetaguarda_Relatorio 
         Caption         =   "Re&taguarda"
         Begin VB.Menu smnrelPendencias 
            Caption         =   "Relat�rio &Pend�ncias"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu smnPortal_Relatorio 
         Caption         =   "&Portal"
         Begin VB.Menu smnRelatorio_Senha 
            Caption         =   "Relat�rio &Senha"
            Enabled         =   0   'False
         End
         Begin VB.Menu smnrelTriagem 
            Caption         =   "Relat�rio &Triagem"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuMovimentacoes 
      Caption         =   "&Movimenta��es"
      Begin VB.Menu smnRetaguarda_Movimentacoes 
         Caption         =   "&Retaguarda"
         Begin VB.Menu smnmovGera_Plano_Completo 
            Caption         =   "Gera Plano Completo"
            Enabled         =   0   'False
         End
         Begin VB.Menu smnmovGeracao_Ordem_Servico 
            Caption         =   "Gera��o de Ordem Servi�o"
            Enabled         =   0   'False
         End
         Begin VB.Menu smnmovIntegracao_retaguarda_portal 
            Caption         =   "Integra��o Retaguarda X Portal"
         End
      End
      Begin VB.Menu smnPortal_Movimentacoes 
         Caption         =   "&Portal"
         Begin VB.Menu smnProcessamento_Senha_Portal 
            Caption         =   "Processamento de &Senha Portal"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuEstatisticas 
      Caption         =   "&Estatisticas"
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
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Ventura                                                        '
' M�dulo.................: Servi�os                                                       '
' Objetivo...............: MDI Principal                                                  '
' Data de Cria��o........: 19/01/2004                                                     '
' Equipe Respons�vel.....: Only Tech Solutions                                            '
' �ltima Manuten��o......:                                                                '
' Desenvolvedor..........:                                                                '
' Data �ltima manuten��o.:   /  /                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim acesso As New DLLSystemManager.Acessibilidade
'------------------------------------------------------------
'Declara��o da variavel do intercomunicador de mensagens
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
    
    'Informa��es Constantes para o log
    log.Usuario = OCXUsuario.Nome
    log.Programa = "M�dulo Servi�os"
    log.Estacao = OCXUsuario.Estacao

    'Informa��es Variaveis para o log
    log.Evento = "Acessibilidade"
    log.Tipo = 4
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")

    log.Descricao = "Inicializando a acessibilidade do M�dulo Servi�os"
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
    Cliente_mensagem_exe.ID_Aplicativo = Me.hWnd
    Cliente_mensagem_exe.Interceptar
End Sub

Private Sub smncadTipo_Atendimento_Servico_Click()
    frmTipo_Atendimento_Servico.Show
End Sub

Private Sub smnMarcha_Click()
    frmMarcha_Analitica.Show
End Sub

Private Sub smnFuncao_Insumo_Click()
    frmFuncao_Insumo.Show
End Sub

Private Sub smnInsumo_Click()
    frmInsumo.Show
End Sub

Private Sub smnmovIntegracao_retaguarda_portal_Click()
    frmIntegracao_retaguarda_portal.Show
End Sub

Private Sub smnPrioridade_Pendencia_Click()
    frmPrioridade_Pendencia.Show
End Sub

Private Sub smnrelTriagem_Click()
    frmRelatorio_Triagem.Show
End Sub

Private Sub smnSoftware_pendencias_Click()
    frmSoftware_Pendencias.Show
End Sub

Private Sub smnStatus_Pendencia_Click()
    frmStatus_Pendencia.Show
End Sub

Private Sub smnTipo_Marcha_Click()
    frmTipo_Marcha.Show
End Sub

Private Sub smnSobre_Click()
    frmSobre.Show
End Sub

Private Sub smnPlano_servicos_Click()
    frmPlano_Servicos.Show
End Sub

Private Sub smnContrato_servicos_Click()
    frmContrato_Servico.Show
End Sub

Private Sub smnmovGera_Plano_Completo_Click()
    frmMovimentacoes_Gera_Plano_Completo.Show
End Sub

Private Sub smnParametros_Servicos_Click()
    frmParametros_Servicos.Show
End Sub

Private Sub smnServicos_Click()
    frmServicos.Show
End Sub

Private Sub smnEquipamento_Laboratorio_Click()
    frmEquipamento_Laboratorio.Show
End Sub

Private Sub smnTipo_Servico_Click()
    frmTipo_Servico.Show
End Sub

Private Sub smnRelatorio_Senha_click()
    frmRelatorio_Senha.Show
End Sub

Private Sub smnProcessamento_Senha_Portal_click()
    frmMovimentacoes_Processamento_Senha_Portal.Show
End Sub

Private Sub smnTriagem_Laboratorio_Click()
    frmTriagem_Laboratorio.Show
End Sub

Private Sub smnmovGeracao_Ordem_Servico_Click()
    frmMovimentacoes_Geracao_Ordem_Servico.Show
End Sub

Private Sub smnrelPendencias_Click()
    frmRelatorio_Pendencias.Show
End Sub

Private Sub Access()

    Dim strMensagem_cliente() As String
    Dim mensagem_design As String
    
    On Error GoTo Erro
   
    'Marcar sempre essa variavel com false quando for compilar e testar no admin
    booDesign_time = True
    
    If booDesign_time = True Then
       'mensagem_design = "ESTA��OTESTE�MARCOS�SENHA_TESTE�marcos�4�AREA_TRABALHO_TESTE�100"
       mensagem_design = "ESTA��OTESTE�MARCOS�SENHA_TESTE�marcos�1�AREA_TRABALHO_TESTE�100"
       strMensagem_cliente = Split(mensagem_design, "�")
       MDIPrincipal.OCXUsuario.Empresa = strMensagem_cliente(6)
    Else
       If Cliente_mensagem_exe.MensagemRecebida = "" Then
          'Tentativa de acessar inf dos usu�rios
          '-------------------------------------------------------------------------------------------
          'Log
          log.Tipo = 4    ' Tipo de Log de uso da Only Tech
          log.Data = Date
          log.Hora = Format(Now, "hh:mm:ss")
          log.Descricao = "Sistema acessando registro da m�quina na tentativa de obter inf. do usu�rio."
          'Gravando o log
          log.Gravar_log "Otica", Me
          '--------------------------------------------------------------------------------------------
          Dim strMensagem_Registro As String
        
          strMensagem_Registro = Movimentacoes.Consulta_Contingencia_Acessibilidade("Otica")
         
          If strMensagem_Registro = "" Then
             'Falha nas 2 primeiras tentativas, sistema impossibilitado de acessar inf. do usu�rio
             '-------------------------------------------------------------------------------------------
             'Log
             log.Tipo = 4    ' Tipo de Log de uso da Only Tech
             log.Data = Date
             log.Hora = Format(Now, "hh:mm:ss")
             log.Descricao = "Falha no acesso a mem�ria e ao registro da m�quina, sistema impossibilitado de acessar inf. do usu�rio."
             'Gravando o log
             log.Gravar_log "Otica", Me
             '--------------------------------------------------------------------------------------------
              MsgBox "Acessibilidade - Ocorreu uma falha de execu��o interna do aplicativo,reexecute o mesmo,se o problema persistir contacte Only Tech Solutions", vbInformation, "Only Tech"
              Exit Sub
          End If
          strMensagem_cliente = Split(strMensagem_Registro, "�")
          MDIPrincipal.OCXUsuario.Empresa = strMensagem_cliente(11)
       Else
          strMensagem_cliente = Split(Cliente_mensagem_exe.MensagemRecebida, "�")
          MDIPrincipal.OCXUsuario.Empresa = strMensagem_cliente(11)
       End If
    End If

    
    OCXUsuario.Nome = strMensagem_cliente(3)
    OCXUsuario.Estacao = strMensagem_cliente(0)
    OCXUsuario.Codigo = strMensagem_cliente(4)
    
    'Informa��es Variaveis para o log
    log.Evento = "Load"
    log.Tipo = 1
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")
   
    log.Descricao = "Inicializando o M�dulo de Servi�os"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    Call Acessibilidade
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Access")
    Exit Sub
End Sub
Private Function Acessibilidade()
    'Servi�os
    Movimentacoes.Acessibilidade_Item_Menu "Servi�os", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnServicos
    'Plano de Servi�os
    Movimentacoes.Acessibilidade_Item_Menu "Plano de Servi�os", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnPlano_servicos
    'Contrato de Servi�os
    Movimentacoes.Acessibilidade_Item_Menu "Contrato de Servi�os", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnContrato_servicos
    'Gera Plano Completo
    Movimentacoes.Acessibilidade_Item_Menu "Gera Plano Completo", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnmovGera_Plano_Completo
    'Insumo
    Movimentacoes.Acessibilidade_Item_Menu "Insumo", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnInsumo
    'Fun��o Insumo
    Movimentacoes.Acessibilidade_Item_Menu "Fun��o Insumo", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnFuncao_Insumo
    'Tipo Marcha
    Movimentacoes.Acessibilidade_Item_Menu "Tipo Marcha", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnTipo_Marcha
    'Marcha Anal�tica
    Movimentacoes.Acessibilidade_Item_Menu "Marcha Anal�tica", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnMarcha
    'Par�metros de Servi�os
    Movimentacoes.Acessibilidade_Item_Menu "Par�metros de Servi�os", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnParametros_Servicos
    'Equipamento Laborat�rio
    Movimentacoes.Acessibilidade_Item_Menu "Equipamento Laborat�rio", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnEquipamento_Laboratorio
    'Tipo Servi�o
    Movimentacoes.Acessibilidade_Item_Menu "Tipo Servi�o", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnTipo_Servico
    'Status Pend�ncia
    Movimentacoes.Acessibilidade_Item_Menu "Status Pend�ncia", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnStatus_Pendencia
    'Prioridade Pend�ncia
    Movimentacoes.Acessibilidade_Item_Menu "Prioridade Pend�ncia", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnPrioridade_Pendencia
    'Relat�rio Senha
    Movimentacoes.Acessibilidade_Item_Menu "Relat�rio Senha", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnRelatorio_Senha
    'Movimenta��es Processamento de senha Portal
    Movimentacoes.Acessibilidade_Item_Menu "Processamento de Senha Portal", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnProcessamento_Senha_Portal
    'Pend�ncias
    Movimentacoes.Acessibilidade_Item_Menu "Pend�ncias", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnSoftware_pendencias
    'Relat�rio Triagem
    Movimentacoes.Acessibilidade_Item_Menu "Relat�rio Triagem", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnrelTriagem
    'Triagem Laborat�rio
    Movimentacoes.Acessibilidade_Item_Menu "Triagem Laborat�rio", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnTriagem_Laboratorio
    'Gera��o de Ordem Servi�o
    Movimentacoes.Acessibilidade_Item_Menu "Gera��o de Ordem Servi�o", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnmovGeracao_Ordem_Servico
    'Relat�rio Pend�ncias
    Movimentacoes.Acessibilidade_Item_Menu "Relat�rio Pend�ncias", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnrelPendencias
    'Integra��o Retaguarda X Portal
    Movimentacoes.Acessibilidade_Item_Menu "Integra��o Retaguarda X Portal", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnmovIntegracao_retaguarda_portal
    'Tipo Atendimento Servi�o
    Movimentacoes.Acessibilidade_Item_Menu "Tipo Atendimento Servi�o", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smncadTipo_Atendimento_Servico
End Function
