VERSION 5.00
Object = "{D0159C1D-A983-4698-8940-3BE45A260C35}#1.0#0"; "SegundoPlanoMDI.ocx"
Object = "{C5014412-BD55-402F-8335-07C273732964}#1.1#0"; "AplicativoUsu�rio.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Concentrador de Vendas"
   ClientHeight    =   8310
   ClientLeft      =   1740
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrincipal.frx":1782
   Begin AplicativoUsu�rioOCX.AplicativoUsu�rio OCXUsuario 
      Left            =   11640
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin OCXSegundoPlano.SegundoPlanoMDI SegundoPlanoMDI 
      Left            =   12240
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastros"
      Begin VB.Menu smnEmpresa 
         Caption         =   "&Empresa"
      End
      Begin VB.Menu smnParametros 
         Caption         =   "&Par�metros"
         Begin VB.Menu smnParametros_fiscal 
            Caption         =   "&Fiscal"
         End
         Begin VB.Menu smnParametros_vendas 
            Caption         =   "&Vendas"
         End
         Begin VB.Menu smnParametros_ecf 
            Caption         =   "&ECF"
         End
      End
      Begin VB.Menu smnPostos 
         Caption         =   "Po&stos de Abastecimento"
         Begin VB.Menu smnTanques 
            Caption         =   "&Tanques"
            Enabled         =   0   'False
         End
         Begin VB.Menu smnBombas 
            Caption         =   "&Bombas"
            Enabled         =   0   'False
         End
         Begin VB.Menu smnEncerrante 
            Caption         =   "&Encerrante"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu smnOperador_Ecf 
         Caption         =   "&Operador ECF"
      End
      Begin VB.Menu smnProduto 
         Caption         =   "P&rodutos"
      End
      Begin VB.Menu smnVendedores 
         Caption         =   "&Vendedores"
      End
      Begin VB.Menu smnFinalizadora 
         Caption         =   "&Finalizadora"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnImpressora_ecf 
         Caption         =   "&Impressora ECF"
      End
      Begin VB.Menu smnPdv 
         Caption         =   "&Pdv"
      End
      Begin VB.Menu smnOperacao_Caixa 
         Caption         =   "Opera&��o Caixa"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnComanda 
         Caption         =   "&Comanda"
      End
      Begin VB.Menu smnCupom 
         Caption         =   "&Cupom Fiscal"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnTabela_Preco 
         Caption         =   "&Tabela de Pre�os"
      End
      Begin VB.Menu smnFechamento_caixa_posto 
         Caption         =   "Fechamento de Caixa"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnNota_Saida_Cupom 
         Caption         =   "Nota Fiscal de Sa�da Cupom"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relat�rios"
      Begin VB.Menu smnrelTira_Teima 
         Caption         =   "Tira Teima"
         Shortcut        =   ^T
      End
      Begin VB.Menu smnrelTesta_impressora 
         Caption         =   "Teste de impressora"
      End
      Begin VB.Menu smnrelBomba 
         Caption         =   "Relat�rio de &Bombas"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnrelEncerrante 
         Caption         =   "Relat�rio de &Encerrantes"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnrelCarteirinha_Gerente 
         Caption         =   "Impress�o de &Carteirinha Gerente"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuMovimentacoes 
      Caption         =   "&Movimenta��es"
      Begin VB.Menu smnmovConsolida_vendas 
         Caption         =   "&Cosolida Vendas "
         Enabled         =   0   'False
      End
      Begin VB.Menu smnmovGera_Integracao_Frente_Loja_Exportacao 
         Caption         =   "Gera Integra��o Frente Loja - Exporta��o"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnmovGera_Integracao_Frente_Loja_importacao 
         Caption         =   "Gera Integra��o Frente Loja - Importa��o"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnmovExportacao_balancas 
         Caption         =   "Exporta��o para balan�as"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnmovEmissao_Nota_Cupom 
         Caption         =   "Emiss�o de Nota Cupom"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnmovEmissao_Nota_Totalizador 
         Caption         =   "Emiss�o de Nota &Totalizador"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuEstatisticas 
      Caption         =   "&Estatisticas"
      Begin VB.Menu smnMovimeneto_Caixa 
         Caption         =   "Fechamento &Opera��o de Caixa"
      End
      Begin VB.Menu smnVenda_Diaria 
         Caption         =   "Fechamento &Di�rio de Vendas"
      End
      Begin VB.Menu smnComanda_Nao_Finalizada 
         Caption         =   "Coman&da N�o Finalizada"
      End
      Begin VB.Menu smnestAnalise_Check_out 
         Caption         =   "�nalise de acompanhamento de check outs"
      End
      Begin VB.Menu smnTira_Teima_Fechamento_Caixa 
         Caption         =   "&Tira Teima Fechamento Caixa"
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
' Only Tech                                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                      '
' M�dulo.................: Concentrador Vendas                                            '
' Objetivo...............: MDI Principal                                                  '
' Data de Cria��o........: 14/01/2005                                                     '
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
    log.Programa = "M�dulo Cadastros Base"
    log.Estacao = OCXUsuario.Estacao

    'Informa��es Variaveis para o log
    log.Evento = "Acessibilidade"
    log.Tipo = 4
    log.Data = Date
    log.Hora = Format(Now, "hh:mm:ss")

    log.Descricao = "Inicializando a acessibilidade do M�dulo Cadastros Base"
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

Private Sub smnTira_Teima_Fechamento_Caixa_Click()
    frmTira_Teima_Fechamento_Caixa.Show
End Sub

Private Sub smnComanda_Click()
    frmComanda.Show
End Sub

Private Sub smnCupom_Click()
    frmCupom.Show
End Sub

Private Sub smnEmpresa_Click()
    frmEmpresa.Show
End Sub

Private Sub smnestAnalise_Check_out_Click()
    frmAnalise_checkouts.Show
End Sub

Private Sub smnFinalizadora_Click()
    frmFinalizadora.Show
End Sub

Private Sub smnImpressora_ecf_Click()
    frmImpressora_Ecf.Show
End Sub

Private Sub smnmovExportacao_balancas_Click()
    frmMovimentacoes_exportacao_balancas.Show
End Sub

Private Sub smnmovGera_Integracao_Frente_Loja_Exportacao_Click()
    frmMovimentacoes_Gera_Integracao_Frente_Loja_Exportacao.Show
End Sub

Private Sub smnmovGera_Integracao_Frente_Loja_importacao_Click()
    frmMovimentacoes_Gera_Integracao_Frente_Loja_Importacao.Show
End Sub

Private Sub smnMovimeneto_Caixa_Click()
    frmRelatorio_Operacao_Caixa.Show
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

Private Sub smnParametros_fiscal_Click()
    frmParametros_Fiscais.Show
End Sub

Private Sub smnPdv_Click()
    frmPdv.Show
End Sub

Private Sub smnrelBomba_Click()
    frmRelatorio_Bombas.Show
End Sub

Private Sub smnrelEncerrante_Click()
    frmRelatorio_Encerrante.Show
End Sub

Private Sub smnrelTesta_impressora_Click()
    frmTeste_impressora.Show
End Sub

Private Sub smnrelTira_Teima_Click()
    FrmTira_Teima.Show
End Sub

Private Sub smnSobre_Click()
    frmSobre.Show
End Sub

Private Sub smnVenda_Diaria_Click()
    frmRelatorio_Fechamento_Diario_Vendas.Show
End Sub

Private Sub smnTanques_Click()
    frmTanques.Show
End Sub

Private Sub smnBombas_Click()
    frmBombas.Show
End Sub

Private Sub smnEncerrante_Click()
    frmEncerrante.Show
End Sub

Private Sub smnFechamento_caixa_posto_Click()
    frmFechamento_caixa_posto.Show
End Sub

Private Sub smnNota_Saida_Cupom_Click()
    frmNota_Saida_Cupom.Show
End Sub

Private Sub smnmovEmissao_Nota_Cupom_Click()
    frmMovimentacoes_Emissao_Nota_Cupom.Show
End Sub

Private Sub smnmovEmissao_Nota_Totalizador_Click()
    frmMovimentacoes_Emissao_Nota_Totalizador.Show
End Sub

Private Sub smnrelCarteirinha_Gerente_Click()
    frmRelatorio_Impressao_Carteirinha_Gerente.Show
End Sub

Private Sub Access()

    Dim strMensagem_cliente() As String
    Dim mensagem_design As String
    
    On Error GoTo Erro
   
    'Marcar sempre essa v ariavel com false quando for compilar e testar no admin
    booDesign_time = True
    
    If booDesign_time = True Then
       'mensagem_design = "ESTA��OTESTE�MARCOS�SENHA_TESTE�marcos�1�AREA_TRABALHO_TESTE�100"
       'mensagem_design = "ESTA��OTESTE�MIX�SENHA_TESTE�mix�1�AREA_TRABALHO_TESTE�100"
       mensagem_design = "ESTA��OTESTE�carlito�SENHA_TESTE�carlito�2�AREA_TRABALHO_TESTE�500"
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
   
    log.Descricao = "Inicializando o M�dulo de Cadastros Base"
    'Gravando o log
    log.Gravar_log "Otica", Me
    
    Call Acessibilidade
    
    Exit Sub
Erro:
    Call Erro.Erro(Me, "OTICA", "Access")
    Exit Sub
    
End Sub
Private Function Acessibilidade()

    'Par�metros ECF
    Movimentacoes.Acessibilidade_Item_Menu "Par�metros ECF", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnParametros_ecf
    'Par�metros Fiscais
    Movimentacoes.Acessibilidade_Item_Menu "Par�metros Fiscais", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnParametros_fiscal
    'Par�metros Venda
    Movimentacoes.Acessibilidade_Item_Menu "Par�metros Venda", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnParametros_vendas
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
    'Opera��o Caixa
    Movimentacoes.Acessibilidade_Item_Menu "Opera��o Caixa", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnOperacao_Caixa
    'Finalizadora
    Movimentacoes.Acessibilidade_Item_Menu "Finalizadora", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnFinalizadora
    'Teste de Impressora
    Movimentacoes.Acessibilidade_Item_Menu "Teste de Impressora", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnrelTesta_impressora
    'Produtos
    Movimentacoes.Acessibilidade_Item_Menu "Produtos", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnProduto
    'Vendedores
    Movimentacoes.Acessibilidade_Item_Menu "Vendedores", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnVendedores
    'Cupom Fiscal
    Movimentacoes.Acessibilidade_Item_Menu "Cupom Fiscal", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnCupom
    'Fechamento Opera��o de Caixa
    Movimentacoes.Acessibilidade_Item_Menu "Fechamento Opera��o de Caixa", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnMovimeneto_Caixa
    'Venda Di�ria
    Movimentacoes.Acessibilidade_Item_Menu "Fechamento Di�rio de Vendas", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnVenda_Diaria
    'Comanda N�o Finalizada
    Movimentacoes.Acessibilidade_Item_Menu "Comanda N�o Finalizada", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnComanda_Nao_Finalizada
    'Tanques
    Movimentacoes.Acessibilidade_Item_Menu "Tanques", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnTanques
    'Bombas
    Movimentacoes.Acessibilidade_Item_Menu "Bombas", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnBombas
    'Encerrante
    Movimentacoes.Acessibilidade_Item_Menu "Encerrante", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnEncerrante
    'Relat�rio de Encerrantes
    Movimentacoes.Acessibilidade_Item_Menu "Relat�rio de Encerrantes", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnrelEncerrante
    'Relat�rio de Bombas
    Movimentacoes.Acessibilidade_Item_Menu "Relat�rio de Bombas", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnrelBomba
    'Gera Integra��o Frente Loja - Exporta��o
    Movimentacoes.Acessibilidade_Item_Menu "Gera Integra��o Frente Loja - Exporta��o", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnmovGera_Integracao_Frente_Loja_Exportacao
    'Gera Integra��o Frente Loja - Importa��o
    Movimentacoes.Acessibilidade_Item_Menu "Gera Integra��o Frente Loja - Importa��o", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnmovGera_Integracao_Frente_Loja_importacao
    'Fechamento de Caixa
    Movimentacoes.Acessibilidade_Item_Menu "Fechamento de Caixa", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnFechamento_caixa_posto
    'Exporta��o para Balan�as
    Movimentacoes.Acessibilidade_Item_Menu "Exporta��o para Balan�as", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnmovExportacao_balancas
    'Nota Fiscal de Sa�da Cupom
    Movimentacoes.Acessibilidade_Item_Menu "Nota Fiscal de Sa�da Cupom", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnNota_Saida_Cupom
    'Emiss�o de Nota Cupom
    Movimentacoes.Acessibilidade_Item_Menu "Emiss�o de Nota Cupom", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnmovEmissao_Nota_Cupom
    'Emiss�o de Nota Totalizador
    Movimentacoes.Acessibilidade_Item_Menu "Emiss�o de Nota Totalizador", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnmovEmissao_Nota_Totalizador
    'Impress�o de Carteirinha Gerente
    Movimentacoes.Acessibilidade_Item_Menu "Impress�o de Carteirinha Gerente", "Otica", "BDRetaguarda", OCXUsuario.Codigo, Me.smnrelCarteirinha_Gerente
End Function
