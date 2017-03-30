VERSION 5.00
Object = "{06DDD466-EE4A-11D6-9F91-000102C349D1}#1.1#0"; "SegundoPlanoMDI.ocx"
Object = "{D3F9E3A8-F26B-11D6-9F91-000102C349D1}#2.2#0"; "AplicativoUsuário.ocx"
Begin VB.MDIForm MDIPrincipal_Mercadolologia 
   BackColor       =   &H8000000C&
   Caption         =   "Mercadologia"
   ClientHeight    =   8190
   ClientLeft      =   3630
   ClientTop       =   2850
   ClientWidth     =   11880
   Icon            =   "MDIPrincipal_Mercadolologia.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin AplicativoUsuárioOCX.AplicativoUsuário UsuárioOCX 
      Left            =   4920
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin SegundoPlano.SegundoPlanoMDI SegundoPlano 
      Left            =   5400
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer tmrIntercomunicador 
      Interval        =   500
      Left            =   5880
      Top             =   360
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastros"
      Begin VB.Menu mnuCategoria 
         Caption         =   "&Categoria"
      End
      Begin VB.Menu mnuEstrutura 
         Caption         =   "&Estrutura Mercadológica"
      End
      Begin VB.Menu mnuEmpresa 
         Caption         =   "&Empresa"
      End
      Begin VB.Menu mnuGrupo 
         Caption         =   "&Grupo"
      End
      Begin VB.Menu mnuproduto 
         Caption         =   "&Produto"
      End
      Begin VB.Menu mnuSubgrupo 
         Caption         =   "&Sub Grupo"
      End
   End
   Begin VB.Menu mnuMovimentacoes 
      Caption         =   "&Movimentações"
   End
   Begin VB.Menu mnuRelatorio 
      Caption         =   "&Relatorios"
   End
End
Attribute VB_Name = "MDIPrincipal_Mercadolologia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comunicacao_cliente As VetorDeMensagens.ClienteDeMensagens
Public strEstação As String
Dim booProjeto As Boolean
Private Sub MDIForm_Load()

    'Atribuindo booprojeto = true, estabeleço o uso da aplicação de forma que ela não use
    'o intercomunicador que está no módulo admin.No caso do False, atribuir quando for
    'compilar o EXE.
    booProjeto = False
    
    If booProjeto = False Then
        Set comunicacao_cliente = New VetorDeMensagens.ClienteDeMensagens
        comunicacao_cliente.Interceptar
        tmrIntercomunicador.Enabled = True
    Else
        tmrIntercomunicador.Enabled = False
        strEstação = "INFO-028"
        UsuárioOCX.NomeReduzido = "Marcos"
        UsuárioOCX.Senha = "1234"
        UsuárioOCX.Nome = "Marcos Baião"
        'Luis insira aqui a parte de privilégios de acesso
    End If
    
    Set SegundoPlano.FormulárioMDI = Me
    SegundoPlano.AutoAtualizar = True
    SegundoPlano.DistânciaDaBorda = 10
    SegundoPlano.Cor(CorEmCima_enPDC) = vbWhite
    SegundoPlano.Cor(CorEmBaixo_enPDC) = vbBlue
    'SegundoPlano.ArquivoDaImagem = "X:\Projetos\Gestor Mil\Logomarca\Gestor_trabalho.jpg"
    SegundoPlano.EstiloDoFundo = FundoGradiente_enMDIF
    SegundoPlano.PosiçãoDaFigura = FiguraAjustada_enPF
         
End Sub

Private Sub mnuCategoria_Click()
    frmCategoria.Show
End Sub

Private Sub mnuCidade_Click()
    frmCidade.Show
End Sub

Private Sub mnuEmpresa_Click()
    frmEmpresa.Show
End Sub

Private Sub mnuEstrutura_mercadologica_Click()
    frmEstrutura.Show
End Sub

Private Sub mnuGrupo_Click()
    frmGrupo.Show
End Sub

Private Sub mnuSecao_Click()
    frmSecao.Show
End Sub

Private Sub mnuSubgrupo_Click()
    frmSubgrupo.Show
End Sub

Private Sub tmrIntercomunicador_Timer()

    'If comunicacao_cliente.MensagemRecebida <> "" Then
        Dim Mensagens() As String
        Mensagens = Split(comunicacao_cliente.MensagemRecebida, "¤")
        tmrIntercomunicador.Enabled = False
        strEstação = Mensagens(0)
        UsuárioOCX.NomeReduzido = Mensagens(1)
        UsuárioOCX.Senha = Mensagens(2)
        UsuárioOCX.Nome = Mensagens(3)
        'UsuárioOCX.ÁreaDeTrabalho = Mensagens(4)
        UsuárioOCX.PrivilégioAcessar = Mensagens(5)
        UsuárioOCX.PrivilégioAlterar = Mensagens(6)
        UsuárioOCX.PrivilégioConsultar = Mensagens(7)
        UsuárioOCX.PrivilégioExcluir = Mensagens(8)
        UsuárioOCX.PrivilégioIncluir = Mensagens(9)
        MsgBox Mensagens
    'End If
  
End Sub
