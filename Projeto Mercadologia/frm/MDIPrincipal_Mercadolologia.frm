VERSION 5.00
Object = "{06DDD466-EE4A-11D6-9F91-000102C349D1}#1.1#0"; "SegundoPlanoMDI.ocx"
Object = "{D3F9E3A8-F26B-11D6-9F91-000102C349D1}#2.2#0"; "AplicativoUsu�rio.ocx"
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
   Begin AplicativoUsu�rioOCX.AplicativoUsu�rio Usu�rioOCX 
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
         Caption         =   "&Estrutura Mercadol�gica"
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
      Caption         =   "&Movimenta��es"
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
Public strEsta��o As String
Dim booProjeto As Boolean
Private Sub MDIForm_Load()

    'Atribuindo booprojeto = true, estabele�o o uso da aplica��o de forma que ela n�o use
    'o intercomunicador que est� no m�dulo admin.No caso do False, atribuir quando for
    'compilar o EXE.
    booProjeto = False
    
    If booProjeto = False Then
        Set comunicacao_cliente = New VetorDeMensagens.ClienteDeMensagens
        comunicacao_cliente.Interceptar
        tmrIntercomunicador.Enabled = True
    Else
        tmrIntercomunicador.Enabled = False
        strEsta��o = "INFO-028"
        Usu�rioOCX.NomeReduzido = "Marcos"
        Usu�rioOCX.Senha = "1234"
        Usu�rioOCX.Nome = "Marcos Bai�o"
        'Luis insira aqui a parte de privil�gios de acesso
    End If
    
    Set SegundoPlano.Formul�rioMDI = Me
    SegundoPlano.AutoAtualizar = True
    SegundoPlano.Dist�nciaDaBorda = 10
    SegundoPlano.Cor(CorEmCima_enPDC) = vbWhite
    SegundoPlano.Cor(CorEmBaixo_enPDC) = vbBlue
    'SegundoPlano.ArquivoDaImagem = "X:\Projetos\Gestor Mil\Logomarca\Gestor_trabalho.jpg"
    SegundoPlano.EstiloDoFundo = FundoGradiente_enMDIF
    SegundoPlano.Posi��oDaFigura = FiguraAjustada_enPF
         
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
        Mensagens = Split(comunicacao_cliente.MensagemRecebida, "�")
        tmrIntercomunicador.Enabled = False
        strEsta��o = Mensagens(0)
        Usu�rioOCX.NomeReduzido = Mensagens(1)
        Usu�rioOCX.Senha = Mensagens(2)
        Usu�rioOCX.Nome = Mensagens(3)
        'Usu�rioOCX.�reaDeTrabalho = Mensagens(4)
        Usu�rioOCX.Privil�gioAcessar = Mensagens(5)
        Usu�rioOCX.Privil�gioAlterar = Mensagens(6)
        Usu�rioOCX.Privil�gioConsultar = Mensagens(7)
        Usu�rioOCX.Privil�gioExcluir = Mensagens(8)
        Usu�rioOCX.Privil�gioIncluir = Mensagens(9)
        MsgBox Mensagens
    'End If
  
End Sub
