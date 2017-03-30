VERSION 5.00
Begin VB.Form frmAdminDeskTopCliente 
   BackColor       =   &H80000001&
   ClientHeight    =   8565
   ClientLeft      =   1710
   ClientTop       =   1950
   ClientWidth     =   11100
   Icon            =   "frmAdminDeskTopCliente.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   11100
   Visible         =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmAdminDeskTopCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarProgramas() As String 'local copy
Private mvarID_Usuário As Integer
Private mvarUsuário As String 'local copy
Private mvarSenha As String 'local copy
Private mvarTotaldeProgramas As Integer 'local copy
Public Property Let TotaldeProgramas(ByVal vData As Integer)
    mvarTotaldeProgramas = vData
End Property

Public Property Get TotaldeProgramas() As Integer
    TotaldeProgramas = mvarTotaldeProgramas
End Property

Public Sub AdicionarPrograma(NomeDoPrograma As String, IDdoPrograma As Long, PIDdoPrograma As Long)
    Dim x As Integer
    x = UBound(mvarProgramas, 2) + 1
    ReDim Preserve mvarProgramas(2, x)
    mvarProgramas(0, x) = NomeDoPrograma
    mvarProgramas(1, x) = IDdoPrograma '+ 1
    mvarProgramas(2, x) = PIDdoPrograma '+ 1
    AtualizarProgramas
End Sub
Public Sub RemoverAplicativo(IDdoPrograma As Long)
    Dim i As Integer
    For i = 1 To UBound(mvarProgramas, 2)
        If mvarProgramas(2, i) = IDdoPrograma Then
            mvarProgramas(2, i) = Empty
            mvarProgramas(1, i) = Empty
            mvarProgramas(0, i) = Empty
        End If
    Next i
    AtualizarProgramas
End Sub
Private Sub AtualizarProgramas()
    Dim i As Integer
    Dim x As Integer
    Dim mtTemp() As String
    ReDim mtTemp(2, 0)
    
    For i = 1 To UBound(mvarProgramas, 2)
        If mvarProgramas(0, i) <> Empty Then
            If mvarProgramas(1, i) <> Empty Then
                x = UBound(mtTemp, 2) + 1
                ReDim Preserve mtTemp(2, x)
                mtTemp(0, x) = mvarProgramas(0, x)
                mtTemp(1, x) = mvarProgramas(1, x)
                mtTemp(2, x) = mvarProgramas(2, x)
            End If
        End If
    Next i
    mvarTotaldeProgramas = UBound(mvarProgramas, 2)
End Sub

Public Property Let Senha(ByVal vData As String)
    mvarSenha = vData
End Property


Public Property Get Senha() As String
    Senha = mvarSenha
End Property

Public Property Let ID_Usuário(ByVal vData As Integer)
    mvarID_Usuário = vData
End Property

Public Property Get ID_Usuário() As Integer
    ID_Usuário = mvarID_Usuário
End Property


Public Property Let Usuário(ByVal vData As String)
    mvarUsuário = vData
End Property

Public Property Get Usuário() As String
    Usuário = mvarUsuário
End Property

'Private Property Let Programas(ByVal vData As String)
    'A propriedade Programas armazenará seus dados em forma de vetor (matriz)
    'Este vetor será composto de duas colunas a iniciar pelo zero, onde serão
    'armazenados na coluna 0 o Nome do Programa Aberto e na coluna 1 o ID
    '(identificador hwnd do Windows da janela aberta).
    'Esta variável matriz estará dimensionada com zero (0) quando não houver
    'programas abertos para que a função Ubound retorne o total de programas
    'abertos sem maiores problemas.
    'No caso, para se obter o nome do programa armazenado na segunda linha da
    'matriz temos a linha seguinte:
    'strNomeDoPrograma = mvarProgramas(0, 2)
    'strIDDoPrograma = mvarProgramas(1, 2)
    'strPIDdoPrograma = mvarProgramas(2, 2)
    'mvarProgramas = vData
'End Property
Private Property Get Programas() As String
    Programas = mvarProgramas
End Property

Function Programa(Coluna As Integer, linha As Integer) As String
    Programa = mvarProgramas(Coluna, linha)
End Function

Private Sub Form_Activate()
    If frmAdminMDI.AplicativoUsuário(0).NomeReduzido = Empty Then
        Me.WindowState = 1
        Me.Enabled = False
        Exit Sub
    End If
    If frmAdminMDI.AplicativoUsuário(0).NomeReduzido = Me.Usuário Then
        frmAdminMDI.AplicativoUsuário(0).Nome = frmAdminMDI.AplicativoUsuário(Me.ID_Usuário).Nome
        frmAdminMDI.AplicativoUsuário(0).NomeReduzido = frmAdminMDI.AplicativoUsuário(Me.ID_Usuário).NomeReduzido
        frmAdminMDI.AplicativoUsuário(0).Senha = frmAdminMDI.AplicativoUsuário(Me.ID_Usuário).Senha
        frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho = frmAdminMDI.AplicativoUsuário(Me.ID_Usuário).ÁreaDeTrabalho
        Set frmAdminMDI.AplicativoUsuário(0).Janela = frmAdminMDI.AplicativoUsuário(Me.ID_Usuário).Janela
    Else
    
        'AT.AlternarPara frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, EDJ_Minimizada
        Me.WindowState = 1
        Me.Enabled = False
        If frmAdminMDI.AplicativoUsuário(0).Janela.WindowState = 0 Then AT.AlternarPara frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, EDJ_Normal
        If frmAdminMDI.AplicativoUsuário(0).Janela.WindowState = 1 Then AT.AlternarPara frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, EDJ_Minimizada
        If frmAdminMDI.AplicativoUsuário(0).Janela.WindowState = 2 Then AT.AlternarPara frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, EDJ_Maximizada
        
        
        'x Form_Deactivate
    End If
    
End Sub

Private Sub Form_Load()
    ReDim mvarProgramas(2, 0)
    
   
    
    'AT.CriarÁreaDeTrabalho Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'UnloadMode
    'vbFormControlMenu 0    O usuário escolheu o comando Close no menu Control do formulário.
    'VbFormCode 1           A instrução Unload é chamada a partir de código.
    'VbAppWindows 2         A sessão de ambiente operacional atual do Microsoft Windows está terminando.
    'VbAppTaskManager 3     O Gerenciador de tarefas do Microsoft Windows está fechando o aplicativo.
    'vbFormMDIForm 4        Um formulário MDI filho está fechando porque o formulário MDI está sendo fechado.
    'vbFormOwner 5          Um formulário está fechando porque seu proprietário está sendo fechado.
    Dim mtzUsuáriosLogados() As String
    
    If UnloadMode = 1 Then
        Exit Sub
        
    End If
    
    If frmAdminMDI.AplicativoUsuário(0).NomeReduzido = Empty Then
        If frmAdminMDI.WindowState = 0 Then AT.AlternarPara frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, EDJ_Normal
        If frmAdminMDI.WindowState = 1 Then AT.AlternarPara frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, EDJ_Minimizada
        If frmAdminMDI.WindowState = 2 Then AT.AlternarPara frmAdminMDI.AplicativoUsuário(0).ÁreaDeTrabalho, EDJ_Maximizada
        Exit Sub
    End If
    
    Dim msg As String
    If Me.Caption <> "" Then
        If frmAdminMDI.AplicativoUsuário(0).Nome = Me.Usuário Then
            msg = msg & "Essa Área de Trabalho está logada com o usuário <"
            msg = msg & Me.Usuário & ">. "
            msg = msg & "Você deseja fechar sua Área de Trabalho e todos os seus "
            msg = msg & "Aplicativos abertos interrompendo os seus processos?"
            
        Else
            msg = msg & "Essa Área de Trabalho está logada com o usuário "
            msg = msg & Me.Usuário & ". "
            msg = msg & "O Usuário nesse momento é " & frmAdminMDI.AplicativoUsuário(0).Nome & ". Você deseja "
            msg = msg & "fechar a Área de Trabalho do usuário " & Me.Usuário & " e todos os Aplicativos abertos interrompendo os seus processos "
            msg = msg & "assim mesmo?"
        End If
        r = MsgBox(msg, vbYesNo)
        If r = vbNo Then
            Cancel = True
        Else
            'FecharProcessosAbertos
            FecharÁreaDeTrabalho Me
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)

    'DestroyWindow mWnd
    'TerminateProcess GetCurrentProcess, 0
End Sub

