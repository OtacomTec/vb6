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
Private mvarID_Usu�rio As Integer
Private mvarUsu�rio As String 'local copy
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

Public Property Let ID_Usu�rio(ByVal vData As Integer)
    mvarID_Usu�rio = vData
End Property

Public Property Get ID_Usu�rio() As Integer
    ID_Usu�rio = mvarID_Usu�rio
End Property


Public Property Let Usu�rio(ByVal vData As String)
    mvarUsu�rio = vData
End Property

Public Property Get Usu�rio() As String
    Usu�rio = mvarUsu�rio
End Property

'Private Property Let Programas(ByVal vData As String)
    'A propriedade Programas armazenar� seus dados em forma de vetor (matriz)
    'Este vetor ser� composto de duas colunas a iniciar pelo zero, onde ser�o
    'armazenados na coluna 0 o Nome do Programa Aberto e na coluna 1 o ID
    '(identificador hwnd do Windows da janela aberta).
    'Esta vari�vel matriz estar� dimensionada com zero (0) quando n�o houver
    'programas abertos para que a fun��o Ubound retorne o total de programas
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
    If frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = Empty Then
        Me.WindowState = 1
        Me.Enabled = False
        Exit Sub
    End If
    If frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = Me.Usu�rio Then
        frmAdminMDI.AplicativoUsu�rio(0).Nome = frmAdminMDI.AplicativoUsu�rio(Me.ID_Usu�rio).Nome
        frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = frmAdminMDI.AplicativoUsu�rio(Me.ID_Usu�rio).NomeReduzido
        frmAdminMDI.AplicativoUsu�rio(0).Senha = frmAdminMDI.AplicativoUsu�rio(Me.ID_Usu�rio).Senha
        frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho = frmAdminMDI.AplicativoUsu�rio(Me.ID_Usu�rio).�reaDeTrabalho
        Set frmAdminMDI.AplicativoUsu�rio(0).Janela = frmAdminMDI.AplicativoUsu�rio(Me.ID_Usu�rio).Janela
    Else
    
        'AT.AlternarPara frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho, EDJ_Minimizada
        Me.WindowState = 1
        Me.Enabled = False
        If frmAdminMDI.AplicativoUsu�rio(0).Janela.WindowState = 0 Then AT.AlternarPara frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho, EDJ_Normal
        If frmAdminMDI.AplicativoUsu�rio(0).Janela.WindowState = 1 Then AT.AlternarPara frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho, EDJ_Minimizada
        If frmAdminMDI.AplicativoUsu�rio(0).Janela.WindowState = 2 Then AT.AlternarPara frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho, EDJ_Maximizada
        
        
        'x Form_Deactivate
    End If
    
End Sub

Private Sub Form_Load()
    ReDim mvarProgramas(2, 0)
    
   
    
    'AT.Criar�reaDeTrabalho Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'UnloadMode
    'vbFormControlMenu 0    O usu�rio escolheu o comando Close no menu Control do formul�rio.
    'VbFormCode 1           A instru��o Unload � chamada a partir de c�digo.
    'VbAppWindows 2         A sess�o de ambiente operacional atual do Microsoft Windows est� terminando.
    'VbAppTaskManager 3     O Gerenciador de tarefas do Microsoft Windows est� fechando o aplicativo.
    'vbFormMDIForm 4        Um formul�rio MDI filho est� fechando porque o formul�rio MDI est� sendo fechado.
    'vbFormOwner 5          Um formul�rio est� fechando porque seu propriet�rio est� sendo fechado.
    Dim mtzUsu�riosLogados() As String
    
    If UnloadMode = 1 Then
        Exit Sub
        
    End If
    
    If frmAdminMDI.AplicativoUsu�rio(0).NomeReduzido = Empty Then
        If frmAdminMDI.WindowState = 0 Then AT.AlternarPara frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho, EDJ_Normal
        If frmAdminMDI.WindowState = 1 Then AT.AlternarPara frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho, EDJ_Minimizada
        If frmAdminMDI.WindowState = 2 Then AT.AlternarPara frmAdminMDI.AplicativoUsu�rio(0).�reaDeTrabalho, EDJ_Maximizada
        Exit Sub
    End If
    
    Dim msg As String
    If Me.Caption <> "" Then
        If frmAdminMDI.AplicativoUsu�rio(0).Nome = Me.Usu�rio Then
            msg = msg & "Essa �rea de Trabalho est� logada com o usu�rio <"
            msg = msg & Me.Usu�rio & ">. "
            msg = msg & "Voc� deseja fechar sua �rea de Trabalho e todos os seus "
            msg = msg & "Aplicativos abertos interrompendo os seus processos?"
            
        Else
            msg = msg & "Essa �rea de Trabalho est� logada com o usu�rio "
            msg = msg & Me.Usu�rio & ". "
            msg = msg & "O Usu�rio nesse momento � " & frmAdminMDI.AplicativoUsu�rio(0).Nome & ". Voc� deseja "
            msg = msg & "fechar a �rea de Trabalho do usu�rio " & Me.Usu�rio & " e todos os Aplicativos abertos interrompendo os seus processos "
            msg = msg & "assim mesmo?"
        End If
        r = MsgBox(msg, vbYesNo)
        If r = vbNo Then
            Cancel = True
        Else
            'FecharProcessosAbertos
            Fechar�reaDeTrabalho Me
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)

    'DestroyWindow mWnd
    'TerminateProcess GetCurrentProcess, 0
End Sub

