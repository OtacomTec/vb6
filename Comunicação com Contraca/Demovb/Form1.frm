VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   6615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   6615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Interrompe comunicação com a catraca"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   6615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4200
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inicia comunicação com a catraca"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "String enviada pelo buffer de saída"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "String recebida no buffer de entrada"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                                         'Declaração das variáveis globais ...
Private MDado As String
Private MCatraca As String
Private MCracha As String
Private MSentido As String
Private Sub Command1_Click()
                                                            'Inicialização dos parâmetros necessários a operação do pooling
    ActiveDll                                               'Ativa a dll alocando memória para sua operação
    SetComm 2                                               'Seleciona a porta serial do microcomputador como sendo a COM2
    SetBaudRate 4800                                        'Configura a velocidade de comunicação para 4800 BPS
    OpenComm                                                'Abre a porta serial
    InsertTerminal 1                                        'Insere a catraca 1 no pooling
    EnableTerminal 1                                        'Habilita a catraca 1
    SetPoolingIntervalTime 100                              'Configura o intervalo de pooling para 100ms
    SetTerminalResponseTime 500                             'Configura o tempo de aguardo da resposta para 500ms
    SetTerminalTimeOut "01,2000"                            'Configura o tempo de aguardo da catraca para 2s
    SetConditionAfterTimeOut "01,B"                         'Configura a catraca para bloquear o acesso caso o computador não responda antes dos 2s
    SetDateTime "01,01/01/2001 12:00:00"                    'Configura a data e hora do relógio da catraca
    SendMessage "01,2000,Terminal 1      ,inicializado... " 'Envia mensagem a catraca
    StartPooling                                            'Inicializa o pooling
    Timer1.Enabled = True                                   'Habilita rotina interna de tratamento do pooling
    
End Sub
Private Sub Command2_Click()
                                                            'Finalização do pooling
    Timer1.Enabled = False                                  'Desabilita rotina interna de tratamento do pooling
    StopPooling                                             'Interrompe o pooling
    CloseComm                                               'Fecha a porta serial
    DeactiveDll                                             'Desativa a dll liberando a memória alocada por esta
    
End Sub
Private Sub Timer1_Timer()
                                                            'Rotina de tratamento do buffer de entrada de dados...
                                                            'Este exemplo pressupõe que uma catraca opere controlando acesso somente
                                                            'como ENTRADA, sem no entanto fazer criticas restritivas ao acesso...
    MDado = Question
    If MDado <> "" Then
        MCatraca = Mid(MDado, 1, 2)
        MCracha = Mid(MDado, 4, 10)
        MSentido = Mid(MDado, 15, 1)
        Text1.Text = MDado
        If MSentido = "E" Then                           'Se o sentido de passagem do crachá for entrada, libera o crachá
            MDado = MDado + ",L,2000," + MCracha + "....OK,Cracha..Liberado"
        Else                                             'Se não, bloqueia o crachá
            MDado = MDado + ",B,2000," + MCracha + "....OK,Cracha.Bloqueado"
        End If
        Text2.Text = MDado
        Answer MDado
    End If
    
End Sub

