VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConexoes_concentrador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexões ao Concentrador Only Tech"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConexoes_concentrador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   3090
      Top             =   3330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   2490
      Top             =   3330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1414
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdExecutar 
      Caption         =   "Parar Servidor"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Frame fraLog 
      Caption         =   "Log de Autenticações no Concentrador"
      Height          =   3195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6975
      Begin RichTextLib.RichTextBox rtbLog 
         Height          =   2805
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4948
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmConexoes_concentrador.frx":1782
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmConexoes_concentrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub StartServer()
    sckListen.Listen 'Aguardando por conexões
    rtbLog.Text = " - Servidor iniciado com sucesso!"  'Adicionando ao Log viewer
End Sub
Private Sub StopServer()
    sckListen.Close
    'Para desconectar todos os usuários conectados
    For i = 0 To MAX_USERS
        If User(i).FreeSocket = False Then 'Socket em uso
        User(i).FreeSocket = True 'Reseta a variavel
        User(i).HasAuthenticated = False
        sckServer(i).Close 'Mata a conexão
    End If
Next i

    rtbLog.Text = " - Servidor paralizado com sucesso!"

End Sub
    
Private Sub cmdSair_Click()
    StopServer
    Unload Me
    End
End Sub
    
Private Sub cmdExecutar_Click()
    If cmdExecutar.Caption = "Iniciar Servidor" Then
        StartServer
        cmdExecutar.Caption = "Parar Servidor"
    Else
        StopServer
        cmdExecutar.Caption = "Iniciar Servidor"
    End If
End Sub

Private Sub Form_Load()
    For i = 1 To MAX_USERS
        User(i).FreeSocket = True
        Load sckServer(i)
    Next i
    
    User(0).FreeSocket = True
    
    StartServer
    
End Sub
    
Private Sub rtbLog_Change()

    rtbLog.SelStart = Len(rtbLog)
    
End Sub
    
Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
    Dim i As Integer
    Dim strAuthString As String
    
    For i = 0 To MAX_USERS
        If User(i).FreeSocket = True Then
            sckServer(i).Accept requestID
            User(i).FreeSocket = False
            rtbLog.Text = " - Servidor aceitou conexão usabdo socket ID: " & i & " e com IP " & sckServer(i).RemoteHostIP & "."
            DoEvents
            
            strAuthString = GenerateAuthString(i)
            
            SendData GenerateAuthString(i), i, Me
            Exit Sub
        End If
    Next i
    
    rtbLog.Text = " - Usuário com IP " & sckListen.RemoteHostIP & " teve sua conexão negada por falta de sockets liberados."
    sckListen.Close
    sckListen.Listen
    
End Sub
    
Private Sub sckServer_Close(Index As Integer)
    sckServer(Index).Close
    User(Index).FreeSocket = True
    User(Index).HasAuthenticated = False
    rtbLog.Text = " - Socket ID: " & Index & " Conexão(ões) fechadas."
End Sub
    
Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String, SplitData() As String, SplitRequest() As String
    Dim strAuthString As String
    On Error GoTo errServer
    
    sckServer(Index).GetData strData
    LogRAW strData
    
    SplitData = Split(strData, DATA_DELIMITER)
    
    For i = 0 To UBound(SplitData) - 1
        SplitRequest = Split(SplitData(i), "|")
        
        If User(Index).HasAuthenticated = False Then
        
            If CheckAuthentication(SplitRequest(0), Index) = True Then
               User(Index).HasAuthenticated = True
               rtbLog.Text = " - Socket ID: " & Index & " Autenticado com sucesso!"
               SendData "Autenticado com sucesso!", Index, Me
            Else
               User(Index).HasAuthenticated = False
               SendData "Autenticação negada!", Index, Me
               rtbLog.Text = " - Socket ID: " & Index & " Não conectado(String de autenticação inválida)"
               DisconnectUser Index, Me
            End If
        End If
    Next i

errServer:

End Sub
    
Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckServer_Close Index
End Sub
    
