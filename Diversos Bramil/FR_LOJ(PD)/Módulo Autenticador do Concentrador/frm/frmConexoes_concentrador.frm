VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConexoes_concentrador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexões ao Concentrador Only Tech"
   ClientHeight    =   3825
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7200
   Icon            =   "frmConexoes_concentrador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pichook 
      Height          =   405
      Left            =   1470
      ScaleHeight     =   345
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   3090
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   2640
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1414
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdToggleServer 
      Caption         =   "&Parar Servidor"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Frame fraLog 
      Caption         =   "Log do Autenticador do Concentrador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   120
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
            Name            =   "AvantGarde Bk BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   900
      Picture         =   "frmConexoes_concentrador.frx":1807
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuAutenticador 
      Caption         =   "&Autenticador"
      Begin VB.Menu smnuAutenticador 
         Caption         =   "&Iniciar"
         Index           =   0
      End
      Begin VB.Menu smnuAutenticador 
         Caption         =   "&Parar"
         Index           =   1
      End
      Begin VB.Menu smnuAutenticador 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu smnuAutenticador 
         Caption         =   "&Log de Autenticações"
         Index           =   3
      End
      Begin VB.Menu smnuAutenticador 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu smnuAutenticador 
         Caption         =   "&Sair"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmConexoes_concentrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim TrayI As NOTIFYICONDATA
Private Sub StartServer()
    sckListen.Listen
    log "Servidor iniciado com sucesso!"
    
End Sub
Private Sub StopServer()
    sckListen.Close
    For I = 0 To MAX_USERS
        If User(I).FreeSocket = False Then
        User(I).FreeSocket = True
        User(I).HasAuthenticated = False
        sckServer(I).Close
    End If
Next I

log "Servidor parado com sucesso!"

End Sub
    
Private Sub cmdQuit_Click()
    StopServer
    Unload Me
    End
End Sub
    
Private Sub cmdToggleServer_Click()
    If cmdToggleServer.Caption = "&Iniciar Servidor" Then
        StartServer
        cmdToggleServer.Caption = "&Parar Servidor"
    Else
        StopServer
        cmdToggleServer.Caption = "&Iniciar Servidor"
    End If
    
End Sub
    
Private Sub Form_Load()

    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd 'Link do icone do systray para o picturebox
    TrayI.uID = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.uCallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = imgIcon(0).Picture
    TrayI.szTip = "Autenticador do Concentrador" & Chr$(0)
    'Criando o icone
    Shell_NotifyIcon NIM_ADD, TrayI
    Me.Hide
    
    Robozinho.ShellIconInitialize Me

    For I = 1 To MAX_USERS
        User(I).FreeSocket = True
        Load sckServer(I)
    Next I
    
    User(0).FreeSocket = True
    
    StartServer
    
End Sub
    
Private Sub rtbLog_Change()

    rtbLog.SelStart = Len(rtbLog)
    
End Sub
    
Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
    Dim I As Integer
    Dim strAuthString As String
    
    Dim rstPDV As New ADODB.Recordset
    
    strSql = Empty
    strSql = "SELECT * FROM TBPDV"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstPDV, "Otica", Me
    
    For I = 0 To MAX_USERS
        If User(I).FreeSocket = True Then
            sckServer(I).Accept requestID
            User(I).FreeSocket = False
            
            rstPDV.MoveFirst
            rstPDV.Find ("DFEndereco_ip_TBPdv = " & sckServer(I).RemoteHostIP & "")
            
            log "Concentrador aceitou conexão do PDV " & rstPDV!PKCodigo_TBPdv & " e com IP " & sckServer(I).RemoteHostIP & "."
            DoEvents
            
            strAuthString = GenerateAuthString(I)
            
            SendData GenerateAuthString(I), I
            Exit Sub
        End If
    Next I
    
    rstPDV.MoveFirst
    rstPDV.Find ("DFEndereco_ip_TBPdv = " & sckServer(I).RemoteHostIP & "")
            
    log "PDV - " & rstPDV!PKCodigo_TBPdv & "  com IP " & sckListen.RemoteHostIP & " foi negada a conexão por não possuir sockets liberadas."
    
    sckListen.Close
    sckListen.Listen
    
    Set rstPDV = Nothing
    
End Sub
    
Private Sub sckServer_Close(Index As Integer)
    sckServer(Index).Close
    User(Index).FreeSocket = True
    User(Index).HasAuthenticated = False
    log "Socket ID: " & Index & " foi fechado a conexão."
End Sub
    
Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String, SplitData() As String, SplitRequest() As String
    Dim strAuthString As String
    Dim rstPDV As New ADODB.Recordset
    
    On Error GoTo errServer
    
    strSql = Empty
    strSql = "SELECT * FROM TBPDV"
    Movimentacoes.Select_geral strSql, "BDRetaguarda", rstPDV, "Otica", Me
    
    sckServer(Index).GetData strData
    LogRAW strData
    
    SplitData = Split(strData, DATA_DELIMITER)
    
    For I = 0 To UBound(SplitData) - 1
        SplitRequest = Split(SplitData(I), "|")
        
        If User(Index).HasAuthenticated = False Then
        
            If CheckAuthentication(SplitRequest(0), Index) = True Then
               User(Index).HasAuthenticated = True
               rstPDV.MoveFirst
               rstPDV.Find ("DFEndereco_ip_TBPdv = " & sckServer(Index).RemoteHostIP & "")
               log "PDV: " & rstPDV!PKCodigo_TBPdv & " foi autenticado com sucesso!"
               SendData "AUTENTICADO", Index
            Else
               User(Index).HasAuthenticated = False
               SendData "NEGADO", Index
               rstPDV.MoveFirst
               rstPDV.Find ("DFEndereco_ip_TBPdv = " & sckServer(Index).RemoteHostIP & "")
               log "PDV: " & rstPDV!PKCodigo_TBPdv & " foi desconectado (String inválida!)"
               DisconnectUser Index
            End If
        End If
Next I

errServer:

End Sub
    
Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckServer_Close Index
End Sub
Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Msg = X / Screen.TwipsPerPixelX
    
    If Msg = WM_LBUTTONDBLCLK Then  'se for dado duplo clique no icone
        smnuAutenticador_Click 0
    ElseIf Msg = WM_RBUTTONUP Then  'clique com o direito
        Me.PopupMenu mnuAutenticador
    End If
    
End Sub
Private Sub smnuAutenticador_Click(Index As Integer)
    Select Case Index
        Case 0
            Call cmdToggleServer_Click
        Case 1
            Call cmdToggleServer_Click
        Case 3
            frmConexoes_concentrador.Show
        Case 5
            Unload Me
    End Select
End Sub
