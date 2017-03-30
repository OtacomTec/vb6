VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmFTP 
   Caption         =   "FTP "
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   8580
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRecebe 
      BackColor       =   &H00C0C0FF&
      Caption         =   "<--"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton cmdEnvia 
      BackColor       =   &H00C0C0FF&
      Caption         =   "-->"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   495
   End
   Begin VB.Frame fraLocal 
      Caption         =   "Local Host"
      Height          =   4215
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   4215
      Begin VB.FileListBox filList 
         Height          =   3600
         Left            =   1920
         MultiSelect     =   2  'Extended
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
      Begin VB.DirListBox dirList 
         Height          =   3015
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.DriveListBox drvList 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraRemote 
      Caption         =   "Host Remoto"
      Height          =   4215
      Left            =   5040
      TabIndex        =   11
      Top             =   2760
      Width           =   3375
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdMkDir 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&MkDir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ListBox lstRemote 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   2595
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   12
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblRemoteDirectory 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   60
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   615
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   5775
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Pronto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.Frame fraFTP 
      Caption         =   "FTP :"
      Height          =   1575
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1620
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1620
         TabIndex        =   6
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Senha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   420
         TabIndex        =   5
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblUsername 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   420
         TabIndex        =   4
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Endereço :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   420
         TabIndex        =   3
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdConecta 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Conectar com o Host Remoto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   5775
   End
   Begin VB.TextBox txtLog 
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7080
      Width           =   8415
   End
   Begin InetCtlsObjects.Inet itcFTP 
      Left            =   7440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdConecta_Click()
    ' Verifica a propriedade Caption para desconectar ou conectar
    If Left(cmdConecta.Caption, 4) = "&Con" Then
        conectaHost
    Else
        desconectaHost
    End If
End Sub

Private Sub cmdDelete_Click()
    
    Dim operacao As String
    Dim nomeArquivo As String
    Dim response As Integer, contador As Integer
    
    response = MsgBox("Confirma a operação ?", vbQuestion + vbYesNo, "Delete")
    If response = vbYes Then
        For contador = 0 To lstRemote.ListCount - 1
            If lstRemote.Selected(contador) = True Then
                nomeArquivo = lstRemote.List(contador)
                ' Verifica se é um diretorio ou arquivo
                If Right(nomeArquivo, 1) = "/" Then
                    operacao = "rmdir " & Left(nomeArquivo, Len(nomeArquivo) - 1)
                Else
                    operacao = "delete " & nomeArquivo
                End If
                executaComando operacao, False
            End If
        Next contador
        listaDir
    End If
End Sub

Private Sub cmdMkDir_Click()
    Dim dir As String, operacao As String
    
    dir = InputBox("Informe o nome da pasta", "Cria Diretório")
    If dir <> "" Then
        operacao = "mkdir " & dir
        executaComando operacao, True
    End If
End Sub

Private Sub cmdRecebe_Click()
    Dim contador As Integer
    Dim operacao As String
    Dim nomeArquivo As String, arquivoSaida As String
    
    For contador = 0 To lstRemote.ListCount - 1
        If lstRemote.Selected(contador) = True Then
            nomeArquivo = lstRemote.List(contador)
            If Len(dirList.Path) > 3 Then
                arquivoSaida = dirList.Path & "\" & nomeArquivo
            Else
                arquivoSaida = dirList.Path & nomeArquivo
            End If
            operacao = "recv " & nomeArquivo & " " & arquivoSaida
            executaComando operacao, False
            lstRemote.Selected(contador) = False
        End If
    Next contador
   filList.Refresh
End Sub

Private Sub cmdEnvia_Click()
    Dim contador As Integer
    Dim operacao As String
    Dim nomeArquivo As String, arquivoSaida As String
    
    For contador = 0 To filList.ListCount - 1
        If filList.Selected(contador) = True Then
            nomeArquivo = filList.List(contador)
            arquivoSaida = lblRemoteDirectory.Caption & "/" & nomeArquivo
            operacao = "send " & nomeArquivo & " " & arquivoSaida
            executaComando operacao, False
        End If
    Next contador
    listaDir
    filList.Refresh
End Sub

Private Sub dirList_Change()
    filList.Path = dirList.Path
End Sub

Private Sub drvList_Change()
    On Error GoTo driveError
    
    dirList.Path = drvList.Drive
    Exit Sub
driveError:
    MsgBox Err.Description, vbExclamation, "Drive Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo trata_erro
    If cmdEnvia.Enabled = True Then
        itcFTP.Execute , "Quit"
    End If
    Exit Sub
trata_erro:
        MsgBox "Erro na aplicação FTP", vbCritical
End Sub

Private Sub itcFTP_StateChanged(ByVal State As Integer)
    Select Case State
        Case icResolvingHost
            lblStatus.Caption = "Resolvendo Host"
        Case icHostResolved
            lblStatus.Caption = "Host Resolvido"
        Case icConnecting
            lblStatus.Caption = "Conectando ..."
        Case icConnected
            lblStatus.Caption = "Conectado"
        Case icRequesting
            lblStatus.Caption = "Requesitando ..."
        Case icRequestSent
            lblStatus.Caption = "Requesição enviada"
        Case icReceivingResponse
            lblStatus.Caption = "Recebendo ..."
        Case icResponseReceived
            lblStatus.Caption = "Resposta recebida"
        Case icDisconnecting
            lblStatus.Caption = "Desconectando ..."
        Case icDisconnected
            lblStatus.Caption = "Desconectado"
        Case icError
            lblStatus.Caption = itcFTP.ResponseInfo
            txtLog.Text = txtLog.Text & itcFTP.ResponseInfo & vbCrLf
        Case icResponseCompleted
            lblStatus.Caption = "operacao Completa"
            txtLog.Text = txtLog.Text & "operacao Completa" & vbCrLf
    End Select
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub lstRemote_DblClick()
    Dim operacao As String, dir As String
    
    ' Se o item é uma pasta muda para a pasta
    If Right(lstRemote.List(lstRemote.ListIndex), 1) = "/" Then
        dir = lstRemote.List(lstRemote.ListIndex)
        operacao = "cd " & Left(dir, Len(dir) - 1)
        executaComando operacao, True
    End If
End Sub

Private Sub txtLog_GotFocus()
    txtAddress.SetFocus
End Sub

Private Sub conectaHost()
    Dim operacao  As String
    
     On Error GoTo connectError
    
    If txtAddress.Text <> "" Then
      itcFTP.URL = txtAddress.Text
      itcFTP.UserName = txtUsername.Text
      itcFTP.Password = txtPassword.Text
      listaDir
      cmdEnvia.Enabled = True
      cmdRecebe.Enabled = True
      cmdMkDir.Enabled = True
      cmdDelete.Enabled = True
      lstRemote.Enabled = True
      cmdConecta.Caption = "D&esconectar do Host Remoto"
    Else
      MsgBox "Informe o nome do servidor FTP.", vbCritical
      txtAddress.SetFocus
    End If
    Exit Sub
    
connectError:
    MsgBox Err.Description, vbExclamation, "FTP"
End Sub

Private Sub desconectaHost()

   On Error GoTo trata_erro

    Dim operacao  As String
    
    itcFTP.Execute , "quit"
    operacao = "quit"
    executaComando operacao, False
    cmdEnvia.Enabled = False
    cmdRecebe.Enabled = False
    cmdMkDir.Enabled = False
    cmdDelete.Enabled = False
    lstRemote.Enabled = False
    cmdConecta.Caption = "&Conectar ao Host Remoto"
    Exit Sub
trata_erro:
    MsgBox "Erro ao efetuar a operacao com : " & txtAddress.Text & vbCrLf & " erro : " & Err.Number
    
End Sub

Private Sub executaComando(ByVal op As String, ByVal ld As Boolean)
 
    On Error GoTo trata_erro

    If itcFTP.StillExecuting Then
        itcFTP.Cancel
    End If
    txtLog.Text = txtLog.Text & "Comando: " & op & vbCrLf
    itcFTP.Execute , op
    terminaComando
    If ld = True Then
        listaDir
        terminaComando
    End If
    Exit Sub
    
trata_erro:
    MsgBox "Não foi possivel efetuar operacao com : " & txtAddress.Text & vbCrLf & " erro : " & Err.Number
End Sub

Private Sub terminaComando()
    Do While itcFTP.StillExecuting
        DoEvents
    Loop
End Sub

Private Sub listaDir()
    Dim operacao As String
    Dim data As Variant, contador As Integer
    Dim inicio As Integer, length As Integer
    
    inicio = 1
    lstRemote.Clear
    operacao = "dir"
    executaComando operacao, False
    Do
        data = itcFTP.GetChunk(1024, icString)
        DoEvents
        For contador = 1 To Len(data)
            If Mid(data, contador, 1) = Chr(13) Then
                If length > 0 And Mid(data, inicio, length) <> "./" Then
                    lstRemote.AddItem Mid(data, inicio, length)
                End If
                inicio = contador + 2
                length = -1
            Else
                length = length + 1
            End If
        Next contador
    Loop While LenB(data) > 0
    operacao = "pwd"
    executaComando operacao, False
    lblRemoteDirectory.Caption = itcFTP.GetChunk(1024, icString)
End Sub

