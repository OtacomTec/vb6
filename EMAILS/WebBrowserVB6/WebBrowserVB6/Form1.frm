VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Navegador VB6"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   7920
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Muito Grande"
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
      Index           =   4
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Grande"
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
      Index           =   3
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Média"
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
      Index           =   2
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Pequena"
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
      Index           =   1
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Muito Pequena"
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
      Index           =   0
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Propriedades da Página"
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
      Index           =   3
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Configurar Página"
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
      Index           =   2
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Visualizar"
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
      Index           =   1
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Imprimir"
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
      Index           =   0
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Procurar"
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
      Index           =   5
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Home"
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
      Index           =   4
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Atualizar"
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
      Index           =   3
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Parar"
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
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Avançar"
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
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Retornar"
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
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtCriterio 
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   7440
      Width           =   11295
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Procurar"
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
      Left            =   240
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txtURL 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Text            =   "www.macoratti.net"
      Top             =   720
      Width           =   11295
   End
   Begin VB.CommandButton cmdNavegar 
      Caption         =   "Navegar"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   12495
      ExtentX         =   22040
      ExtentY         =   11033
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Tamanho da  Fonte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   20
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNavegar_Click()
On Error GoTo trata_erro
   WebBrowser1.Navigate Trim(txtURL.Text)
   Exit Sub
trata_erro:
   MsgBox Err.Description
End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Dim frm As Form1
    Set frm = New Form1
    Set ppDisp = frm.WebBrowser1.Object
    frm.Show
End Sub
Private Sub cmdLocalizar_Click()
    Dim strTexto As String
    strTexto = txtCriterio.Text
    If PaginaWebContem(strTexto) = True Then 'check if the word is in page
         MsgBox "A página contém o texto "
    Else
         MsgBox "A página não contém esse texto" 'string is not in page
    End If
End Sub
Private Function PaginaWebContem(ByVal s As String) As Boolean
    Dim i As Long, EHTML
    For i = 1 To WebBrowser1.Document.All.length
        Set EHTML = WebBrowser1.Document.All.Item(i)
          If Not (EHTML Is Nothing) Then
            If InStr(1, EHTML.innerHTML, s, vbTextCompare) > 0 Then
                WebPageContains = True
                Exit Function
        End If
    End If
Next i
End Function
Private Sub Form_Load()
      WebBrowser1.Navigate "about:blank"
      'criaPagina
       ProgressBar1.Appearance = ccFlat
       ProgressBar1.scrolling = ccScrollingSmooth
End Sub
Private Sub criaPagina()
        Dim HTML As String
            '----------código HTML----------
        HTML = "<HTML>" & _
                "<TITLE>Pagina Carrega no evento Load</TITLE>" & _
                "<BODY>" & _
                "<FONT COLOR = BLUE>" & _
                "Este página foi feita  " & _
                "<FONT SIZE = 5>" & _
                "<B>" & _
                "via código por Macoratti.net " & _
                "</B>" & _
                "</FONT SIZE>" & _
                "</BR >" & _
                "www.macoratti.net" & _
                "</FONT>" & _
                "</BODY>" & _
                "</HTML>"
                '----------HTML fim ---------------
        WebBrowser1.Document.Write HTML
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error Resume Next
                       
        Select Case Index
            Case 0 'botão Go Back
                WebBrowser1.GoBack
            Case 1 'Botão Go Forward
                WebBrowser1.GoForward
            Case 2
                WebBrowser1.Stop
            Case 3
                WebBrowser1.Refresh 'atualiza pagina
            Case 4
                WebBrowser1.GoHome 'botão Go to home
            Case 5
                WebBrowser1.GoSearch 'botão Search
        End Select
End Sub
Private Sub Command2_Click(Index As Integer)
On Error Resume Next
                       
        Select Case Index
            Case 0 'botão Imprimir
               WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
            Case 1 'Botão Visualizar
                WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
            Case 2 'botão Configurar
                WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
            Case 3 'botão propriedades
                WebBrowser1.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
        End Select
End Sub
Private Sub Command3_Click(Index As Integer)
On Error Resume Next
Select Case Index
            Case 0 'Muto pequena
               WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
            Case 1 'pequena
                WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
            Case 2 'Medida
                WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
            Case 3 'Grande
                WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
            Case 4 'Muito Grande
                WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
        End Select
End Sub
Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)

Select Case Command
        Case 1 'Avançar
            Command1(0).Enabled = Enable
        Case 2 'Retornar
            Command1(1).Enabled = Enable
    End Select

End Sub
Private Sub WebBrowser1_DownloadComplete()
On Error Resume Next
    'voce precisa adicionar o "Microsoft HTML Object Library"!!!!!!!!!
    Dim HTMLdoc As HTMLDocument
        Dim HTMLlinks As HTMLAnchorElement
            Dim STRtxt As String
    ' Lista os links links.
    Set HTMLdoc = WebBrowser1.Document
    For Each HTMLlinks In HTMLdoc.links
           STRtxt = STRtxt & HTMLlinks.href & vbCrLf
    Next HTMLlinks
    'exibe os links em um textbox
    'Text1.Text = STRtxt
    'salva os links em um arquivo texto
    'Open "C:\dados\linksLog.txt" For Append As #1
    'Print #1, STRtxt
    'Close #1
End Sub
Private Sub WebBrowser1_ProgressChange(ByVal Progresso As Long, ByVal ProgressoMax As Long)
On Error Resume Next
    If Progresso = -1 Then ProgressBar1.Value = 100
        Me.Caption = "100%"
    If Progresso > 0 And ProgressoMax > 0 Then
        ProgressBar1.Value = Progresso * 100 / ProgressoMax
        Me.Caption = Int(Progresso * 100 / ProgressoMax) & "%"
    End If
    Exit Sub
End Sub
