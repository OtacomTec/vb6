VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6945
      Left            =   1740
      ScaleHeight     =   6885
      ScaleWidth      =   9165
      TabIndex        =   6
      Top             =   210
      Width           =   9225
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   525
      Left            =   270
      TabIndex        =   5
      Top             =   5190
      Width           =   1245
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   525
      Left            =   240
      TabIndex        =   4
      Top             =   4350
      Width           =   1245
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   525
      Left            =   270
      TabIndex        =   3
      Top             =   3510
      Width           =   1245
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   525
      Left            =   270
      TabIndex        =   2
      Top             =   2670
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   525
      Left            =   270
      TabIndex        =   1
      Top             =   1800
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   270
      TabIndex        =   0
      Top             =   1080
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
      ' Capture the entire screen.

      Private Sub Command1_Click()
      
      SavePicture CaptureScreen(), "c:\marcos.jpg"
      
      End Sub

      ' Capture the entire form including title and border.
      Private Sub Command2_Click()
         Set Picture1.Picture = CaptureForm(Me)
      End Sub

      ' Capture the client area of the form.
      Private Sub Command3_Click()
         Set Picture1.Picture = CaptureClient(Me)
      End Sub

      ' Capture the active window after two seconds.
      Private Sub Command4_Click()
         MsgBox "Two seconds after you close this dialog " & _
            "the active window will be captured."

         ' Wait for two seconds.
         Dim EndTime As Date
         EndTime = DateAdd("s", 2, Now)
         Do Until Now > EndTime
            DoEvents
         Loop

         Set Picture1.Picture = CaptureActiveWindow()

         ' Set focus back to form.
         Me.SetFocus
      End Sub

      ' Print the current contents of the picture box.
      Private Sub Command5_Click()
         PrintPictureToFitPage Printer, Picture1.Picture
         Printer.EndDoc
      End Sub

      ' Clear out the picture box.
      Private Sub Command6_Click()
         Set Picture1.Picture = Nothing
      End Sub

      ' Initialize the form and controls.
      Private Sub Form_Load()
         Me.Caption = "Capture and Print Example"
         Command1.Caption = "&Screen"
         Command2.Caption = "&Form"
         Command3.Caption = "&Client"
         Command4.Caption = "&Active"
         Command5.Caption = "&Print"
         Command6.Caption = "C&lear"
         Picture1.AutoSize = True
      End Sub
      '--------------------------------------------------------------------

Function Grava_Imagem(Path As String, Campo_Imagem As String, Tabela_Imagem As String, Campo_Chave_Tabela_Imagem As String, Valor_Campo_Chave_Tabela_Imagem As String, Banco As String, Form As Object)

    'OPEN RECORDSET FOR WRITING
    Dim rs As New ADODB.Recordset
    Dim srmImagem As New ADODB.Stream
    Dim CN As New DLLConexao_Sistema.Conexao
    Dim strsql As String
    
    CN.Initial_Catalog = Banco
    CN.Abrir_conexao ("Otica")

    strsql = Empty
    strsql = "SELECT " & Campo_Imagem & " FROM " & Tabela_Imagem & " WHERE " & Campo_Chave_Tabela_Imagem & " = " & Valor_Campo_Chave_Tabela_Imagem
    rs.Open strsql, CN.CNConexao, adOpenStatic, adLockOptimistic

    If IsNull(rs.Fields(Campo_Imagem)) = False Then
        
        srmImagem.Type = adTypeBinary
        srmImagem.Open
        srmImagem.LoadFromFile Path
        
        rs.Fields(0) = srmImagem.Read
        rs.Update
        srmImagem.Close

    End If
    
    rs.Close

    Set srmImagem = Nothing
    Set rs = Nothing
    
End Function
