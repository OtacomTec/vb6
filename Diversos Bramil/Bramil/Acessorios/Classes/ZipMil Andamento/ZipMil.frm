VERSION 5.00
Begin VB.Form FormZipMil 
   Caption         =   "Exemplo ZipMil"
   ClientHeight    =   3105
   ClientLeft      =   3825
   ClientTop       =   2490
   ClientWidth     =   4515
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   4515
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   225
      TabIndex        =   4
      Top             =   900
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Descompactando"
      Height          =   465
      Left            =   2280
      TabIndex        =   3
      Top             =   225
      Width           =   1905
   End
   Begin VB.FileListBox File2 
      Height          =   1845
      Left            =   2295
      TabIndex        =   2
      Top             =   855
      Width           =   1905
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   225
      TabIndex        =   1
      Top             =   1260
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compactando"
      Height          =   465
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Nome do Zip:"
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   720
      Width           =   1725
   End
End
Attribute VB_Name = "FormZipMil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selecionado As String

Private Sub Command1_Click()
    On Error GoTo vbErrorHandler
    Dim Zip As cZipMil
    Set Zip = New cZipMil
        
    caminho = App.Path & "\"
    NomeDoZip = Trim(Text1.Text) & ".Zip"
    
    With Zip
        .NomeDoZip = caminho & NomeDoZip
        .AtualizarZip = False       'Sempre vai criar um novo
        .IncluiSubPastas = False
        For i = 0 To File1.ListCount - 1
            .AdicionarArquivo caminho & File1.List(i)
        Next i
        
'        .AdicionarArquivo App.Path & "\*.*" 'adciona todos os arquivos na pasta
        If .Zipar <> 0 Then   'Cria o zip e exibe erros
            MsgBox .ÚltimaMensagem ' any errors
        End If
    End With
    
    MsgBox "O arquivo " & Zip.NomeDoZip & " foi criado com sucesso!"
    File2.Refresh
    
    Set Zip = Nothing
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & "Form1:Zipar" & " " & Err.Description
End Sub

Private Sub Command2_Click()

On Error GoTo vbErrorHandler
    Dim UnZip As cUnZipMil
    Set UnZip = New cUnZipMil
    
    caminho = App.Path & "\"
    
    With UnZip
        .NomeDoZip = selecionado
        '.PastaExtração = caminho
        .SobregravarArquivos = True
        .HonorDirectories = False
        .SensívelCaracter = False        'True torna o soft sensível a maiúsculos e minúsculos
        '.HonorDirectories = False       'True descompacta e recria todos os subdiretórios na pasta atual
                
        If .UnZipar <> 0 Then
            MsgBox .ÚltimaMensagem
        End If
    End With
    
    MsgBox "Extração do arquivo " & UnZip.NomeDoZip & " completa com sucesso!"
    Set UnZip = Nothing
    File1.Refresh
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & "Form1:Unzip" & " " & Err.Description
End Sub

Private Sub File2_Click()
    selecionado = File2.List(File2.ListIndex)
End Sub

Private Sub Form_Load()
    File1.Path = App.Path
    File1.Pattern = "*.doc"
    
    File2.Path = App.Path
    File2.Pattern = "*.zip"
End Sub
