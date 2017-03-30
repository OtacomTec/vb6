VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROCESSA IMAGENS"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESSA IMAGENS"
      BeginProperty Font 
         Name            =   "Nina"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1110
      TabIndex        =   0
      Top             =   630
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1020
      TabIndex        =   1
      Top             =   2430
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call SalvaImagem
End Sub

Private Sub Form_Load()
'''''''
'''''''
'''''''      'CREATE CONNECTION OBJECT AND ASSIGN CONNECTION STRING
'      Dim conn As ADODB.Connection
'      Set conn = New ADODB.Connection
'
'      conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDRetaguarda;Data Source=ONLYTECH-03"
'      conn.CursorLocation = adUseClient
'
'      conn.Open
'
'      'OPEN RECORDSET FOR WRITING
'      Dim rs As ADODB.Recordset
'      Set rs = New ADODB.Recordset
'
'        Dim mystream As ADODB.Stream
'        Set mystream = New ADODB.Stream
'
'       mystream.Type = adTypeBinary
'       rs.Open "SELECT * FROM TBPRODUTO", conn, adOpenStatic, adLockOptimistic
'
'       Do While rs.EOF = False
'
'          '  mystream.Open
'          '  mystream.LoadFromFile "c:\logo_only_para_messenger.jpg"
'
'            'rs!file_size = mystream.Size
'            rs!DFTeste2 = mystream.Read
'            rs.Update
'
'           ' mystream.Close
'
'            rs.MoveNext
'
'       Loop
'    rs.Close

  '  Call SalvaImagem

End Sub

Private Function carregaImagem(ByVal fileName As String) As Byte()
      
      Dim conn As ADODB.Connection
      Set conn = New ADODB.Connection

      conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDRetaguarda;Data Source=ONLYTECH-03"
      conn.CursorLocation = adUseClient

      conn.Open

      'OPEN RECORDSET TO READ BLOB
      rs.Open "Select * from TBPRODUTO WHERE IXCODIGO_TBPRODUTO = 2", conn
      mystream.Open
      mystream.Write rs!file_blob
      
      'mystream.SaveToFile "c:\newimage.jpg", adSaveCreateOverWrite
      mystream.Close
      rs.Close

      'OPEN RECORDSET FOR UPDATE OF BLOB COLUMN
      rs.Open "Select * from files WHERE files.file_id = 1", conn, adOpenStatic, adLockOptimistic
      mystream.Open
      mystream.LoadFromFile "c:\adaptor.jpg"
      rs!file_blob = mystream.Read
      rs.Update
      mystream.Close
      rs.Close
      
      'OPEN RECORDSET TO READ UPDATED IMAGE
      rs.Open "Select * from files WHERE files.file_id = 1", conn
      mystream.Open
      mystream.Write rs!file_blob
      mystream.SaveToFile "c:\newupdatedimage.jpg", adSaveCreateOverWrite
      mystream.Close
      rs.Close
      conn.Close
      MsgBox "Success! Check your C:\ directory for newimage.jpg and newupdatedimage.jpg"



'
'    Dim br As BinaryReader = New BinaryReader(fs)
'    Return (br.ReadBytes(Convert.ToInt32(br.BaseStream.Length)))

End Function
Function SalvaImagem()

    Dim oTabela As ADODB.Recordset
    Set oTabela = New ADODB.Recordset
      
    Dim CN As New DLLConexao_Sistema.Conexao
    
    CN.Abrir_conexao ("Otica")
    
    oTabela.Source = "SELECT * FROM TBPRODUTO"
    oTabela.Open , CN.CNConexao, adOpenDynamic, adLockOptimistic
    
    Dim lRet    As Boolean
    Dim oStream As ADODB.Stream
    
    lRet = True
    
    Set oStream = New ADODB.Stream
    Dim intCont As Long
    Dim strIMAGEM As String
    
    intCont = 1
    
    Do While oTabela.EOF = False
        With oStream
            .Type = adTypeBinary                               'tipo da leitura, para campos Image/BLOB/OLEDB deve ser binário
            .Open
            strIMAGEM = oTabela!DFPath_imagem_TBProduto
            .LoadFromFile strIMAGEM
            oTabela.Fields("DFImagem_stream_TBProduto").Value = .Read  'atribui o conteúdo ao campo passado como referência
            oTabela.Update                                         'salva a informação.
            .Close                                                        'e fecha o arquivo, para permitir que seja eliminado, claro.
        End With
        intCont = intCont + 1
        Me.Label1.Caption = intCont
        oTabela.MoveNext
        DoEvents
        Me.Refresh
    Loop
    
    MsgBox "EXECUTADO COM SUCESSO"
    
    GoTo SIExit:                                                  'se chegou até aqui, ok.
    
SIErr:
    MsgBox Err.Description                               'mostra erro - desnecessário, se você tratar o retorno da função.
    Err.Clear                                                       'limpa a matriz de erros
    lRet = False                                                  'retorna falso
    
SIExit:
    SalvaImagem = lRet                                     'Envia o retorno da função
    Set oStream = Nothing                                'limpa a instância ativa, eliminando o espaço ocupado pela variável.
    lRet = Empty                                                'idem, para a var. de retorno.
    
End Function
