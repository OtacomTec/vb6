VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   2370
      Top             =   690
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDteste;Data Source=."
      OLEDBString     =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDteste;Data Source=."
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gravar"
      Height          =   345
      Left            =   540
      TabIndex        =   4
      Top             =   1920
      Width           =   915
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   540
      TabIndex        =   1
      Top             =   1260
      Width           =   2235
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   570
      TabIndex        =   0
      Top             =   420
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Preço Total"
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   1050
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   570
      TabIndex        =   2
      Top             =   210
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim strSql As String
    Dim cn As New ADODB.Connection
    cn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDteste;Data Source=."
    strSql = "INSERT INTO Tabela1(codigo,preco_total) VALUES (" & Me.Text1.Text & ", " & Funcoes_Gerais.Grava_Moeda(Me.Text2) & ")"
    cn.Execute strSql
    
End Sub

