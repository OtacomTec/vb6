VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   1560
   ClientTop       =   1950
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Etiqueta Tripa"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Etiqueta Amarela"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim lclDataMax As clDataMax
    Set lclDataMax = New clDataMax
    Set lclDataMax.MsComm = MSComm1
    Dim txtLinha
    Dim liMargem  As Integer
    liMargem = 15
    txtLinha = "FIGO FATIADO COM NEGAO DO ACOUGUE DO MERCADO INANUGURADO"
    
    lclDataMax.PortaCom = 1
    lclDataMax.PreparaImpressão 10
    
       
    lclDataMax.ImprimeTexto Left(txtLinha, 19), , "2", "00", , "0060", "0015"
    lclDataMax.ImprimeTexto Mid(txtLinha, 20, 19), , "2", "00", , "0048", "0015"
    lclDataMax.ImprimeTexto Mid(txtLinha, 39, 19), , "2", "00", , "0036", "0015"
    lclDataMax.ImprimeBarra "7891099643338", , , , , , "0015"
    
    lclDataMax.ImprimeTexto Left(txtLinha, 19), , "2", "00", , "0060", "0155"
    lclDataMax.ImprimeTexto Mid(txtLinha, 20, 19), , "2", "00", , "0048", "0155"
    lclDataMax.ImprimeTexto Mid(txtLinha, 39, 19), , "2", "00", , "0036", "0155"
    lclDataMax.ImprimeBarra "7891099643338", , , , , , "0155"
    
    lclDataMax.ImprimeTexto Left(txtLinha, 19), , "2", "00", , "0060", "0295"
    lclDataMax.ImprimeTexto Mid(txtLinha, 20, 19), , "2", "00", , "0048", "0295"
    lclDataMax.ImprimeTexto Mid(txtLinha, 39, 19), , "2", "00", , "0036", "0295"
    lclDataMax.ImprimeBarra "7891099643338", , , , , , "0295"
   
    
    
    lclDataMax.FinalizaImpressão
    
    
    Set lclDataMax = Nothing
End Sub

Private Sub Command3_Click()
    Dim lclDataMax As clDataMax
    Set lclDataMax = New clDataMax
    Set lclDataMax.MsComm = MSComm1
    Dim txtLinha
    txtLinha = "FIGO FATIADO COM NEGAO DO ACOUGUE DO MERCADO INANUGURADO"
    
    lclDataMax.PortaCom = 1
    lclDataMax.PreparaImpressão
    lclDataMax.ImprimeTexto Left(txtLinha, 36), , , , , "0070"
    lclDataMax.ImprimeTexto Mid(txtLinha, 37, 20), , , , , "0040"
    lclDataMax.ImprimeTexto "039373  E/36", , 3, , , "0040", "0280"
    lclDataMax.ImprimeTexto "000,50", , 6, "21", , , "0170"
    lclDataMax.ImprimeTexto "R$", , 5, "11", , , "0115"
    lclDataMax.ImprimeBarra "7891099643338"
    lclDataMax.FinalizaImpressão
    Set lclDataMax = Nothing
    
End Sub

Private Sub Form_Load()

End Sub
