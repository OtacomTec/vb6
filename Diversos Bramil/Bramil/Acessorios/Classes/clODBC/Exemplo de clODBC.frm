VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   1125
   ClientTop       =   2235
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   10200
   Begin VB.CommandButton Command3 
      Caption         =   "Cria ODBC"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   5340
      Width           =   1275
   End
   Begin VB.TextBox TextDescrição 
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Text            =   "Teste de Descrição"
      Top             =   4980
      Width           =   2715
   End
   Begin VB.TextBox TextSenha 
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   4650
      Width           =   2715
   End
   Begin VB.TextBox TextUsuário 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Text            =   "Admin"
      Top             =   4290
      Width           =   2715
   End
   Begin VB.TextBox TextServidor 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   3945
      Width           =   2715
   End
   Begin VB.TextBox TextBaseDeDados 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Text            =   "H:\aplicvb\bd\bdgms002.mdb"
      Top             =   3600
      Width           =   2715
   End
   Begin VB.TextBox TextDriver 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "Microsoft Access Driver (*.mdb)"
      Top             =   3270
      Width           =   2715
   End
   Begin VB.TextBox TextNomeDSN 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "Teste "
      Top             =   2910
      Width           =   2715
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2565
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4524
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      FormatString    =   "DNS                                    |Drive                                              "
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2565
      Left            =   5100
      TabIndex        =   17
      Top             =   300
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4524
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "Drives Instalados                                                                         "
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   675
      Left            =   5220
      TabIndex        =   18
      Top             =   3030
      Width           =   4815
   End
   Begin VB.Label Label8 
      Caption         =   "Descrição"
      Height          =   225
      Left            =   150
      TabIndex        =   15
      Top             =   5010
      Width           =   1875
   End
   Begin VB.Label Label7 
      Caption         =   "Senha"
      Height          =   225
      Left            =   150
      TabIndex        =   13
      Top             =   4680
      Width           =   1875
   End
   Begin VB.Label Label6 
      Caption         =   "Usuário"
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   4320
      Width           =   1875
   End
   Begin VB.Label Label5 
      Caption         =   "Servidor"
      Height          =   225
      Left            =   150
      TabIndex        =   9
      Top             =   3960
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "Base de Dados"
      Height          =   225
      Left            =   150
      TabIndex        =   8
      Top             =   3630
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Driver"
      Height          =   225
      Left            =   150
      TabIndex        =   7
      Top             =   3300
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "Nome do DSN"
      Height          =   225
      Left            =   150
      TabIndex        =   6
      Top             =   2970
      Width           =   1875
   End
   Begin VB.Label Label1 
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   2625
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lclODBC As clODBC
Dim mclReg As clRegedit


Private Sub Command3_Click()
    Dim lstrODBC As Variant
    lstrODBC = mclReg.ListarChaves("HKEY_LOCAL_MACHINE\SoftWare\ODBC\ODBC.INI")
    
    For i = 1 To UBound(lstrODBC)
        If TextNomeDSN.Text = (lstrODBC(i)) Then
            lboAchou = True
            Exit For
        End If
    Next i
    
    If lboAchou = False Then
        lclODBC.CriarODBCSistema TextNomeDSN.Text, _
                             TextDriver.Text, _
                             TextDescrição.Text, _
                             TextBaseDeDados.Text, _
                             TextUsuário.Text, _
                             TextSenha.Text
    End If
                             
    ExibeDrivers
    ExibeDSN
    Label1.Caption = lclODBC.TotalDNS & " DSN's Instalados"
End Sub

Private Sub Form_Load()
    Set lclODBC = New clODBC
    Set mclReg = New clRegedit
    ExibeDSN
    Label1.Caption = lclODBC.TotalDNS & " DSN's Instalados"
    ExibeDrivers
End Sub
Sub ExibeDSN()
    Me.MSFlexGrid.Rows = 1
    
    For i = 1 To UBound(lclODBC.DSNInstalados())
        MSFlexGrid.AddItem ""
        MSFlexGrid.TextMatrix(MSFlexGrid.Rows - 1, 0) = lclODBC.DSNInstalados(i)
        MSFlexGrid.TextMatrix(MSFlexGrid.Rows - 1, 1) = lclODBC.DRVdosDSNInstalados(i)
        
    Next i
End Sub
Sub ExibeDrivers()
    Me.MSFlexGrid1.Rows = 1
    For i = 1 To UBound(lclODBC.DriversInstalados())
        MSFlexGrid1.AddItem ""
        MSFlexGrid1.TextMatrix(i, 0) = lclODBC.DriversInstalados(i) & " | - " & lclODBC.DriversAtributos(i)
    Next i
End Sub
Private Sub MSFlexGrid1_Click()
    TextDriver.Text = Trim(Mid(MSFlexGrid1.Text, 1, InStr(1, MSFlexGrid1.Text, "|") - 1))
    Label9.Caption = MSFlexGrid1.Text
End Sub
