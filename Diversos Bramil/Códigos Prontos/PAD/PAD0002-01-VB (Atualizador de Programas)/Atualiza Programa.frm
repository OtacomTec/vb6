VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Atualizador de Programas"
   ClientHeight    =   5130
   ClientLeft      =   2505
   ClientTop       =   2520
   ClientWidth     =   9240
   Icon            =   "Atualiza Programa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9240
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8310
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atualiza Programa.frx":0A8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atualiza Programa.frx":0DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atualiza Programa.frx":1132
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Atualiza Programa.frx":1486
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   4845
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox CheckDesligarWindows 
      Caption         =   "Desligar o Computador ao Finalizar"
      Height          =   225
      Left            =   90
      TabIndex        =   8
      Top             =   4290
      Width           =   2895
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   953
      ButtonWidth     =   1349
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Atualizar"
            Key             =   "Atualizar"
            Object.ToolTipText     =   "Atualiza os Programas Selecionados"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Parar"
            Key             =   "Parar"
            Object.ToolTipText     =   "Interrompe a execução da atualização corrente"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Adicionar"
            Key             =   "Adicionar"
            Object.ToolTipText     =   "Adiciona um arquivo para ser atualizado"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remover"
            Key             =   "Remover"
            Object.ToolTipText     =   "Remove um arquivo para ser atualizado"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3210
      Left            =   4530
      TabIndex        =   6
      Top             =   990
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5662
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FormatString    =   "   |Arquivo                               |Pasta                                 "
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2370
      Pattern         =   "*.exe"
      TabIndex        =   3
      Top             =   990
      Width           =   2115
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   3870
      Width           =   2235
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   90
      TabIndex        =   0
      Top             =   990
      Width           =   2235
   End
   Begin VB.Label Label5 
      Caption         =   "Atualizar os Arquivos Abaixo"
      Height          =   225
      Left            =   4530
      TabIndex        =   5
      Top             =   780
      Width           =   2505
   End
   Begin VB.Label Label4 
      Caption         =   "Arquivos Novos"
      Height          =   225
      Left            =   2430
      TabIndex        =   4
      Top             =   780
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "Pasta a ser atualizada"
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   780
      Width           =   2145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Private boParar As Boolean
Private pclDesligarWindows As eWin

Private Function AtualizarArquivo(NomeDoArquivo As String, PastaDoArquivo As String) As Integer
    On Error GoTo rotErro
    ChDrive Mid(PastaDoArquivo, 1, 1)
    ChDir PastaDoArquivo
    
    NomeDoArquivoAtual = Replace(NomeDoArquivo, ".", ".")
    NomeDoArquivoNovo = NomeDoArquivo
    Name NomeDoArquivoAtual As NomeDoArquivoAtual & " " & Format(FileDateTime(NomeDoArquivoAtual), "yyyy-mm-dd hhmmss")
    
    If AtualizarArquivo <> 0 Then
        Err.Clear
    Else
        Name NomeDoArquivoNovo As NomeDoArquivoAtual
        AtualizarArquivo = 0
    End If
    DoEvents
    Exit Function
    
rotErro:
    Select Case Err.Number
        Case 75
            AtualizarArquivo = Err.Number
            Resume Next
        Case Else
            AtualizarArquivo = Err.Number
    End Select

End Function



Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    'MsgBox File1.List(File1.ListIndex)
    
End Sub

Private Sub File1_DblClick()
    Dim Selecionado As Boolean
    If MSFlexGrid1.Rows <> 1 Then
        For i = 1 To Me.MSFlexGrid1.Rows - 1
            MSFlexGrid1.Row = i
            If MSFlexGrid1.TextMatrix(i, 1) = File1.List(File1.ListIndex) Then
                If MSFlexGrid1.TextMatrix(i, 2) = Dir1.Path Then
                    Selecionado = True
                    Exit For
                End If
            End If
        Next i
    End If
    If Selecionado = False Then
        MSFlexGrid1.AddItem ""
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = File1.List(File1.ListIndex)
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = Dir1.Path
    End If
End Sub

Private Sub Form_Load()
    Set pclDesligarWindows = New eWin
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set pclDesligarWindows = Nothing
    End
End Sub

Private Sub MSFlexGrid1_DblClick()
    If MSFlexGrid1.Row = 0 Then Exit Sub
    If MSFlexGrid1.Row = 1 And MSFlexGrid1.Rows = 2 Then
        MSFlexGrid1.Rows = 1
    Else
        MSFlexGrid1.RemoveItem MSFlexGrid1.Row
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Atualizar"
            Dim Atualizado As Boolean
            Do While Atualizado = False
                Atualizado = True
                For i = 1 To MSFlexGrid1.Rows - 1
                    If boParar = True Then
                        boParar = False
                        StatusBar1.Panels(1).Text = "Processo Abortado "
                        StatusBar1.Panels(2).Text = Time
                        Exit Sub
                    End If
                    MSFlexGrid1.Col = 0
                    MSFlexGrid1.Row = i
                    If MSFlexGrid1.CellBackColor <> vbGreen And MSFlexGrid1.CellBackColor <> vbRed Then
                        Atualizado = False
                        StatusBar1.Panels(1).Text = "Tentando Atualizar... "
                        StatusBar1.Panels(2).Text = MSFlexGrid1.TextMatrix(i, 2) & "\" & MSFlexGrid1.TextMatrix(i, 1) & " "
                        ret = AtualizarArquivo(MSFlexGrid1.TextMatrix(i, 1), MSFlexGrid1.TextMatrix(i, 2))
                        DoEvents
                        If ret = 0 Then
                            MSFlexGrid1.Col = 0
                            MSFlexGrid1.Row = i
                            MSFlexGrid1.CellBackColor = vbGreen
                            StatusBar1.Panels(1).Text = "Atualizado "
                            StatusBar1.Panels(2).Text = MSFlexGrid1.TextMatrix(i, 2) & "\" & MSFlexGrid1.TextMatrix(i, 1) & " "
                            File1.Refresh
                        ElseIf ret = 53 Then
                            MSFlexGrid1.CellBackColor = vbRed
                            StatusBar1.Panels(1).Text = "Erro... "
                            StatusBar1.Panels(2).Text = MSFlexGrid1.TextMatrix(i, 2) & "\" & MSFlexGrid1.TextMatrix(i, 1) & " "
                        
                        ElseIf ret = 75 Then
                            MSFlexGrid1.TextMatrix(i, 0) = "U"
                            StatusBar1.Panels(1).Text = "Arquivo sendo utilizado... "
                            StatusBar1.Panels(2).Text = MSFlexGrid1.TextMatrix(i, 2) & "\" & MSFlexGrid1.TextMatrix(i, 1) & " "
                            
                        End If
                    End If
                    DoEvents
                Next i
            Loop
            StatusBar1.Panels(1).Text = "Processo Concluído "
            StatusBar1.Panels(2).Text = Time
            If CheckDesligarWindows.Value = 1 Then
                If pclDesligarWindows.ID_SO > emb_VerSOWindows98SE Then
                    pclDesligarWindows.SairDoWindows SDW_POWEROFF
                Else
                    pclDesligarWindows.SairDoWindows SDW_SHUTDOWN
                End If
                'pclDesligarWindows.Desligar (Win_Desligar)
            End If
            
        Case "Parar"
            boParar = True
        Case "Adicionar"
            If File1.List(File1.ListIndex) = "" Then Exit Sub
            File1_DblClick
        Case "Remover"
            MSFlexGrid1_DblClick
    End Select
End Sub
