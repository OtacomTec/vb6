VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Visualizador de Arquivos com lay-out"
   ClientHeight    =   5565
   ClientLeft      =   1725
   ClientTop       =   1605
   ClientWidth     =   9450
   Icon            =   "Visualizar Arquivo Binário.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   9450
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4305
      Left            =   60
      TabIndex        =   0
      Top             =   690
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   7594
      _Version        =   393216
      Cols            =   14
      AllowUserResizing=   1
      FormatString    =   $"Visualizar Arquivo Binário.frx":3542
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1875
      Left            =   2670
      TabIndex        =   3
      Top             =   1800
      Width           =   4785
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   90
         TabIndex        =   4
         Top             =   840
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Carregando Grid. Por favor aguarde um momento ..."
         Height          =   225
         Left            =   150
         TabIndex        =   5
         Top             =   570
         Width           =   4665
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9090
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Visualizar Arquivo Binário.frx":35DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   1111
      ButtonWidth     =   1058
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abrir Arquivo Binário"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Alterar"
            Key             =   "Alterar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   5280
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Nome do Arquivo"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Registro Atual / Total de Registros"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Data e Hora do Arquivo"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   5100
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1965
      Left            =   510
      TabIndex        =   6
      Top             =   3090
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3466
      _Version        =   393216
      Cols            =   14
      FormatString    =   $"Visualizar Arquivo Binário.frx":36EE
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Label2_Click()

End Sub


Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        If MSFlexGrid2.Visible = False Then
            If Me.Height < 4200 Then Me.Height = 4200
            If Me.Width < 6000 Then Me.Width = 6000
            Me.MSFlexGrid1.Height = Me.Height - Me.StatusBar1.Height - Me.Toolbar1.Height - 500
            Me.MSFlexGrid1.Width = Me.Width - 250
            Me.Frame1.Top = (Me.Height / 2) - (Frame1.Height / 2)
            Me.Frame1.Left = (Me.Width / 2) - (Frame1.Width / 2)
        Else
            altura = (Form1.StatusBar1.Top - Form1.Toolbar1.Height) - 130
        
            If Me.Height < 4200 Then Me.Height = 4200
            If Me.Width < 6000 Then Me.Width = 6000
            Me.MSFlexGrid1.Height = altura / 2
            
            Me.MSFlexGrid2.Top = MSFlexGrid1.Top + MSFlexGrid1.Height + 50
            Me.MSFlexGrid2.Height = Me.MSFlexGrid1.Height
            Me.MSFlexGrid1.Width = Me.Width - 250
            Me.MSFlexGrid2.Width = Me.MSFlexGrid1.Width
            
            Me.Frame1.Top = (Me.Height / 2) - (Frame1.Height / 2)
            Me.Frame1.Left = (Me.Width / 2) - (Frame1.Width / 2)
        
        End If
    End If
End Sub

Private Sub MSFlexGrid1_Click()
    'Me.LabelRegistroAtual.Caption = "Registro Atual : " & MSFlexGrid1.Row
    Form1.StatusBar1.Panels(2).Text = MSFlexGrid1.Row & " / " & MSFlexGrid1.Rows - 1
End Sub

Private Sub MSFlexGrid1_DblClick()
    Form2.Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
    Form2.Show 1
    
End Sub

Private Sub MSFlexGrid2_Click()
    Form1.StatusBar1.Panels(2).Text = MSFlexGrid1.Row + 9000 & " / " & MSFlexGrid1.Rows - 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Abrir"
            CommonDialog.Filter = "Todos Arq Textos Binários (txtGMS002/ZB2/ZB3_??_???.txt)|txtZB2_??_???.txt; txtZB3_??_???.txt; txtGMS002_??_???.txt" & _
                          "|Arquivos Textos Binários Tipo1 (*.txt)|txtGMS002_??_???.txt" & _
                          "|Arquivos Textos Binários ZB2 (*.txt)|txtZB2_??_???.txt" & _
                          "|Arquivos Textos Binários ZB2 Geral (*G.txt)|txtZB2_??_???G.txt" & _
                          "|Arquivos Textos Binários ZB3 (*.txt)|txtZB3_??_???.txt" & _
                          "|Arquivos Textos Binários GMS005 (*.txt)|txtGMS005_??_???.txt" & _
                          "|Arquivos Textos Visual Mix (CADPROD.txt)|CADPROD.txt"
            CommonDialog.DialogTitle = "Abrir Arquivos Textos Binários"
            CommonDialog.FilterIndex = 1
            'CommonDialog.FileName = "txtGMS002_??_???.txt"
            CommonDialog.CancelError = False
            CommonDialog.ShowOpen
            If Len(CommonDialog.FileName) > 0 Then
                If Dir(CommonDialog.FileName) <> "" Or InStr(CommonDialog.FileName, "*") = 0 Or InStr(CommonDialog.FileName, "?") = 0 Then
                    If Dir(CommonDialog.FileName) <> "" Then
                        Me.Frame1.Visible = True
                        CaminhoDoArquivo = CommonDialog.FileName
                        Me.MSFlexGrid1.Visible = False
                        ExibirTXTBinário CommonDialog.FileName         'RegAntigo)
                        Me.MSFlexGrid1.Visible = True
                    End If
                End If
            Else
                Exit Sub
            End If
            Exit Sub
    End Select
    
End Sub
