VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormMyDebug 
   Caption         =   "MyDebug"
   ClientHeight    =   5355
   ClientLeft      =   4530
   ClientTop       =   2325
   ClientWidth     =   6585
   Icon            =   "FRM0000-01-F1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6585
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1470
      Top             =   2670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRM0000-01-F1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRM0000-01-F1.frx":0556
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salva em arquivo texto"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   510
      Width           =   6405
   End
End
Attribute VB_Name = "FormMyDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        If Me.Height < 2100 Then
            Me.Height = 2100
        End If
        If Me.Width < 3100 Then
            Me.Width = 3100
        End If
        
        Text1.Height = Me.Height - Text1.Top - 500
        Text1.Width = Me.Width - Text1.Left - 200
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyPress = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lbCanal As Byte
    Select Case Button.Key
        Case "Salvar"
            
            lbCanal = FreeFile
            lstrNomeDoArquivo = "Debug " & Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hhmmss") & ".deb"
            Open App.Path & "\" & lstrNomeDoArquivo For Output Shared As #lbCanal
            Print #lbCanal, Text1.Text
            Close #lbCanal
            MsgBox "Arquivo " & lstrNomeDoArquivo & " foi Salvo!"
        Case "Imprimir"
            Printer.Print Text1.Text
            Printer.EndDoc
    End Select
End Sub


