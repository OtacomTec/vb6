VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3413
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15583915
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDCAAB&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceitar"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Escreva o Nome da máquina"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label2"
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Os direitos autorais foram mantidos!!!
'
'Caso melhore ou encontre erros neste projeto, favor poste para mim
'no endereço:
'anderson_afn@hotmail.com
'ou
'anderson_afn@ig.com.br
'ou
'poste no site da ScriptBrasil
'
'Valeu!!!!

'---------------------------------------------------------------------------------------
' Empresa    : VSoft,Lda.
' Projecto   : ProgMáquina
' Data/Hora  : 29-05-2007 14:21
' Autor      : Morpheus
' Descrição  : Diz quais os programas instalados na nossa máquina ou remota.
'---------------------------------------------------------------------------------------


Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Obtener_Software(Optional ByVal computer_name As String = ".")
    
    Dim Listado As String
    Dim objWMIService As Object
    Dim objsoftware As Object
    Dim colSoftware As Object
    
    Dim i As Integer
    Me.MousePointer = vbHourglass
    
    On Error GoTo Error_sub
    
    If computer_name = "" Then computer_name = "."
    
    Set objWMIService = GetObject("winmgmts:" & _
    "{impersonationLevel=impersonate}!\\" & _
    computer_name & _
    "\root\cimv2")
    
    Set colSoftware = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Product")
    
    
    If colSoftware.Count > 0 Then
        Dim item As ListItem
        For Each objsoftware In colSoftware
            Set item = ListView1.ListItems.Add(, , _
            ChequearNULO(objsoftware.Caption))
            item.SubItems(1) = ChequearNULO(objsoftware.version)
            item.SubItems(2) = ChequearNULO(objsoftware.installlocation)
            i = i + 1
        Next
    Else
        MsgBox "Não há software instalado neste computador", vbInformation
    End If
    
    Set objWMIService = Nothing
    Me.MousePointer = 0
    Exit Sub
Error_sub:
    MsgBox Err.Description, vbCritical
    Me.MousePointer = 0
    On Error Resume Next
    Set objWMIService = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ChequearNULO(Valor As Variant) As String
    If IsNull(Valor) Then
        ChequearNULO = vbNullString
    Else
        ChequearNULO = Valor
    End If
End Function

Private Sub Command1_Click()
    Call Obtener_Software(Text1.Text)
End Sub

Private Sub Form_Load()
    Command1.Caption = "Aceitar"
    '[usando WMI]
    Me.Caption = "VSoft, Lda. - Quais os programas que estão instalados em sua máquina."
    Text1.Text = ""
    
    With ListView1
        .ColumnHeaders.Add , , "Software"
        .ColumnHeaders.Add , , "Versão"
        .ColumnHeaders.Add , , "Local"
        .View = lvwReport
        .GridLines = True
    End With
    
End Sub



