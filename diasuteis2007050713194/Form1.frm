VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calcular Dias Úteis"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQtFeriados 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3690
      TabIndex        =   11
      Top             =   2475
      Width           =   495
   End
   Begin VB.TextBox txtDiasNaoUteis 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2340
      TabIndex        =   10
      Top             =   2475
      Width           =   495
   End
   Begin VB.TextBox txtDiasUteis 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   855
      TabIndex        =   9
      Top             =   2475
      Width           =   495
   End
   Begin VB.TextBox txtFeriados 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   244
      TabIndex        =   6
      Text            =   "01/01/2007;20/02/2007;06/04/2007;01/05/2007"
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox txtDtFinal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2411
      MaxLength       =   10
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtDtIni 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   244
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "01/01/2007"
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "CALCULAR"
      Height          =   615
      Left            =   1290
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Resultado em dias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Feriados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2865
      TabIndex        =   12
      Top             =   2520
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Não Úteis:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1380
      TabIndex        =   8
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Úteis:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   315
      TabIndex        =   7
      Top             =   2520
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Feriados (Separados por ""ponto e virgula"" = ;)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   251
      TabIndex        =   5
      Top             =   1320
      Width           =   3945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2411
      TabIndex        =   3
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   244
      TabIndex        =   1
      Top             =   480
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalcular_Click()
    Dim intUteis As Integer, intNaoUteis As Integer, intQtFeriados As Integer
    
    calcularDiasUteis txtDtIni.Text, txtDtFinal.Text, txtFeriados.Text, intUteis, intNaoUteis, intQtFeriados
    
    txtDiasUteis.Text = CStr(intUteis)
    txtDiasNaoUteis.Text = CStr(intNaoUteis)
    txtQtFeriados.Text = CStr(intQtFeriados)
    
    MsgBox "Calculo concluído!", vbInformation, "Calcular dias úteis!"

End Sub

Private Sub Form_Load()
    txtDtFinal.Text = Format(Now, "dd/MM/yyyy")
End Sub

Public Function calcularDiasUteis(ByVal dataInicial As String, ByVal dataFinal As String, ByVal Feriados As String, ByRef intDiasUteis As Integer, intDiasNaoUteis As Integer, ByRef intFeriados As Integer)
    
    Dim DiasUteis As Long, DiasNaoUteis As Integer, DiaAtual As Integer, qtFeriados As Integer
    Dim xLoop As Integer, vFeriados As Variant, I As Long

    On Error Resume Next

    '===========================================================================================
    '   calcular dias úteis entre um intervalo de datas, tirando os feriados, se for o caso.
    '   por Gustavo de Almeida Rodrigues - gustavo_rodrigues@terra.com.br
    '   Maio de 2007
    '===========================================================================================
    
    If Not IsDate(dataInicial) Or dataInicial <> Format(dataInicial, "dd/MM/yyyy") Then
        MsgBox "Data inicial invalida!" & vbNewLine & dataInicial, vbExclamation, "dd/MM/yyyy"
        Exit Function
    ElseIf Not IsDate(dataFinal) Or dataFinal <> Format(dataFinal, "dd/MM/yyyy") Then
        MsgBox "Data final invalida!" & vbNewLine & dataFinal, vbExclamation, "dd/MM/yyyy"
        Exit Function
    ElseIf InStr(1, Feriados, ";") > 0 And Feriados <> "" Then
        vFeriados = Split(Feriados, ";")
        For xLoop = 0 To UBound(vFeriados) - 1
            If Not IsDate(vFeriados(xLoop)) Or vFeriados(xLoop) <> Format(vFeriados(xLoop), "dd/MM/yyyy") Then
                MsgBox "Data " & CStr(vFeriados(xLoop)) & " invalida!", vbExclamation, "dd/MM/yyyy"
                Exit Function
            End If
            DoEvents
        Next
    ElseIf (Not IsDate(Feriados) And Feriados <> "") Or (Feriados <> Format(Feriados, "dd/MM/yyyy") And Feriados <> "") Then
        MsgBox "Data " & CStr(Feriados) & " invalida!", vbExclamation, "dd/MM/yyyy"
        Exit Function
    End If
    DoEvents
    DoEvents
    For I = CDate(dataInicial) To CDate(dataFinal)
        DiaAtual = Weekday(I)
        Select Case DiaAtual
            Case vbSunday, vbSaturday
                DiasNaoUteis = DiasNaoUteis + 1
            Case Else
                DiasUteis = DiasUteis + 1
        End Select
        If Feriados <> "" Then
            If (InStr(1, Feriados, CStr(Format(CDate(I), "dd/MM/yyyy"))) > 0) Then
                DiasNaoUteis = DiasNaoUteis + 1
                qtFeriados = qtFeriados + 1
            End If
        End If
        DoEvents
    Next
    DoEvents
    intDiasNaoUteis = DiasNaoUteis
    intDiasUteis = DiasUteis - qtFeriados
    intFeriados = qtFeriados
    

End Function

