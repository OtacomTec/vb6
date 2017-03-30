VERSION 5.00
Object = "{F871B372-BD4B-4283-A10E-0AB1C61FA941}#1.0#0"; "JanChart.ocx"
Begin VB.Form frmMain 
   Caption         =   "JanChart Demo"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin prjJanChart.JanChart JanChart1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11456
      BackColor       =   -2147483633
      ChartWidth      =   8640
      ChartHeight     =   6240
      CaptionNumber   =   1
      CaptionBackColor=   8438015
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WorksheetNumber =   1
      WorksheetGroupNumber=   1
      WorksheetWidth  =   400
      BeginProperty DataFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Index           =   2
      Left            =   5520
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   2
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdRandomize 
      Caption         =   "&Randomize"
      Height          =   495
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   6720
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************************
' Hi,
' This is VB project for JanChart demo
' The chart will have:
'   2 column for caption (IO Number, Date In)
'   4 column for data (1, 2, 3, 4)
'   20 row for data (1, 2, ..., 20)
'
' If you need some help or have some question or bug report
' please feel free to contact me
' My email address is: jimmi_kembaren@sancerta.com
'
' Btw, I live in Bandung - Java
'
' Chrs,
' Jimmi A. Kembaren
'**********************************************************************


Private Sub cmdAbout_Click(Index As Integer)
    JanChart1.About
End Sub

Private Sub cmdExit_Click(Index As Integer)
    End
End Sub

Private Sub cmdRandomize_Click(Index As Integer)
    Draw_Data
End Sub

Private Sub Form_Load()
    'draw chart
    Init_Chart
    
    'draw data
    Draw_Data
End Sub



Private Sub Init_Chart()
    
    'init gantt chart
    
    JanChart1.CaptionNumber = 2                     'set caption column 2
    JanChart1.Set_CaptionWidth 1, 1400              'set 1st caption width
    JanChart1.Set_CaptionWidth 2, 1200              'set 2nd caption width
    
    JanChart1.Set_CaptionName 1, "IO Number"        'set 1st caption label
    JanChart1.Set_CaptionName 2, "Date In"          'set 2nd caption label
    
    JanChart1.WorksheetNumber = 8                   'set worksheet column number
    JanChart1.WorksheetWidth = 1300                 'set worksheet column width
    
    JanChart1.WorksheetGroupNumber = 8              'set worksheet column group
    JanChart1.Set_WorksheetGroupLabel 1, "Week"     'set worksheet column group label
    
    JanChart1.Set_WorksheetLabel 1, "1"             'set 1st worksheet column label
    JanChart1.Set_WorksheetLabel 2, "2"             'set 2st worksheet column label
    JanChart1.Set_WorksheetLabel 3, "3"             'set 3st worksheet column label
    JanChart1.Set_WorksheetLabel 4, "4"             'set 4st worksheet column label
    
    JanChart1.Refresh
    
End Sub 'init_chart

Private Sub Draw_Data()
    'draw chart data
    Dim i As Integer
    Dim intNumData As Integer
    Dim strIONo As String, strDateIn As String
    Dim dblBarWidth1 As Double, dblBarStart1 As Double
    Dim dblBarWidth2 As Double, dblBarStart2 As Double
    
    'set number of data (number of row)
    intNumData = 20
    JanChart1.DataNumber = intNumData
    
    'draw every rows
    For i = 1 To intNumData
        strIONo = "JK-" & Year(Now()) & "-" & Right("0000" & i, 5)      'create io number
        strDateIn = Right("0" & i, 2) & " " & Format(Date, "mmm yyyy")  'create date in
        
        JanChart1.Set_DataCaptionLabel i, 1, strIONo        'set io number for this row
        JanChart1.Set_DataCaptionLabel i, 2, strDateIn      'set date in for this row
        
        dblBarStart1 = (3 * Rnd) + 0        'create randomize data
        dblBarWidth1 = (2 * Rnd) + 0.2      'create randomize data
        
        If (dblBarStart1 < JanChart1.WorksheetNumber) Then
            'draw bar
            JanChart1.Set_DataWorksheet i, dblBarStart1, dblBarStart1 + dblBarWidth1, &HFF8080
        End If
        
        dblBarStart2 = dblBarStart1 + dblBarWidth1  'create randomize data
        dblBarWidth2 = (1 * Rnd) + 0.1              'create randomize data
        
        If (dblBarStart2 < JanChart1.WorksheetNumber) Then
            'draw bar
            JanChart1.Set_DataWorksheet i, dblBarStart2, dblBarStart2 + dblBarWidth2, &HFF&
        End If
                
    Next i
    
    JanChart1.Refresh
        
End Sub 'draw_data
