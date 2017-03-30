VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4125
      Left            =   60
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   90
      Width           =   10395
   End
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   4125
      Left            =   60
      OleObjectBlob   =   "Form1.frx":2356
      TabIndex        =   1
      Top             =   4350
      Width           =   10395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Values() As String
Private Sub Form_Load()
    cn.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Administrador\Meus documentos\grafico.mdb;Persist Security Info=False")
    rs.CursorLocation = adUseClient
    rs.Open "SELECT DFNome_Candidato,DFPercentual FROM Candidato", cn, adOpenKeyset, adLockOptimistic
    Dim max_colunas As Integer
    Dim max_linhas As Integer
    
    max_colunas = rs.Fields.Count
    max_linhas = rs.RecordCount
    
    ReDim Values(1 To max_linhas, 1 To max_colunas)
    rs.MoveFirst

    L = 1
    C = 1

    Do While rs.EOF = False
        C = 1
        Values(L, C) = rs!DFNome_Candidato
        C = C + 1
        Values(L, C) = rs!DFPercentual
        rs.MoveNext
        L = L + 1
    Loop

    ' Send the data to the chart.
    MSChart1.RowCount = 1
    MSChart1.ColumnCount = 1
    MSChart1.ChartData = Values
    MSChart1.chartType = VtChChartType2dBar
    MSChart1.Plot.SeriesCollection(1).LegendText = "VOTOS"
    
    '------------------------------------------
    
    
    ReDim Values(1 To max_linhas)
    rs.MoveFirst
    
    L = 1
    C = 1
    
    Do While rs.EOF = False
        Values(L) = rs!DFPercentual
    '    MSChart2.Plot.SeriesCollection(L).LegendText = rs!DFNome_Candidato
        rs.MoveNext
        L = L + 1
    Loop
    
       
    ' Send the data to the chart.
    MSChart2.RowCount = 1
    MSChart2.ColumnCount = 1
    MSChart2.ChartData = Values
    MSChart2.chartType = VtChChartType2dPie
    MSChart2.Title = "TESTE"
    MSChart2.RowLabel = "testinho"
    
    rs.MoveFirst
    
    L = 1
    
    Do While rs.EOF = False
        MSChart2.Plot.SeriesCollection(L).LegendText = rs!DFNome_Candidato
        rs.MoveNext
        L = L + 1
    Loop

    With MSChart2.Legend
        .Location.Visible = True
        .Location.LocationType = VtChLocationTypeRight
        .TextLayout.HorzAlignment = VtHorizontalAlignmentRight


    End With
    
End Sub

