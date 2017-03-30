VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   630
      TabIndex        =   1
      Top             =   450
      Width           =   2220
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5505
      Left            =   585
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   1035
      Width           =   9195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
    If Combo1.ListIndex = 0 Then
        MSChart1.Title = "Gráfico de Área"
        MSChart1.chartType = VtChChartType2dArea
    ElseIf Combo1.ListIndex = 1 Then
        MSChart1.Title = "Gráfico de Barras"
        MSChart1.chartType = VtChChartType2dBar
    Else
        MSChart1.Title = "Gráfico de Linhas"
        MSChart1.chartType = VtChChartType2dLine
    End If
    MSChart1.Refresh
End Sub


Private Sub Form_Load()
    
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = "SELECT TOP 10 TITLE, INT(Description) AS Vendas FROM TITLES WHERE DESCRIPTION LIKE '2%' AND INT(DESCRIPTION) < 100"
    rs.Open sql, com, 1, 1
    
    
    
    Combo1.AddItem ("Área")
    Combo1.AddItem ("Barras")
    Combo1.AddItem ("Linhas")
    Combo1.ListIndex = 0
    
    Set MSChart1.DataSource = rs
    MSChart1.Title = "Gráfico de Área"
    MSChart1.ColumnCount = 1
    MSChart1.RowCount = rs.RecordCount
    MSChart1.chartType = VtChChartType2dArea
    


    MSChart1.ShowLegend = True
    MSChart1.ColumnLabelIndex = 0
    
    i = 1
    While Not rs.EOF() And Not rs.BOF

            MSChart1.Row = i
            MSChart1.RowLabel = rs("Vendas")
            rs.MoveNext

            i = i + 1
    Wend
 
  
    MSChart1.Refresh
End Sub
