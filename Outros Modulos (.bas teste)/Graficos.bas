Attribute VB_Name = "Graficos"
Public Function Monta_Grafico_Barra(strSQL As String, Nome_MSchart As Object, Indice_X As Integer, Indice_Y As Integer, Titulo_Grafico As String, Optional banco As String, Optional Aplicacao As String) As String
    
    'Os Campos: Indice_X e Indice_Y Indicam quais os campos do eixo X e eixo Y respectivamente
    'de acordo com a Query passada para recordset
    
    'OBS: 1º Campo da Query o indice é igual a 0 e assim sucessivamente
    
    Dim max_colunas As Integer
    Dim max_linhas As Integer
    Dim rstGrafico As New ADODB.Recordset
    Dim conexao_Grafico As New DLLConexao_Sistema.conexao
    
    'Indicando o banco à conectar-se
    conexao_Grafico.Initial_Catalog = "BDSupervisor"
    
    conexao_Grafico.Abrir_conexao ("PDV")
       
    rstGrafico.CursorLocation = adUseClient
    rstGrafico.Open strSQL, conexao_Grafico.CNConexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    max_colunas = rstGrafico.Fields.Count
    max_linhas = rstGrafico.RecordCount
    
    ReDim Values(1 To max_linhas, 1 To max_colunas)
    rstGrafico.MoveFirst

    L = 1
    C = 1

    Do While rstGrafico.EOF = False
        C = 1
        Values(L, C) = rstGrafico.Fields(Indice_X)
        C = C + 1
        Values(L, C) = rstGrafico.Fields(Indice_Y)
        rstGrafico.MoveNext
        L = L + 1
    Loop

    'Nome_MSchart.RowCount = 1
    'Nome_MSchart.ColumnCount = 1
    Nome_MSchart.ChartData = Values
    Nome_MSchart.chartType = VtChChartType2dBar
    Nome_MSchart.Title = Titulo_Grafico
    Nome_MSchart.Plot.SeriesCollection(1).LegendText = "Finalizadoras"
    Nome_MSchart.Plot.SeriesCollection(2).Position.Hidden = True
    Nome_MSchart.Plot.SeriesCollection(3).Position.Hidden = True
    
End Function

Public Function Monta_Grafico_Pizza(strSQL As String, Nome_MSchart As Object, Indice_X As Integer, Titulo_Grafico As String, Optional banco As String, Optional Aplicacao As String) As String
        
    Dim max_linhas As Integer
    Dim rstGrafico As New ADODB.Recordset
    Dim conexao_Grafico As New DLLConexao_Sistema.conexao
    
    'Indicando o banco à conectar-se
    conexao_Grafico.Initial_Catalog = "BDSupervisor"
    
    conexao_Grafico.Abrir_conexao ("PDV")
       
    rstGrafico.CursorLocation = adUseClient
    rstGrafico.Open strSQL, conexao_Grafico.CNConexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    max_linhas = rstGrafico.RecordCount
    
    ReDim Values(1 To max_linhas)
    rstGrafico.MoveFirst
    
    L = 1
    C = 1
    
    Do While rstGrafico.EOF = False
        Values(L) = rstGrafico.Fields(Indice_X)
        rstGrafico.MoveNext
        L = L + 1
    Loop
           
    ' Send the data to the chart.
    'Nome_MSchart.RowCount = 1
    'Nome_MSchart.ColumnCount = 1
    Nome_MSchart.ChartData = Values
    Nome_MSchart.chartType = VtChChartType2dPie
    Nome_MSchart.Title = Titulo_Grafico
        
    Form1.MSChart1.RowLabel = "testinho"
    
    rstGrafico.MoveFirst
    Nome_MSchart.Plot.SeriesCollection(1).LegendText = rstGrafico.Fields(1) & " " & rstGrafico.Fields(3)
    rstGrafico.MoveNext
    Nome_MSchart.Plot.SeriesCollection(2).LegendText = rstGrafico.Fields(1) & " " & rstGrafico.Fields(3)
        
       
End Function
