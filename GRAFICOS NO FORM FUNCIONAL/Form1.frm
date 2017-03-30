VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.0#0"; "MSOWC.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin OWC.ChartSpace ChartSpace1 
      Height          =   2835
      Left            =   1710
      TabIndex        =   3
      Top             =   2520
      Width           =   6825
      XMLData         =   $"Form1.frx":0000
      ScreenUpdating  =   -1  'True
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Utilizar folha de cálculo"
      Height          =   525
      Left            =   7260
      TabIndex        =   2
      Top             =   1920
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Utilizar o conjunto de registos ADO"
      Height          =   825
      Left            =   5490
      TabIndex        =   1
      Top             =   1590
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Utilizar matrizes"
      Height          =   525
      Left            =   3630
      TabIndex        =   0
      Top             =   1890
      Width           =   1245
   End
   Begin OWC.DataSourceControl DataSourceControl1 
      Left            =   840
      Top             =   930
      XMLData         =   $"Form1.frx":00C5
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Create arrays for the x-values and the y-values
    Dim xValues As Variant, yValues1 As Variant, yValues2 As Variant
    xValues = Array("Beverages", "Condiments", "Confections", _
                    "Dairy Products", "Grains & Cereals", _
                    "Meat & Poultry", "Produce", "Seafood")
    yValues1 = Array(104737, 50952, 78128, 117797, 52902, 80160, 47491, _
                     62435)
    yValues2 = Array(20000, 15000, 36000, 56000, 40000, 18000, 20000, _
                     33000)
    
    'Create a new chart
    Dim oChart As WCChart
    ChartSpace1.Clear
    ChartSpace1.Refresh
    Set oChart = ChartSpace1.Charts.Add
    
    'Add a title to the chart
    oChart.HasTitle = True
    oChart.Title.Caption = "Sales Per Category"
    
    'Add a series to the chart with the x-values and y-values
    'from the arrays and set the series type to a column chart
    Dim oSeries As WCSeries
    Set oSeries = oChart.SeriesCollection.Add
    With oSeries
        .Caption = "1995"
        .SetData chDimCategories, chDataLiteral, xValues
        .SetData chDimValues, chDataLiteral, yValues1
        .Type = chChartTypeColumnClustered
    End With
    
    'Add another series to the chart with the x-values and y-values
    'from the arrays and set the series type to a line chart
    Set oSeries = oChart.SeriesCollection.Add
    With oSeries
        .Caption = "1996"
        .SetData chDimCategories, chDataLiteral, xValues
        .SetData chDimValues, chDataLiteral, yValues2
        .Type = chChartTypeLineMarkers
    End With
    
    'Add a value axis to the right of the chart for the second series
    oChart.Axes.Add oChart.Axes(chAxisPositionLeft).Scaling, _
        chAxisPositionRight, chValueAxis

    'Format the Value Axes
    oChart.Axes(chAxisPositionLeft).NumberFormat = "$#,##0"
    oChart.Axes(chAxisPositionRight).NumberFormat = "0"
    oChart.Axes(chAxisPositionLeft).MajorUnit = 20000
    oChart.Axes(chAxisPositionRight).MajorUnit = 20000
    
    'Show the legend at the bottom of the chart
    oChart.HasLegend = True
    oChart.Legend.Position = chLegendPositionBottom


End Sub

Private Sub Command2_Click()
'Set up the DataSourceControl for the Chartspace
    Dim rsd As RecordsetDef
    DataSourceControl1.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\Arquivos de programas\Microsoft Visual Studio\VB98\nwind.mdb"
    Set rsd = DataSourceControl1.RecordsetDefs.AddNew( _
             "Select * from [Category Sales for 1995]", 3)
    With ChartSpace1
        .Clear
        .Refresh
        .DataSource = DataSourceControl1
        .DataMember = rsd.Name
    End With
     
    'This Chartspace will contain 2 charts. Make the layout so that the
    'charts are positioned horizontally
    ChartSpace1.ChartLayout = chChartLayoutHorizontal
    
    'Create a new bar chart from the query
    Dim oBarChart As WCChart
    Set oBarChart = ChartSpace1.Charts.Add
    With oBarChart
        .Type = chChartTypeBarClustered
        .SetData chDimCategories, 0, 0  'Categories are first field
        .SetData chDimValues, 0, 1      'Values are second field
    
        'Format the value axis for the bar chart so that it
        'shows values in thousands (i.e., 45000 displays as 45) and
        'in increments of 25000. Remove the gridlines
        With .Axes(chAxisPositionBottom)
            .NumberFormat = "0,"
            .MajorUnit = 25000
            .HasMajorGridlines = False
        End With
        
        'Change the color of the series and the plot area
        .SeriesCollection(0).Interior.Color = RGB(150, 0, 150)
        .PlotArea.Interior.Color = RGB(240, 240, 10)
    End With
    
    'Create a new exploded pie chart from the query
    Dim oPieChart As WCChart
    Set oPieChart = ChartSpace1.Charts.Add
    With oPieChart
        .Type = chChartTypePie
        .SetData chDimCategories, 0, 0  'Categories are first field
        .SetData chDimValues, 0, 1      'Values are second field
        .SeriesCollection(0).Explosion = 20
        
        'Add a legend to the bottom of the pie chart
        .HasLegend = True
        .Legend.Position = chLegendPositionBottom
        
        'Add a title to the chart
        .HasTitle = True
        .Title.Caption = "Sales by Category for 1995"
        .Title.Font.Bold = True
        .Title.Font.Size = 11
        
        'Make the chart width 50% the size of the bar chart's width
        .WidthRatio = 50
        
        'Show data labels on the slices as percentages
        With .SeriesCollection(0).DataLabelsCollection.Add
            .HasValue = False
            .HasPercentage = True
            .Font.Size = 8
            .Interior.Color = RGB(255, 255, 255)
        End With
        
    End With
    

End Sub

Private Sub Command3_Click()
   'Dynamically add a spreadsheet control to the form
   Dim oSheet As Spreadsheet
   Me.Controls.Add "OWC.Spreadsheet", "Sheet"
   Set oSheet = Me!Sheet
   
   'Fill the Sheet with data
   With oSheet
        .Range("A1:A10").Formula = "=Row()"
        .Range("B1:B10").Formula = "=A1^2"
        .Range("A12").Formula = "=Max(A1:A10)"
        .Range("B12").Formula = "=Max(B1:B10)"
   End With
   
   'Create an xy-scatter chart using the data in the spreadsheet
   Dim oChart As WCChart
   With ChartSpace1
        .Clear
        .Refresh
        .DataSource = oSheet.object
        Set oChart = .Charts.Add
        oChart.Type = chChartTypeScatterSmoothLineMarkers
        oChart.SetData chDimXValues, 0, "a1:a10"
        oChart.SetData chDimYValues, 0, "b1:b10"
   End With
   
   With oChart
        'Display the Axes Titles and
        'set the major units for the axes
        With .Axes(chAxisPositionBottom)
            .HasTitle = True
            .Title.Caption = "X"
            .Title.Font.Size = 8
            .MajorUnit = 1
        End With
        With .Axes(chAxisPositionLeft)
            .HasTitle = True
            .Title.Caption = "X Squared"
            .Title.Font.Size = 8
            .MajorUnit = 10
        End With
        
        'Set the maximum and minimum axis values
        .Scalings(chDimXValues).Maximum = oSheet.Range("A12").Value
        .Scalings(chDimXValues).Minimum = 1
        .Scalings(chDimYValues).Maximum = oSheet.Range("B12").Value
        
        'Change the marker and line styles for the series
        With .SeriesCollection(0)
            .Marker.Style = chMarkerStyleDot
            .Marker.Size = 6
            .Line.Weight = 1
            .Line.Color = RGB(255, 0, 0)
        End With
   End With
   
   'Remove the spreadsheet
   Me.Controls.Remove "Sheet"
   

End Sub

