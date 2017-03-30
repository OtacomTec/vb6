VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gráfico"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboTipoGrafico 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   75
      Width           =   9270
   End
   Begin MSChart20Lib.MSChart mscGrafico 
      Height          =   7260
      Left            =   0
      OleObjectBlob   =   "Form1.frx":000C
      TabIndex        =   0
      Top             =   480
      Width           =   9285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'Autor:  Rodrigo dos S. Gomes
'Função: Gera gráficos
'Data:   30/11/2008
'-------------------------------------------------------------------------------------------

Option Explicit

'Variáveis
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private mavarValor()        As Variant
    Private mctTipoGrafico      As VtChChartType
    Private mblnDadosEmLinha    As Boolean

Private Sub cboTipoGrafico_Click()
    
    mblnDadosEmLinha = True
    Me.MousePointer = vbHourglass
    
    Select Case cboTipoGrafico.ItemData(cboTipoGrafico.ListIndex)
        Case 1: mctTipoGrafico = VtChChartType3dBar
        Case 2: mctTipoGrafico = VtChChartType2dBar
        Case 3: mctTipoGrafico = VtChChartType3dLine: mblnDadosEmLinha = False
        Case 4: mctTipoGrafico = VtChChartType2dLine: mblnDadosEmLinha = False
        Case 5: mctTipoGrafico = VtChChartType3dArea
        Case 6: mctTipoGrafico = VtChChartType2dArea: mblnDadosEmLinha = False
        Case 7: mctTipoGrafico = VtChChartType3dStep
        Case 8: mctTipoGrafico = VtChChartType2dStep: mblnDadosEmLinha = False
        Case 9: mctTipoGrafico = VtChChartType3dCombination
        Case 10: mctTipoGrafico = VtChChartType2dCombination
        Case 11: mctTipoGrafico = VtChChartType2dPie
        Case 12: mctTipoGrafico = VtChChartType2dXY
    End Select
    
    sCarregaDados
    sGeraGrafico
    Me.MousePointer = vbDefault
    
End Sub

Private Sub sGeraGrafico()
    
    Dim intContador As Integer
    
    With mscGrafico
        .Title = "Teste de Gráfico"
        .chartType = mctTipoGrafico
        .ChartData = mavarValor
        .Plot.DataSeriesInRow = True
        .Plot.UniformAxis = True
        
        '*** FORMATA O CAMPO DO GRÁFICO A ESQUERDA *************************************
        For intContador = 1 To .Plot.Axis(1).Labels.Count
           .Plot.Axis(1).Labels(intContador).Format = "0""0%"""
           .Plot.Axis(1).Labels(intContador).VtFont.Name = "Tahoma"
           .Plot.Axis(1).Labels(intContador).VtFont.Size = 8
        Next
        '*******************************************************************************
        
        '*** COLOCA PORCENTAGEM NO GRÁFICO *********************************************
        For intContador = 1 To .Plot.SeriesCollection.Count
            With .Plot.SeriesCollection(intContador).DataPoints.Item(-1).DataPointLabel
                .LocationType = VtChLabelLocationTypeOutside
                .Component = VtChLabelComponentPercent
                .PercentFormat = "0%"
                .VtFont.Size = 8
            End With
            DoEvents
        Next intContador
        '*******************************************************************************
        
        '*** TIRA A PORCENTAGEM DO GRÁFICO - SOMENTE NÚMERO ****************************
        'For intContador = 1 To .Plot.SeriesCollection.Count
        '    With .Plot.SeriesCollection(intContador).DataPoints.Item(-1).DataPointLabel
        '        .LocationType = VtChLabelLocationTypeOutside
        '        .Component = VtChLabelComponentValue
        '        .VtFont.Size = 7
        '    End With
        'Next intContador
        '*******************************************************************************
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'Área de plotagem de acordo com o respectivo tipo de gráfico
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            .Plot.DataSeriesInRow = mblnDadosEmLinha
            .Plot.Axis(VtChAxisIdX).AxisScale.Hide = mblnDadosEmLinha
            .Plot.Axis(VtChAxisIdZ).AxisScale.Hide = True
            .Plot.UniformAxis = False
            
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            'Propriedades que alteram somente gráficos 3D
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                .Plot.Projection = VtProjectionTypeOblique
                .Plot.View3d.Set 75, 30
                .Plot.DepthToHeightRatio = 2
                .Plot.WidthToHeightRatio = 2
                .Plot.xGap = 0    'Espaço entre divisões X-Axis (Comprimento)
                .Plot.zGap = 0.8  'Espaço entre divisões Z-Axis (Profundidade)
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'Retira a etiqueta do gráfico de pizza
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Select Case mctTipoGrafico
                Case VtChChartType2dPie
                    .ColumnLabelIndex = 1
                    .ColumnLabel = " "
            End Select
            
        .Footnote.Text = "Texo de Rodapé"
        .AllowDynamicRotation = True
        .ShowLegend = mblnDadosEmLinha

   End With
   
End Sub

Private Sub Form_Load()
    
    ReDim mavarValor(1 To 6, 1) As Variant
    
    With cboTipoGrafico
        .AddItem "Barras 3D" 'VtChChartType3dBar
        .ItemData(.NewIndex) = 1
        
        .AddItem "Barras 2D" 'VtChChartType2dBar
        .ItemData(.NewIndex) = 2
        
        .AddItem "Linhas 3D" 'VtChChartType3dLine
        .ItemData(.NewIndex) = 3
        
        .AddItem "Linhas 2D" 'VtChChartType2dLine
        .ItemData(.NewIndex) = 4
        
        .AddItem "Área 3D" 'VtChChartType3dArea
        .ItemData(.NewIndex) = 5
        
        .AddItem "Área 2D" 'VtChChartType2dArea
        .ItemData(.NewIndex) = 6
        
        .AddItem "Passo 3D" 'VtChChartType3dStep
        .ItemData(.NewIndex) = 7
        
        .AddItem "Passo 2D" 'VtChChartType2dStep
        .ItemData(.NewIndex) = 8
        
        .AddItem "Combinação 3D" 'VtChChartType3dCombination
        .ItemData(.NewIndex) = 9
        
        .AddItem "Combinação 2D" 'VtChChartType2dCombination
        .ItemData(.NewIndex) = 10
        
        .AddItem "Pizza 2D" 'VtChChartType2dPie
        .ItemData(.NewIndex) = 11
        
        .ListIndex = 0
        mctTipoGrafico = 0
    End With
    
End Sub

Private Sub sCarregaDados()

    Dim rs              As New Recordset
    Dim intContador     As Integer
    
    Set rs = sp_Transacoes("PES")
    
    If rs.EOF Then
        ReDim mavarValor(1 To 1, 1)
        Exit Sub
    End If
    
    ReDim mavarValor(1 To rs.RecordCount, 1)
    
    For intContador = 1 To rs.RecordCount
        
        mavarValor(intContador, 0) = "Codigo " & rs!orderid&
        mavarValor(intContador, 1) = rs!EmployeeId
        
        rs.MoveNext
        
    Next intContador

End Sub
