Attribute VB_Name = "Module2"
Function MinChartNumber() As Double
    
    Dim srs As Series
    Dim FirstTime  As Boolean
    Dim MinNumber As Double
    
    'First Time Looking at This Chart?
    FirstTime = True
    
    'Determine Chart's Overall Min From Connected Data Source
    For Each srs In ActiveChart.SeriesCollection
        
        'Determine Minimum value in Series
        MinNumber = Application.WorksheetFunction.Min(srs.Values)
        
        'Store value if currently the overall Minimum Value
        If FirstTime = True Then
            MinChartNumber = MinNumber
        ElseIf MinNumber < MinChartNumber Then
            MinChartNumber = MinNumber
        End If
        
        'First Time Looking at This Chart?
        FirstTime = False
    
    Next srs

End Function
