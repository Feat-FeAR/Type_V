Attribute VB_Name = "Module3"
Function MaxChartNumber() As Double
    
    Dim srs As Series
    Dim FirstTime  As Boolean
    Dim MaxNumber As Double
    
    'First Time Looking at This Chart?
    FirstTime = True
    
    'Determine Chart's Overall Max From Connected Data Source
    For Each srs In ActiveChart.SeriesCollection
    
        'Determine Maximum value in Series
        MaxNumber = Application.WorksheetFunction.Max(srs.Values)
        
        'Store value if currently the overall Maximum Value
        If FirstTime = True Then
            MaxChartNumber = MaxNumber
        ElseIf MaxNumber > MaxChartNumber Then
            MaxChartNumber = MaxNumber
        End If
        
        'First Time Looking at This Chart?
        FirstTime = False
    
    Next srs

End Function
