Attribute VB_Name = "Module_BC"
Sub MakeBoxCool()
'VBA code: Change Box Plot layout 
'NOTE: This version does not handle secondary y-axis for negative values
	
	Dim AnyChartSelected As Object
	Dim yAxisLabel As String
	Dim FullColorCode As Long
	Dim FullColorCodeComp As Long 'Comp for Complementary
	Dim RedComp As Integer 'Comp for Component
	Dim GreenComp As Integer
	Dim BlueComp As Integer
	Dim SeriesNum As Long
	Dim LineThickness As Double
	Dim PointSize As Double
	
	'Initial values
	LineThickness = 1.5
	PointSize = 7
	
	'Check for chart selection
	Set AnyChartSelected = ActiveWorkbook.ActiveChart
	
	If AnyChartSelected Is Nothing Then
		MsgBox ("No chart selected!" & Chr(10) & Chr(13) & "The subroutine has been aborted"), vbExclamation, "Warning"
		Exit Sub 'Stop any further execution of this Sub
	End If
	
	'Count the number of series in the ActiveChart
	SeriesNum = ActiveChart.SeriesCollection.Count
	
	'Open the ColorPicker dialog box, applying the RGB color (200,0,0) as the default
	'then store the selected color by editing the 50th color position
	'NOTE: there are 56 color positions available in Excel
	If Application.Dialogs(xlDialogEditColor).Show(50, 200, 0, 0) = True Then
		'Set the variable FullColorCode equal to the value selected the DialogBox
		FullColorCode = ActiveWorkbook.Colors(50)
	Else
		Exit Sub 'Stop any further execution of this Sub if the user selected cancel
	End If
	
	'Retrieve RGB components
	RedComp = FullColorCode Mod 256
	GreenComp = FullColorCode \ 256 Mod 256 'Operator \ is the integer-division-operator in VBA
	BlueComp = FullColorCode \ 65536 Mod 256
	
	'Define Complementary Color
	FullColorCodeComp = RGB(Abs(RedComp - 150), Abs(GreenComp - 150), Abs(BlueComp - 150))
	
	'InputBox for user-defined y-axis label
	yAxisLabel = InputBox("Insert Y-axis label", "Y-axis label", "Y-Axis")
	If Len(yAxisLabel) = 0 Then
		MsgBox ("No string inserted!" & Chr(10) & Chr(13) & "The subroutine has been aborted"), vbExclamation, "Warning"
		Exit Sub 'Stop any further execution of this Sub if the user selected cancel or inserted a 0-character string
	End If
	
	'----------------
	'Do the restyling
	'----------------
	
	'Remove unwanted elements
	With ActiveChart
		.HasTitle = False 'Remove Title
		.Axes(xlValue).HasMajorGridlines = False 'Remove Grid lines
		.Axes(xlValue).HasMinorGridlines = False 'Remove Grid lines
		'.ChartArea.Format.Fill.Visible = msoFalse 'No background color fill
		'.PlotArea.Format.Fill.Visible = msoFalse 'No background color fill
	End With
	
	'Change Box color
	For i = 3 To 4
		With ActiveChart.SeriesCollection(i).Format
			.Fill.Visible = msoFalse
			.Line.ForeColor.RGB = FullColorCode
			.Line.Weight = LineThickness
		End With
	Next i
	
	'Change Whisker color
	For i = 1 To 2
		With ActiveChart.SeriesCollection(2 * i).ErrorBars.Format.Line
			.ForeColor.RGB = FullColorCode
			.Weight = LineThickness
		End With
	Next i
	
	'Change Mean Marker color
	'NOTE: you have to format the line first, then the markers
	With ActiveChart.SeriesCollection(5)
		.Format.Line.Visible = msoTrue
		.Format.Line.Weight = LineThickness
		.Format.Line.Visible = msoFalse
		.Format.Fill.Visible = msoFalse
		.MarkerSize = PointSize
		.MarkerForegroundColor = FullColorCodeComp
	End With
	
	'Change Outlier Marker color
	'NOTE: If (SeriesNum > 5) outlier points are present
	If SeriesNum > 5 Then
		For i = 6 To SeriesNum
			With ActiveChart.SeriesCollection(i)
				.Format.Line.Visible = msoTrue
				.Format.Line.Weight = LineThickness
				.Format.Line.Visible = msoFalse
				.Format.Fill.Visible = msoFalse
				.MarkerSize = PointSize - 0.5
				.MarkerForegroundColor = FullColorCodeComp
			End With
		Next i
	End If
	
	'Axis Labels
	With ActiveChart
		.Axes(xlCategory, xlPrimary).HasTitle = False 'Remove x-axis Label
		.Axes(xlValue, xlPrimary).HasTitle = True 'Add y-axis Label
		.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = yAxisLabel 'y-axis Label name
		.Axes(xlValue, xlPrimary).AxisTitle.Characters.Font.Size = 14 'y-axis Label size
	End With
	With ActiveChart.Axes(xlCategory).TickLabels.Font 'x-axis Tick font
		.Bold = msoTrue
		.Size = 12
	End With
	With ActiveChart.Axes(xlValue).TickLabels.Font 'y-axis Tick font
		.Bold = msoTrue
		.Size = 10
	End With
	
	'Axis Thickness and Color
	With ActiveChart.Axes(xlCategory, xlPrimary).Format.Line
		.Visible = msoTrue
		.Weight = LineThickness 'x-axis Thickness
		.ForeColor.RGB = RGB(0, 0, 0) 'x-axis Color
		.ForeColor.TintAndShade = 0
		.ForeColor.Brightness = 0
		.Transparency = 0
	End With
	With ActiveChart.Axes(xlValue, xlPrimary).Format.Line
		.Visible = msoTrue
		.Weight = LineThickness 'y-axis Thickness
		.ForeColor.RGB = RGB(0, 0, 0) 'y-axis Color
		.ForeColor.TintAndShade = 0
		.ForeColor.Brightness = 0
		.Transparency = 0
	End With
	
End Sub
