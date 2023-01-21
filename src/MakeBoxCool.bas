Attribute VB_Name = "Module_BC"
Option Explicit	'Force explicit variable declaration

Sub MakeBoxCool()
'VBA code: Change Box Plot layout
'NOTE: This version does handle secondary y-axis for negative values
	
	Dim MyChart As Object			'User-selected chart
	Dim SeriesNum As Integer		'Number of series in the chart
	Dim FullColorCode As Long		'RGB color code (0-16,777,215)
	Dim FullColorCodeComp As Long	'RGB complementary color code (0-16,777,215)
	Dim RedComp As Integer			'Red component (0-255)
	Dim GreenComp As Integer		'Green component (0-255)
	Dim BlueComp As Integer			'Blue component (0-255)
	Dim yAxisLabel As String		'User-defined y-axis label
	Dim AxisFlag As Integer			'Auxiliary variable for axis selection
	Dim LineThickness As Double		'Line Thickness
	Dim PointSize As Double			'Mean marker and outlier size
	Dim AxisSize As Double			'Y-axis Label size
	Dim i As Integer				'For-loop auxiliary variable
	
	'Initialization
	AxisFlag = 1
	LineThickness = 1.5
	PointSize = 7
	AxisSize = 14
	
	'Check for chart selection
	Set MyChart = ActiveWorkbook.ActiveChart
	
	'Chart selection check
	If MyChart Is Nothing Then
		MsgBox "No chart selected!" & vbNewLine _
		& "The subroutine has been aborted", vbOKOnly + vbExclamation, "Warning"
		Exit Sub 'Stop Sub execution
	End If
	
	'Count the number of series in the ActiveChart
	SeriesNum = MyChart.SeriesCollection.Count
	
	'Open the ColorPicker dialog box, applying the RGB color (200,0,0) as the default
	'then store the selected color by editing the 50th color position
	'NOTE: there are 56 color positions available in Excel
	If Application.Dialogs(xlDialogEditColor).Show(50, 200, 0, 0) = True Then
		'Set the variable FullColorCode equal to the value selected the DialogBox
		FullColorCode = ActiveWorkbook.Colors(50)
	Else
		Exit Sub 'Stop Sub execution if the user selected Cancel
	End If
	
	'Retrieve RGB components:
	'RGB(RedComp,GreenComp,BlueComp) =
	'	RedComp*(256^0) + GreenComp*(256^1) + BlueComp*(256^2)
	RedComp = FullColorCode Mod 256
	GreenComp = FullColorCode \ 256 Mod 256 'Operator \ is the integer-division-operator in VBA
	BlueComp = FullColorCode \ 65536 Mod 256
	
	'Define Complementary Color
	FullColorCodeComp = RGB(Abs(RedComp - 150), Abs(GreenComp - 150), Abs(BlueComp - 150))
	
	'InputBox for user-defined y-axis label
	yAxisLabel = Application.InputBox( _
		Prompt:="Insert Y-axis label", _
		Title:="Y-axis label", _
		Default:="Y-Axis", _
		Type:=2)
	
	'Cancel option returns Boolean False, here casted to String beacouse of initial Dim
	If yAxisLabel = "False" Then Exit Sub
	
	'----------------
	'Do the restyling
	'----------------
	
	'Remove unwanted elements
	With MyChart
		.HasTitle = False 'Remove Title
		.Axes(xlValue).HasMajorGridlines = False 'Remove Grid lines
		.Axes(xlValue).HasMinorGridlines = False 'Remove Grid lines
		'.ChartArea.Format.Fill.Visible = msoFalse 'No background color fill
		'.PlotArea.Format.Fill.Visible = msoFalse 'No background color fill
	End With
	
	'Change Box color
	For i = 3 To 4
		With MyChart.SeriesCollection(i).Format
			.Fill.Visible = msoFalse
			.Line.ForeColor.RGB = FullColorCode
			.Line.Weight = LineThickness
		End With
	Next i
	
	'Change Whisker color
	For i = 1 To 2
		With MyChart.SeriesCollection(2 * i).ErrorBars.Format.Line
			.ForeColor.RGB = FullColorCode
			.Weight = LineThickness
		End With
	Next i
	
	'Change Mean Marker color
	'NOTE: you have to format the line first, then the markers
	With MyChart.SeriesCollection(5)
		.Format.Line.Visible = msoTrue
		.Format.Line.Weight = LineThickness
		.Format.Line.Visible = msoFalse
		.Format.Fill.Visible = msoFalse
		.MarkerSize = PointSize
		.MarkerForegroundColor = FullColorCodeComp
	End With
	
	'Change Outlier Marker Color
	'NOTE: If (SeriesNum > 5) outlier points are present
	If SeriesNum > 5 Then
		For i = 6 To SeriesNum
			With MyChart.SeriesCollection(i)
				.Format.Line.Visible = msoTrue
				.Format.Line.Weight = LineThickness
				.Format.Line.Visible = msoFalse
				.Format.Fill.Visible = msoFalse
				.MarkerSize = PointSize - 0.5
				.MarkerForegroundColor = FullColorCodeComp
			End With
		Next i
	End If
	
	'Axes restyling
	
	'Check if a secondary y-axis is present for negative values
	'NOTE: Values 1 and 2 can be used in place of Group Names xlPrimary and
	'xlSecondary, respectively
	If MyChart.Axes.Count = 3 Then
		AxisFlag = 2
	End If
	
	'Axis Labels
	With MyChart
		.Axes(xlCategory, xlPrimary).HasTitle = False 'Remove x-axis Label
		.Axes(xlValue, AxisFlag).HasTitle = True 'Add y-axis Label
		.Axes(xlValue, AxisFlag).AxisTitle.Characters.Text = yAxisLabel 'y-axis Label name
		.Axes(xlValue, AxisFlag).AxisTitle.Characters.Font.Size = AxisSize 'y-axis Label size
	End With
	With MyChart.Axes(xlCategory, xlPrimary).TickLabels.Font 'x-axis Tick font
		.Bold = msoTrue
		.Size = AxisSize - 2
	End With
	With MyChart.Axes(xlValue, AxisFlag).TickLabels.Font 'y-axis Tick font
		.Bold = msoTrue
		.Size = AxisSize - 4
	End With
	
	'Axis Thickness and Color
	With MyChart.Axes(xlCategory, xlPrimary).Format.Line
		.Visible = msoTrue
		.Weight = LineThickness 'x-axis Thickness
		.ForeColor.RGB = RGB(0, 0, 0) 'x-axis Color
		.ForeColor.TintAndShade = 0
		.ForeColor.Brightness = 0
		.Transparency = 0
	End With
	With MyChart.Axes(xlValue, AxisFlag).Format.Line
		.Visible = msoTrue
		.Weight = LineThickness 'y-axis Thickness
		.ForeColor.RGB = RGB(0, 0, 0) 'y-axis Color
		.ForeColor.TintAndShade = 0
		.ForeColor.Brightness = 0
		.Transparency = 0
	End With
	
	If AxisFlag = 2 Then
		'In the presence of negative data, swap primary and secondary y-axes (*see bottom note)
		With MyChart
			.SetElement (msoElementSecondaryCategoryAxisShow) 'Introduce secondary x-axis
			.Axes(xlCategory, xlPrimary).Crosses = xlMaximum 'Move primary y-axis on the right
			.Axes(xlCategory, xlSecondary).Crosses = xlAutomatic 'Move secondary y-axis on the left 
			.SetElement (msoElementSecondaryCategoryAxisNone) 'Remove secondary x-axis
			.Axes(xlValue, xlPrimary).Format.Line.Visible = msoFalse 'Remove primary y-axis
			.Axes(xlValue, xlPrimary).TickLabels.Font.Color = RGB(255, 255, 255) 'Whiten primary y-axis labels 
		End With
	End If
	
End Sub

'*NOTE:
'To switch primary and secondary y-axes you have to temporarily introduce
'the secondary x-axis (which Excel doesn't add by default).

'Go to Chart menu > Chart Options > Axes tab, check the same option of
'Secondary Category (X) Axis that is checked for Primary Category (X) Axis.

'Double click the primary X axis (bottom of the chart) and on the Scale tab,
'check Value (Y) Axis Crosses at Maximum Category.

'Double click the secondary X axis (top of the chart) and on the Scale tab,
'UN-check Value (Y) Axis Crosses at Maximum Category.

'Go to Chart menu > Chart Options > Axes tab, UN-check the option of
'Secondary Category (X) Axis that you selected above.
