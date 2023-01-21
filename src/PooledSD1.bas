Attribute VB_Name = "Module_ES"
Option Explicit

Sub PooledSD()
'VBA code: Compute pooled standard deviation and effect size
	
	Dim selRange As Range	'Current selection will be used for output
	Dim MyWS As Worksheet	'The ActiveSheet in the ActiveWorkbook
	Dim MyCell As Range		'For-loop auxiliary variable
	Dim rng1 As Range		'Group_1 containing-data range
	Dim rng2 As Range		'Group_2 containing-data range
	
	Dim var1 As Double		'Group_1 (sample) variance
	Dim var2 As Double		'Group_2 (sample) variance
	Dim n1 As Integer		'Group_1 sample size (-1)
	Dim n2 As Integer		'Group_2 sample size (-1)
	Dim pSD As Double		'Pooled Standard Deviation (output)
	Dim d As Double			'Cohen's d (output)
	
	'Cell range selection check
	If TypeName(Selection) = "Range" Then
		'Put the uppermost left cell of the current selection into a Range variable
		Set selRange = Selection.Cells(1, 1) 'ActiveWorkbook.ActiveSheet implied
	Else
		MsgBox "Please select a range of cells before running this macro!"
		Exit Sub
	End If
	
	Set MyWS = ActiveWorkbook.ActiveSheet
	
	'Check if the output range (a 5x2 grid) is empty or not
	For Each MyCell In MyWS.Range(selRange.Cells(1, 1), selRange.Cells(5, 2))
		If IsEmpty(MyCell) = False Then
			MsgBox ("Output will overwrite existing data!" _
			& vbNewLine _
			& "Press Cancel in the next dialog window if this is not acceptable."), vbExclamation, "Warning"
			Exit For
		End If
	Next MyCell
	
	'Error-Handling
	On Error Resume Next 'Continue executing the code immediately after the statement that generated the error
		
		'InputBox for user-defined cell range
		Set rng1 = Application.InputBox( _
			Title:="Data Selection - Group 1", _
			Prompt:="Select the range containing Group_1 data", _
			Type:=8) 'A cell reference, as a Range object
		
		'Stop Sub execution in case of Cancel
		If rng1 Is Nothing Then Exit Sub
		
		'InputBox for user-defined cell range
		Set rng2 = Application.InputBox( _
			Title:="Data Selection - Group 2", _
			Prompt:="Select the range containing Group_2 data", _
			Type:=8) 'A cell reference, as a Range object
		
		'Stop Sub execution in case of Cancel
		If rng2 Is Nothing Then Exit Sub
	
	'Disable error handler
	On Error GoTo 0
	
	'Compute Output
	var1 = Application.WorksheetFunction.Var_S(rng1)
	n1 = rng1.Count - 1
	var2 = Application.WorksheetFunction.Var_S(rng2)
	n2 = rng2.Count - 1
	
	pSD = Sqr((n1 * var1 + n2 * var2) / (n1 + n2))
	d = Abs(Application.WorksheetFunction.Average(rng1) - Application.WorksheetFunction.Average(rng2)) / pSD
	
	'Print Output
	selRange.Cells(1, 1).Value = "Effect Size Analysis (FeAR)"
	MyWS.Range(selRange.Cells(2, 1), selRange.Cells(2, 2)).Borders(xlEdgeTop).LineStyle = xlContinuous
	selRange.Cells(2, 1).Value = "Size n1"
	selRange.Cells(2, 2).Value = n1 + 1
	selRange.Cells(3, 1).Value = "Size n2"
	selRange.Cells(3, 2).Value = n2 + 1
	selRange.Cells(4, 1).Value = "Pooled SD"
	selRange.Cells(4, 2).Value = pSD
	selRange.Cells(5, 1).Value = "Cohen's d"
	selRange.Cells(5, 2).Value = d
	MyWS.Range(selRange.Cells(5, 1), selRange.Cells(5, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
	MyWS.Range(selRange.Cells(2, 1), selRange.Cells(5, 1)).Font.Italic = True
	selRange.Cells(5, 1).Characters(Start:=9, Length:=1).Font.FontStyle = "Regular"
	selRange.Cells(2, 1).Characters(Start:=7, Length:=1).Font.Subscript = True
	selRange.Cells(3, 1).Characters(Start:=7, Length:=1).Font.Subscript = True
	
End Sub