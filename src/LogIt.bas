Attribute VB_Name = "Module_LI"
Option Explicit	'Force explicit variable declaration

Sub LogIt()
'VBA code: Log-transform cell content: y=log(x)
	
	Dim selRange As Range	'User-selected range
	Dim r As Integer		'Number of selected rows
	Dim c As Integer		'Number of selected columns
	Dim i As Integer		'Row index
	Dim j As Integer		'Column index
	Dim b As Variant		'Log base - Variant can hold any type of value
	Dim val As Variant		'Cell value
	Dim natural As Boolean	'Flag for natural base (e)
	
	'Put user Selection into a Range variable
	Set selRange = Selection 'ActiveWorkbook.ActiveSheet implied
	
	r = selRange.Rows.Count    'Number of selected rows
	c = selRange.Columns.Count 'Number of selected columns
	
	'Make entries static to get rid of possible inner references
	selRange.Copy
	selRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone 'Paste Values
	
	'InputBox for user-input log base
	b = InputBox("y = log(x)" _
		& vbNewLine & vbNewLine _
		& "Choose Logarithm Base", _
		"Logarithmic Transformation", "e") 'Cancel option returns a zero-length string ("")
	If b = "e" Then
			natural = True
			b = 2 'Assign a positive number !=1 just to get through the next If control statement without Warning
	End If
	
	'Check the base - NOTE: Len(b)=0 is True in the case of both empty input and Cancel option
	If Len(b) = 0 Or b <= 0 Or b = 1 Or Not IsNumeric(b) Then 'NOTE: VBA doesn't short-circuit, but "Type Mismatch" can be easily avoided using Variant
			MsgBox ("Invalid Logarithm Base" _
				& vbNewLine _
				& "The subroutine has been aborted"), _
				vbExclamation, "Warning"
			Exit Sub 'Stop any further execution of this Sub in case of: Cancel option, empty input, non-numeric entry, b<=0, or b=1
	End If
	
	'Log-transform values by overwriting
	For i = 1 To r
		For j = 1 To c
			val = selRange.Cells(i, j).Value
			If IsNumeric(val) Then
				If val > 0 Then
					If natural Then
						selRange.Cells(i, j).Value = Application.WorksheetFunction.Ln(val) 'Application can be omitted
					Else
						selRange.Cells(i, j).Value = Application.WorksheetFunction.Log(val, b)
					End If
				Else
					selRange.Cells(i, j).Value = "#NUM!"
				End If
			Else
				selRange.Cells(i, j).Value = "#VALUE!"
			End If
		Next j
	Next i

End Sub