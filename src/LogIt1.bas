Attribute VB_Name = "Module_LI"
Option Explicit	'Force explicit variable declaration

Sub LogIt()
'VBA code: Log-transform cell content
	
	Dim selRange As Range	'User-selected range
	Dim r As Integer		'Number of selected rows
	Dim c As Integer		'Number of selected columns
	Dim i As Integer		'Row index
	Dim j As Integer		'Column index
	Dim b As Variant		'Log base - Variant can hold any type of value
	Dim natural As Boolean	'Flag for natural base (e)
	
	'Put user selection into a Range variable
	Set selRange = Selection
	
	r = selRange.Rows.Count    'Number of selected rows
	c = selRange.Columns.Count 'Number of selected columns
	
	'Make fields static to get rid of possible inner references
	Selection.Copy
	selRange.Cells(1, 1).Select 'Select the uppermost left cell
	Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone 'Paste Values
	
	'InputBox for user-input log base
	b = InputBox("Choose Logarithm base", "Logarithm base", "e") 'Cancel option returns a zero-length string ("")
	If b = "e" Then
		natural = True
		b = 2 'Assign any positive number (other than 1) just to go through the If control statement with no error messages
	End If
	
	'Check the base - NOTE: Len(b)=0 is True in the case of both empty input and Cancel option
	If Len(b) = 0 Or b <= 0 Or b = 1 Or Not IsNumeric(b) Then 'NOTE: VBA doesn't short-circuit, but "Type Mismatch" can be easily avoided using Variant
		MsgBox ("Invalid Logarithm Base" & Chr(10) & Chr(13) & "The subroutine has been aborted"), vbExclamation, "Warning"
		Exit Sub 'Stop any further execution of this Sub in case of: Cancel option, empty input, non-numeric entry, b<=0, or b=1
	End If
	
	'Log-transform values by overwriting
	For i = 1 To r
		For j = 1 To c
			If IsNumeric(selRange.Cells(i, j)) Then
				If selRange.Cells(i, j) > 0 Then
					If natural Then
						selRange.Cells(i, j) = Application.WorksheetFunction.Ln(selRange.Cells(i, j))
					Else
						selRange.Cells(i, j) = Application.WorksheetFunction.Log(selRange.Cells(i, j), b)
					End If
				Else
					selRange.Cells(i, j) = "#NUM!"
				End If
			Else
				selRange.Cells(i, j) = "#VALUE!"
			End If
		Next j
	Next i
	
End Sub
