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
	
	'Cell range selection check
	If TypeName(Selection) = "Range" Then
		'Put user Selection into a Range variable
		Set selRange = Selection 'ActiveWorkbook.ActiveSheet implied
	Else
		MsgBox "Invalid selection!", vbOKOnly + vbExclamation, "Warning"
		Exit Sub
	End If
	
	r = selRange.Rows.Count    'Number of selected rows
	c = selRange.Columns.Count 'Number of selected columns
	
	'InputBox for user-input log base
	b = Application.InputBox( _
		Prompt:="y = log(x)" & vbNewLine & vbNewLine & "Choose Logarithm Base", _
		Title:="Logarithmic Transformation", _
		Default:="e", _
		Type:=1 + 2)
	
	'Cancel option returns Boolean False
	If b = "False" Then Exit Sub
	'NOTE: Using the String "False" in place of the Boolean False
	'allows distinguishing between False and 0 !!
	
	'Special case - natural base
	If b = "e" Then
		natural = True
		b = 2 'Assign a positive number !=1 just to pass the next If-statement with no Warnings
	End If
	
	'Base Check
	'NOTE: VBA doesn't short-circuit,
	'but "Type Mismatch" can be easily avoided using Variant
	If Len(b) = 0 Or b <= 0 Or b = 1 Or Not IsNumeric(b) Then 
		MsgBox "Invalid Logarithm Base" _
			& vbNewLine _
			& "The subroutine has been aborted", _
			vbOKOnly + vbExclamation, "Warning"
		Exit Sub 'Stop Sub execution in case of: empty input, non-numeric entry, b<=0, or b=1
	End If
	
	'Make entries static to get rid of possible inner references
	selRange.Copy
	selRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone 'Paste Values
	
	'Log-transform values by overwriting
	For i = 1 To r '--For Each cell In selRange-- as an alternative, with --Dim cell As Range-- (no need for i,j,r,c)
		For j = 1 To c
			With selRange.Cells(i, j)
				val = .Value
				If IsEmpty(val) Then
					.Value = ""
				Else
					If IsNumeric(val) Then
						If val > 0 Then
							If natural Then
								.Value = Application.WorksheetFunction.Ln(val) 'Application can be omitted
							Else
								.Value = Application.WorksheetFunction.Log(val, b)
							End If
						Else
							.Value = "#NUM!"
						End If
					Else
						.Value = "#VALUE!"
					End If
				End If
			End With
		Next j
	Next i

End Sub
