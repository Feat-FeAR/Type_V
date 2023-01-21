Attribute VB_Name = "Module_AV"
Sub SaveAsValue()
'VBA code: Save all worksheets as values only
	
	'NOTE:
	'ThisWorkbook refers to the workbook in which Excel VBA code is being executed
	'ActiveWorkbook refers to the Excel Workbook that currently has focus, meaning is the front-facing Excel Window
	'Here ActiveWorkbook is used to make this subroutine suitable to be turned into an Add-in
	'Otherwise, using ThisWorkbook, the subroutine would work on FeAR_MacroSet.xlam file (!!!)
	
	Dim pos As Integer
	Dim outName As String
	Dim pathName As String
	Dim wsh As Worksheet
	
	'Check for filename extension and, in case, remove it
	'InStrRev returns the position of an occurrence of one string within another...
	'...from the end of the string
	'If string-match is not found InStrRev returns 0
	pos = InStrRev(ActiveWorkbook.Name, ".", -1, vbTextCompare)
	
	If pos = 0 Then
		MsgBox ("Current workbook still unsaved!" & Chr(10) & Chr(13) & "The subroutine has been aborted"), vbExclamation, "Warning"
		Exit Sub 'Stop any further execution of this Sub
	End If
	
	'Left returns a Variant containing a specified number of characters from the left side of a string
	outName = Left(ActiveWorkbook.Name, (pos - 1))
	outName = (outName & " - Static.xlsx")
	pathName = ActiveWorkbook.Path & Application.PathSeparator & outName
	
	'Make a copy of the file...
	ActiveWorkbook.SaveCopyAs (pathName)
	Workbooks.Open (pathName)
	Workbooks(outName).Activate
	
	'...now make the copy static
	For Each wsh In ActiveWorkbook.Worksheets
		wsh.Cells.Copy
		wsh.Cells.PasteSpecial xlPasteValues
	Next
	
	'Quit Copy (or Cut) mode and remove the moving border
	Application.CutCopyMode = False
	
	'Remove selection and save the static copy
	For Each wsh In ActiveWorkbook.Worksheets
		wsh.Activate
		ActiveSheet.Cells(1, 1).Select
	Next
	ActiveWorkbook.Worksheets(1).Activate 'Make the first sheet active
	
	ActiveWorkbook.Save
	
	MsgBox "A static copy has been saved", vbInformation, "Information"
	
End Sub
