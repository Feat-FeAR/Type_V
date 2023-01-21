Attribute VB_Name = "Module_AV"
Option Explicit	'Force explicit variable declaration

Sub SaveAsValue()
'VBA code: Save all worksheets as values only
	
	Dim pos As Integer			'Auxiliary variable to catch file extension
	Dim outName As String		'Name of the static copy
	Dim pathName As String		'Name of the static copy with its complete path
	
	Dim MyWB As Workbook		'Active Workbook (original)
	Dim StaticWB As Workbook	'Static Workbook (copy)
	Dim MyWS As Worksheet		'General sheet of the Static Workbook
	
	'NOTE:
	'ThisWorkbook:   refers to the workbook in which Excel VBA code is
	'				 being executed
	'ActiveWorkbook: refers to the Excel Workbook that currently has focus,
	'                meaning is the front-facing Excel Window
	'Here, ActiveWorkbook is used to make this subroutine suitable to be turned
	'into an Add-in, otherwise, using ThisWorkbook, the subroutine would work on
	'FeAR_MacroSet.xlam file (!!!)
	Set MyWB = ActiveWorkbook
	
	'Check for filename extension and, in case, remove it
	'InStrRev returns the position of an occurrence of one string within another
	'from the end of the string
	'If string-match is not found InStrRev returns 0
	pos = InStrRev(MyWB.Name, ".", -1, vbTextCompare)
	
	If pos = 0 Then
		MsgBox "The current workbook is still unsaved!" _
		& vbNewLine _
		& "The subroutine has been aborted", _
		vbOKOnly + vbExclamation, "Warning"
		Exit Sub 'Stop Sub execution
	End If
	
	'Left returns a Variant containing a specified number of characters from the
	'left side of a string
	outName = Left(MyWB.Name, (pos - 1)) 'Implicit downcast Variant-to-String
	outName = outName & " - Static.xlsx"
	pathName = MyWB.Path & Application.PathSeparator & outName
	
	'Make a copy of the file...
	MyWB.SaveCopyAs pathName
	Workbooks.Open pathName
	Set StaticWB = Workbooks(outName)
	
	'...now make the copy static
	For Each MyWS In StaticWB.Worksheets
		MyWS.Cells.Copy
		MyWS.Cells.PasteSpecial xlPasteValues
	Next
	
	'Quit Copy (or Cut) mode and remove the moving border
	Application.CutCopyMode = False
	
	'Remove whole-sheet selection and make the first sheet active
	For Each MyWS In StaticWB.Worksheets
		MyWS.Activate
		ActiveSheet.Cells(1, 1).Select
	Next
	StaticWB.Worksheets(1).Activate
	
	'Save the static copy
	StaticWB.Save
	
	MsgBox "A static copy has been saved", vbOKOnly + vbInformation, "Information"
	
End Sub
