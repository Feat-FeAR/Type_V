Attribute VB_Name = "Module_GV"
Option Explicit	'Force explicit variable declaration

Sub GetVersions()
'VBA code: Get Version of FeAR Macros and RealStats Add-in
	
	'NOTE:
	'The versioning convention for FeAR MacroSet is #Subs.#Release, where:
	'	#Subs is incremented whenever the number of Sub is changed
	'		(NOTE: #Subs not necessarily is the current number of Subs)
	'		Any #Subs increment resets #Release value to 0
	'	#Release is incremented after any change in a Sub code
	
	Dim MyMacroVer As String
	Dim TestedWith As String
	Dim RealStatsVer As String
	
	Dim MyWB As Workbook
	Dim MyWS As Worksheet
	Dim MyCell As Range
	
	'Initialization
	MyMacroVer = "4.1 - 17-Sep-2020"
	TestedWith = "7.3"
	
	Set MyWB = ActiveWorkbook
	MyWB.Worksheets.Add.Name = "Versions" 'Add a temporary worksheet
	Set MyWS = MyWB.Worksheets("Versions")
	Set MyCell = MyWS.Range("A1")
	
	'To bypass the alert message when the temporary sheet will be deleted
	Application.DisplayAlerts = False 'Switch off the alert button
	
	MyCell.Formula = "=VER()"
	'NOTE: Ver() is not a worksheet-functions available to VBA,
	'so you can't directly use something like this:
	'RealStatsVer = Application.WorksheetFunction.Ver()
	
	If IsError(MyCell.Value) = True Then
		RealStatsVer = "NONE!"
	Else
		RealStatsVer = MyCell.Value
	End If
	
	'Delete the temporary worksheet
	MyWS.Delete
	
	'Restore the alert message
	Application.DisplayAlerts = True 'Switch on the alert button
	
	MsgBox "FeAR MacroSet version " & MyMacroVer & vbNewLine _
	& "Tested with RealStats release " & TestedWith & vbNewLine & vbNewLine _
	& "Installed RealStats release " & vbNewLine & RealStatsVer, _
	vbOKOnly + vbInformation, "Version Info"
	
End Sub
