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
	
	'Initial values
	MyMacroVer = "4.0 - 15-Aug-2020"
	TestedWith = "6.8.1"
	
	'To bypass the alert message when the temporary sheet is deleted
	Application.DisplayAlerts = False 'Switch off the alert button
	
	Worksheets.Add.Name = "Versions"
	Worksheets("Versions").Cells(1, 1).Formula = "=VER()"
	'NOTE: Ver() is not a worksheet-functions available to VBA,
	'so you can't directly use something like this:
	'RealStatsVer = Application.WorksheetFunction.Ver()
	
	If IsError(Cells(1, 1).Value) = True Then
		RealStatsVer = "NONE!"
	Else
		RealStatsVer = Cells(1, 1).Value
	End If
	
	Worksheets("Versions").Delete
	
	'Restore the alert message
	Application.DisplayAlerts = True 'Switch on the alert button
	
	MsgBox ("FeAR MacroSet version " & MyMacroVer & vbNewLine _
	& "Tested with RealStats release " & TestedWith & vbNewLine & vbNewLine _
	& "Installed RealStats release " & vbNewLine & RealStatsVer), _
	vbInformation, "Version Info"

End Sub
