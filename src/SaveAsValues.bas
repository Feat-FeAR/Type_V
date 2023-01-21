Attribute VB_Name = "Module_AV"
Sub SaveAsValue()
'VBA code: Save all worksheets as values only
    
    Dim wsh As Worksheet
    Dim outName As String
    
    'Check for filename extension and remove it
    'InStrRev returns the position of an occurrence of one string within another...
    '...from the end of the string
    'If stringmatch is not found InStrRev returns 0
    pos = InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare)
    
    If pos = 0 Then
        MsgBox ("Workbook still unsaved!" & Chr(10) & Chr(13) & "Aborted"), vbExclamation, "Warning"
        Exit Sub 'Stop any further execution of this Sub
    End If
    
    'Do the job
    For Each wsh In ThisWorkbook.Worksheets
        wsh.Cells.Copy
        wsh.Cells.PasteSpecial xlPasteValues
    Next
    
    'This cancels Copy (or Cut) mode and removes the moving border
    Application.CutCopyMode = False
    
    For Each wsh In ThisWorkbook.Worksheets
        wsh.Activate
        ActiveSheet.Cells(1, 1).Select
    Next
    
    'Now save a static copy of the file
    'Left returns a Variant containing a specified number of characters from the left side of a string
    outName = Left(ThisWorkbook.Name, (pos - 1))
    outName = (outName & " - Static.xlsx")
    
    ThisWorkbook.SaveCopyAs (ThisWorkbook.Path & Application.PathSeparator & outName)
    
    MsgBox "A static copy has been saved", vbInformation, "Information"
        
End Sub
