Attribute VB_Name = "Module_AV"
Sub SaveAsValue()
'VBA code: Save all worksheets as values only
        
        Dim wsh As Worksheet
        
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
        
End Sub
