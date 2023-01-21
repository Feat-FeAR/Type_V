Attribute VB_Name = "Module_SE"
Option Explicit 'Force explicit variable declaration

Sub SuperExp()
'VBA code: Transform numbers to a text string formatted as scientific notation 
'          with the power of ten superscripted
    
    Dim selRange As Range   'User-selected range
    Dim r As Integer        'Number of selected rows
    Dim c As Integer        'Number of selected columns
    Dim i As Integer        'Row index
    Dim j As Integer        'Column index
    Dim val As Variant      'Cell value
    Dim s As Variant        'InputBox returned value
    Dim mult As String      'Multiplication symbol used for Sci-Exp Notation
    Dim strt As Integer     'Exponent position in the final string

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
    
    'VBA MsgBox doesn't know Radio Buttons but you can do it with an InputBox:
    s = Application.InputBox( _
        Title:="To Superscripted Sci-Exp", _
        Prompt:="Select the multiplication symbol to be used" & _
            vbCrLf & _
            vbCrLf & vbTab & "1 - Times Sign:   " & Chr(215) & _
            vbCrLf & vbTab & "2 - Middle Dot:   " & Chr(183) & _
            vbCrLf & vbTab & "3 - Whitespace:   " & _
            vbCrLf & vbTab, _
        Default:=1, _
        Type:=1)
    
    Select Case s
        Case 1
            mult = " " & Chr(215) & " " 'Multiplication Sign (also the default)
            strt = 10
        Case 2
            mult = " " & Chr(183) & " " 'Middle Dot
            strt = 10
        Case 3
            mult = " "                  'White Space
            strt = 8
        Case "False" 'Cancel option returns Boolean False...
            Exit Sub '(...but better use "False" to distinguish from 0)
        Case Else
            MsgBox "Invalid Option" _
                & vbNewLine _
                & "The subroutine has been aborted", _
                vbOKOnly + vbExclamation, "Warning"
        Exit Sub 
    End Select

    'Make entries static to get rid of possible inner references
    selRange.Copy
    selRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone 'Paste Values
    
    'Format entries as Scientific Notation Numbers
    selRange.NumberFormat = "0.00E+00"
    
    'String-transform values by overwriting
    For i = 1 To r '--For Each cell In selRange-- as an alternative, with --Dim cell As Range-- (no need for i,j,r,c)
        For j = 1 To c
            With selRange.Cells(i, j)
                val = .Value
                If IsEmpty(val) Then
                    .Value = ""
                Else
                    If IsNumeric(val) Then
                        'Convert to String using scientific format "0.00E+00"
                        val = Format(val, "Scientific")
                        
                        'Convert to exponential format with superscript exponent
                        .Value = Application.WorksheetFunction.Substitute(val,"E",mult & "10")
                        .Characters(Start:=strt, Length:=3).Font.Superscript = True
                    Else
                        .Value = "#VALUE!"
                    End If
                End If
            End With
        Next j
    Next i
        
End Sub
