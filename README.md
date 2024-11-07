Sub VlookupOneRowBelow()
    Dim lookupValue As String
    Dim lookupRange As Range
    Dim resultCell As Range
    Dim foundRow As Long

    ' Set your lookup value and range here
    lookupValue = "YourLookupValue"
    Set lookupRange = ThisWorkbook.Sheets("Sheet1").Range("A2:B10") ' Adjust range as needed

    ' Find the row where the VLOOKUP would match
    On Error Resume Next
    foundRow = Application.Match(lookupValue, lookupRange.Columns(1), 0)
    On Error GoTo 0

    If foundRow > 0 Then
        ' Get the cell one row below and in the second column
        Set resultCell = lookupRange.Cells(foundRow + 1, 2)
        
        ' Check if the cell is not empty, then perform action
        If Not IsEmpty(resultCell.Value) Then
            ' Replace this line with the action you want to perform
            MsgBox "Cell below is: " & resultCell.Value
        End If
    Else
        MsgBox "Lookup value not found."
    End If
End Sub
