Sub CopyMatchingValuesToColumnS()
    Dim ws As Worksheet
    Dim searchValue As String
    Dim lastRow As Long
    Dim i As Long
    Dim outputRow As Long

    ' Set the worksheet to the active sheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Get the value to search from cell S1
    searchValue = ws.Range("S1").Value

    ' Clear column S from row 7 to 22
    ws.Range("S7:S22").ClearContents

    ' Initialize the output row counter
    outputRow = 7

    ' Get the last used row in column D
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    ' Loop through column D starting from row 7
    For i = 7 To lastRow
        If ws.Cells(i, "D").Value = searchValue Then
            ' Copy value from column A to column S
            ws.Cells(outputRow, "S").Value = ws.Cells(i, "A").Value
            outputRow = outputRow + 1
        End If
    Next i
End Sub
