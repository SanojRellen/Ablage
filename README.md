Sub ProcessData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Template")
    
    ' Find the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Step 1: Split column A by ":"
    For i = 1 To lastRow
        If InStr(ws.Cells(i, "A").Value, ":") > 0 Then
            ws.Cells(i, "B").Value = Split(ws.Cells(i, "A").Value, ":")(1)
            ws.Cells(i, "A").Value = Split(ws.Cells(i, "A").Value, ":")(0)
        End If
    Next i
    
    ' Step 2: Remove spaces before the first letter or number in column B
    For i = 1 To lastRow
        ws.Cells(i, "B").Value = Trim(ws.Cells(i, "B").Value)
    Next i
    
    ' Step 3: Put the last 6 characters of A16 into B16
    ws.Cells(16, "B").Value = Right(ws.Cells(16, "A").Value, 6)
    
    ' Step 4: Replace all "/" with "." in cells B4 to B7
    For i = 4 To 7
        ws.Cells(i, "B").Value = Replace(ws.Cells(i, "B").Value, "/", ".")
    Next i
End Sub
