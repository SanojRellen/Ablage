Sub CopyAndFilterData()

    ' Step 1: Copy Data from Sheet1 to Sheet2
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long, copyRange As Range
    Dim destRow As Long

    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")

    ' Find the last filled row in Sheet1, column G
    lastRow = ws1.Cells(ws1.Rows.Count, "G").End(xlUp).Row

    ' Set the range to be copied (from A5 to last row in column G)
    Set copyRange = ws1.Range("A5:G" & lastRow)

    ' Paste the range in Sheet2 starting at A9 (including formatting)
    destRow = 9
    copyRange.Copy
    ws2.Range("A" & destRow).PasteSpecial Paste:=xlPasteAll

    ' Step 2: Filter and Delete Rows based on Date Range
    Dim startDate As Date, endDate As Date
    Dim row As Long
    Dim cell As Range

    ' Get the start and end dates from Sheet2, B1 and B2
    startDate = ws2.Range("B1").Value
    endDate = ws2.Range("B2").Value

    ' Loop through the rows in Sheet2, starting at row 9
    For row = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row To 9 Step -1
        Set cell = ws2.Cells(row, "B")
        
        ' Check if the cell contains a date and if it's outside the range
        If IsDate(cell.Value) Then
            If cell.Value < startDate Or cell.Value > endDate Then
                ws2.Rows(row).Delete
            End If
        End If
    Next row

End Sub
