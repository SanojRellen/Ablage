Sub SortBlocksByColumnE_Corrected()

    Dim ws2 As Worksheet
    Dim lastRow As Long
    Dim startRow As Long, endRow As Long
    Dim i As Long

    Set ws2 = ThisWorkbook.Sheets("Sheet2")
    
    ' Find the last filled row in column E
    lastRow = ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row

    startRow = 9 ' Starting from row 9 in Sheet2

    ' Loop through the rows to identify blocks and sort them
    For i = startRow To lastRow + 1
        
        ' If the current row is empty or it's the end of the sheet (indicating the end of a block)
        If Application.WorksheetFunction.CountA(ws2.Rows(i)) = 0 Or i = lastRow + 1 Then
            ' End of a block, so sort the previous block
            If i > startRow Then
                ' Set endRow as one row before the empty row (or lastRow if it's the end of the sheet)
                endRow = i - 1

                ' Check if there's something to sort (i.e., if the block has more than one row)
                If endRow > startRow Then
                    ' Sort the block from startRow to endRow based on column E (largest to smallest)
                    ws2.Sort.SortFields.Clear
                    ws2.Sort.SortFields.Add Key:=ws2.Range("E" & startRow & ":E" & endRow), _
                                            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                    With ws2.Sort
                        .SetRange ws2.Range("A" & startRow & ":F" & endRow)
                        .Header = xlNo
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                    End With
                End If
            End If
            
            ' Set the next block's start row as the row after the empty row
            startRow = i + 1
        End If
    Next i

End Sub
