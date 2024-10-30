Dim i As Integer
For i = 0 To 5 ' Loop from 0 to 5 (for a maximum of 6 items)
    If i <= UBound(dstelist) Then ' Check if there is a value at this index
        ws.Range("C14").Offset(0, i).Value = dstelist(i)
    Else
        ws.Range("C14").Offset(0, i).ClearContents ' Clear cell if no corresponding list value
    End If
Next i




Dim datelist As Variant
Dim i As Integer

' Split values in B15 into an array (assuming comma-separated values)
datelist = Split(ws.Range("B15").Value, ",")

' Loop to copy all but the last value of datelist into J15 to N15
For i = 0 To UBound(datelist) - 1 ' Stop at the second-to-last item
    ws.Range("J15").Offset(0, i).Value = datelist(i) ' Fill cell with value from datelist
Next i

' Clear any remaining cells in the range J15:N15 if datelist has fewer than 5 values
For i = UBound(datelist) To 4
    ws.Range("J15").Offset(0, i).ClearContents
Next i

