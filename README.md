Dim i As Integer
For i = 0 To 5 ' Loop from 0 to 5 (for a maximum of 6 items)
    If i <= UBound(dstelist) Then ' Check if there is a value at this index
        ws.Range("C14").Offset(0, i).Value = dstelist(i)
    Else
        ws.Range("C14").Offset(0, i).ClearContents ' Clear cell if no corresponding list value
    End If
Next i
