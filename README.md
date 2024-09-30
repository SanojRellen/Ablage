Sub SeparateAndCopyValues()
    Dim inputString As String
    Dim values() As String
    Dim i As Integer
    
    ' Get the input from cell C15
    inputString = Range("C15").Value
    
    ' Split the string by slashes
    values = Split(inputString, "/")
    
    ' Loop through the array and copy values to cells F15 to K15
    For i = LBound(values) To UBound(values)
        ' Check if the index is within the range of F15 to K15 (0 to 5)
        If i < 6 Then
            Range("F" & 15).Offset(0, i).Value = Trim(values(i))
        End If
    Next i
End Sub







Dim cell As Range
For Each cell In Range("C20:H20")
    If Not IsEmpty(cell.Value) Then
        ' Your operation here
    End If
Next cell
