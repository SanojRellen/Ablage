Dim cell As Range

' Loop through the range C20:H20
For Each cell In templateWs.Range("C20:H20")
    ' Remove the last 4 characters
    cell.Value = Left(cell.Value, Len(cell.Value) - 4)
    
    ' Reformat as number and divide by 100
    cell.Value = CDbl(cell.Value) / 100
Next cell
