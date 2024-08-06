
Dim cell As Range
Dim multiplier As Double

' Get the value from J20
multiplier = templateWs.Range("J20").Value

' Loop through the range C20:H20
For Each cell In templateWs.Range("C20:H20")
    ' Multiply each cell's value by the value in J20 and put the result back into the cell
    cell.Value = cell.Value * multiplier
    
    ' Append ",00 EUR" and format as text
    cell.Value = Format(cell.Value, "0") & ",00 EUR"
    cell.NumberFormat = "@"
Next cell
