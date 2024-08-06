
Dim cellValue As String
Dim newValue As String

' Read the value from C21
cellValue = templateWs.Range("C21").Value

' Replace the comma with a dot, convert to a numeric value, then format as "0%"
newValue = Format(CDbl(Replace(cellValue, ",", ".")), "0%")

' Set the new value back to C21
templateWs.Range("C21").Value = newValue
