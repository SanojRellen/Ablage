Dim cellValue As String

' Format C21 as text
templateWs.Range("C21").NumberFormat = "@"

' Read the value from C21
cellValue = templateWs.Range("C21").Value

' Delete the last 4 characters and add % sign
cellValue = Left(cellValue, Len(cellValue) - 4) & "%"

' Set the new value back to C21
templateWs.Range("C21").Value = cellValue
