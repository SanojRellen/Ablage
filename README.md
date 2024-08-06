Dim modifiedValue As String

' Read the value from B31 and remove the last 7 characters
modifiedValue = Left(templateWs.Range("B31").Value, Len(templateWs.Range("B31").Value) - 7)

' Copy the result to J20 and format as number
templateWs.Range("J20").Value = CDbl(modifiedValue)
templateWs.Range("J20").NumberFormat = "0.00" ' Optional: format as number with 2 decimal places
