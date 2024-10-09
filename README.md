
Dim formattedValue As String
Dim result As String

' Assuming the formatted value is already in the form "1.000,00"
formattedValue = "1.000,00"

' Store all but the last three characters
result = Left(formattedValue, Len(formattedValue) - 3)

' Now "result" will store "1.000"
