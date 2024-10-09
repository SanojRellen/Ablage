Range("A1").Value = Replace(Replace(Format(Range("A1").Value, "#,##0.00"), ".", "X"), ",", ".")
Range("A1").Value = Replace(Range("A1").Value, "X", ",")


Dim value As Double
value = 0.1234 ' This would be 12.34%

' Apply percentage format and then replace the separators
Dim formattedValue As String
formattedValue = Replace(Replace(Format(value, "0.00%"), ".", "X"), ",", ".")
formattedValue = Replace(formattedValue, "X", ",")

' Now formattedValue will be in German format, e.g., "12,34%"



Dim value As Double
value = 1234.56

' Apply numeric format and then replace the separators
Dim formattedValue As String
formattedValue = Replace(Replace(Format(value, "0.00"), ".", "X"), ",", ".")
formattedValue = Replace(formattedValue, "X", ",")

' Now formattedValue will be in German format, e.g., "1.234,56"

