Sub ExtractPercentage()
    Dim value As String
    Dim result As Double

    ' Get the value from C23 and remove any non-numeric characters
    value = Range("C23").Value
    result = Val(Replace(value, "% p.a.", ""))

    ' Place the result in D24
    Range("D24").Value = result
End Sub
