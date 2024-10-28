Sub ExtractPercentage()
    Dim value As String
    Dim result As Double

    ' Get the value from C23 and remove any non-numeric characters
    value = Range("C23").Value
    result = Val(Replace(value, "% p.a.", ""))

    ' Place the result in D24
    Range("D24").Value = result
End Sub



Sub FormatPercentageString()
    Dim mittlere_Rendite As String
    Dim formattedRendite As String

    ' Example value for mittlere_Rendite
    mittlere_Rendite = "0.0185"
    
    ' Convert to percentage format with a comma as decimal separator
    formattedRendite = Format(CDbl(mittlere_Rendite) * 100, "0,00") & "%"

    ' Output the formatted result
    MsgBox formattedRendite  ' or assign to a cell, e.g., Range("D24").Value = formattedRendite
End Sub

