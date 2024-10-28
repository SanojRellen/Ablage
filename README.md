
Sub ExtractNumber()
    Dim originalText As String
    Dim extractedNumber As Double

    ' Example text
    originalText = "2.35% p.a."
    
    ' Extract the numeric part and convert it to a double
    extractedNumber = Val(Replace(originalText, "% p.a.", ""))
    
    ' Output the result (e.g., assign to a cell)
    Range("D24").Value = extractedNumber
End Sub

