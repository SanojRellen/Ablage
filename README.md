Sub FormatNumberToGerman()
    Dim cell As Range
    Set cell = Range("C16")
    
    ' Format the number in English format with thousand separators
    cell.NumberFormat = "#,##0.00"
    
    ' Convert the number to German format (uses regional settings)
    cell.Value = Replace(cell.Text, ",", ";") ' Temporarily replace commas
    cell.Value = Replace(cell.Value, ".", ",") ' Replace dots with commas
    cell.Value = Replace(cell.Value, ";", ".") ' Replace temporary semicolons with dots
End Sub
