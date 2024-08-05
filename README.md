# Ablage
Zum abgreifen

Dim baseValueText As String
Dim baseValue As Double
Dim relativeValue As String
Dim absoluteValue As Double

' Read the base value from B31 (text format)
baseValueText = templateWs.Range("B31").Value
baseValueText = Replace(baseValueText, ".", "") ' Remove the thousand separator if any
baseValue = CDbl(Replace(baseValueText, ",", ".")) ' Convert to double

' Convert relative coupon values to absolute values based on B31
' For cell C20
relativeValue = templateWs.Range("C20").Value
relativeValue = CDbl(Replace(Replace(relativeValue, "%", ""), ",", ".")) / 100 ' Convert to decimal
If baseValue = 100 Then
    absoluteValue = relativeValue * baseValue
ElseIf baseValue = 1000 Then
    absoluteValue = relativeValue * (baseValue / 10)
End If
templateWs.Range("C20").Value = Format(absoluteValue, "0.00") & " EUR"

' For cell D20
relativeValue = templateWs.Range("D20").Value
relativeValue = CDbl(Replace(Replace(relativeValue, "%", ""), ",", ".")) / 100 ' Convert to decimal
If baseValue = 100 Then
    absoluteValue = relativeValue * baseValue
ElseIf baseValue = 1000 Then
    absoluteValue = relativeValue * (baseValue / 10)
End If
templateWs.Range("D20").Value = Format(absoluteValue, "0.00") & " EUR"

' For cell E20
relativeValue = templateWs.Range("E20").Value
relativeValue = CDbl(Replace(Replace(relativeValue, "%", ""), ",", ".")) / 100 ' Convert to decimal
If baseValue = 100 Then
    absoluteValue = relativeValue * baseValue
ElseIf baseValue = 1000 Then
    absoluteValue = relativeValue * (baseValue / 10)
End If
templateWs.Range("E20").Value = Format(absoluteValue, "0.00") & " EUR"

' For cell F20
relativeValue = templateWs.Range("F20").Value
relativeValue = CDbl(Replace(Replace(relativeValue, "%", ""), ",", ".")) / 100 ' Convert to decimal
If baseValue = 100 Then
    absoluteValue = relativeValue * baseValue
ElseIf baseValue = 1000 Then
    absoluteValue = relativeValue * (baseValue / 10)
End If
templateWs.Range("F20").Value = Format(absoluteValue, "0.00") & " EUR"

' For cell G20
relativeValue = templateWs.Range("G20").Value
relativeValue = CDbl(Replace(Replace(relativeValue, "%", ""), ",", ".")) / 100 ' Convert to decimal
If baseValue = 100 Then
    absoluteValue = relativeValue * baseValue
ElseIf baseValue = 1000 Then
    absoluteValue = relativeValue * (baseValue / 10)
End If
templateWs.Range("G20").Value = Format(absoluteValue, "0.00") & " EUR"

' For cell H20
relativeValue = templateWs.Range("H20").Value
relativeValue = CDbl(Replace(Replace(relativeValue, "%", ""), ",", ".")) / 100 ' Convert to decimal
If baseValue = 100 Then
    absoluteValue = relativeValue * baseValue
ElseIf baseValue = 1000 Then
    absoluteValue = relativeValue * (baseValue / 10)
End If
templateWs.Range("H20").Value = Format(absoluteValue, "0.00") & " EUR"


 
