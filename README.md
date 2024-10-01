Sub CalculateFirstPaymentDate()
    Dim startDate As String
    Dim day As String
    Dim month As String
    Dim year As String
    Dim firstPaymentDate As Date
    Dim paymentMonth As Integer
    Dim paymentYear As Integer

    ' Get the start date from cell C17 as text (assuming format is DD/MM/YYYY)
    startDate = Range("C17").Value
    
    ' Extract day, month, and year from the start date
    day = Left(startDate, 2)
    month = Mid(startDate, 4, 2)
    year = Right(startDate, 4)
    
    ' Calculate the first payment month (add 3 months)
    paymentMonth = (CInt(month) + 3 - 1) Mod 12 + 1
    paymentYear = CInt(year)
    
    ' Adjust the year if the resulting month is less than the original month (i.e., crossed a year boundary)
    If paymentMonth <= CInt(month) Then
        paymentYear = paymentYear + 1
    End If
    
    ' Create the first payment date
    firstPaymentDate = DateSerial(paymentYear, paymentMonth, CInt(day))
    
    ' Paste the first payment date in cell D11
    Range("D11").Value = firstPaymentDate
End Sub
