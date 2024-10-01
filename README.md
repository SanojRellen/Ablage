Sub CalculatePayoffSchedule()
    Dim startDate As String
    Dim day As String
    Dim month As String
    Dim formattedDay As String
    Dim payoffMonths(1 To 4) As Integer
    Dim monthNames(1 To 4) As String
    Dim i As Integer
    
    ' Get the start date from cell C17 as text
    startDate = Range("C17").Value
    
    ' Extract day and month from the start date (assuming format is DD/MM/YYYY)
    day = Left(startDate, 2)
    month = Mid(startDate, 4, 2)
    
    ' Convert day to a formatted day with suffix (e.g., 27 -> 27th)
    Select Case CInt(day)
        Case 1, 21, 31
            formattedDay = day & "st"
        Case 2, 22
            formattedDay = day & "nd"
        Case 3, 23
            formattedDay = day & "rd"
        Case Else
            formattedDay = day & "th"
    End Select
    
    ' Paste the formatted day in cell D17
    Range("D17").Value = formattedDay
    
    ' Paste the month in cell E17
    Range("E17").Value = CInt(month)
    
    ' Determine the payoff months (every 3 months from the start month)
    For i = 0 To 3
        payoffMonths(i + 1) = (CInt(month) + (i * 3) - 1) Mod 12 + 1
    Next i
    
    ' Paste the payoff months in cells F17 to I17
    For i = 1 To 4
        Range("F17").Offset(0, i - 1).Value = payoffMonths(i)
    Next i
    
    ' Convert the payoff months to full month names and paste in cells F18 to I18
    For i = 1 To 4
        monthNames(i) = MonthName(payoffMonths(i))
        Range("F18").Offset(0, i - 1).Value = monthNames(i)
    Next i
End Sub
