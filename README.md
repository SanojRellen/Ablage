=IF(WEEKDAY(EDATE(L9,12),2)>5,WORKDAY(EDATE(L9,12)-1,1),EDATE(L9,12))




Dim inputDate As Date

' Set your input date (for example, the date could be in a cell like Range("A1").Value)
inputDate = Range("A1").Value

' Check if the date is a Saturday or Sunday
If Weekday(inputDate, vbMonday) > 5 Then
    ' If Saturday (6) or Sunday (7), add the appropriate number of days to get to Monday
    inputDate = inputDate + (8 - Weekday(inputDate, vbMonday))
End If

' Store the adjusted date back in a cell, e.g., in A2
Range("A2").Value = inputDate

