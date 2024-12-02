Sub OrganizeDataByCriteria()
    Dim wsAll As Worksheet, wsByClients As Worksheet, wsByCurrencyPair As Worksheet, wsByTradeType As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim currencyPairs As Object, tradeTypes As Object
    Dim i As Long, currentValue As String

    ' Set worksheets
    Set wsAll = ThisWorkbook.Sheets("ALL")
    Set wsByClients = ThisWorkbook.Sheets("By Clients")
    Set wsByCurrencyPair = ThisWorkbook.Sheets("By Currency Pair")
    Set wsByTradeType = ThisWorkbook.Sheets("By Trade Type")

    ' Initialize dictionaries to track unique values
    Set currencyPairs = CreateObject("Scripting.Dictionary")
    Set tradeTypes = CreateObject("Scripting.Dictionary")

    ' Find the last row in ALL sheet
    lastRow = wsAll.Cells(wsAll.Rows.Count, "A").End(xlUp).Row

    ' Copy rows ending with "Total" to "By Clients"
    targetRow = 2
    For i = 1 To lastRow
        If Right(wsAll.Cells(i, 1).Value, 5) = "Total" Then
            wsAll.Rows(i).Copy wsByClients.Cells(targetRow, 2)
            targetRow = targetRow + 1
        End If
    Next i

    ' Sort the data in "By Clients"
    With wsByClients
        .Range("B2", .Cells(targetRow - 1, 256)).Sort Key1:=.Range("B2"), Order1:=xlAscending, Header:=xlNo
        ' Remove blank rows between last filled row and row 100
        .Range("B" & targetRow & ":ER100").ClearContents
    End With

    ' Extract unique currency pairs in column D to "By Currency Pair"
    targetRow = 2
    For i = 1 To lastRow
        currentValue = wsAll.Cells(i, 4).Value
        If Not currencyPairs.Exists(currentValue) And currentValue <> "" Then
            currencyPairs.Add currentValue, True
            wsByCurrencyPair.Cells(targetRow, 1).Value = currentValue
            targetRow = targetRow + 1
        End If
    Next i

    ' Extract unique trade types in column B ending with "Total" to "By Trade Type"
    targetRow = 2
    For i = 1 To lastRow
        currentValue = wsAll.Cells(i, 2).Value
        If Right(currentValue, 5) = "Total" And Not tradeTypes.Exists(currentValue) Then
            tradeTypes.Add currentValue, True
            wsByTradeType.Cells(targetRow, 1).Value = currentValue
            targetRow = targetRow + 1
        End If
    Next i

    MsgBox "Data organized successfully!"
End Sub






Sub OrganizeDataByCriteria()
    Dim wsAll As Worksheet, wsByClients As Worksheet, wsByCurrencyPair As Worksheet, wsByTradeType As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim currencyPairs As Object, tradeTypes As Object
    Dim i As Long, currentValue As String

    ' Set worksheets
    Set wsAll = ThisWorkbook.Sheets("ALL")
    Set wsByClients = ThisWorkbook.Sheets("By Clients")
    Set wsByCurrencyPair = ThisWorkbook.Sheets("By Currency Pair")
    Set wsByTradeType = ThisWorkbook.Sheets("By Trade Type")

    ' Initialize dictionaries to track unique values
    Set currencyPairs = CreateObject("Scripting.Dictionary")
    Set tradeTypes = CreateObject("Scripting.Dictionary")

    ' Find the last row in ALL sheet
    lastRow = wsAll.Cells(wsAll.Rows.Count, "A").End(xlUp).Row

    ' Copy values ending with "Total" in column A to "By Clients" sheet starting from B2
    targetRow = 2
    For i = 1 To lastRow
        If Right(wsAll.Cells(i, 1).Value, 5) = "Total" Then
            wsByClients.Cells(targetRow, 2).Value = wsAll.Cells(i, 1).Value
            targetRow = targetRow + 1
        End If
    Next i

    ' Sort the data in "By Clients"
    With wsByClients
        .Range("B2", .Cells(targetRow - 1, 2)).Sort Key1:=.Range("B2"), Order1:=xlAscending, Header:=xlNo
        ' Remove blank rows between last filled row and row 100
        .Range("B" & targetRow & ":B100").ClearContents
    End With

    ' Extract unique currency pairs in column D to "By Currency Pair"
    targetRow = 2
    For i = 1 To lastRow
        currentValue = wsAll.Cells(i, 4).Value
        If Not currencyPairs.Exists(currentValue) And currentValue <> "" Then
            currencyPairs.Add currentValue, True
            wsByCurrencyPair.Cells(targetRow, 1).Value = currentValue
            targetRow = targetRow + 1
        End If
    Next i

    ' Extract unique trade types in column B ending with "Total" to "By Trade Type"
    targetRow = 2
    For i = 1 To lastRow
        currentValue = wsAll.Cells(i, 2).Value
        If Right(currentValue, 5) = "Total" And Not tradeTypes.Exists(currentValue) Then
            tradeTypes.Add currentValue, True
            wsByTradeType.Cells(targetRow, 1).Value = currentValue
            targetRow = targetRow + 1
        End If
    Next i

    MsgBox "Data organized successfully!"
End Sub



=INDEX(ALL!D3:D10000, MATCH(LARGE(ALL!E3:E10000, 1), ALL!E3:E10000, 0))


