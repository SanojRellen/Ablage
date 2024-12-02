
Sub ClearData()
    Dim wsAll As Worksheet, wsByClients As Worksheet, wsByCurrencyPair As Worksheet, wsByTradeType As Worksheet
    Dim lastRowAll As Long, lastRowByClients As Long, lastRowByCurrency As Long, lastRowByTradeType As Long
    
    ' Set worksheets
    Set wsAll = ThisWorkbook.Sheets("ALL")
    Set wsByClients = ThisWorkbook.Sheets("By Clients")
    Set wsByCurrencyPair = ThisWorkbook.Sheets("By Currency Pair")
    Set wsByTradeType = ThisWorkbook.Sheets("By Trade Type")
    
    ' Clear range A2:N in "ALL"
    lastRowAll = wsAll.Cells(wsAll.Rows.Count, "N").End(xlUp).Row
    If lastRowAll >= 2 Then
        wsAll.Range("A2:N" & lastRowAll).ClearContents
    End If

    ' Clear column B starting from row 16 in "By Clients"
    lastRowByClients = wsByClients.Cells(wsByClients.Rows.Count, "B").End(xlUp).Row
    If lastRowByClients >= 16 Then
        wsByClients.Range("B16:B" & lastRowByClients).ClearContents
    End If

    ' Clear column B starting from row 15 in "By Currency Pair"
    lastRowByCurrency = wsByCurrencyPair.Cells(wsByCurrencyPair.Rows.Count, "B").End(xlUp).Row
    If lastRowByCurrency >= 15 Then
        wsByCurrencyPair.Range("B15:B" & lastRowByCurrency).ClearContents
    End If

    ' Clear column B starting from row 3 in "By Trade Type"
    lastRowByTradeType = wsByTradeType.Cells(wsByTradeType.Rows.Count, "B").End(xlUp).Row
    If lastRowByTradeType >= 3 Then
        wsByTradeType.Range("B3:B" & lastRowByTradeType).ClearContents
    End If

    MsgBox "Data cleared successfully!"
End Sub


