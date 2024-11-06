Sub CopyAndPasteWithDelete()
    Dim wsCopy As Worksheet
    Dim wsOutput As Worksheet
    
    ' Define worksheets
    Set wsCopy = ThisWorkbook.Sheets("Copy Table from NGT")
    Set wsOutput = ThisWorkbook.Sheets("Output Sheet")
    
    ' Clear all cells in row 7 of Output Sheet except columns O and K
    wsOutput.Rows(7).ClearContents
    wsOutput.Range("O7").Value = wsOutput.Range("O7").Value
    wsOutput.Range("K7").Value = wsOutput.Range("K7").Value
    
    ' Paste specific values from "Copy Table from NGT" to "Output Sheet" row 7
    wsOutput.Range("C7").Value = "FBB"
    wsOutput.Range("E7").Value = wsCopy.Range("F2").Value
    wsOutput.Range("F7").Value = wsCopy.Range("D2").Value
    wsOutput.Range("G7").Value = "EC"
    wsOutput.Range("I7").Value = wsCopy.Range("AB2").Value
    wsOutput.Range("J7").Value = wsCopy.Range("AC2").Value
    wsOutput.Range("L7").Value = "Open"
    wsOutput.Range("M7").Value = "SCHILMIK"
    wsOutput.Range("T7").Value = wsCopy.Range("G2").Value
    wsOutput.Range("V7").Value = wsCopy.Range("I2").Value
    
    ' Delete row 2 in the "Copy Table from NGT" sheet
    wsCopy.Rows(2).Delete
End Sub
