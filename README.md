Sub CopyDataFromFiles()
    Dim wbTarget As Workbook
    Dim wbSource As Workbook
    Dim folderPath As String
    Dim fileName As String
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet

    ' Set the folder path containing the source files
    folderPath = "C:\Path\To\Your\Folder\" ' Change to your folder path

    ' Open the target workbook
    Set wbTarget = Workbooks.Open(folderPath & "v2_iBoxx Covered Indices Historycopy.xlsm")

    ' Array of source files and corresponding target sheets
    Dim sourceFiles(1 To 2) As String
    Dim targetSheets(1 To 2) As String

    sourceFiles(1) = "Spreads_Basic_Engine.xlsm"
    sourceFiles(2) = "YTM and Duration basic engine.xlsm"
    
    targetSheets(1) = "Raw_Data"
    targetSheets(2) = "Raw_Yield"

    ' Loop through source files and copy data
    Dim i As Integer
    For i = 1 To 2
        ' Open source workbook
        Set wbSource = Workbooks.Open(folderPath & sourceFiles(i))
        Set wsSource = wbSource.Sheets("Adhoc")

        ' Set target worksheet
        Set wsTarget = wbTarget.Sheets(targetSheets(i))

        ' Copy data
        wsSource.Range("EU12:KL6000").Copy
        wsTarget.Range("A7").PasteSpecial Paste:=xlPasteValues

        ' Close source workbook
        wbSource.Close False
    Next i

    ' Handle Duration data
    Set wbSource = Workbooks.Open(folderPath & "YTM and Duration basic engine.xlsm")
    Set wsSource = wbSource.Sheets("Adhoc")
    Set wsTarget = wbTarget.Sheets("Raw_Duration")

    wsSource.Range("EU12:KL6000").Copy
    wsTarget.Range("A7").PasteSpecial Paste:=xlPasteValues
    wbSource.Close False

    ' Save and close the target workbook
    wbTarget.Save
    wbTarget.Close

    MsgBox "Data copied successfully!"

End Sub
