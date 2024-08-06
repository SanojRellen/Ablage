Sub CopyMostRecentFile()
    Dim folderPath As String
    Dim recentFile As String
    Dim File As String
    Dim fileDate As Date
    Dim recentDate As Date
    Dim ws As Worksheet
    Dim templateWs As Worksheet
    Dim recentWb As Workbook

    ' Set folder path
    folderPath = "C:\Your\Folder\Path\"

    ' Find the most recent Excel file in the folder
    recentFile = ""
    recentDate = DateSerial(1900, 1, 1)
    File = Dir(folderPath & "*.xls*")
    Do While File <> ""
        fileDate = FileDateTime(folderPath & File)
        If fileDate > recentDate Then
            recentDate = fileDate
            recentFile = folderPath & File
        End If
        File = Dir
    Loop

    ' Exit if no file is found
    If recentFile = "" Then
        MsgBox "No Excel files found in the specified folder."
        Exit Sub
    End If

    ' Open the most recent file and set reference
    Set recentWb = Workbooks.Open(recentFile)
    Set ws = recentWb.Sheets("Sheet1")
    Set templateWs = ThisWorkbook.Sheets("Template")
    
    ' Copy the contents
    ws.Cells.Copy Destination:=templateWs.Cells

    ' Close the recent workbook without saving
    recentWb.Close SaveChanges:=False
End Sub
