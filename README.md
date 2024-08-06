Sub CheckCellAndUpdateWord()
    Dim ws As Worksheet
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim bookmarkName As String
    Dim table As Object
    Dim rowToDelete As Object
    Dim cellValue As Variant
    Dim i As Integer

    ' Set your worksheet and bookmark name
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your actual sheet name
    bookmarkName = "YourBookmarkName" ' Change to your actual bookmark name

    ' Get the value from B34
    cellValue = ws.Range("B34").Value

    ' Initialize Word application
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open("C:\Your\Word\Document\Path\YourDocument.docx") ' Update the path

    ' Check if the value in B34 is 1
    If cellValue = 1 Then
        ' Find the table containing the bookmark and delete the row
        For Each table In wdDoc.Tables
            For i = 1 To table.Rows.Count
                If table.Cell(i, 1).Range.Bookmarks.Exists(bookmarkName) Then
                    ' Set reference to the row containing the bookmark
                    Set rowToDelete = table.Rows(i)
                    ' Delete the row
                    rowToDelete.Delete
                    Exit For
                End If
            Next i
            If Not rowToDelete Is Nothing Then Exit For
        Next table
    Else
        ' Copy the value of B34 into the specified bookmark
        wdDoc.Bookmarks(bookmarkName).Range.Text = cellValue
    End If

    ' Save changes and close Word
    wdDoc.Save
    wdDoc.Close SaveChanges:=True
    wdApp.Quit

    ' Clean up
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
