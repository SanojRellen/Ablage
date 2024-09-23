Sub DeleteEmptyRowsInWordTable()

    Dim wordApp As Object
    Dim wordDoc As Object
    Dim table As Object
    Dim i As Long
    Dim bookmarkName As String
    Dim rowHasContent As Boolean
    
    ' Set Word application and document
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Open("C:\Your\Word\Document.docx") ' Replace with your document path

    ' Assume the table is the first table in the document (adjust if necessary)
    Set table = wordDoc.Tables(1)
    
    ' Loop through each row in the table (assuming 2 columns, 6 rows)
    For i = 6 To 1 Step -1 ' Loop backwards to avoid issues when deleting rows
        rowHasContent = False
        
        ' Replace with your bookmark naming pattern
        bookmarkName = "Bookmark" & i ' Assuming your bookmarks are named Bookmark1, Bookmark2, etc.
        
        ' Check if the bookmark exists and contains text
        If wordDoc.Bookmarks.Exists(bookmarkName) Then
            If Trim(wordDoc.Bookmarks(bookmarkName).Range.Text) <> "" Then
                rowHasContent = True
            End If
        End If
        
        ' If the row is empty (bookmark has no content), delete the row
        If Not rowHasContent Then
            table.Rows(i).Delete
        End If
    Next i
    
    ' Save and close the Word document
    wordDoc.Save
    wordDoc.Close
    wordApp.Quit

    ' Cleanup
    Set wordApp = Nothing
    Set wordDoc = Nothing
    Set table = Nothing

End Sub




Dim example As String
Dim example2 As String

' Example string containing the date
example = "27 March 2024"

' Remove the last 5 characters (the year and the space before it)
example2 = Left(example, Len(example) - 5)

' Output the result
MsgBox example2 ' This will display "27 March"

