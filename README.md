
Sub CheckValuesCalculateDifferenceAndInsertText()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim bookmarkName As String
    Dim c31Value As Double
    Dim c32Value As Double
    Dim difference As Double
    Dim formattedDifference As String
    Dim textToInsert As String

    ' Define the bookmark name in the Word document
    bookmarkName = "YourBookmarkName"

    ' Get the values from C31 and C32
    c31Value = templateWs.Range("C31").Value
    c32Value = templateWs.Range("C32").Value

    ' Determine the text to insert based on the condition
    If c32Value > c31Value Then
        ' Calculate the difference
        difference = c32Value - c31Value
        
        ' Format the difference as a number with two decimals
        formattedDifference = Format(difference, "0.00")
        
        ' Construct the text to insert
        textToInsert = "ein in Höhe von EUR " & formattedDifference
    Else
        ' If C32 is not greater than C31, insert "kein"
        textToInsert = "kein"
    End If

    ' Assuming you have already set up Word application and document
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Open("C:\path\to\your\word\document.docx")
    
    ' Insert the text into the Word document at the bookmark
    If wordDoc.Bookmarks.Exists(bookmarkName) Then
        wordDoc.Bookmarks(bookmarkName).Range.Text = textToInsert
    Else
        MsgBox "Bookmark not found in Word document.", vbExclamation
    End If
    
    ' Save and close the Word document
    wordDoc.Save
    wordDoc.Close
    wordApp.Quit
    
    ' Clean up
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub
