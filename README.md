Sub CalculateAndTransferToWord()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim bookmarkName As String
    Dim rate As Double
    Dim years As Double
    Dim principal As Double
    Dim resultValue As Double
    Dim formattedResult As String

    ' Define the bookmark name in the Word document
    bookmarkName = "YourBookmarkName"

    ' Extract the numeric part of the percentage from C23
    rate = Val(Replace(templateWs.Range("C23").Value, "%", "")) / 100
    
    ' Extract the numeric part of the years from C16
    years = Val(templateWs.Range("C16").Value)
    
    ' Get the value from C31 (the principal amount)
    principal = templateWs.Range("C31").Value
    
    ' Perform the calculation: years * rate * principal
    resultValue = years * rate * principal
    
    ' Format the result as a number with two decimals
    formattedResult = Format(resultValue, "#,##0.00")
    
    ' Assuming you have already set up Word application and document
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Open("C:\path\to\your\word\document.docx")
    
    ' Insert the formatted result into the Word document at the bookmark
    If wordDoc.Bookmarks.Exists(bookmarkName) Then
        wordDoc.Bookmarks(bookmarkName).Range.Text = formattedResult
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
