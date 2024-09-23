Dim wordApp As Object
Dim wordDoc As Object
Dim bookmarkName As String
Dim formattedDate As String

' Set Word application and document
Set wordApp = CreateObject("Word.Application")
Set wordDoc = wordApp.Documents.Open("C:\Your\Word\Document.docx") ' Replace with your document path

' Get the formatted date from Excel
formattedDate = Format(Range("A1").Value, "dd mmmm yyyy") ' This ensures the date is formatted as 25 March 2024

' Replace the bookmark with the formatted date in Word
bookmarkName = "YourBookmarkName" ' Replace with your actual bookmark name
If wordDoc.Bookmarks.Exists(bookmarkName) Then
    wordDoc.Bookmarks(bookmarkName).Range.Text = formattedDate
End If

' Save and close the Word document
wordDoc.Save
wordDoc.Close
wordApp.Quit

Set wordApp = Nothing
Set wordDoc = Nothing
