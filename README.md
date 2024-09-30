Sub SeparateAndCopyValues()
    Dim inputString As String
    Dim values() As String
    Dim i As Integer
    
    ' Get the input from cell C15
    inputString = Range("C15").Value
    
    ' Split the string by slashes
    values = Split(inputString, "/")
    
    ' Loop through the array and copy values to cells F15 to K15
    For i = LBound(values) To UBound(values)
        ' Check if the index is within the range of F15 to K15 (0 to 5)
        If i < 6 Then
            Range("F" & 15).Offset(0, i).Value = Trim(values(i))
        End If
    Next i
End Sub







Dim cell As Range
For Each cell In Range("C20:H20")
    If Not IsEmpty(cell.Value) Then
        ' Your operation here
    End If
Next cell








Sub FillWordBookmark()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim bookmarkName As String
    Dim cellValue As String
    Dim fixedText As String
    
    ' Initialize variables
    bookmarkName = "YourBookmarkName" ' Replace with the name of your bookmark
    fixedText = "Peter"
    cellValue = ThisWorkbook.Sheets("Sheet1").Range("A3").Value
    
    ' Start Word and open the document
    On Error Resume Next
    Set wdApp = GetObject(Class:="Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject(Class:="Word.Application")
    End If
    On Error GoTo 0
    
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open("C:\Path\To\Your\Document.docx") ' Update the path to your document
    
    ' Check if bookmark exists and fill it
    If wdDoc.Bookmarks.Exists(bookmarkName) Then
        wdDoc.Bookmarks(bookmarkName).Range.Text = cellValue & " " & fixedText
    Else
        MsgBox "Bookmark not found!", vbExclamation
    End If

    ' Optional: Save and close the document
    ' wdDoc.Save
    ' wdDoc.Close
    ' wdApp.Quit

    ' Clean up
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub

