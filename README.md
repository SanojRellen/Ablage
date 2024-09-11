  
Sub SendEmailWithTable()

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim MailBody As String
    Dim Rng As Range
    Dim Ws As Worksheet
    Dim EmailBody As String
    Dim TableContent As String
    
    ' Define the worksheet and the range for the table
    Set Ws = ThisWorkbook.Sheets("Sheet1") ' Change Sheet1 to your sheet name
    Set Rng = Ws.Range("R6:T14")
    
    ' Convert the table range to HTML format
    TableContent = RangeToHTML(Rng)
    
    ' Construct the body of the email
    EmailBody = "Hey, here are the finding levels:" & "<br><br>"
    EmailBody = EmailBody & TableContent & "<br><br>"
    ' You can remove the abstract section by commenting or deleting this line
    ' EmailBody = EmailBody & "Abstract: [Your abstract here]" & "<br><br>"
    EmailBody = EmailBody & "Please amend as necessary." & "<br><br>"
    EmailBody = EmailBody & "Best regards,<br>Dennis"
    
    ' Create Outlook object
    On Error Resume Next
    Set OutlookApp = GetObject(class:="Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject(class:="Outlook.Application")
    End If
    On Error GoTo 0
    
    ' Create an Outlook mail item
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Setup the email details
    With OutlookMail
        .To = "recipient@example.com" ' Add recipient email address
        .Subject = "Findings Levels and Abstract"
        .HTMLBody = EmailBody
        .Display ' To display the email before sending (for review)
        ' .Send ' Uncomment this to send the email directly without reviewing
    End With
    
    ' Clean up
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

End Sub

' Function to convert range to HTML for pasting into the email
Function RangeToHTML(Rng As Range) As String
    Dim TempWorkbook As Workbook
    Dim TempWorksheet As Worksheet
    Dim FileName As String
    Dim HtmlString As String
    Dim Cell As Range

    ' Copy the range to a new temporary workbook
    Set TempWorkbook = Workbooks.Add(1)
    Set TempWorksheet = TempWorkbook.Sheets(1)
    Rng.Copy
    TempWorksheet.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    TempWorksheet.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats

    ' Save the temporary workbook as an HTML file
    FileName = Environ$("temp") & "\TempFile.htm"
    TempWorksheet.Parent.SaveAs FileName, xlHtml
    
    ' Open the saved HTML file and get the text
    With CreateObject("Scripting.FileSystemObject")
        HtmlString = .GetFile(FileName).OpenAsTextStream(1, -2).ReadAll
    End With
    
    ' Delete the temporary file and workbook
    Kill FileName
    TempWorkbook.Close False
    Set TempWorkbook = Nothing
    
    ' Return the HTML content
    RangeToHTML = HtmlString
End Function
