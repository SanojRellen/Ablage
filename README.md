

Sub SaveRangeAsNewExcelAndAttachToMail()

    Dim ws As Worksheet
    Dim newWorkbook As Workbook
    Dim newSheet As Worksheet
    Dim mypath As String
    Dim filename As String
    Dim dateStr As String
    Dim rngToCopy As Range
    Dim currentWorkbook As Workbook
    
    ' Define your worksheet
    Set currentWorkbook = ThisWorkbook
    Set ws = currentWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name

    ' Define the path and file name
    mypath = "C:\Your\Path\Here\" ' Change this to your desired path
    dateStr = Format(Date, "dd_mm_yyyy")
    filename = mypath & dateStr & "_Funding_Levels.xlsx"
    
    ' Define the range to copy
    Set rngToCopy = ws.Range("R7:T18") ' Adjust range as needed

    ' Create a new workbook and copy the range
    Set newWorkbook = Workbooks.Add
    Set newSheet = newWorkbook.Sheets(1)
    
    ' Paste the range into the new workbook
    rngToCopy.Copy
    newSheet.Range("A1").PasteSpecial Paste:=xlPasteAll ' Paste with formatting
    
    ' Save the new workbook
    Application.DisplayAlerts = False ' Suppress overwrite prompt
    newWorkbook.SaveAs filename
    Application.DisplayAlerts = True
    
    ' Close the new workbook
    newWorkbook.Close
    
    ' Call the function to send an email with the attachment
    Call SendEmailWithAttachment(filename)
    
End Sub






Sub SendEmailWithAttachment(filename As String)

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim EmailBody As String
    
    ' Create Outlook object
    On Error Resume Next
    Set OutlookApp = GetObject(class:="Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject(class:="Outlook.Application")
    End If
    On Error GoTo 0
    
    ' Create an Outlook mail item
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Construct the email body (example text)
    EmailBody = "Hey, here are the finding levels. Please see the attached file." & "<br><br>" & "Best regards,<br>Your Name"
    
    ' Setup the email details
    With OutlookMail
        .To = "recipient@example.com" ' Add recipient email address
        .Subject = "Funding Levels " & Format(Date, "dd_mm_yyyy")
        .HTMLBody = EmailBody
        .Attachments.Add filename ' Attach the saved file
        .Display ' To display the email before sending (for review)
        ' .Send ' Uncomment this line if you want to send directly without displaying
    End With
    
    ' Clean up
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

End Sub

