Sub GetClientEmails()
    Dim OutlookApp As Object
    Dim OutlookNamespace As Object
    Dim OutlookFolder As Object
    Dim OutlookMailItem As Object
    Dim EmailItem As Object
    Dim EmailBody As String
    Dim i As Integer
    Dim EmailAddress As String
    Dim LastRow As Long
    
    ' Get the email address from the active Excel cell
    EmailAddress = ActiveCell.Value
    
    ' Initialize Outlook application
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    ' Access the default inbox
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set OutlookFolder = OutlookNamespace.GetDefaultFolder(6) ' 6 refers to the Inbox
    
    ' Loop through emails in the inbox
    i = 1
    For Each EmailItem In OutlookFolder.Items
        If TypeName(EmailItem) = "MailItem" Then
            If InStr(EmailItem.SenderEmailAddress, EmailAddress) > 0 Then
                ' Found an email from the specified client
                LastRow = ThisWorkbook.Sheets("Client-Fund Relationships").Cells(Rows.Count, 1).End(xlUp).Row + 1
                
                ' Log the communication in Excel
                With ThisWorkbook.Sheets("Client-Fund Relationships")
                    .Cells(LastRow, 1).Value = ActiveCell.Offset(0, -1).Value ' Client ID
                    .Cells(LastRow, 2).Value = "Related Fund ID" ' You can customize this
                    .Cells(LastRow, 3).Value = EmailItem.ReceivedTime
                    .Cells(LastRow, 4).Value = EmailItem.Subject
                    .Cells(LastRow, 5).Value = "Email Content: " & Left(EmailItem.Body, 100) ' Short summary
                End With
            End If
        End If
    Next EmailItem
End Sub

