Sub SendEmail()
    Dim OutlookApp As Object
    Dim OutlookMail As Object

    ' Create Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")
    ' Create a new email item
    Set OutlookMail = OutlookApp.CreateItem(0)

    With OutlookMail
        .To = "recipient@example.com"  ' Replace with the recipient's email address
        .Subject = "Your Subject Here"
        .Body = "This is the body of the email. Customize this text as needed."
        .Send  ' Sends the email immediately
    End With

    ' Clean up
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub
