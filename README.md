
Sub SendEmailWithFormattedTable()

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
    
    ' Convert the table range to HTML format directly with formatting
    TableContent = ConvertRangeToFormattedHTML(Rng)
    
    ' Construct the body of the email
    EmailBody = "Hey, here are the finding levels:" & "<br><br>"
    EmailBody = EmailBody & TableContent & "<br><br>"
    ' Remove the abstract row if unnecessary
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
        .Subject = "Findings Levels"
        .HTMLBody = EmailBody
        .Display ' To display the email before sending (for review)
        ' .Send ' Uncomment this to send the email directly without reviewing
    End With
    
    ' Clean up
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

End Sub

' Function to convert a range to an HTML table string with formatting
Function ConvertRangeToFormattedHTML(Rng As Range) As String
    Dim Cell As Range
    Dim TableHTML As String
    Dim Row As Range
    Dim TempVal As String
    Dim CellColor As String
    Dim FontColor As String
    Dim BoldTag As String
    Dim ItalicTag As String
    
    ' Start the HTML table
    TableHTML = "<table border=""1"" cellpadding=""5"" cellspacing=""0"">"
    
    ' Loop through each row in the range
    For Each Row In Rng.Rows
        TableHTML = TableHTML & "<tr>"
        ' Loop through each cell in the row
        For Each Cell In Row.Cells
            TempVal = Cell.Value
            
            ' Get font color in RGB and convert to HEX
            FontColor = RGBToHex(Cell.Font.Color)
            
            ' Get fill (background) color in RGB and convert to HEX
            CellColor = RGBToHex(Cell.Interior.Color)
            
            ' Handle bold and italic font styles
            If Cell.Font.Bold Then
                BoldTag = "<b>"
            Else
                BoldTag = ""
            End If
            
            If Cell.Font.Italic Then
                ItalicTag = "<i>"
            Else
                ItalicTag = ""
            End If
            
            ' Format numeric values
            If IsNumeric(TempVal) Then
                TempVal = Format(TempVal, "0.00")
            End If
            
            ' Add cell with font and fill color
            TableHTML = TableHTML & "<td style=""background-color:" & CellColor & "; color:" & FontColor & """>"
            TableHTML = TableHTML & BoldTag & ItalicTag & TempVal & "</b></i></td>"
        Next Cell
        TableHTML = TableHTML & "</tr>"
    Next Row
    
    ' Close the HTML table
    TableHTML = TableHTML & "</table>"
    
    ' Return the HTML string
    ConvertRangeToFormattedHTML = TableHTML
End Function

' Function to convert RGB to HEX for HTML color codes
Function RGBToHex(RGBColor As Long) As String
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    
    Red = RGBColor Mod 256
    Green = (RGBColor \ 256) Mod 256
    Blue = (RGBColor \ 65536) Mod 256
    
    RGBToHex = "#" & Right("0" & Hex(Blue), 2) & Right("0" & Hex(Green), 2) & Right("0" & Hex(Red), 2)
End Function
