# Ablage
Zum abgreifen


Sub CopySheetAndModify()
    Dim folderPath As String
    Dim recentFile As String
    Dim File As String
    Dim fileDate As Date
    Dim recentDate As Date
    Dim ws As Worksheet
    Dim templateWs As Worksheet
    Dim cellValue As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim pdfPath As String
    Dim outApp As Object
    Dim outMail As Object
    Dim emailAddress As String
    Dim emailBody As String
    Dim pdfName As String

    ' Set folder path and email address
    folderPath = "C:\Your\Folder\Path\"
    emailAddress = "recipient@example.com"

    ' Find the most recent Excel file in the folder
    recentFile = ""
    recentDate = DateSerial(1900, 1, 1)
    File = Dir(folderPath & "*.xls*")
    Do While File <> ""
        fileDate = FileDateTime(folderPath & File)
        If fileDate > recentDate Then
            recentDate = fileDate
            recentFile = folderPath & File
        End If
        File = Dir
    Loop

    ' Exit if no file is found
    If recentFile = "" Then
        MsgBox "No Excel files found in the specified folder."
        Exit Sub
    End If

    ' Copy Sheet1 from the most recent file to the Template sheet in the current workbook
    Set ws = Workbooks.Open(recentFile).Sheets("Sheet1")
    Set templateWs = ThisWorkbook.Sheets("Template")
    ws.Cells.Copy Destination:=templateWs.Cells
    Workbooks(recentFile).Close SaveChanges:=False

    ' Check cell B5 and replace if necessary
    cellValue = templateWs.Range("B5").Value
    If cellValue = "SX5R Index" Then
        templateWs.Range("B5").Value = "EUROSTOXX 50"
    End If

    ' Open Word and insert values into bookmarks
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open("C:\Your\Word\Document\Path\YourDocument.docx")

    With wdDoc
        .Bookmarks("first").Range.Text = templateWs.Range("B5").Value
        .Bookmarks("second").Range.Text = templateWs.Range("B6").Value

        ' Construct PDF name
        pdfName = templateWs.Range("B6").Value & "_" & templateWs.Range("B1").Value & "_" & templateWs.Range("B5").Value & ".pdf"
        pdfPath = "C:\Your\PDF\Path\" & pdfName
        
        ' Print PDF path to Immediate Window
        Debug.Print pdfPath

        ' Save as PDF
        .SaveAs2 pdfPath, 17 ' 17 represents the wdFormatPDF constant
    End With

    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing

    ' Prepare email body
    emailBody = "Hi Dennis," & vbCrLf & vbCrLf & _
                "anteil das PIB zum " & templateWs.Range("B3").Value & vbCrLf & vbCrLf & _
                "Viele Grüße"

    ' Send email with PDF attachment
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)
    
    With outMail
        .To = emailAddress
        .Subject = "Subject of Email"
        .Body = emailBody
        .Attachments.Add pdfPath
        .Send
    End With

    Set outMail = Nothing
    Set outApp = Nothing

    MsgBox "Process Completed Successfully"
End Sub

