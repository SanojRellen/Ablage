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
    Dim docPath As String
    Dim outApp As Object
    Dim outMail As Object
    Dim emailAddress As String
    Dim emailBody As String
    Dim currentDate As String
    Dim percentageValue As String
    Dim monthYear As String
    Dim dateList As Variant
    Dim i As Integer

    ' Set folder path and email address
    folderPath = "C:\Your\Folder\Path\"
    emailAddress = "recipient@example.com"
    currentDate = Format(Date, "dd/mm/yyyy")

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
    If cellValue = "SX5E Index" Then
        templateWs.Range("B5").Value = "EUROSTOXX 50"
    End If

    ' Prepare additional values
    percentageValue = Format(templateWs.Range("B23").Value, "0.00%")
    monthYear = Format(templateWs.Range("B13").Value, "mm/yyyy")

    ' Remove all spaces from B14 and split dates into an array
    cellValue = Replace(templateWs.Range("B14").Value, " ", "")
    dateList = Split(cellValue, "/")

    ' Copy dates into cells C14, D14, E14
    templateWs.Range("C14").Value = dateList(0)
    templateWs.Range("D14").Value = dateList(1)
    templateWs.Range("E14").Value = dateList(2)

    ' Open Word and insert values into bookmarks
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open("C:\Your\Word\Document\Path\YourDocument.docx")

    With wdDoc
        .Bookmarks("First").Range.Text = templateWs.Range("B5").Value
        .Bookmarks("Second").Range.Text = templateWs.Range("B6").Value
        .Bookmarks("Third").Range.Text = currentDate
        .Bookmarks("Fourth").Range.Text = percentageValue
        .Bookmarks("Fifth").Range.Text = monthYear

        ' Insert dates into bookmarks Date1, Date2, Date3
        .Bookmarks("Date1").Range.Text = templateWs.Range("C14").Value
        .Bookmarks("Date2").Range.Text = templateWs.Range("D14").Value
        .Bookmarks("Date3").Range.Text = templateWs.Range("E14").Value

        ' Construct Word document path
        docPath = "C:\Your\Word\Path\" & Format(Date, "yyyy-mm-dd") & "_" & templateWs.Range("B1").Value & "_" & templateWs.Range("B5").Value & ".docx"
        
        ' Print Word document path to Immediate Window
        Debug.Print docPath

        ' Save as Word document
        .SaveAs2 docPath, 16 ' 16 represents the wdFormatDocumentDefault constant
    End With

    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing

    ' Prepare email body
    emailBody = "Hi Dennis," & vbCrLf & vbCrLf & _
                "anteil das PIB zum " & templateWs.Range("B3").Value & vbCrLf & vbCrLf & _
                "Viele Grüße"

    ' Send email with Word document attachment
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)
    
    With outMail
        .To = emailAddress
        .Subject = "Subject of Email"
        .Body = emailBody
        .Attachments.Add docPath
        .Send
    End With

    Set outMail = Nothing
    Set outApp = Nothing

    MsgBox "Process Completed Successfully"
End Sub

 
