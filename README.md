

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
