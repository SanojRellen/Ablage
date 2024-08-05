# Ablage
Zum abgreifen


Dim fileName As String
Dim parts As Variant
Dim filePath As String

' Set the file path (example path, adjust accordingly)
filePath = "C:\Your\Word\Document\Path\abc_def.docx"

' Extract the file name from the file path (removes the path)
fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

' Remove the file extension
fileName = Left(fileName, InStrRev(fileName, ".") - 1)

' Split the file name into parts
parts = Split(fileName, "_")

' Open Word and insert the parts into bookmarks
Set wdApp = CreateObject("Word.Application")
wdApp.Visible = False
Set wdDoc = wdApp.Documents.Open(filePath)

With wdDoc
    .Bookmarks("FirstPart").Range.Text = parts(0)
    .Bookmarks("SecondPart").Range.Text = parts(1)

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








Dim fullValue As String
Dim datePart As String
Dim timePart As String

' Get the full value from B7
fullValue = templateWs.Range("B7").Value

' Split the value based on the space
datePart = Trim(Left(fullValue, InStr(fullValue, " ") - 1)) & "." & Trim(Mid(fullValue, InStr(fullValue, " ") + 1, 4))
timePart = Trim(Mid(fullValue, InStrRev(fullValue, " ") + 1))

' Update B7 and C7
templateWs.Range("B7").Value = datePart
templateWs.Range("C7").Value = timePart






Dim valueB31 As Double
Dim valueB34 As Double
Dim sumValue As Double

' Remove " EUR" and convert text to numeric values
valueB31 = CDbl(Replace(Replace(templateWs.Range("B31").Value, " EUR", ""), ",", "."))
valueB34 = CDbl(Replace(Replace(templateWs.Range("B34").Value, " EUR", ""), ",", "."))

' Calculate the sum
sumValue = valueB31 + valueB34

' Convert the sum back to text format with comma and " EUR"
templateWs.Range("C31").Value = Replace(Format(sumValue, "0.00"), ".", ",") & " EUR"
