Sub AdjustAndFillTemplate()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim sourceItemNodes As IXMLDOMNodeList
    Dim templateDatumNodes As IXMLDOMNodeList
    Dim zinsfestlegungstageNode As IXMLDOMElement
    Dim i As Integer
    Dim sourceItemCount As Integer
    Dim templateDatumCount As Integer
    Dim newDatumNode As IXMLDOMElement
    
    ' Load the XML documents
    Set xmlDocSource = New MSXML2.DOMDocument60
    Set xmlDocTemplate = New MSXML2.DOMDocument60
    
    xmlDocSource.Load "C:\path\to\your\source.xml" ' Update with your source XML file path
    xmlDocTemplate.Load "C:\path\to\your\template.xml" ' Update with your template XML file path
    
    ' Get the list of item nodes from the source XML
    Set sourceItemNodes = xmlDocSource.SelectNodes("//couponEvents/schedule/item")
    sourceItemCount = sourceItemNodes.Length
    
    ' Get the list of Datum nodes from the template XML
    Set templateDatumNodes = xmlDocTemplate.SelectNodes("//zinsfestlegungstage/Datum")
    templateDatumCount = templateDatumNodes.Length
    
    ' Get the zinsfestlegungstage node in the template to add/remove Datum nodes
    Set zinsfestlegungstageNode = xmlDocTemplate.SelectSingleNode("//zinsfestlegungstage")
    
    ' Adjust the number of Datum nodes in the template to match the number of item nodes in the source
    If sourceItemCount > templateDatumCount Then
        ' Add additional Datum nodes
        For i = templateDatumCount + 1 To sourceItemCount
            Set newDatumNode = xmlDocTemplate.createElement("Datum")
            zinsfestlegungstageNode.appendChild newDatumNode
        Next i
    ElseIf sourceItemCount < templateDatumCount Then
        ' Remove extra Datum nodes
        For i = templateDatumCount To sourceItemCount + 1 Step -1
            zinsfestlegungstageNode.removeChild templateDatumNodes.Item(i - 1)
        Next i
    End If
    
    ' Now, re-select the Datum nodes since they may have been added/removed
    Set templateDatumNodes = xmlDocTemplate.SelectNodes("//zinsfestlegungstage/Datum")
    
    ' Fill the Datum nodes with the corresponding dates from the source
    For i = 0 To sourceItemCount - 1
        templateDatumNodes.Item(i).Text = sourceItemNodes.Item(i).SelectSingleNode("couponDate").Text
    Next i
    
    ' Save the modified template XML
    xmlDocTemplate.Save "C:\path\to\your\filled_template.xml" ' Update with your output file path
    
    MsgBox "Template has been adjusted and filled successfully!", vbInformation
End Sub









Sub CalculateAndInsertInterest()
    Dim principal As Double
    Dim rate As Double
    Dim startDate As Date
    Dim endDate As Date
    Dim duration1 As Double
    Dim duration2 As Double
    Dim interest1 As Double
    Dim interest2 As Double
    Dim totalInterest As Double
    
    ' Get values from the Excel sheet
    principal = Range("C14").Value ' Assuming principal is in C14
    rate = Range("C16").Value ' Assuming rate is in C16
    startDate = Range("C15").Value
    endDate = Range("C17").Value
    
    ' Calculate duration in 30/360 convention
    Dim year1 As Double, year2 As Double
    
    ' Calculate full years first
    year1 = 1 ' First full year is always 360/360
    year2 = 359 / 360 ' Second period has 359 days (missing one day)

    ' Calculate interest
    interest1 = principal * rate * year1
    interest2 = principal * rate * year2
    totalInterest = interest1 + interest2

    ' Open Word and insert the result in the bookmark
    Dim wdApp As Object
    Dim wdDoc As Object
    
    ' Open Word application
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True ' Set to False if you don't want Word to be visible
    
    ' Open the Word document
    Set wdDoc = wdApp.Documents.Open("C:\path\to\your\document.docx")
    
    ' Insert totalInterest into the bookmark
    With wdDoc.Bookmarks("YourBookmarkName").Range
        .Text = Format(totalInterest, "0.00") ' Format to 2 decimal places
    End With
    
    ' Save and close the Word document
    wdDoc.Save
    wdDoc.Close
    
    ' Clean up
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
