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











Sub InsertBulletPointsFromExcel()
    ' Define variables
    Dim ExcelApp As Object
    Dim Workbook As Object
    Dim CouponRate As String
    Dim StartDate As Date
    Dim EndDate As Date
    Dim NumberOfPeriods As Integer
    Dim i As Integer
    Dim CurrentCouponDate As Date
    Dim BookmarkRange As Range

    ' Initialize Excel and open the workbook
    Set ExcelApp = CreateObject("Excel.Application")
    Set Workbook = ExcelApp.Workbooks.Open("Path to your Excel file")

    ' Read values from Excel
    CouponRate = Workbook.Sheets("Sheet1").Range("D23").Value
    StartDate = Workbook.Sheets("Sheet1").Range("C15").Value
    EndDate = Workbook.Sheets("Sheet1").Range("C14").Value

    ' Calculate the number of periods (years) between the end date and start date
    NumberOfPeriods = Year(EndDate) - Year(StartDate)

    ' Get the range of the bookmark
    Set BookmarkRange = ActiveDocument.Bookmarks("Bookmark1").Range

    ' Clear any existing text in the bookmark
    BookmarkRange.Text = ""

    ' Start with the last coupon date (End Date)
    CurrentCouponDate = EndDate

    ' Loop through each period and insert bullet points
    For i = 0 To NumberOfPeriods - 1
        ' Insert the bullet point with the coupon rate and the current coupon date
        BookmarkRange.InsertAfter (i + 1) & ") " & CouponRate & " " & Format(CurrentCouponDate, "dd.mm.yyyy") & vbCrLf

        ' Move to the previous year's coupon date
        CurrentCouponDate = DateAdd("yyyy", -1, CurrentCouponDate)
    Next i

    ' Re-apply the bookmark to the updated range
    ActiveDocument.Bookmarks.Add "Bookmark1", BookmarkRange

    ' Clean up
    Workbook.Close False
    ExcelApp.Quit
    Set ExcelApp = Nothing
    Set Workbook = Nothing
End Sub

