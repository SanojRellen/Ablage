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











Sub CalculateCumulativeCoupon()

    Dim principal As Double
    Dim rate As Double
    Dim startDate As Date
    Dim endDate As Date
    Dim currentYearStart As Date
    Dim currentYearEnd As Date
    Dim fullYears As Integer
    Dim remainingDays As Integer
    Dim i As Integer
    Dim couponPayment As Double
    Dim Kupon_kumuliert_Zwischenergebnis As Double
    
    ' Example values
    principal = 100000 ' Assign the principal amount here
    rate = 0.05 ' Assign the annual interest rate here (e.g., 5% as 0.05)
    startDate = Range("C15").Value ' Start date in cell C15
    endDate = Range("C17").Value ' End date in cell C17
    
    ' Initialize cumulative coupon
    Kupon_kumuliert_Zwischenergebnis = 0
    
    ' Calculate the number of full years
    fullYears = Year(endDate) - Year(startDate)
    
    ' Loop through each full year
    For i = 0 To fullYears - 1
        ' Set the start and end of the current year in the loop
        currentYearStart = DateSerial(Year(startDate) + i, Month(startDate), Day(startDate))
        currentYearEnd = DateSerial(Year(startDate) + i + 1, Month(startDate), Day(startDate)) - 1
        
        ' Check if it's the last year in the range
        If i = fullYears - 1 Then
            ' Adjust the end date for the last period
            currentYearEnd = endDate
        End If
        
        ' Calculate remaining days in the final year, or use 360 for full years
        If currentYearEnd < endDate Then
            remainingDays = 360
        Else
            remainingDays = Day360(currentYearStart, currentYearEnd)
        End If
        
        ' Calculate coupon payment for the current year
        couponPayment = principal * rate * (remainingDays / 360)
        Kupon_kumuliert_Zwischenergebnis = Kupon_kumuliert_Zwischenergebnis + couponPayment
    Next i
    
    ' Output the cumulative coupon payment result
    MsgBox "Kupon_kumuliert_Zwischenergebnis: " & Format(Kupon_kumuliert_Zwischenergebnis, "0.00")
    
End Sub

Function Day360(start_date As Date, end_date As Date) As Integer
    ' Day count convention 30/360 calculation
    Dim d1 As Integer, d2 As Integer, m1 As Integer, m2 As Integer, y1 As Integer, y2 As Integer
    
    d1 = Day(start_date)
    d2 = Day(end_date)
    m1 = Month(start_date)
    m2 = Month(end_date)
    y1 = Year(start_date)
    y2 = Year(end_date)
    
    ' Adjust day for 30/360 convention
    If d1 = 31 Then d1 = 30
    If d2 = 31 And d1 = 30 Then d2 = 30
    
    Day360 = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
End Function


