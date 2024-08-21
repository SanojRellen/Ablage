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













Sub AdjustToLastBusinessDay()
    Dim cell As Range
    Dim checkDate As Date
    Dim dayOfWeek As Integer

    ' Loop through each cell in the range G23:L23
    For Each cell In Range("G23:L23")
        ' Check if the cell is not empty
        If Not IsEmpty(cell.Value) Then
            ' Get the date from the cell
            checkDate = cell.Value
            
            ' Get the day of the week (1 = Sunday, 2 = Monday, ..., 7 = Saturday)
            dayOfWeek = Weekday(checkDate, vbMonday)
            
            ' If the date is Saturday (dayOfWeek = 6) or Sunday (dayOfWeek = 7)
            If dayOfWeek > 5 Then
                ' Adjust to the previous Friday
                cell.Value = checkDate - (dayOfWeek - 5)
            End If
        End If
    Next cell
End Sub
