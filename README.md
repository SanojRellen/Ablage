Sub InsertLastBarrierObservationDate()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim sourceItemNodes As IXMLDOMNodeList
    Dim lastBarrierObservationDateNode As IXMLDOMNode
    Dim templateNode As IXMLDOMNode
    Dim lastObservationDate As String
    Dim lastSourceItem As IXMLDOMNode
    
    ' Load the XML documents
    Set xmlDocSource = New MSXML2.DOMDocument60
    Set xmlDocTemplate = New MSXML2.DOMDocument60
    
    If Not xmlDocSource.Load("C:\path\to\your\source.xml") Then ' Update with your source XML file path
        MsgBox "Failed to load source XML file.", vbExclamation
        Exit Sub
    End If
    
    If Not xmlDocTemplate.Load("C:\path\to\your\template.xml") Then ' Update with your template XML file path
        MsgBox "Failed to load template XML file.", vbExclamation
        Exit Sub
    End If
    
    ' Get the list of 'item' nodes from the 'schedule' node in the source XML
    Set sourceItemNodes = xmlDocSource.SelectNodes("//schedule/item")
    
    If sourceItemNodes Is Nothing Or sourceItemNodes.Length = 0 Then
        MsgBox "No 'item' nodes found in source XML.", vbExclamation
        Exit Sub
    End If
    
    ' Get the last 'item' node in the source
    Set lastSourceItem = sourceItemNodes.Item(sourceItemNodes.Length - 1)
    
    ' Get the 'item' node under 'barrierEventObservationDates' in the last 'item' block
    Set lastBarrierObservationDateNode = lastSourceItem.SelectSingleNode("barrierEventObservationDates/item")
    
    If lastBarrierObservationDateNode Is Nothing Then
        MsgBox "No 'item' node found under 'barrierEventObservationDates' in the last block.", vbExclamation
        Exit Sub
    End If
    
    ' Extract the text value of the last 'item' node, which is the date we want
    lastObservationDate = lastBarrierObservationDateNode.Text
    
    ' Find the target node in the template XML where we want to insert the date
    Set templateNode = xmlDocTemplate.SelectSingleNode("//zinsfestlegungstage/LastObservationDate") ' Update with your actual target node
    
    If templateNode Is Nothing Then
        MsgBox "'LastObservationDate' node not found in template XML.", vbExclamation
        Exit Sub
    End If
    
    ' Fill the target node in the template with the last observation date from the source
    templateNode.Text = lastObservationDate
    
    ' Save the modified template XML
    xmlDocTemplate.Save "C:\path\to\your\filled_template.xml" ' Update with your output file path
    
    MsgBox "Last observation date has been inserted into the template successfully!", vbInformation
End Sub

