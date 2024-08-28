Sub InsertWKNandISIN()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim sourceISINNode As IXMLDOMNode
    Dim templateWKNNode As IXMLDOMNode
    Dim templateISINNode As IXMLDOMNode
    Dim isinValue As String
    Dim wknValue As String
    Dim idNodes As IXMLDOMNodeList
    
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
    
    ' Find the ISIN node in the source XML
    Set sourceISINNode = xmlDocSource.SelectSingleNode("//underlyings/item/isin")
    
    If sourceISINNode Is Nothing Then
        MsgBox "'isin' node not found in source XML.", vbExclamation
        Exit Sub
    End If
    
    ' Get the ISIN value
    isinValue = sourceISINNode.Text
    
    ' Derive the WKN from the ISIN (assuming WKN is characters 3-8)
    wknValue = Mid(isinValue, 3, 6)
    
    ' Get all 'Id' nodes under 'Underlyings/ids' in the template XML
    Set idNodes = xmlDocTemplate.SelectNodes("//Underlyings/ids/Id")
    
    If idNodes Is Nothing Or idNodes.Length < 2 Then
        MsgBox "'Id' nodes not found or not enough 'Id' nodes in template XML.", vbExclamation
        Exit Sub
    End If
    
