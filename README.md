Sub PartiallyFillNode()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim sourceNode As IXMLDOMNode
    Dim templateNode As IXMLDOMNode
    Dim extractedDate As String
    
    ' Load the XML documents
    Set xmlDocSource = New MSXML2.DOMDocument60
    Set xmlDocTemplate = New MSXML2.DOMDocument60
    
    xmlDocSource.Load "C:\path\to\your\source.xml" ' Update with your source XML file path
    xmlDocTemplate.Load "C:\path\to\your\template.xml" ' Update with your template XML file path
    
    ' Find the date node in the source XML
    Set sourceNode = xmlDocSource.SelectSingleNode("//Date")
    
    If Not sourceNode Is Nothing Then
        ' Extract the date value
        extractedDate = sourceNode.Text
        
        ' Find the target node in the template XML
        Set templateNode = xmlDocTemplate.SelectSingleNode("//TargetNode") ' Update TargetNode with the actual node name
        
        If Not templateNode Is Nothing Then
            ' Construct the desired text
            templateNode.Text = "EXPR.ZT SD A " & extractedDate & " SX5E"
        Else
            MsgBox "Target node not found in template XML.", vbExclamation
        End If
    Else
        MsgBox "Date node not found in source XML.", vbExclamation
    End If
    
    ' Save the modified template XML
    xmlDocTemplate.Save "C:\path\to\your\filled_file.xml" ' Update with your output file path
End Sub
