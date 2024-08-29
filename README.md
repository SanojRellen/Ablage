Sub InsertPreviousDayDate()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim sourceDateNode As IXMLDOMNode
    Dim templateDateNode As IXMLDOMNode
    Dim sourceDate As String
    Dim previousDate As String
    
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
    
    ' Get the date node from the source XML
    Set sourceDateNode = xmlDocSource.SelectSingleNode("//yourSourceNodePath") ' Update with your actual XPath
    
    If sourceDateNode Is Nothing Then
        MsgBox "'sourceDateNode' not found in source XML.", vbExclamation
        Exit Sub
    End If
    
    ' Extract the date as a string
    sourceDate = sourceDateNode.Text
    
    ' Convert the date string to a VBA date and subtract one day
    previousDate = Format(DateAdd("d", -1, CDate(sourceDate)), "dd-mm-yyyy")
    
    ' Get the target node in the template XML
    Set templateDateNode = xmlDocTemplate.SelectSingleNode("//yourTemplateNodePath") ' Update with your actual XPath
    
    If templateDateNode Is Nothing Then
        MsgBox "'templateDateNode' not found in template XML.", vbExclamation
        Exit Sub
    End If
    
    ' Insert the previous date into the template XML
    templateDateNode.Text = previousDate
    
    ' Save the modified template XML
    xmlDocTemplate.Save "C:\path\to\your\filled_template.xml" ' Update with your output file path
    
    MsgBox "Previous date has been inserted into the template successfully!", vbInformation
End Sub
