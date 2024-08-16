Sub FillKurznameWithYear()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim sourceNode As IXMLDOMNode
    Dim templateNode As IXMLDOMNode
    Dim extractedYear As String
    
    ' Load the XML documents
    Set xmlDocSource = New MSXML2.DOMDocument60
    Set xmlDocTemplate = New MSXML2.DOMDocument60
    
    xmlDocSource.Load "C:\path\to\your\source.xml" ' Update with your source XML file path
    xmlDocTemplate.Load "C:\path\to\your\template.xml" ' Update with your template XML file path
    
    ' Find the maturity date node in the source XML
    Set sourceNode = xmlDocSource.SelectSingleNode("//MaturityDate") ' Update with the actual node name
    
    If Not sourceNode Is Nothing Then
        ' Extract the year (last two digits) from the date
        extractedYear = Right(sourceNode.Text, 2)
        
        ' Find the Kurzname node in the template XML
        Set templateNode = xmlDocTemplate.SelectSingleNode("//Kurzname") ' Update with the actual node name
        
        If Not templateNode Is Nothing Then
            ' Construct the desired text
            templateNode.Text = "BARCL.BK EXP.Z" & extractedYear & " SX5E"
        Else
            MsgBox "Kurzname node not found in template XML.", vbExclamation
        End If
    Else
        MsgBox "MaturityDate node not found in source XML.", vbExclamation
    End If
    
    ' Save the modified template XML
    xmlDocTemplate.Save "C:\path\to\your\filled_file.xml" ' Update with your output file path
End Sub
