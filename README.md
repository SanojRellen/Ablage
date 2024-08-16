 Sub DeleteCommentAfterTelefon()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim telefonNode As IXMLDOMNode
    Dim commentNode As IXMLDOMNode
    
    ' Load the XML document
    Set xmlDoc = New MSXML2.DOMDocument60
    xmlDoc.Load "C:\path\to\your\file.xml" ' Update with your XML file path
    
    ' Find the Telefon node
    Set telefonNode = xmlDoc.SelectSingleNode("//Telefon")
    
    If Not telefonNode Is Nothing Then
        ' Find the comment node after the Telefon node
        Set commentNode = telefonNode.NextSibling
        
        ' Ensure the comment is the one we're targeting
        If Not commentNode Is Nothing And commentNode.NodeType = MSXML2.NODE_COMMENT Then
            If InStr(commentNode.Text, "bitte einfügen") > 0 Then
                ' Remove the comment node
                telefonNode.ParentNode.RemoveChild commentNode
            End If
        End If
    End If
    
    ' Save the modified XML
    xmlDoc.Save "C:\path\to\your\filled_file.xml" ' Update with your output file path
End Sub
