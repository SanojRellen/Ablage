Sub AdjustAndFillTemplate()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim sourceItemNodes As IXMLDOMNodeList
    Dim templateDatumNodes As IXMLDOMNodeList
    Dim zinsfestlegungstageNode As IXMLDOMElement
    Dim i As Integer
    Dim sourceItemCount As Integer
    
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
    
    ' Get the list of item nodes from the source XML
    Set sourceItemNodes = xmlDocSource.SelectNodes("//couponEvents/schedule/item")
    
    If sourceItemNodes Is Nothing Then
        MsgBox "No 'item' nodes found in source XML.", vbExclamation
        Exit Sub
    End If
    
    sourceItemCount = sourceItemNodes.Length
    
    ' Get the list of Datum nodes from the template XML
    Set templateDatumNodes = xmlDocTemplate.SelectNodes("//zinsfestlegungstage/Datum")
    
    If templateDatumNodes Is Nothing Then
        MsgBox "No 'Datum' nodes found in template XML.", vbExclamation
        Exit Sub
    End If
    
    ' Get the zinsfestlegungstage node in the template to remove Datum nodes if needed
    Set zinsfestlegungstageNode = xmlDocTemplate.SelectSingleNode("//zinsfestlegungstage")
    
    If zinsfestlegungstageNode Is Nothing Then
        MsgBox "'zinsfestlegungstage' node not found in template XML.", vbExclamation
        Exit Sub
    End If
    
    ' Check if the number of item nodes is less than 6
    If sourceItemCount < 6 Then
        ' Remove extra Datum nodes in the template
        For i = 6 To sourceItemCount + 1 Step -1
            zinsfestlegungstageNode.removeChild templateDatumNodes.Item(i - 1)
        Next i
    End If
    
    ' Re-select the Datum nodes after deletion
    Set templateDatumNodes = xmlDocTemplate.SelectNodes("//zinsfestlegungstage/Datum")
    
    ' Check again after modification
    If templateDatumNodes Is Nothing Then
        MsgBox "Error after modifying template XML.", vbExclamation
        Exit Sub
    End If
    
    ' Fill the Datum nodes with the corresponding valuationDates from the source
    For i = 0 To sourceItemCount - 1
        ' Ensure both source item and template datum nodes exist
        If Not sourceItemNodes.Item(i) Is Nothing And Not templateDatumNodes.Item(i) Is Nothing Then
            ' Get the valuationDate node from the source item node
            Dim valuationDateNode As IXMLDOMNode
            Set valuationDateNode = sourceItemNodes.Item(i).SelectSingleNode("valuationDate")
            
            ' Check if the valuationDate node exists
            If Not valuationDateNode Is Nothing Then
                ' Fill the template Datum node with the valuationDate value
                templateDatumNodes.Item(i).Text = valuationDateNode.Text
            Else
                MsgBox "valuationDate node not found in source item node " & (i + 1) & ".", vbExclamation
            End If
        Else
            MsgBox "Error accessing source item or template Datum nodes at index " & (i + 1) & ".", vbExclamation
        End If
    Next i
    
    ' Save the modified template XML
    xmlDocTemplate.Save "C:\path\to\your\filled_template.xml" ' Update with your output file path
    
    MsgBox "Template has been adjusted and filled successfully!", vbInformation
End Sub
