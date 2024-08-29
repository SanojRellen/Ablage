Sub AdjustAndFillTemplateWithDatesAndValues()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim callEvent_item_nodes As IXMLDOMNodeList
    Dim templateKuendigungNodes As IXMLDOMNodeList
    Dim zinsfestlegungstageNode As IXMLDOMElement
    Dim i As Integer
    Dim callEventItemCount As Integer
    Dim unitSizeNode As IXMLDOMNode
    Dim unitSizeValue As String
    
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
    Set callEvent_item_nodes = xmlDocSource.SelectNodes("//callevents/schedule/item")
    
    If callEvent_item_nodes Is Nothing Or callEvent_item_nodes.Length = 0 Then
        MsgBox "No 'item' nodes found in source XML.", vbExclamation
        Exit Sub
    End If
    
    callEventItemCount = callEvent_item_nodes.Length
    
    ' Get the constant unitSize value from the source XML
    Set unitSizeNode = xmlDocSource.SelectSingleNode("//unitSize")
    
    If unitSizeNode Is Nothing Then
        MsgBox "'unitSize' node not found in source XML.", vbExclamation
        Exit Sub
    End If
    
    unitSizeValue = unitSizeNode.Text
    
    ' Get the list of kuendigung nodes from the template XML
    Set templateKuendigungNodes = xmlDocTemplate.SelectNodes("//zahlungen/kuendigung")
    
    If templateKuendigungNodes Is Nothing Then
        MsgBox "No 'kuendigung' nodes found in template XML.", vbExclamation
        Exit Sub
    End If
    
    ' Adjust the template to match the number of source items
    If callEventItemCount < 5 Then
        ' Remove extra kuendigung nodes in the template
        For i = templateKuendigungNodes.Length To callEventItemCount + 1 Step -1
            templateKuendigungNodes.Item(i - 1).ParentNode.removeChild templateKuendigungNodes.Item(i - 1)
        Next i
    End If
    
    ' Re-select the kuendigung nodes after deletion
    Set templateKuendigungNodes = xmlDocTemplate.SelectNodes("//zahlungen/kuendigung")
    
    ' Fill each kuendigung node with the corresponding data from the source
    For i = 0 To callEventItemCount - 1
        ' Get the relevant nodes from the source XML
        Dim barrierDateNode As IXMLDOMNode
        Dim settlementDateNode As IXMLDOMNode
        Dim barrierLevelValueNode As IXMLDOMNode
        
        Set barrierDateNode = callEvent_item_nodes.Item(i).SelectSingleNode("barrierEventObservationDates/item")
        Set settlementDateNode = callEvent_item_nodes.Item(i).SelectSingleNode("settlementDate")
        Set barrierLevelValueNode = callEvent_item_nodes.Item(i).SelectSingleNode("barrierLevelRelative/value")
        
        ' Ensure nodes exist before proceeding
        If barrierDateNode Is Nothing Or settlementDateNode Is Nothing Or barrierLevelValueNode Is Nothing Then
            MsgBox "Required nodes not found in source item node " & (i + 1) & ".", vbExclamation
            Exit Sub
        End If
        
        ' Format the barrier level value
        Dim barrierLevelFormatted As String
        barrierLevelFormatted = Format(CDbl(barrierLevelValueNode.Text) * 100, "0.00")
        
        ' Fill the template nodes with the extracted data
        templateKuendigungNodes.Item(i).SelectSingleNode("beobachtungstag").Text = barrierDateNode.Text
        templateKuendigungNodes.Item(i).SelectSingleNode("rückzahlungsvaluta").Text = settlementDateNode.Text
        templateKuendigungNodes.Item(i).SelectSingleNode("Kuendigungskurs").Text = unitSizeValue
        templateKuendigungNodes.Item(i).SelectSingleNode("Tilgungslevelprozent").Text = barrierLevelFormatted
        
    Next i
    
    ' Save the modified template XML
    xmlDocTemplate.Save "C:\path\to\your\filled_template.xml" ' Update with your output file path
    
    MsgBox "Template has been adjusted and filled successfully!", vbInformation
End Sub












.
' Fill each kuendigung node with the corresponding data from the source
For i = 0 To callEventItemCount - 1
    ' Get the current kuendigung node in the template XML
    Dim currentKuendigungNode As IXMLDOMElement
    Set currentKuendigungNode = templateKuendigungNodes.Item(i)
    
    ' Ensure the current kuendigung node exists
    If Not currentKuendigungNode Is Nothing Then
        ' Get the relevant nodes from the source XML
        Dim barrierDateNode As IXMLDOMNode
        Dim settlementDateNode As IXMLDOMNode
        Dim barrierLevelValueNode As IXMLDOMNode
        
        Set barrierDateNode = callEvent_item_nodes.Item(i).SelectSingleNode("barrierEventObservationDates/item")
        Set settlementDateNode = callEvent_item_nodes.Item(i).SelectSingleNode("settlementDate")
        Set barrierLevelValueNode = callEvent_item_nodes.Item(i).SelectSingleNode("barrierLevelRelative/value")
        
        ' Ensure nodes exist before proceeding
        If Not barrierDateNode Is Nothing Then
            Dim beobachtungstagNode As IXMLDOMNode
            Set beobachtungstagNode = currentKuendigungNode.SelectSingleNode("beobachtungstag")
            If Not beobachtungstagNode Is Nothing Then
                beobachtungstagNode.Text = barrierDateNode.Text
            Else
                MsgBox "'beobachtungstag' node not found in kuendigung node " & (i + 1) & ".", vbExclamation
            End If
        Else
            MsgBox "'barrierEventObservationDates/item' node not found in source item node " & (i + 1) & ".", vbExclamation
            Exit Sub
        End If
        
        If Not settlementDateNode Is Nothing Then
            Dim rueckzahlungsvalutaNode As IXMLDOMNode
            Set rueckzahlungsvalutaNode = currentKuendigungNode.SelectSingleNode("rückzahlungsvaluta")
            If Not rueckzahlungsvalutaNode Is Nothing Then
                rueckzahlungsvalutaNode.Text = settlementDateNode.Text
            Else
                MsgBox "'rückzahlungsvaluta' node not found in kuendigung node " & (i + 1) & ".", vbExclamation
            End If
        Else
            MsgBox "'settlementDate' node not found in source item node " & (i + 1) & ".", vbExclamation
            Exit Sub
        End If
        
        If Not unitSizeNode Is Nothing Then
            Dim kuendigungskursNode As IXMLDOMNode
            Set kuendigungskursNode = currentKuendigungNode.SelectSingleNode("Kuendigungskurs")
            If Not kuendigungskursNode Is Nothing Then
                kuendigungskursNode.Text = unitSizeValue
            Else
                MsgBox "'Kuendigungskurs' node not found in kuendigung node " & (i + 1) & ".", vbExclamation
            End If
        Else
            MsgBox "'unitSize' node not found in source XML.", vbExclamation
            Exit Sub
        End If
        
        If Not barrierLevelValueNode Is Nothing Then
            Dim tilgungslevelprozentNode As IXMLDOMNode
            Set tilgungslevelprozentNode = currentKuendigungNode.SelectSingleNode("Tilgungslevelprozent")
            If Not tilgungslevelprozentNode Is Nothing Then
                ' Format the barrier level value
                Dim barrierLevelFormatted As String
                barrierLevelFormatted = Format(CDbl(barrierLevelValueNode.Text) * 100, "0.00")
                tilgungslevelprozentNode.Text = barrierLevelFormatted
            Else
                MsgBox "'Tilgungslevelprozent' node not found in kuendigung node " & (i + 1) & ".", vbExclamation
            End If
        Else
            MsgBox "'barrierLevelRelative/value' node not found in source item node " & (i + 1) & ".", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "No kuendigung node found at index " & (i + 1) & " in the template.", vbExclamation
        Exit Sub
    End If
Next i
