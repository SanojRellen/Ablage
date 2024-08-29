Sub AdjustAndFillTemplateWithDatesAndValues()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim callEvent_item_nodes As IXMLDOMNodeList
    Dim templateCouponNodes As IXMLDOMNodeList
    Dim i As Integer
    Dim callEventItemCount As Integer
    Dim templateCouponCount As Integer
    Dim interestCommencementDateNode As IXMLDOMNode
    Dim interestCommencementDate As String
    
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
    
    ' Get the interest commencement date from the source XML
    Set interestCommencementDateNode = xmlDocSource.SelectSingleNode("//interestCommencementDate")
    
    If interestCommencementDateNode Is Nothing Then
        MsgBox "'interestCommencementDate' node not found in source XML.", vbExclamation
        Exit Sub
    End If
    
    interestCommencementDate = interestCommencementDateNode.Text
    
    ' Get the list of item nodes from the source XML
    Set callEvent_item_nodes = xmlDocSource.SelectNodes("//couponEvents/schedule/item")
    
    If callEvent_item_nodes Is Nothing Or callEvent_item_nodes.Length = 0 Then
        MsgBox "No 'item' nodes found in source XML.", vbExclamation
        Exit Sub
    End If
    
    callEventItemCount = callEvent_item_nodes.Length
    
    ' Get the list of Coupon nodes from the template XML
    Set templateCouponNodes = xmlDocTemplate.SelectNodes("//Zahlungen/Coupon")
    
    If templateCouponNodes Is Nothing Then
        MsgBox "No 'Coupon' nodes found in template XML.", vbExclamation
        Exit Sub
    End If
    
    templateCouponCount = templateCouponNodes.Length
    
    ' Adjust the template to match the number of source items
    If callEventItemCount < 6 Then
        ' Remove extra Coupon nodes in the template
        For i = templateCouponCount To callEventItemCount + 1 Step -1
            templateCouponNodes.Item(i - 1).ParentNode.removeChild templateCouponNodes.Item(i - 1)
        Next i
    End If
    
    ' Re-select the Coupon nodes after deletion
    Set templateCouponNodes = xmlDocTemplate.SelectNodes("//Zahlungen/Coupon")
    
    ' Fill each Coupon node with the corresponding data from the source
    For i = 0 To templateCouponNodes.Length - 1
        ' Get the current Coupon node in the template XML
        Dim currentCouponNode As IXMLDOMElement
        Set currentCouponNode = templateCouponNodes.Item(i)
        
        ' Ensure the current Coupon node exists
        If Not currentCouponNode Is Nothing Then
            ' Get the relevant nodes from the source XML
            Dim paymentDateNode As IXMLDOMNode
            Dim fixedAmountRelativeValueNode As IXMLDOMNode
            
            Set paymentDateNode = callEvent_item_nodes.Item(i).SelectSingleNode("paymentDate")
            Set fixedAmountRelativeValueNode = callEvent_item_nodes.Item(i).SelectSingleNode("fixedAmountRelative/value")
            
            ' Ensure nodes exist before proceeding
            If Not paymentDateNode Is Nothing And Not fixedAmountRelativeValueNode Is Nothing Then
                ' Convert fixed amount value and fill Bonusbetrag/Wert
                Dim bonusbetragValue As Double
                bonusbetragValue = CDbl(fixedAmountRelativeValueNode.Text) * 1000
                currentCouponNode.SelectSingleNode("Bonusbetrag/Wert").Text = CStr(bonusbetragValue)
                
                ' Fill Zinsperiode/Beginn and Zinsperiode/Ende
                If i = 0 Then
                    ' For the first Coupon, use interestCommencementDate for Beginn and paymentDate - 1 for Ende
                    currentCouponNode.SelectSingleNode("Zinsperiode/Beginn").Text = interestCommencementDate
                    currentCouponNode.SelectSingleNode("Zinsperiode/Ende").Text = Format(DateAdd("d", -1, CDate(paymentDateNode.Text)), "yyyy-mm-dd")
                Else
                    ' For subsequent Coupons, use the previous paymentDate for Beginn and paymentDate - 1 for Ende
                    currentCouponNode.SelectSingleNode("Zinsperiode/Beginn").Text = callEvent_item_nodes.Item(i - 1).SelectSingleNode("paymentDate").Text
                    currentCouponNode.SelectSingleNode("Zinsperiode/Ende").Text = Format(DateAdd("d", -1, CDate(paymentDateNode.Text)), "yyyy-mm-dd")
                End If
            Else
                MsgBox "Required nodes not found in source item node " & (i + 1) & ".", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "No Coupon node found at index " & (i + 1) & " in the template.", vbExclamation
            Exit Sub
        End If
    Next i
    
    ' Save the modified template XML
    xmlDocTemplate.Save "C:\path\to\your\filled_template.xml" ' Update with your output file path
    
    MsgBox "Template has been adjusted and filled successfully!", vbInformation
End Sub
