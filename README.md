Sub FormatAndAppendEUR()
    Dim cell As Range
    Dim formattedValue As String
    
    ' Loop through each cell in the range C20:H20
    For Each cell In Range("C20:H20")
        ' Convert the cell value to a number format with two decimal places
        cell.NumberFormat = "0.00"
        
        ' Get the current value of the cell, ensuring it shows two decimal places
        formattedValue = Format(cell.Value, "0.00")
        
        ' Append " EUR"
        formattedValue = formattedValue & " EUR"
        
        ' Update the cell with the new formatted value
        cell.Value = formattedValue
        
        ' Set the cell format to Text to keep the " EUR" part
        cell.NumberFormat = "@"
    Next cell
End Sub

















Sub ParseAndFillMostRecentXML()
    Dim xmlDocSource As MSXML2.DOMDocument60
    Dim xmlDocTemplate As MSXML2.DOMDocument60
    Dim sourceNode As IXMLDOMNode
    Dim templateNode As IXMLDOMNode
    
    ' Define paths
    Dim pickupPath As String
    Dim storagePath As String
    Dim templatePath As String
    Dim inputFileName As String
    Dim outputFileName As String
    
    pickupPath = "C:\path\to\input\files\" ' Change to your input files directory
    storagePath = "C:\path\to\output\files\" ' Change to your output files directory
    templatePath = "C:\path\to\template\template.xml" ' Change to your template file path
    
    ' Find the most recent file in the pickup path
    inputFileName = ""
    Dim fileName As String
    Dim mostRecentFile As String
    Dim mostRecentDate As Date
    Dim fileDate As Date
    
    fileName = Dir(pickupPath & "*.xml")
    If fileName <> "" Then
        mostRecentFile = fileName
        mostRecentDate = FileDateTime(pickupPath & fileName)
        
        Do While fileName <> ""
            fileDate = FileDateTime(pickupPath & fileName)
            If fileDate > mostRecentDate Then
                mostRecentDate = fileDate
                mostRecentFile = fileName
            End If
            fileName = Dir
        Loop
        
        inputFileName = mostRecentFile
    End If
    
    If inputFileName = "" Then
        MsgBox "No XML files found in the pickup path.", vbExclamation
        Exit Sub
    End If
    
    ' Create the output file name
    outputFileName = "filled_" & inputFileName
    
    ' Create XML document objects
    Set xmlDocSource = New MSXML2.DOMDocument60
    Set xmlDocTemplate = New MSXML2.DOMDocument60
    
    ' Load the template XML file
    If Not xmlDocTemplate.Load(templatePath) Then
        MsgBox "Failed to load template XML file.", vbExclamation
        Exit Sub
    End If
    
    ' Load the input XML file
    If xmlDocSource.Load(pickupPath & inputFileName) Then
        ' Find the nameLong node in the source XML
        Set sourceNode = xmlDocSource.SelectSingleNode("//nameLong")
        If Not sourceNode Is Nothing Then
            ' Find the Name node in the template XML
            Set templateNode = xmlDocTemplate.SelectSingleNode("//Name")
            If Not templateNode Is Nothing Then
                ' Replace the text content
                templateNode.Text = sourceNode.Text
                
                ' Save the modified template XML to the storage path
                xmlDocTemplate.Save storagePath & outputFileName
            Else
                MsgBox "Name node not found in template XML.", vbExclamation
            End If
        Else
            MsgBox "nameLong node not found in input XML.", vbExclamation
        End If
    Else
        MsgBox "Failed to load input XML file.", vbExclamation
    End If
    
    MsgBox "XML Processing Completed!", vbInformation
End Sub
