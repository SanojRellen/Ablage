
Dim ws As Worksheet
Dim cellValue As String

' Reference the "Template" sheet and cell B4
Set ws = ThisWorkbook.Sheets("Template")
cellValue = ws.Range("B4").Value

With wddDoc
    If Right(cellValue, 5) = "Index" Then
        ' Delete the table containing the "Stock_Box" bookmark
        Dim tbl As Table
        For Each tbl In .Tables
            If tbl.Range.Includes(.Bookmarks("Stock_Box").Range) Then
                tbl.Delete
                Exit For
            End If
        Next tbl
    Else
        ' Delete the table containing the "Index_Box" bookmark
        Dim tblIndex As Table
        For Each tblIndex In .Tables
            If tblIndex.Range.Includes(.Bookmarks("Index_Box").Range) Then
                tblIndex.Delete
                Exit For
            End If
        Next tblIndex
        
        ' Delete the table containing the "Index_Disclaimer" bookmark
        Dim tblDisclaimer As Table
        For Each tblDisclaimer In .Tables
            If tblDisclaimer.Range.Includes(.Bookmarks("Index_Disclaimer").Range) Then
                tblDisclaimer.Delete
                Exit For
            End If
        Next tblDisclaimer
    End If
End With
