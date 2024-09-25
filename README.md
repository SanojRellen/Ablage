
Sub CopyFormatting()

    Dim sourceRange As Range
    Dim targetRange As Range
    Dim cellSource As Range
    Dim cellTarget As Range
    Dim rowOffset As Long
    Dim colOffset As Long

    ' Define the source range (A1:C4)
    Set sourceRange = Range("A1:C4")
    
    ' Define the target range (A5:C8)
    Set targetRange = Range("A5:C8")
    
    ' Loop through each cell in the source range
    For Each cellSource In sourceRange
        ' Calculate the row and column offset from the source cell
        rowOffset = cellSource.Row - sourceRange.Row
        colOffset = cellSource.Column - sourceRange.Column
        
        ' Map the corresponding cell in the target range
        Set cellTarget = targetRange.Cells(1 + rowOffset, 1 + colOffset)
        
        ' Apply the formatting from the source cell to the target cell
        cellTarget.Font = cellSource.Font
        cellTarget.Interior.Color = cellSource.Interior.Color
        cellTarget.Borders.LineStyle = cellSource.Borders.LineStyle
        cellTarget.NumberFormat = cellSource.NumberFormat
    Next cellSource

End Sub

