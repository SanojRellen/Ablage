Sub FilterAndCopyAppointments()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim copyRange As Range
    Dim pasteRange As Range

    ' Set the worksheet to "Appointments"
    Set ws = ThisWorkbook.Sheets("Appointments")

    ' Find the last row in column AE
    lastRow = ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row

    ' Apply the filter on column AE for value 1
    ws.Range("A1:AE" & lastRow).AutoFilter Field:=31, Criteria1:="1"  ' Field 31 = Column AE

    ' Define the range to copy (only visible cells in column D)
    Set copyRange = ws.Range("D2:D" & lastRow).SpecialCells(xlCellTypeVisible)

    ' Define the starting cell to paste (AG19) and resize the range to match the copy range
    Set pasteRange = ws.Range("AG19").Resize(copyRange.Rows.Count, 1)

    ' Copy values from column D to column AG starting from AG19
    copyRange.Copy
    pasteRange.PasteSpecial xlPasteValues

    ' Remove the filter
    ws.AutoFilterMode = False

    ' Clear the clipboard
    Application.CutCopyMode = False
End Sub
