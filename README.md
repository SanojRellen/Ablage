
Sub SetCalculationModeAndSaveSettings()
    ' Store original settings to revert back later
    Dim originalCalculation As XlCalculation
    Dim originalRecalculateBeforeSave As Boolean

    ' Save current settings
    originalCalculation = Application.Calculation
    originalRecalculateBeforeSave = Application.CalculateBeforeSave

    ' Change Excel settings
    Application.Calculation = xlCalculationManual   ' Set to manual calculation
    Application.CalculateBeforeSave = False          ' Untick Recalculate workbook before saving

    ' Your macro code goes here
    ' For example:
    MsgBox "Performing operations with manual calculation mode."

    ' After your code, revert to the original settings
    Application.Calculation = originalCalculation
    Application.CalculateBeforeSave = originalRecalculateBeforeSave

    MsgBox "Settings reverted back to original."
End Sub





Sub DeleteZeroValues()
    Dim targetSheets As Variant
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    ' Define target sheets
    targetSheets = Array("Raw_Data", "Raw_Duaration", "Raw_Yield")

    ' Loop through each sheet in the array
    For Each sheetName In targetSheets
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Define the range A8:GC6000
        Set rng = ws.Range("A8:GC6000")
        
        ' Loop through each cell in the range
        For Each cell In rng
            If cell.Value = 0 Then
                cell.ClearContents ' Clear the cell if it contains exactly 0
            End If
        Next cell
    Next sheetName
    
    MsgBox "Zero values deleted."
End Sub

