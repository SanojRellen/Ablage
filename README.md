
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
