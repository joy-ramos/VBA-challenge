Attribute VB_Name = "YearlySummary"
Sub YearlySummary()

Dim ws As String
ws = ActiveSheet.Name

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

'Summary

Dim ColYearlyChange As Range
Dim ColTotalStockVolume As Range

Dim GreatestIncrease As Double
Dim StrGreatestIncrease As String
Dim GreatestDecrease As Double
Dim StrGreatestDecrease As String

Dim GreatestTotalStockVolume As Double

 
'Set range from which to determine largest value
Set ColYearlyChange = Sheets(ws).Range("K2:K" & lastrow)
Set ColTotalStockVolume = Sheets(ws).Range("L2:L" & lastrow)
    
    
'Worksheet function MIN/MAX returns the largest value in a range

GreatestIncrease = Application.WorksheetFunction.Max(ColYearlyChange)
GreatestDecrease = Application.WorksheetFunction.Min(ColYearlyChange)
GreatestTotalStockVolume = Application.WorksheetFunction.Max(ColTotalStockVolume)

'Sets values to string for .Find
StrGreatestIncrease = Str(Round(GreatestIncrease * 100, 2)) + "%"
StrGreatestIncrease = Application.WorksheetFunction.Trim(StrGreatestIncrease)
'MsgBox "'" + StrGreatestIncrease + "'"
StrGreatestDecrease = Str(Round(GreatestDecrease * 100, 2)) + "%"
StrGreatestDecrease = Application.WorksheetFunction.Trim(StrGreatestDecrease)
    

Dim oRange As Range
Dim TickerCol As Integer
Dim TickerRow As Integer

Dim RoundedVal As Double
    
'Finds greatest increase value
Set oRange = Sheets(ws).Range("K2:K" & lastrow).Find(what:=StrGreatestIncrease, lookat:=xlPart)
'Finds Ticker position in spreadsheet
TickerCol = (oRange.Column) - 2
TickerRow = oRange.Row
    
'--DISPLAY SUMMARY--
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"

    'Greatest % Increase
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 16).Value = Cells(TickerRow, TickerCol).Value
    Cells(2, 17).Value = StrGreatestIncrease

    'Greatest % Decrease
    Set oRange = Sheets(ws).Range("K2:K" & lastrow).Find(what:=StrGreatestDecrease, lookat:=xlPart)
    TickerCol = (oRange.Column) - 2
    TickerRow = oRange.Row
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 16).Value = Cells(TickerRow, TickerCol).Value
        RoundedVal = Round((Cells(TickerRow, TickerCol + 2).Value) * 100, 2)
    Cells(3, 17).Value = StrGreatestDecrease

    'Greatest Total Volumne
    Set oRange = Sheets(ws).Range("L2:L" & lastrow).Find(what:=GreatestTotalStockVolume, lookat:=xlPart)
    TickerCol = (oRange.Column) - 3
    TickerRow = oRange.Row
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 16).Value = Cells(TickerRow, TickerCol).Value
    Cells(4, 17).Value = Str(GreatestTotalStockVolume)
    
End Sub
