Attribute VB_Name = "VBAcode_withSheetLoop"
Sub SheetLoop()

Dim ws As Worksheet

For Each ws In Worksheets
    
    ws.Activate
    VBAhw
    
Next

End Sub

Sub VBAhw()

Dim ws As String
ws = ActiveSheet.Name

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

Sheets(ws).Range("A2:A" & lastrow).AdvancedFilter _
    Action:=xlFilterCopy, _
    CopyToRange:=Sheets(ws).Range("I1"), _
    Unique:=True
    
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
    
'--Select row count for tickers
Dim TickerLastRow As Long
TickerLastRow = Cells(Rows.Count, "I").End(xlUp).Row
'--MsgBox TickerLastRow


For a = 2 To TickerLastRow

    Application.ActiveSheet.UsedRange

    Dim DateAndValuesRange As Range
    Dim TickerString As String
    TickerString = Cells(a, 9).Value
    
    
    Dim lastrow2 As Long
    Dim xRg As Range, yRg As Range
        
        '--Range for date and values for ticker

        With ThisWorkbook.Worksheets(ws)
            lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
            Application.ScreenUpdating = False
            
            For Each xRg In .Range("A1:A" & lastrow)
                If UCase(xRg.Text) = TickerString Then
                    If (yRg Is Nothing) And (DateAndValuesRange Is Nothing) Then
                        Set yRg = .Range("B" & xRg.Row).Resize(, 1)
                        Set DateAndValuesRange = .Range("B" & xRg.Row).Resize(, 6)
                        
                    Else
                        Set yRg = Union(yRg, .Range("B" & xRg.Row).Resize(, 1))
                        Set DateAndValuesRange = Union(DateAndValuesRange, .Range("B" & xRg.Row).Resize(, 6))
                    End If
                End If
            Next xRg
            Application.ScreenUpdating = True
            
  
        End With
        

        '--If Not yRg Is Nothing Then yRg.Select
        'If Not DateAndValuesRange Is Nothing Then DateAndValuesRange.Select
        

    '--Count number of rows for the current range
    
        Dim TickerRangeCount As Long
        
        TickerRangeCount = 0
        
        TickerRangeCount = DateAndValuesRange.Count
        TickerRangeCount = TickerRangeCount / 6
        'MsgBox TickerRangeCount
    
    '--Select min date and max date for each ticker
    
        Dim MinDate As Double
        Dim MaxDate As Double
        
        MinDate = Application.WorksheetFunction.Min(yRg)
        MaxDate = Application.WorksheetFunction.Max(yRg)
        
        'MsgBox MinDate
        'MsgBox MaxDate
        
    '--Select opening and closing value of the ticker
        Dim YearOpenValue As Double
        Dim YearCloseValue As Double
        
        YearOpenValue = 0
        YearCloseValue = 0
        
        
        Dim TotalVolume As Double
        TotalVolume = 0
        
        YearOpenValue = DateAndValuesRange.Cells(1, 2).Value
        YearCloseValue = DateAndValuesRange.Cells(TickerRangeCount, 5).Value
            
            
            
        For i = 1 To TickerRangeCount
            TotalVolume = TotalVolume + DateAndValuesRange.Cells(i, 6).Value
        Next i
        
        
        
        'MsgBox YearCloseValue
        'MsgBox TotalVolume
        
    'Calculate Yearly Change and Percent Change
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    YearlyChange = 0
    PercentChange = 0
    
    
    YearlyChange = YearCloseValue - YearOpenValue
    YearlyChange = Round(YearlyChange, 2)
    
    PercentChange = (YearlyChange / YearOpenValue) * 100
    PercentChange = Round(PercentChange, 2)

    'Formats to 2 decimal places and colors the cell
    
    Cells(a, 10).Value = YearlyChange
        If Cells(a, 10).Value >= 0 Then
            Cells(a, 10).Interior.Color = vbGreen
        Else
            Cells(a, 10).Interior.Color = vbRed
        End If
    
    Cells(a, 11).Value = Str(PercentChange) + "%"
    Cells(a, 12).Value = TotalVolume
    
    Set DateAndValuesRange = Nothing
    Set yRg = Nothing
    

Next a

'SUMMARY SECTION

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


'--THE END!!! Woohooo :)

End Sub


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
