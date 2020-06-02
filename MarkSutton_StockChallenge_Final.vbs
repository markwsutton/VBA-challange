Sub StockChallenge()

'FinalSubmission
'Set Variables
Dim TickerName As String
Dim TickerNum As Single
Dim SheetNum As Single
Dim TickerFirstRow As Single
Dim TickerLastRow As Single
Dim TickerRowCount As Single
Dim TickerFirstOpen As Single
Dim TickerLastClose As Single
Dim TickerVolume As Single
Dim i As Single
Dim LastRow As Single
Dim LastRow2 As Single
Dim MaxPercentRow As Double
Dim MinPercentRow As Double
Dim MaxVolumeRow As Double

'Intialize variables
SheetNum = 0
TickerFirstRow = 0
TickerLastRow = 0
TickerRowCount = 0
TickerFirstOpen = 0
TickerLastClose = 0
TickerVolume = 0

'Loop through each sheet in Workbook
For Each StockSheet In Worksheets

    TickerNum = 0
    'LastRow = StockSheet.Cells(Rows.Count, 1).End(xlUp).Row

    'Specify worksheet to paste all output
    SheetNum = SheetNum + 1
    Set Worksheet_Current = Worksheets(SheetNum)
    LastRow = Worksheet_Current.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create column labels in each sheet
    Worksheet_Current.Cells(1, 9) = "Ticker"
    
    Worksheet_Current.Cells(1, 10) = "Yearly Change"
    Worksheet_Current.Columns(10).Style = "Currency"
    
    Worksheet_Current.Cells(1, 11) = "Percent Change"
    Worksheet_Current.Columns(11).NumberFormat = "0.00%"
    
    Worksheet_Current.Cells(1, 12) = "Total Stock Volume"
    'Worksheet_Current.Columns(12).Style = "Number"
    
    Worksheet_Current.Columns("I:L").AutoFit

    'Populate Ticker Column with all ticker names on that particular sheet
    'Creat For Next loop for all rows, for each new name add it to column
    'Reset ticker name for each sheet
    TickerName = ""
    
    For i = 2 To LastRow + 1
        
        'If this row has a new ticker name
        If TickerName <> Worksheet_Current.Cells(i, 1) Then
        
            'Finish up previous ticker calculations except for first ticker of each sheet
            If i <> 2 Then
                TickerLastRow = i - 1
                TickerFirstOpen = Worksheet_Current.Cells(TickerFirstRow, 3)
                TickerLastClose = Worksheet_Current.Cells(TickerLastRow, 6)
                Worksheet_Current.Cells(TickerNum + 1, 10).Value = TickerLastClose - TickerFirstOpen
                    
                    'If then to color cell
                    If Worksheet_Current.Cells(TickerNum + 1, 10).Value > 0 Then
                        Worksheet_Current.Cells(TickerNum + 1, 10).Interior.ColorIndex = 4
                    ElseIf Worksheet_Current.Cells(TickerNum + 1, 10).Value = 0 Then
                        Worksheet_Current.Cells(TickerNum + 1, 10).Interior.ColorIndex = 0
                    Else: Worksheet_Current.Cells(TickerNum + 1, 10).Interior.ColorIndex = 3
                    End If
                    
                'Add if thens for if volume = 0, TickerFirstOpen = 0
                If TickerVolume = 0 Or TickerFirstOpen = 0 Then
                    Worksheet_Current.Cells(TickerNum + 1, 11).Value = 0
                    Worksheet_Current.Cells(TickerNum + 1, 12).Value = 0
                Else: Worksheet_Current.Cells(TickerNum + 1, 11).Value = (TickerLastClose - TickerFirstOpen) / TickerFirstOpen
                        TickerVolume = TickerVolume + Worksheet_Current.Cells(i, 7)
                        Worksheet_Current.Cells(TickerNum + 1, 12).Value = TickerVolume
                End If
            End If
        
            'Start New Ticker Process
            TickerNum = TickerNum + 1
            TickerName = Worksheet_Current.Cells(i, 1)
            Worksheet_Current.Cells(TickerNum + 1, 9).Value = TickerName
            TickerFirstRow = i
            TickerVolume = 0
            
        'If this is just a next row within a ticker, simply add to the volume and go to next row
        ElseIf TickerName = Worksheet_Current.Cells(i, 1) Then
            TickerVolume = TickerVolume + CLng(Worksheet_Current.Cells(i, 7))
    
        End If
        
    Next i
    
    'Create column labels in each sheet
    Worksheet_Current.Cells(1, 16) = "Ticker"
    Worksheet_Current.Cells(1, 17) = "Value"
    Worksheet_Current.Cells(2, 15) = "Greatest % Increase"
    Worksheet_Current.Cells(3, 15) = "Greatest % Decrease"
    Worksheet_Current.Cells(4, 15) = "Greatest Total Volume"
    
    'Find the three values, associated Tickers, and assign them to the new table
    LastRow2 = Worksheet_Current.Cells(Rows.Count, 9).End(xlUp).Row
    'Get Max and Min of Column 11 and Max of 12 and assign in new table
    Worksheet_Current.Cells(2, 17) = Application.WorksheetFunction.Max(Range(Worksheet_Current.Cells(2, 11), Worksheet_Current.Cells(LastRow2, 11)))
    Worksheet_Current.Cells(3, 17) = Application.WorksheetFunction.Min(Range(Worksheet_Current.Cells(2, 11), Worksheet_Current.Cells(LastRow2, 11)))
    Worksheet_Current.Cells(4, 17) = Application.WorksheetFunction.Max(Range(Worksheet_Current.Cells(2, 12), Worksheet_Current.Cells(LastRow2, 12)))
    
    'Use WorksheetFunction.Match to get rows of the above and then paste the ticker symbols
    MaxPercentRow = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow2)), Range("K2:K" & LastRow2), 0)
    Worksheet_Current.Cells(2, 16) = Worksheet_Current.Cells(MaxPercentRow + 1, 9)
    MinPercentRow = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow2)), Range("K2:K" & LastRow2), 0)
    Worksheet_Current.Cells(3, 16) = Worksheet_Current.Cells(MinPercentRow + 1, 9)
    MaxVolumeRow = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow2)), Range("L2:L" & LastRow2), 0)
    Worksheet_Current.Cells(4, 16) = Worksheet_Current.Cells(MaxVolumeRow + 1, 9)
    
    'Autofit the columns after they are populated
    Worksheet_Current.Columns("O:Q").AutoFit
    

Next StockSheet

End Sub



