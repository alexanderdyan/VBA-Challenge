' Steps:
' ----------------------------------------------------------------------------
' 1. For every worksheet in book
' 2. Go through each row
' 3. Find every ticker's first and last row
' 4. For every new ticker that is encountered
'    Store the corresponding values and
'    Set the computed values in new columns

Sub Extract_Tickers()

    'Loop thru all Sheets -
    For Each ws In Worksheets
        Dim WorksheetYear As String
        WorksheetYear = ws.Name
        
        'Count unique ticker values
        Dim UniqueTickerRow As Integer
        UniqueTickerRow = 1
        
        'Count # of Rows in Sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        
        'Store Stock Prices for every unique ticker
        Dim StartOpeningPrice As Double
        Dim StartClosingPrice As Double
        Dim TotalStockVolume As Double
        
       'Find start prices for first ticker
       StartOpeningPrice = ws.Cells(2, 3).Value
       StartClosingPrice = ws.Cells(2, 6).Value
       
       'stock volume will be updated in loop below
       TotalStockVolume = 0
       
       Dim IsLastTicker As Boolean
       
       'Search every row of sheet (except first)
       For i = 2 To LastRow
            
            ticker = ws.Cells(i, 1).Value
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            'Determine if row is the last one before next ticker
            IsLastTicker = Not ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value
            
            If i = LastRow Then
                IsLastTicker = True
            End If
            
            'add column to unique ticker
            If IsLastTicker Then
                UniqueTickerRow = UniqueTickerRow + 1
                
                'Set unique ticker name in Ticker Column J = 10
                ws.Cells(UniqueTickerRow, 10).Value = ticker
                
                'Create variables for get values from current row and
                'calculate corresponding values in new columns
                Dim EndOpeningPrice As Double
                Dim EndClosingPrice As Double
                Dim YearlyChange As Double
                Dim PercentChange As Double
                
                'get end prices for last date of ticker
                EndOpeningPrice = ws.Cells(i, 3).Value
                EndClosingPrice = ws.Cells(i, 6).Value
                
                'Calculate yearly change
                YearlyChange = EndClosingPrice - StartOpeningPrice
                
                'eliminate division by zero to avoid error
                If StartOpeningPrice = 0 Then
                    StartOpeningPrice = 0.01
                End If
                
                'calculate percentage change
                PercentChange = YearlyChange / (StartOpeningPrice)
                
                ws.Cells(UniqueTickerRow, 11).Value = YearlyChange
                ws.Cells(UniqueTickerRow, 12).Value = Format(PercentChange, "0.00%")
                ws.Cells(UniqueTickerRow, 13).Value = TotalStockVolume
                
                'set color for yearly change cell
                If EndClosingPrice > StartOpeningPrice Then
                    ws.Cells(UniqueTickerRow, 11).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(UniqueTickerRow, 11).Interior.Color = RGB(255, 0, 0)
                End If
                
                'Get StartOpeningPrice for next ticker
                StartOpeningPrice = ws.Cells(i + 1, 3).Value
                
                'Set stock volume for next ticker to 0, gets updated at start of loop
                TotalStockVolume = 0
            End If
            
            'move to next row
        Next i
    'move to next ws
    Next ws
       
            
        

End Sub


