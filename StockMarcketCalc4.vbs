Sub StockMarketCalc():
    'declare variables
    Dim ws as Worksheet
    
    For Each ws In Worksheets
        Dim ResultTicker as String
        Dim TickerTotal as Double
        Dim RowTicker as String
        Dim RowVolume as Double
        Dim ResultRow as Long
        Dim LastRow as Long

        Dim RowDate as Long
        Dim FirstDay as Long
        Dim LastDay as Long
        Dim OpenPrice as Double
        Dim ClosePrice as Double
        Dim YrChange as Double
        Dim PctChange as Double

        'initialize variables
        ResultTicker = Cells(2,1).Value
        TickerTotal = 0
        ResultRow = 2
        FirstDay = Cells(2,2).Value
        OpenPrice = Cells(2,3).Value
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        'iterate through each row to check if ticker of each row is the same as ResultTicker
        For i = 2 to LastRow
            RowTicker = Cells(i,1).Value
            RowVolume = Cells(i,7).Value
            RowDate = Cells(i,2).Value
    
            'If current row ticker is the same as ResultTicker add volume to total
            If RowTicker = ResultTicker Then
                TickerTotal = TickerTotal + RowVolume
                Cells(ResultRow, 10).Value = ResultTicker
                Cells(ResultRow, 13).Value = TickerTotal
                
                'Find last day and closing price, calculate yrchange and pctchange
                If RowDate > FirstDay Then

                    LastDay = Cells(i,2).Value
                    ClosePrice = Cells(i,6).Value
                    YrChange = ClosePrice - OpenPrice
                    Cells(ResultRow, 11).Value = YrChange
                    
                    'if the open price was 0, do not calculate pctchange
                    If OpenPrice = 0 Then
                        Cells(ResultRow, 12).Value = "N/A"
                    Else
                        PctChange = YrChange / OpenPrice * 100
                        Cells(ResultRow, 12).Value = PctChange
                    End If

                    'Change color formatting 
                    If YrChange >= 0 Then
                        Cells(ResultRow, 11).Interior.ColorIndex = 4
                    Else
                        Cells(ResultRow, 11).Interior.ColorIndex = 3
                    End If

                End If

            'if not same, go to next row in result table
            Else 
                ResultRow = ResultRow + 1
                
            End If

            'if the next row ticker is different from result ticker, reset values
            If Cells(i + 1, 1).Value <> ResultTicker Then
                ResultTicker = RowTicker
                TickerTotal = 0
                FirstDay = Cells(i, 2).Value
                OpenPrice = Cells(i, 3).Value

            End if

        Next i

        'Declare variables for determining greatest total volume/%increase/%decrease
        Dim GrtIncrease as Variant
        Dim GrtDecrease as Variant
        Dim GrtVolume as Double
        Dim LastResultRow as Long
        Dim PercentChange as Variant
        Dim TotalVolume as Double
        Dim SummaryTkr as String

        'initialize variables
        GrtIncrease = 0
        GrtDecrease = 0
        GrtVolume = 0
        LastResultRow = Cells(Rows.Count, 10).End(xlUp).Row
        
        'iterate through rows of summary table to find greatest values
        For j = 2 to LastResultRow

            SummaryTkr = Cells(j, 10).Value
            PercentChange = Cells(j, 12).Value

            If PercentChange > GrtIncrease Then
                GrtIncrease = PercentChange
                Cells(2, 16).Value = SummaryTkr
                Cells(2, 17).Value = GrtIncrease
            ElseIf PercentChange < GrtDecrease Then 
                GrtDecrease = PercentChange
                Cells(3, 16).Value = SummaryTkr
                Cells(3, 17).Value = GrtDecrease
            End If

            TotalVolume = Cells(j, 13).Value

            If TotalVolume > GrtVolume Then 
                GrtVolume = TotalVolume
                Cells(4, 16).Value = SummaryTkr
                Cells(4, 17).Value = GrtVolume
            End If

        Next j
    
    Next

End Sub