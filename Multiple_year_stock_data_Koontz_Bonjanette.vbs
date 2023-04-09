Sub stock_challenge()
    Dim WS_Count As Integer
    Dim j As Integer
    
    'get number of sheets in workbook
    WS_Count = ActiveWorkbook.Worksheets.Count
    For j = 1 To WS_Count

        'Get length of worksheet
        Dim lastRow As Long
        lastRow = Worksheets(j).Cells(Rows.Count, "A").End(xlUp).Row + 1
    
        'insert row headings
        Worksheets(j).Cells(1, 9).Value = "Ticker"
        Worksheets(j).Cells(1, 10).Value = "Yearly Change"
        Worksheets(j).Cells(1, 11).Value = "Percent Change"
        Worksheets(j).Cells(1, 12).Value = "Total Stock Volume"
        Worksheets(j).Cells(1, 16).Value = "Ticker"
        Worksheets(j).Cells(1, 17).Value = "Value"
        Worksheets(j).Cells(2, 15).Value = "Greatest % Increase"
        Worksheets(j).Cells(3, 15).Value = "Greatest % Decrease"
        Worksheets(j).Cells(4, 15).Value = "Greatest Total Volume"
    
        Dim i As Long
        Dim ticker As String
        Dim beginningStockPrice As Double
        Dim endingStockPrice As Double
        Dim stockPriceChange As Double
        Dim stockPercentChange As Double
        Dim stockVolume As Double
        Dim dataTableIndex As Long
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim greatestIncreaseTicker As String
        Dim greatestDecreaseTicker As String
        Dim greatestVolumeTicker As String
        
        'initialize index for placing stock into table
        dataTableIndex = 2
    
        'get info for first stock on sheet
        ticker = Worksheets(j).Cells(2, 1).Value
        beginningStockPrice = Worksheets(j).Cells(2, 3).Value
        stockVolume = Worksheets(j).Cells(2, 7).Value
    
        For i = 3 To lastRow
            If ticker = Worksheets(j).Cells(i, 1).Value Then    'Check for new stock
                stockVolume = stockVolume + Worksheets(j).Cells(i, 7).Value
            Else
                endingStockPrice = Worksheets(j).Cells(i - 1, 6).Value
            
                'Calculate final values for stock
                stockPriceChange = endingStockPrice - beginningStockPrice
                stockPercentChange = stockPriceChange / beginningStockPrice
            
                'insert stock data into summary table
                Worksheets(j).Cells(dataTableIndex, 9).Value = ticker
                Worksheets(j).Cells(dataTableIndex, 10).Value = stockPriceChange
                Worksheets(j).Cells(dataTableIndex, 11).Value = stockPercentChange
                Worksheets(j).Cells(dataTableIndex, 12).Value = stockVolume
            
                'format change cells
                If stockPriceChange > 0 Then
                    Worksheets(j).Cells(dataTableIndex, 10).Interior.ColorIndex = 4
                    Worksheets(j).Cells(dataTableIndex, 11).Interior.ColorIndex = 4
                ElseIf stockPriceChange < 0 Then
                    Worksheets(j).Cells(dataTableIndex, 10).Interior.ColorIndex = 3
                    Worksheets(j).Cells(dataTableIndex, 11).Interior.ColorIndex = 3
                End If
            
                ' update dataTableIndex
                dataTableIndex = dataTableIndex + 1
            
                'update greatests
                If stockPercentChange > greatestIncrease Then
                    greatestIncrease = stockPercentChange
                    greatestIncreaseTicker = ticker
                ElseIf stockPercentChange < greatestDecrease Then
                    greatestDecrease = stockPercentChange
                    greatestDecreaseTicker = ticker
                End If
                If stockVolume > greatestVolume Then
                    greatestVolume = stockVolume
                    greatestVolumeTicker = ticker
                End If
            
                'store new ticker and beginning data
                If i <> lastRow Then
                    ticker = Worksheets(j).Cells(i, 1).Value
                    beginningStockPrice = Worksheets(j).Cells(i, 3).Value
                    stockVolume = Worksheets(j).Cells(i, 7).Value
                End If
            
            End If
        
        Next i
    
        'enter greatests into table
        Worksheets(j).Cells(2, 16).Value = greatestIncreaseTicker
        Worksheets(j).Cells(3, 16).Value = greatestDecreaseTicker
        Worksheets(j).Cells(4, 16).Value = greatestVolumeTicker
        Worksheets(j).Cells(2, 17).Value = greatestIncrease
        Worksheets(j).Cells(3, 17).Value = greatestDecrease
        Worksheets(j).Cells(4, 17).Value = greatestVolume
    
        'format
        Worksheets(j).Range("i1:Q1").EntireColumn.AutoFit
        Worksheets(j).Range("k:k").Style = "Percent"
        Worksheets(j).Range("k:k").NumberFormat = "0.00%"
        Worksheets(j).Range("q2:q3").Style = "Percent"
        Worksheets(j).Range("q2:q3").NumberFormat = "0.00%"
    Next j

End Sub
