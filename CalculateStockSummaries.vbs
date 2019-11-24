Attribute VB_Name = "Module1"
Sub CalculateSummary()
    Dim ws As Worksheet
    
    Dim i As Long
    Dim tickerName As String
    Dim totalStockVolume As LongLong
    Dim maxRow As Long
    
    Dim summaryTickerCol As Integer
    Dim summaryYearlyChgCol As Integer
    Dim summaryPctChgCol As Integer
    Dim summaryRow As Integer
    
    Dim totalStockVolCol As Integer
    Dim tickerCount As Long
    Dim tickerOpenValue As Double
    Dim tickerCloseValue As Double
        
    summaryTickerCol = 9
    summaryYearlyChgCol = 10
    summaryPctChgCol = 11
    totalStockVolCol = 12
    summaryRow = 2
    
    For Each ws In Worksheets
        'Calculate Header for Summary
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        Debug.Print ws.Name
        
        tickerName = " "
        tickerCount = 0
        totalStockVolume = 0
        summaryRow = 2
       
        maxRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Debug.Print maxRow
        
        For i = 2 To maxRow
            If tickerCount = 0 Then
                'Store Opening Stock Value for ticker
                tickerOpenValue = ws.Cells(i, 3).Value
                tickerName = ws.Cells(i, 1).Value
            End If
        
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                'Tickers are the same
                tickerCount = tickerCount + 1
            Else
                'Tickers are different
                'Write out summary line for current ticker symbol
                tickerCloseValue = ws.Cells(i, 6)
                ws.Cells(summaryRow, summaryTickerCol).Value = tickerName
                ws.Cells(summaryRow, summaryYearlyChgCol).Value = tickerCloseValue - tickerOpenValue
                If tickerOpenValue = 0 Then
                    ws.Cells(summaryRow, summaryPctChgCol).Value = 0
                Else
                    ws.Cells(summaryRow, summaryPctChgCol).Value = (tickerCloseValue / tickerOpenValue) - 1
                End If
                ws.Cells(summaryRow, totalStockVolCol).Value = totalStockVolume
            
                'Change yearly change cell color, red for negative, green for >= 0
                If ws.Cells(summaryRow, summaryYearlyChgCol).Value < 0 Then
                    ws.Cells(summaryRow, summaryYearlyChgCol).Interior.Color = vbRed
                Else
                    ws.Cells(summaryRow, summaryYearlyChgCol).Interior.Color = vbGreen
                End If
            
                summaryRow = summaryRow + 1
            
                'Reset summary values for next ticker
                tickerName = " "
                tickerCount = 0
                totalStockVolume = 0
                tickerCloseValue = 0
                tickerOpenValue = 0
            
            End If
            
        Next i
    
        'Format summary columns and header rows
        ws.Columns("I:L").AutoFit
        ws.Range("J2:J" & summaryRow).NumberFormat = "$#,##0.00_);($#,##0.00)"
        ws.Range("K2:K" & summaryRow).NumberFormat = "0.00%"
        ws.Range("L2:L" & summaryRow).NumberFormat = "#,##0"

    Next ws

End Sub
Sub CalculateGreatests()
    Dim ws As Worksheet
    
    Dim titleColNum As Integer
    Dim tickerColNum As Integer
    Dim valueColNum As Integer
    
    Dim greatPctIncRow As Integer
    Dim greatPctDecRow As Integer
    Dim greatVolRow As Integer
    
    Dim pctRange As String
    Dim volRange As String
    
    Dim maxRow As Long
    Dim foundCell As Range
    Dim i As Long
    Dim maxPctRow As Long
    Dim minPctRow As Long
    Dim maxVolRow As Long
    
    Dim summaryTickerCol As Integer
    Dim summaryPctChgCol As Integer
    Dim totalStockVolCol As Integer
    
    'Set Column Numbers to pull percent and volume
    summaryTickerCol = 9
    summaryPctChgCol = 11
    totalStockVolCol = 12

    'Set Column numbers for greatest table
    titleColNum = 15
    tickerColNum = 16
    valueColNum = 17
    
    'Set Row numbers for greatest table
    greatPctIncRow = 2
    greatPctDecRow = 3
    greatVolRow = 4
    
    For Each ws In Worksheets
        'Set Titles for Greatest Values
        ws.Cells(greatPctIncRow, titleColNum).Value = "Greatest % Increase"
        ws.Cells(greatPctDecRow, titleColNum).Value = "Greatest % Decrease"
        ws.Cells(greatVolRow, titleColNum).Value = "Greatest Total Volume"
        
        'Set Column names for Greatest Values
        ws.Cells(1, titleColNum).Value = "Highlights/Lowlights: "
        ws.Cells(1, tickerColNum).Value = "Ticker"
        ws.Cells(1, valueColNum).Value = "Value"
        
        maxRow = ws.Range("I1").End(xlDown).Row
        indexRange = "I2:I" & maxRow
        pctRange = "K2:K" & maxRow
        volRange = "L2:L" & maxRow
        
        maxPctValue = WorksheetFunction.Max(ws.Range(pctRange))
        minPctValue = WorksheetFunction.Min(ws.Range(pctRange))
        maxVolValue = WorksheetFunction.Max(ws.Range(volRange))
        
        maxPctRow = 0
        minPctRow = 0
        maxVolRow = 0
        
        For i = 2 To maxRow
            If ws.Cells(i, summaryPctChgCol).Value = maxPctValue Then
                maxPctRow = i
            ElseIf ws.Cells(i, summaryPctChgCol).Value = minPctValue Then
                minPctRow = i
            End If
            If ws.Cells(i, totalStockVolCol) = maxVolValue Then
                maxVolRow = i
            End If
        Next i
            
        ws.Cells(greatPctIncRow, tickerColNum).Value = ws.Cells(maxPctRow, summaryTickerCol)
        ws.Cells(greatPctIncRow, valueColNum).Value = maxPctValue
        
        ws.Cells(greatPctDecRow, tickerColNum).Value = ws.Cells(minPctRow, summaryTickerCol)
        ws.Cells(greatPctDecRow, valueColNum).Value = minPctValue
        
        ws.Cells(greatVolRow, tickerColNum).Value = ws.Cells(maxVolRow, summaryTickerCol)
        ws.Cells(greatVolRow, valueColNum).Value = maxVolValue
        
        'Adjust Column Width for Greatest Values
        ws.Columns("O:Q").AutoFit
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4:Q4").NumberFormat = "#,##0"
    
    Next
End Sub

Sub RunSummaries()
    Call CalculateSummary
    Call CalculateGreatests
End Sub
