Sub StockSummary()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    
    ' Variables to store greatest percent increase, decrease, and total volume
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    
    ' Initialize variables
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        ' Initialize variables for each worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        summaryRow = 2 ' Starting row for summary table
        totalVolume = 0 ' Reset total volume
        
        ' Output headers for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Loop through each row in the worksheet
        For i = 2 To lastRow
            
            ' Check if current row is the first row of a new stock
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                ' If it's not the first stock, calculate and output the summary for the previous stock
                If i > 2 Then
                    ws.Cells(summaryRow, 9).Value = ticker
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    ws.Cells(summaryRow, 11).Value = percentChange
                    ws.Cells(summaryRow, 12).Value = totalVolume
                    summaryRow = summaryRow + 1 ' Move to next row for summary
                End If
                
                ' Reset variables for the new stock
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                totalVolume = 0 ' Reset total volume
                yearlyChange = 0 ' Reset yearly change
                percentChange = 0 ' Reset percent change
            End If
            
            ' Accumulate total volume for the stock
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Calculate yearly change
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = closingPrice - openingPrice
            
            ' Calculate percent change
            If openingPrice <> 0 Then
                percentChange = (yearlyChange / openingPrice) * 100
            Else
                percentChange = 0 ' To avoid division by zero
            End If
            
            ' Update greatest percent increase, decrease, and total volume
            If percentChange > greatestPercentIncrease Then
                greatestPercentIncrease = percentChange
                greatestPercentIncreaseTicker = ticker
            End If
            
            If percentChange < greatestPercentDecrease Then
                greatestPercentDecrease = percentChange
                greatestPercentDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestTotalVolume Then
                greatestTotalVolume = totalVolume
                greatestTotalVolumeTicker = ticker
            End If
            
        Next i
        
        ' Output summary for the last stock in the worksheet
        ws.Cells(summaryRow, 9).Value = ticker
        ws.Cells(summaryRow, 10).Value = yearlyChange
        ws.Cells(summaryRow, 11).Value = percentChange
        ws.Cells(summaryRow, 12).Value = totalVolume
        
        ' Apply conditional formatting
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        ' Apply conditional formatting to highlight negative yearly change values in column J (red)
        With ws.Range("J2:J" & lastRow).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(255, 0, 0) ' Set the background color to red for negative values
        End With
        
        ' Apply conditional formatting to highlight positive yearly change values in column J (green)
        With ws.Range("J2:J" & lastRow).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = RGB(0, 255, 0) ' Set the background color to green for positive values
        End With
        
        ' Output greatest percent increase, decrease, and total volume
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestPercentIncreaseTicker
        ws.Cells(2, 17).Value = greatestPercentIncrease
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestPercentDecreaseTicker
        ws.Cells(3, 17).Value = greatestPercentDecrease
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestTotalVolumeTicker
        ws.Cells(4, 17).Value = greatestTotalVolume
        
    Next ws

End Sub
