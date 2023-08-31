Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim stockVolume As Double
    Dim greatestPI As Double
    Dim greatesPD As Double
    Dim greatestVolume As Double
    

    Dim outputRow As Long
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
                
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables
        ticker = ws.Cells(2, "A").Value
        openingPrice = ws.Cells(2, "C").Value
        closingPrice = ws.Cells(2, "F").Value
        yearlyChange = closingPrice - openingPrice
        percentChange = yearlyChange / openingPrice
        stockVolume = ws.Cells(2, "G").Value
        outputRow = 2
        
        ' Output headers starting from row 1 in columns I through L
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        
        ' Loop through each row starting from row 2
        For i = 2 To lastRow
            ' Check if the year has changed
            If ws.Cells(i, "A").Value <> ticker Then
                ' Output the information for the previous stock
                ws.Cells(outputRow, "I").Value = ticker
                ws.Cells(outputRow, "J").Value = yearlyChange
                ws.Cells(outputRow, "K").Value = percentChange
                ws.Cells(outputRow, "L").Value = stockVolume
                
                ' Reset variables for the new stock
                ticker = ws.Cells(i, "A").Value
                openingPrice = ws.Cells(i, "C").Value
                closingPrice = ws.Cells(i, "F").Value
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
                stockVolume = ws.Cells(i, "G").Value
                
                
                ' Move to the next output row
                outputRow = outputRow + 1
            Else
                ' Accumulate the stock volume for the same year and symbol
                stockVolume = stockVolume + ws.Cells(i, "G").Value
                
                ' Update the closing price for the current stock
                closingPrice = ws.Cells(i, "F").Value
                
                ' Update the yearly change and percent change
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
            End If
        Next i
        
        ' Output the information for the last stock
        ws.Cells(outputRow, "I").Value = ticker
        ws.Cells(outputRow, "J").Value = yearlyChange
        ws.Cells(outputRow, "K").Value = percentChange
        ws.Cells(outputRow, "L").Value = stockVolume
        
        ' Format percentage row
        ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
        
        
        'Colorcode columns
        For j = 2 To lastRow
        'Check if value is negative
            If ws.Cells(j, "J").Value < 0 Then
            'Set the color index to red
            ws.Cells(j, "J").Interior.ColorIndex = 3
            'Check if value is positive
            ElseIf ws.Cells(j, "J").Value > 0 Then
            'Set the color index to green
            ws.Cells(j, "J").Interior.ColorIndex = 4
            'Check if value is zero
            ElseIf ws.Cells(j, "J").Value = 0 Then
            ws.Cells(j, "J").Interior.ColorIndex = 2
        
        End If
        Next j
        
     'Remove gridline
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    
    ' Output headers
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Volume"
        
    'Check for greatest % increase
        greatestPI = Cells(2, "K").Value
        For Each cell In ws.Range("K2:K" & lastRow)
        If cell.Value > greatestPI Then
        greatestPI = cell.Value
        End If
        Next cell
    
    ' Output for greatest % increase
        ws.Cells(2, "Q").Value = greatestPI
        ws.Cells(2, "P").Value = ticker
        'Format
        ws.Cells(2, "Q").NumberFormat = "0.00%"
        
    'Check for greatest % decrease
        ticker = ws.Cells(2, "A").Value
        greatestPD = Cells(2, "K").Value
        For Each cell In ws.Range("K2:K" & lastRow)
        If cell.Value < greatestPD Then
        greatestPD = cell.Value
        End If
        Next cell
    
    ' Output for greatest % decrease
        ws.Cells(3, "Q").Value = greatestPD
        ws.Cells(3, "P").Value = ticker
        'Format
        ws.Cells(3, "Q").NumberFormat = "0.00%"
        
       'Check for greatest volume
       ticker = ws.Cells(2, "A").Value
        greatestVolume = Cells(2, "L").Value
        For Each cell In ws.Range("L2:L" & lastRow)
        If cell.Value > greatestVolume Then
        greatestVolume = cell.Value
        End If
        Next cell
    
    ' Output for greatest volume
        ws.Cells(4, "Q").Value = greatestVolume
        ws.Cells(4, "P").Value = ticker
        
     'Autofit columns
        ws.Columns("I:Q").AutoFit
      
    Next ws
End Sub

