Sub Stock_Summary()

'To get total number of rows
Dim totalRows As Long

'To get row number for where to print
Dim rowNumer As Long

'Variable for open Price, close price and total stocks
Dim openPrice As Double
Dim closePrice As Double
Dim totalStock As Double

'To get Greatest increase, Greatest decrease and Total volume of stocks with its Ticket resp.
Dim maxIncrease As Double
Dim maxIncreaseTicker As String

Dim maxDecrease As Double
Dim maxDecreaseTicker As String

Dim maxTotalVolume As Double
Dim maxTotalVolumeTicker As String

'Loop for each worksheet one by one
For Each ws In Worksheets

    'Reset variables to 0
    openPrice = 0
    closePrice = 0
    totalStock = 0
    maxIncrease = 0
    maxDecrease = 0
    maxTotalVolume = 0
    
    'Total no. of rows in worksheet
    totalRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'To display Header for Summary
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Row no. from where it starts to print summary
    '1st Row is for Header so starts with 2nd
    rowNumber = 2
    
    'Loop from 2nd rows to total no. of rows
    'First row has Header so starts with 2nd row
    For i = 2 To totalRows
    
        'To get first open price from first row
        If i = 2 Then
            openPrice = ws.Cells(i, 3).Value
        End If
            
        'To get Total Stock Volume
        totalStock = totalStock + ws.Cells(i, 7).Value
           
        'For next New Ticker value
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                   
            'To get close price from last row of same Ticker value
            closePrice = ws.Cells(i, 6).Value
                    
            'To print Ticker, Yearly Change, Percentage Change, Total Stock Volume values
            ws.Range("I" & rowNumber).Value = ws.Cells(i, 1).Value
            ws.Range("J" & rowNumber).Value = closePrice - openPrice
            
            'To avoid divide by zero error
            If (openPrice > 0) Then
                ws.Range("K" & rowNumber).Value = (ws.Range("J" & rowNumber).Value / openPrice) * 100 & "%"
            Else
                ws.Range("K" & rowNumber).Value = "NA"
            End If
            
            ws.Range("L" & rowNumber).Value = totalStock
                    
            'Conditional formatting - positive increase display with Green color and negative with Red color
            If (ws.Range("J" & rowNumber).Value > 0) Then
                ws.Range("J" & rowNumber).Interior.ColorIndex = 4
            ElseIf (ws.Range("J" & rowNumber).Value < 0) Then
                ws.Range("J" & rowNumber).Interior.ColorIndex = 3
            End If
            
            'To find Greatest increase, decrease in Stocks with its Ticker resp.
            If (openPrice > 0) Then
                If ws.Range("K" & rowNumber).Value > maxIncrease Then
                    maxIncrease = ws.Range("K" & rowNumber).Value
                    maxIncreaseTicker = ws.Cells(i, 1).Value
                ElseIf ws.Range("K" & rowNumber).Value < maxDecrease Then
                    maxDecrease = ws.Range("K" & rowNumber).Value
                    maxDecreaseTicker = ws.Cells(i, 1).Value
                End If
             End If
             
            'To find Greatest Total Volume in Stocks with its Ticker resp.
            If ws.Range("L" & rowNumber).Value > maxTotalVolume Then
                maxTotalVolume = ws.Range("L" & rowNumber).Value
                maxTotalVolumeTicker = ws.Cells(i, 1).Value
            End If
            
            'To get next row number to print
            rowNumber = rowNumber + 1
            
            'To get open price for next Ticker
            openPrice = ws.Cells(i + 1, 3).Value
            
            'Reset Total Stock value for next Ticker
            totalStock = 0
        End If
        
    Next i
    
    'To display Header for Greatest increase, decrease and Total volume of stocks
    ws.Range("N2").Value = "Greatest % increase"
    ws.Range("N3").Value = "Greatest % decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'To set values
    ws.Range("O2").Value = maxIncreaseTicker
    ws.Range("O3").Value = maxDecreaseTicker
    ws.Range("O4").Value = maxTotalVolumeTicker
    
    ws.Range("P2").Value = maxIncrease * 100 & "%"
    ws.Range("P3").Value = maxDecrease * 100 & "%"
    ws.Range("P4").Value = maxTotalVolume

Next ws


End Sub




