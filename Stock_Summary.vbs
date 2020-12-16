Sub Stock_Summary()

'To get total number of rows
Dim totalRows As Long

'To get row number for where to print
Dim rowNumer As Long

Dim openPrice As Double
Dim closePrice As Double
Dim totalStock As Double

'To get Greatest increase, decrease and Total volume of stocks with its Ticket resp.
Dim maxIncreaseTicker As String
Dim maxDecreaseTicker As String
Dim maxTotalVolumeTicker As String

Dim maxIncrease As Long
Dim maxDecrease As Long
Dim maxTotalVolume As Double

'Total no. of rows of data
totalRows = Cells(Rows.Count, 1).End(xlUp).Row

'To display Header for Summary
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"

'Row no. from where it starts to print summary
'1st Row is for Header so starts with 2nd
rowNumber = 2

'Loop from 2nd rows to total no. of rows
'First row has Header so starts with 2nd row
For i = 2 To totalRows

    'To get first open price from first row
    If i = 2 Then
        openPrice = Cells(i, 3).Value
    End If
        
    'To get Total Stock Volume
    totalStock = totalStock + Cells(i, 7).Value
       
    'For next New Ticker value
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
               
        'To get close price from last row of same Ticker value
        closePrice = Cells(i, 6).Value
                
        'To print Ticker, Yearly Change, Percentage Change, Total Stock Volume values
        Range("I" & rowNumber).Value = Cells(i, 1).Value '& openPrice & " " & closePrice & "="
        Range("J" & rowNumber).Value = openPrice - closePrice
        Range("K" & rowNumber).Value = (Range("J" & rowNumber).Value / openPrice) * 100 & "%"
        Range("L" & rowNumber).Value = totalStock
                
        'Conditional formatting - positive increase display with Green color and negative with Red color
        If (Range("J" & rowNumber).Value > 0) Then
            Range("J" & rowNumber).Interior.ColorIndex = 4
        ElseIf (Range("J" & rowNumber).Value < 0) Then
            Range("J" & rowNumber).Interior.ColorIndex = 3
        End If
        
        'To find Greatest increase, decrease in Stocks with its Ticker resp.
        If Range("J" & rowNumber).Value > maxIncrease Then
            maxIncrease = Range("J" & rowNumber).Value
            maxIncreaseTicker = Cells(i, 1).Value
        ElseIf Range("J" & rowNumber).Value < maxDecrease Then
            maxDecrease = Range("J" & rowNumber).Value
            maxDecreaseTicker = Cells(i, 1).Value
        End If
            
        'To find Greatest Total Volume in Stocks with its Ticker resp.
        If Range("L" & rowNumber).Value > maxTotalVolume Then
            maxTotalVolume = Range("L" & rowNumber).Value
            maxTotalVolumeTicker = Cells(i, 1).Value
        End If
        
        'To get next row number to print
        rowNumber = rowNumber + 1
        
        'To get open price for next Ticker
        openPrice = Cells(i + 1, 3).Value
        
        'Reset Total Stock value for next Ticker
        totalStock = 0
    End If
    
Next i

'To display Header Greatest increase, decrease and Total volume of stocks
Range("N2").Value = "Greatest % increase"
Range("N3").Value = "Greatest % decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

'To set values
Range("O2").Value = maxIncreaseTicker
Range("O3").Value = maxDecreaseTicker
Range("O4").Value = maxTotalVolumeTicker

Range("P2").Value = maxIncrease
Range("P3").Value = maxDecrease
Range("P4").Value = maxTotalVolume

End Sub

