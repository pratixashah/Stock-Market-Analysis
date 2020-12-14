Sub Stock_Summary()

Dim totalRows As Long
Dim rowNumer As Long

Dim openPrice As Double
Dim closePrice As Double
Dim totalStock As Double

'Total no. of rows of data
totalRows = Cells(Rows.Count, 1).End(xlUp).Row

'Row no. from where it starts to print summary
'1st Row is for Header so starts with 2nd
rowNumber = 2

'To display Header
Range("I1").Value = "<ticker>"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"

'Loop from 2nd rows to total no. of rows
'First row has Header so starts with 2nd row
For i = 2 To totalRows

    'To get first open price from first row
    If i = 2 Then
        openPrice = Cells(i, 3).Value
    End If
        
    'To get Total Stock Volume
    totalStock = totalStock + Cells(i, 7).Value
       
    'For New Ticker value
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
               
        'To get close price from last row of same Ticker value
        closePrice = Cells(i, 6).Value
                
        'To print Ticker, Yearly Change, Percentage Change, Total Stock Volume values
        Range("I" & rowNumber).Value = Cells(i, 1).Value '& openPrice & " " & closePrice & "="
        Range("J" & rowNumber).Value = openPrice - closePrice
        Range("K" & rowNumber).Value = (Range("J" & rowNumber).Value / openPrice) * 100 & "%"
        Range("L" & rowNumber).Value = totalStock
                
        'To get next row number to print
        rowNumber = rowNumber + 1
        
        'To get open price for next Ticker
        openPrice = Cells(i + 1, 3).Value
        
        'Reset Total Stock value for next Ticker
        totalStock = 0
    End If
    
Next i

End Sub

