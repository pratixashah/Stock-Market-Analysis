Sub Stock_Summary()

Dim totalRows As Long
Dim rowNumer As Long

Dim openPrice As Double
Dim closePrice As Double
Dim totalStock As Long


totalRows = Cells(Rows.Count, 1).End(xlUp).Row
rowNumer = 2

openPrice = 0

'MsgBox (lastRow)

Range("I1").Value = "<ticker>"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"

For i = 2 To totalRows - 1

    If i = 2 Then
            openPrice = Cells(i, 3).Value
            totalStock = Cells(i, 7).Value
        End If
        
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
               
        closePrice = Cells(i, 6).Value
        totalStock = 0
                
        Range("I" & rowNumer).Value = Cells(i, 1).Value & openPrice & " " & closePrice & "="
        Range("J" & rowNumer).Value = openPrice - closePrice
        Range("K" & rowNumer).Value = (Range("J" & rowNumer).Value / openPrice) * 100 & "%"
        Range("L" & rowNumer).Value = totalStock
                
        rowNumer = rowNumer + 1
        openPrice = Cells(i + 1, 3).Value
        
    Else
        'totalStock = totalStock + Cells(i, 7).Value
    End If
    

Next i

End Sub
