Attribute VB_Name = "Module1"
Sub AlphabetSorting():

For Each ws In Worksheets

Dim worksheetname As String
Dim i As Long
Dim j As Long
Dim tickcount As Long
Dim lastrow1 As Long
Dim lastrow2 As Long
Dim percentcalc As Double
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As Double

worksheetname = ws.Name

ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"

ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest Total Volume"

tickcount = 2

j = 2

lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To lastrow1

If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
ws.Cells(tickcount, 9) = ws.Cells(i, 1)
ws.Cells(tickcount, 10) = ws.Cells(i, 6) - ws.Cells(j, 3)

If ws.Cells(tickcount, 10) < 0 Then
ws.Cells(tickcount, 10).Interior.ColorIndex = 3

Else
ws.Cells(tickcount, 10).Interior.ColorIndex = 4

End If

If ws.Cells(j, 3) <> 0 Then
percentcalc = ((ws.Cells(i, 6) - ws.Cells(j, 3)) / ws.Cells(j, 3))

ws.Cells(tickcount, 11) = Format(percentcalc, "percent")


Else
ws.Cells(tickcount, 11) = Format(0, "percent")

End If

ws.Cells(tickcount, 12) = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))


tickcount = tickcount + 1

j = i + 1

End If

Next i


lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

greatestvolume = ws.Cells(2, 12)
greatestincrease = ws.Cells(2, 11)
greatestdecrease = ws.Cells(2, 11)
        

For i = 2 To lastrow2
            

If ws.Cells(i, 12) > greatestvolume Then
greatestvolume = ws.Cells(i, 12)
ws.Cells(4, 16) = ws.Cells(i, 9)
                                
    
ElseIf ws.Cells(i, 11) > greatestincrease Then
greatestincrease = ws.Cells(i, 11)
ws.Cells(2, 16) = ws.Cells(i, 9)
                
                                
ElseIf ws.Cells(i, 11) < greatestdecrease Then
greatestdecrease = ws.Cells(i, 11)
ws.Cells(3, 16) = ws.Cells(i, 9)
                

End If
                                
ws.Cells(2, 17) = Format(greatestincrease, "Percent")
ws.Cells(3, 17) = Format(greatestdecrease, "Percent")
ws.Cells(4, 17) = Format(greatestvolume, "Scientific")
            
Next i
            
            
Next ws
    
End Sub
