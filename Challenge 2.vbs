Sub Multiple_year_stock_data()
    
    
    For Each ws In ThisWorkbook.Worksheets
    With ws
        
        
        
        
        
    Dim count As Integer
    count = 2
    Dim i As Long
    lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
   
   Dim openprice As Double
   Dim closeprice As Double
   Dim yearlychange As Double
   Dim percentchange As Double
   Dim totalstockvolume As Double
   
   openprice = ws.Cells(2, 3).Value
   
    For i = 2 To lastrow
        
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
        
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


        closeprice = ws.Cells(i, 6).Value
        yearlychange = closeprice - openprice
        percentchange = yearlychange / openprice
        


        ' PRINT TICKER
        ws.Cells(count, 9).Value = ws.Cells(i, 1).Value
        
        ' PRINT PERCENT CHANGE
        ws.Cells(count, 11).Value = FormatPercent(percentchange)
        
        ' PRINT TOTAL STOCK VOLUME
        ws.Cells(count, 12).Value = totalstockvolume
        
        ' PRINT CHANGE
        ws.Cells(count, 10).Value = yearlychange
        If yearlychange >= 0 Then
        ws.Cells(count, 10).Interior.ColorIndex = 4
        ElseIf yearlychange < 0 Then
        ws.Cells(count, 10).Interior.ColorIndex = 3
        End If
        ' END IF FOR COLOR CHANGE
        

        ' set NEXT open price && RESET VARIABLES
        openprice = ws.Cells(i + 1, 3).Value
        totalstockvolume = 0
        count = count + 1
        
        
    End If
        
    Next i
    
    
    Dim j As Long
    Dim greatestpercentincrease As Double
    Dim greatestpercentdecrease As Double
    Dim greatesttotalstockvolume As Double
    
    
    Dim greatestpercentincreasetick As String
    Dim greatestpercentdecreasetick As String
    Dim greatesttotalstockvolumetick As String
    
    ' CHECK COLUMN FOR GREATEST
    For j = 2 To lastrow
        
        ' If this row's variable is larger than the current largest value Then
        ' overwrite the variable with this row's value
    
        If ws.Cells(j, 11).Value > greatestpercentincrease Then
        greatestpercentincrease = ws.Cells(j, 11).Value
        greatestpercentincreasetick = ws.Cells(j, 9).Value
        End If
        
        If ws.Cells(j, 11).Value < greatestpercentdecrease Then
        greatestpercentdecrease = ws.Cells(j, 11).Value
        greatestpercentdecreasetick = ws.Cells(j, 9).Value
        End If
        
        If ws.Cells(j, 12).Value > greatesttotalstockvolume Then
        greatesttotalstockvolume = ws.Cells(j, 12).Value
        greatesttotalstockvolumetick = ws.Cells(j, 9).Value
        End If
    
    Next j
    
    
    ' PRINT HEADERS
    ws.Cells(2, 15).Value = "Greatest Percent Increase"
    ws.Cells(3, 15).Value = "Greatest Percent Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    
        ' PRINT GREASTPERCENTINCREASE & TICKER
        ws.Cells(2, 16).Value = greatestpercentincreasetick
        ws.Cells(2, 17).Value = FormatPercent(greatestpercentincrease)
        
        ' PRINT greatestpercentdecrease & TICKER
        ws.Cells(3, 16).Value = greatestpercentdecreasetick
        ws.Cells(3, 17).Value = FormatPercent(greatestpercentdecrease)
        
        ' PRINT greatesttotalstockvolume & TICKER
        ws.Cells(4, 16).Value = greatesttotalstockvolumetick
        ws.Cells(4, 17).Value = greatesttotalstockvolume
    
    
     End With
    Next ws
    
    
    
    
   
   
   
   

End Sub

