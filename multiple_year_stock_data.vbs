
Sub StockData()

    'Create the variables
        Dim YearOpening As Single
        Dim YearClosing As Single
        Dim ws As Worksheet
        Dim SelectIndex As Double
        Dim Row As Integer
        Dim LastRow As Long
        Dim Volume As Double
     
        'Loop through all sheets
            For Each ws In Sheets
            Worksheets(ws.Name).Activate
            SelectIndex = 2
            Row = 2
        'Determine the last row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Volume = 0
        
    'Assign headers to columns and rows
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
    'Loop through all the tickers to find the different tickers and add them to the summary table
        For i = 2 To LastRow
        
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            Cells(Row, 9).Value = Cells(i, 1).Value
            Row = Row + 1
            End If
        
        Next i
    
    'Loop through all the tickers to find the same tickers and add the total volumes to summary table
        For i = 2 To LastRow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Volume = Volume + Cells(i, 7).Value
            Cells(SelectIndex, 12).Value = Volume
                SelectIndex = SelectIndex + 1
                Volume = 0
  
        Else
                Volume = Volume + Cells(i, 7).Value
            End If
            
        Next i
        

          
    ' Loop to find opening and closing price
        SelectIndex = 2
        For i = 2 To LastRow
     
        'If next ticker is different this is closing
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            YearClosing = Cells(i, 6).Value
    
        'If previous ticker is different this is opening
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            YearOpening = Cells(i, 3).Value
    
            End If
    
        'Calculate yearly percentage change and format it
            If YearOpening > 0 And YearClosing > 0 Then
            Increase = YearClosing - YearOpening
            per_increase = Increase / YearOpening
            Cells(SelectIndex, 10).Value = Increase
            Cells(SelectIndex, 11).Value = FormatPercent(per_increase)
            YearClosing = 0
            YearOpening = 0
            SelectIndex = SelectIndex + 1
    
            End If
    
    Next i
    
    'Find min and max values and add them to the summary table
        MaxPercentage = WorksheetFunction.Max(ActiveSheet.Columns("K"))
        MinPercentage = WorksheetFunction.Min(ActiveSheet.Columns("K"))
        MaxVolume = WorksheetFunction.Max(ActiveSheet.Columns("L"))
        
        Range("Q2").Value = FormatPercent(MaxPercentage)
        Range("Q3").Value = FormatPercent(MinPercentage)
        Range("Q4").Value = MaxVolume
        
        
    'Loops through columns K and L to find the max, min percentages and max volume

        For i = 2 To LastRow
    
            If MaxPercentage = Cells(i, 11).Value Then
            Range("P2").Value = Cells(i, 9).Value
                
            ElseIf MinPercentage = Cells(i, 11).Value Then
            Range("P3").Value = Cells(i, 9).Value
    
            ElseIf MaxVolume = Cells(i, 12).Value Then
            Range("P4").Value = Cells(i, 9).Value
    
            End If
    
        Next i
    
    'Conditional formating for the yearly change column

    For i = 2 To LastRow

        If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        Else
        Cells(i, 10).Interior.ColorIndex = 3
        End If
    
    Next i

 'Conditional formating for the percentage change column

    For i = 2 To LastRow

        If Cells(i, 11).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
        Else
        Cells(i, 11).Interior.ColorIndex = 3
        End If
    
    Next i
Next ws
    
End Sub
