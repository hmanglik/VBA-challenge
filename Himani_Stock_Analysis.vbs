Sub Stock()

'Iterate through worksheets
For Each ws In Worksheets

    ' Set initial variables
    Dim Ticker_Symbol As String
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Price_Change As Double
    Dim Percent_Change As Double
    Dim Volume As Double
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Initial Opening Price
    Opening_Price = ws.Cells(2, 3).Value
    
    ' Find the last row in the data
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set the headers as "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "Ticker", "Value", "Greatest % Increase", "Greatest % Decrease" and "Greatest Total Volume"
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' Loop through all stocks
    For i = 2 To lastrow
    
        ' Check if opening price is non zero
        If Opening_Price = 0 Then
            GoTo skipthisiteration
        End If
        
        ' Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Set the Ticker Symbol name
            Ticker_Symbol = ws.Cells(i, 1).Value
            
            ' Set the Closing Price
            Closing_Price = ws.Cells(i, 6).Value
            
            ' Calculate the Price Change
            Price_Change = Closing_Price - Opening_Price
            
            ' Calculate the Percent Change
            Percent_Change = Price_Change / Opening_Price
            
            ' Format into Percent
            Percent = FormatPercent(Percent_Change, 2)
            
            ' Update the Opening Price
            Opening_Price = ws.Cells(i + 1, 3).Value
            
            ' Add to the Volume Total
            Volume = Volume + ws.Cells(i, 7).Value
            
            ' Print the Ticker Symbol, Price_Change, Percent Change, and Volume values in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            ws.Range("J" & Summary_Table_Row).Value = Price_Change
            ws.Range("K" & Summary_Table_Row).Value = Percent
            ws.Range("L" & Summary_Table_Row).Value = Volume
            
            ' Color Formatting in Column J
            If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the volume
            Volume = 0
            
            ' If the cell immediately following a row is the same ticker...
        Else
        
            ' Add to the Volume Total
            Volume = Volume + ws.Cells(i, 7).Value
        End If

skipthisiteration:
    Next i
    
    'Set initial variables and counters
    Dim Last_Row As Integer
    Dim Ticker As String
    Max = 0
    Min = 0
    Greatest_Volume = 0
    
    'Find last row of column 11
    Last_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row

    'Iterate from second row to last row
    For i = 2 To Last_Row

    'Find the greatest percent increase and corresponding ticker
        If ws.Cells(i, 11).Value > Max Then
            Max = ws.Cells(i, 11).Value
            Ticker = ws.Cells(i, 9).Value
            In_Percent = FormatPercent(Max, 2)
            
            'Input into Challenge table
            ws.Range("Q2").Value = In_Percent
            ws.Range("P2").Value = Ticker
            
        End If
        
        'Find the greatest percent decrease and corresponding ticker
        If ws.Cells(i, 11).Value < Min Then
            Min = ws.Cells(i, 11).Value
            Ticker = ws.Cells(i, 9).Value
            In_Percent = FormatPercent(Min, 2)
            
            'Input into Challenge table
            ws.Range("Q3").Value = In_Percent
            ws.Range("P3").Value = Ticker
        End If
        
        'Find the greatest total volume and corresponding ticker
        If ws.Cells(i, 12).Value > Greatest_Volume Then
            Greatest_Volume = ws.Cells(i, 12).Value
            Ticker = ws.Cells(i, 9).Value
            
            'Input into Challenge table
            ws.Range("Q4").Value = Greatest_Volume
            ws.Range("P4").Value = Ticker
            
        End If

    Next i

Next ws

End Sub



