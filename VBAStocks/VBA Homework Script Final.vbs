Attribute VB_Name = "Module1"
Sub Stocks()

'Loop through each worksheet
For Each ws In Worksheets

    'Variable declaration to gather all pertinent information for homework
    Dim Ticker As String
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    
    Dim Total_volume As LongLong
    Total_volume = 0
    
    'Tally used for referencing Opening_Price relative to line i when Closing Price determined
    Dim TickerCount As Long
    TickerCount = 0
    
    'Used to determine placeholder of Ticker Summary info
    Dim TargetRow As Integer
    TargetRow = 2
    
    'Finds the last row for each worksheet
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Inserting headers/rows for summary tables
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    'Check every row calculate total volume for a ticker and difference between opening and closing price after a year
    For i = 2 To LastRow
        'Conditional to trigger ticker summary info when reaching that tickers last line - compares between existing cell and cell below to see if ticker has changed
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker = ws.Cells(i, 1).Value
            
            Opening_Price = ws.Cells((i) - TickerCount, 3).Value
            
            Closing_Price = ws.Cells(i, 6)
            
            Yearly_Change = Closing_Price - Opening_Price
        
            Total_volume = Total_volume + ws.Cells(i, 7).Value
        
            ws.Range("I" & TargetRow).Value = Ticker
            
            ws.Range("J" & TargetRow).Value = Yearly_Change
                'Conditonal Formatting for Yearly Change if positive (green) or negative (red)
                If ws.Range("J" & TargetRow).Value < 0 Then
                    ws.Range("J" & TargetRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & TargetRow).Interior.ColorIndex = 4
                End If
                'Address tickers with 0 as opening price causing Div 0 error for Percentage calculation
                If Opening_Price <> 0 Then
                    Percentage_Change = (Yearly_Change / Opening_Price)
                Else
                    Percentage_Change = 0
                End If
            
            ws.Range("K" & TargetRow).Value = Percentage_Change
        
            ws.Range("L" & TargetRow).Value = Total_volume
            'Adjust placement of next ticker for the summary table
            TargetRow = TargetRow + 1
            'reset volume and tickercounts for referencing next Ticker
            Total_volume = 0
            
            TickerCount = 0
        
        Else
            
            Total_volume = Total_volume + ws.Cells(i, 7).Value
            TickerCount = TickerCount + 1
            
        End If
    
    Next i
    
    'Determine Last Row in the summary table - arbitrarily chose to use Percent column
    Dim LastRow_Percentage As Integer
    LastRow_Percentage = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'Format the Yearly Change # and % cells
    For i = 2 To LastRow_Percentage
        
        ws.Range("J" & i).NumberFormat = "$#,##0.00"
        ws.Range("K" & i).NumberFormat = "0.00%"
        
    Next i
    
    'Challenge portion of homework to determine greatest/lowest percent change and largest volume - separate loops scanning summary table
    Dim MaxChangeTicker As String
    Dim Max As Double
    
    Max = 0
    'Loop through Percent Change rows to find greatest value/corresponding ticker
    For i = 2 To LastRow_Percentage
        
        If ws.Cells(i, 11).Value > Max Then
            
            Max = ws.Cells(i, 11).Value
            
            MaxChangeTicker = ws.Cells(i, 9).Value
        
        End If
    
    Next i
    'Greatest Percent Increase information placement
    ws.Cells(2, 16).Value = MaxChangeTicker
    ws.Cells(2, 17).Value = Max
    
    Dim MinChangeTicker As String
    Dim Min As Double
    
    Min = 0
    'Loop through Percent Change rows to find least value/corresponding ticker
    For i = 2 To LastRow_Percentage
        
        If ws.Cells(i, 11).Value < Min Then
            
            Min = ws.Cells(i, 11).Value
            
            MinChangeTicker = ws.Cells(i, 9).Value
        
        End If
    
    Next i
    'Greatest Percent Decrease info placement
    ws.Cells(3, 16).Value = MinChangeTicker
    ws.Cells(3, 17).Value = Min
    
    Dim MaxVolTicker As String
    Dim MaxVol As LongLong
    
    MaxVol = 0
    'Loop through Total Volume rows to find greatest total volume/corresponding ticker
    For i = 2 To LastRow_Percentage
        
        If ws.Cells(i, 12).Value > MaxVol Then
            
            MaxVol = ws.Cells(i, 12).Value
            
            MaxVolTicker = ws.Cells(i, 9).Value
        
        End If
    
    Next i
    'Largest Volume info placement
    ws.Cells(4, 16).Value = MaxVolTicker
    ws.Cells(4, 17).Value = MaxVol
    
  
Next ws

End Sub

