Attribute VB_Name = "Module1"

'I tried my best with a busy work schedule this past week...
Sub StockTicker():

For Each ws In Worksheets

    Dim Ticker As String

    Dim Ticker_Total As Double
        Ticker_Total = 0

    Dim Ticker_Table As Double
        Ticker_Table = 2
    'Volume total variable
    
    Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0

    'Here is the heading names

        ws.Range("I1").Value = "Ticker"
    
        ws.Range("J1").Value = "Opening Price"
    
        ws.Range("K1").Value = "Close Price"
    
        ws.Range("L1").Value = "Yearly Change"
        
        ws.Range("M1").Value = "Total Stock Volume"
        
        ws.Range("K:K").NumberFormat = "0.00"
        ws.Range("J:J").NumberFormat = "0.00"
        
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("P:P").Columns.AutoFit
        
        
    Dim LR1 As Long
        LR1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
    For i = 2 To LR1
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                Ticker = ws.Cells(i, 1).Value
                     
        Dim Close_Price As Double
                Close_Price = ws.Cells(i, 6).Value
                    
                ws.Range("I" & Ticker_Table).Value = Ticker
            
                ws.Range("K" & Ticker_Table).Value = Close_Price
                
                'Add Total stock volume.
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                'This outputs the total stock volume based on each stock ticker
                
                ws.Range("M" & Ticker_Table).Value = Total_Stock_Volume
                      
                Ticker_Table = Ticker_Table + 1
                
           'The following lines will return the opening pricing for each given year.
            Dim Open_Price As Double
                Open_Price = 0
        
            ElseIf Open_Price = 0 Then
                      Open_Price = ws.Cells(i, 3).Value
                      
                      ws.Range("J" & Ticker_Table).Value = Open_Price
        
            End If
      
      Next i
    ' This part will hopefully calculate the percent year change.
    Dim LR As Long
        LR = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For j = 2 To LR
    
            Dim op As Double
                op = ws.Cells(j, 10).Value
            Dim cp As Double
                cp = ws.Cells(j, 11).Value
                
        Dim per_change As Double
    
                per_change = ((cp - op) / op) * 100

                ws.Cells(j, 12).Value = per_change
        If per_change > 0 Then
            ws.Cells(j, 12).Interior.ColorIndex = 4
        Else
            ws.Cells(j, 12).Interior.ColorIndex = 3
        End If
        

   With ws
    'Finding the greatest, lowest percentage change, and greatest volumn.
    Dim max_per As Double
        'max_per = WorksheetFunction.Max(ws.Range("L:L"))
        max_per = Application.WorksheetFunction.Max(ws.Range("L:L").Value)
        ws.Range("R2").Value = max_per
        
    Dim min_per As Double
        min_per = Application.WorksheetFunction.Min(ws.Range("L:L").Value)
        ws.Range("R3").Value = min_per
        
    Dim max_volumn As Double
        max_volumn = Application.WorksheetFunction.Max(ws.Range("M:M").Value)
        ws.Range("R4").Value = max_volumn
    
    Dim ticker_max As Integer
        ticker_max = Application.WorksheetFunction.Match(ws.Range("R2").Value, ws.Range("L:L").Value, 0)
        ws.Range("Q2").Value = ticker_max
        ws.Range("Q2").Value = ws.Cells(ticker_max, 9)
        
    Dim ticker_min As Integer
        ticker_min = Application.WorksheetFunction.Match(ws.Range("R3").Value, ws.Range("L:L").Value, 0)
        ws.Range("Q3").Value = ticker_min
        ws.Range("Q3").Value = ws.Cells(ticker_min, 9)
        
    Dim greatest_total As Integer
        greatest_total = Application.WorksheetFunction.Match(ws.Range("R4").Value, ws.Range("M:M").Value, 0)
        ws.Range("Q4").Value = greatest_total
        ws.Range("Q4").Value = ws.Cells(greatest_total, 9)
   End With
  Next j
Next ws


End Sub

