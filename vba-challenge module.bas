Attribute VB_Name = "Module1"
Sub vba_challenge():

'loop thru all worksheets

Dim ws As Worksheet

For Each ws In Worksheets

    'create a loop and output the following
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' create headers
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ' Declare variables
    Dim Ticker_symbol As String
    Ticker_symbol_row = 2
    Dim column As Integer
    column = 1
    
    For i = 2 To lastrow

        ' check if same ticker value
        If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
        Ticker_symbol = ws.Cells(i, 1).Value
    
            ' print ticker symbol in I
            ws.Range("I" & Ticker_symbol_row).Value = Ticker_symbol
            Ticker_symbol_row = Ticker_symbol_row + 1
        
        End If
        
    Next i
    
    ' calc yearly change, percent change, and total stock volume
 
    ' Declare variables
    Dim yearly_change_row As Integer
    yearly_change_row = 2
    Dim yearly_change As String
    Dim open_price As Double
    Dim close_price As Double
    Dim delta_price As Double
    Dim per_change As Double
    Dim Total_volume As Double
    
    open_price = ws.Cells(2, 3).Value
  
    ' check if we are within same ticker, if not-
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        close_price = ws.Cells(i, 6).Value
        delta_price = close_price - open_price
        
        
            ' print delta price total in J
            ws.Range("J" & yearly_change_row).Value = delta_price
              
            ' ignore open_prices with a zero value
            If ws.Cells(i, 3).Value = 0 Then
                ws.Cells(i, 3).Value = Null
                
            ' calculate percent change
            Else
                per_change = delta_price / open_price
            End If
            
            ' print percent change in K
            ws.Range("K" & yearly_change_row).Value = per_change
             
            'calculate total stock volume
            Total_volume = Total_volume + ws.Cells(i, 7).Value
             
            ' Print total volume in L
            ws.Range("L" & yearly_change_row).Value = Total_volume
            
            'add one to the row
            yearly_change_row = yearly_change_row + 1
              
            ' Reset totals
            open_price = 0
            open_price = ws.Cells(i + 1, 3).Value
            Total_volume = 0
          
        Else
            delta_price = close_price - open_price
             
        End If
        
    Next i

Next ws

End Sub
