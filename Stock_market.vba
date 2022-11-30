Sub Stock_market()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets


'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

'Create variable to hold stock volume
'Dim stock_volume As Double
'stock_volume = 0

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Set new variables for prices and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0
Dim TickerRow As Long
TickerRow = 1
Dim total_volume As Double
total_volume = 0


'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow
    
    'Ticker symbol output
    total_volume = total_volume + ws.Cells(i, 7).Value
    
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        open_price = ws.Cells(i, 3).Value
        Ticker = ws.Cells(i, 1).Value
    End If

    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        close_price = ws.Cells(i, 6).Value
    End If
    

    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TickerRow = TickerRow + 1
    
        price_change = close_price - open_price
        price_change_percent = (price_change / open_price) * 100
        ws.Cells(TickerRow, "I").Value = Ticker
        ws.Cells(TickerRow, "J").Value = price_change
        If price_change >= 0 Then
           ws.Cells(TickerRow, "J").Interior.ColorIndex = 4
        Else
            ws.Cells(TickerRow, "J").Interior.ColorIndex = 3
        End If
        ws.Cells(TickerRow, "K").Value = price_change_percent
        ws.Cells(TickerRow, "K").NumberFormat = "0.00%"
        ws.Cells(TickerRow, "L").Value = total_volume
        total_volume = 0
    End If
    Next i
    
Lastrow_bonus = ws.Cells(Rows.Count, 9).End(xlUp).Row

Dim Greatest_increase As Double
Greatese_increase = 0
Dim Greatest_decrease As Double
Greatest_decrease = 0
Dim Greatest_TotalVolume As Double
Greatest_TotalVolume = 0
Dim Greatest_increase_ticker As String
Dim Greatest_decrease_ticker As String
Dim Greatest_TotalVolume_ticker As String

For j = 2 To Lastrow_bonus
    If ws.Cells(j, 11).Value >= Greatest_increase Then
        Greatest_increase = ws.Cells(j, 11).Value
        Greatest_increase_ticker = ws.Cells(j, 9).Value
    End If
    If ws.Cells(j, 11).Value < Greatest_decrease Then
        Greatest_decrease = ws.Cells(j, 11).Value
        Greatest_decrease_ticker = ws.Cells(j, 9).Value
    End If
     If ws.Cells(j, 12).Value >= Greatest_TotalVolume Then
        Greatest_TotalVolume = ws.Cells(j, 12).Value
        Greatest_TotalVolume_ticker = ws.Cells(j, 9).Value
    End If
Next j

ws.Cells(2, "P").Value = Greatest_increase_ticker
ws.Cells(2, "Q").Value = Greatest_increase
ws.Cells(2, "Q").NumberFormat = "0.00%"
ws.Cells(3, "P").Value = Greatest_decrease_ticker
ws.Cells(3, "Q").Value = Greatest_decrease
ws.Cells(3, "Q").NumberFormat = "0.00%"
ws.Cells(4, "P").Value = Greatest_TotalVolume_ticker
ws.Cells(4, "Q").Value = Greatest_TotalVolume

Next ws

End Sub


