Attribute VB_Name = "Module1"
Sub vba_challenge_2_stock()

'set up the collection of variables
Dim ws As Worksheet

Dim ticker_count As Integer
Dim Summary_table_row As Integer
Dim stock_volume As Double
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim last_row As Long
Dim ticker_name As String

For Each ws In ActiveWorkbook.Worksheets

'give those variables hard coded starting values as needed
    ticker_count = 2
    stock_volume = 0
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Summary_table_row = 2
    opening_price = ws.Cells(2, 3).Value

'create the loop
     For i = 2 To last_row
     
     stock_volume = stock_volume + ws.Cells(i, 7).Value
     ws.Range("M" & Summary_table_row).Value = stock_volume
     
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker_name = ws.Cells(i, 1).Value
            closing_price = ws.Cells(i, 6).Value
            yearly_change = (closing_price - opening_price)
            percent_change = ((closing_price - opening_price) / opening_price)
           
            'print locations for values
            ws.Range("J" & Summary_table_row).Value = ticker_name
            ws.Range("K" & Summary_table_row).Value = yearly_change
            ws.Range("L" & Summary_table_row).Value = percent_change
            
            'reset variables and values for next loop
            Summary_table_row = Summary_table_row + 1
            stock_volume = 0
            ticker_count = ticker_count + 1
            opening_price = ws.Cells(i + 1, 3)
        
        End If
        
        'Conditional formatting
        If ws.Cells(i, 11).Value >= 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 11).Interior.ColorIndex = 3
        End If
             If ws.Cells(i, 12).Value >= 0 Then
            ws.Cells(i, 12).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 12).Interior.ColorIndex = 3
        End If
    Next i

'formatting the tables
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"

ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"

Next ws

End Sub
