Attribute VB_Name = "Module1"
Sub homework():

Dim last_row As Long
Dim last_column As Integer
Dim ticker_row As Long
Dim open_price_row As Long
Dim open_price As Double
Dim close_price As Double
Dim sum As Long
Dim ws As Worksheet
Dim max As Double
Dim min As Double
Dim max_vol As Double



last_row = Cells(Rows.Count, "A").End(xlUp).Row
last_column = Cells(1, Columns.Count).End(xlToRight).Column

For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total stock Volume"
    ticker_row = 2
    open_price_row = 2
    max = 0
    min = 0
    max_vol = 0
    MsgBox (ws.Name)

    For i = 2 To (last_row)


        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ws.Cells(ticker_row, 9).Value = ws.Cells(i, 1).Value
        
        open_price = ws.Cells(open_price_row, 3).Value
        
        close_price = ws.Cells(i, 6).Value
        
            ''filter function to check if open price = 0
     
            For j = i To (last_row + 1)
    
            If open_price = 0 Then
        
            open_price = ws.Cells(open_price_row + 1, 3).Value
        
            
            End If
        
            Next j
    
        
        ws.Cells(ticker_row, 10).Value = close_price - open_price
    
        If ws.Cells(ticker_row, 10).Value > 0 Then
        
        ws.Cells(ticker_row, 10).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(ticker_row, 10).Value < 0 Then
        
        ws.Cells(ticker_row, 10).Interior.ColorIndex = 3
        
        End If
        
        ''second function to check if close_price=0 or open price =0
        
        If close_price <> 0 And open_price <> 0 Then
        
        ws.Cells(ticker_row, 11).Value = ws.Cells(ticker_row, 10).Value / open_price
    
        ws.Cells(ticker_row, 11).NumberFormat = "0.00%"
    
        ws.Cells(ticker_row, 12) = Application.sum(ws.Range("G" & open_price_row & ":" & "G" & i))
        
        End If
        
        open_price_row = i + 1

        ticker_row = ticker_row + 1

        End If

    Next i
    
    MsgBox (ticker_row)
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    For k = 2 To (ticker_row - 1)
    
    If ws.Cells(k, 11) > max Then
    max = ws.Cells(k, 11)
    ws.Cells(2, 16) = ws.Cells(k, 9)
    End If
    ws.Cells(2, 17) = max
    ws.Cells(2, 17).NumberFormat = "0.00%"
    Next k
    
    For l = 2 To (ticker_row - 1)

    If ws.Cells(l, 11) < min Then
    min = ws.Cells(l, 11)
    ws.Cells(3, 16) = ws.Cells(l, 9)
    End If
    ws.Cells(3, 17) = min
    ws.Cells(3, 17).NumberFormat = "0.00%"
    Next l
    
    For m = 2 To (ticker_row - 1)
    If ws.Cells(m, 12) > max_vol Then
    max_vol = ws.Cells(m, 12)
    ws.Cells(4, 16) = ws.Cells(m, 9)
    End If
    ws.Cells(4, 17) = max_vol
    Next m
    
    
    
    
    
    
    
    
    
    

Next ws

End Sub

