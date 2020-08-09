Attribute VB_Name = "Module1"
Sub Stock_Loop()

Dim ws As Worksheet
Dim ticker_count As Integer
Dim opening_value As Double
Dim closing_value As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim ticker_initial_index As Double
Dim ticker_final_index As Double
Dim total_stock_volume As Variant
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double
Dim greatest_increase_index As Double
Dim greatest_decrease_index As Double
Dim greatest_volume_index As Double

For Each ws In Worksheets

    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
    ticker_count = 2
    
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
                ws.Cells(ticker_count, 9) = ws.Cells(i, 1)
                opening_value = ws.Cells(i, 3)
                ticker_initial_index = i
            ElseIf ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                closing_value = ws.Cells(i, 6)
                ticker_final_index = i
                yearly_change = closing_value - opening_value
                If opening_value = 0 Then
                    ws.Cells(ticker_count, 11) = 0
                Else
                    percent_change = yearly_change / opening_value
                    ws.Cells(ticker_count, 10) = yearly_change
                    ws.Cells(ticker_count, 11) = percent_change
                    ws.Cells(ticker_count, 11).NumberFormat = "0.00%"
                End If
                If yearly_change < 0 Then
                    ws.Cells(ticker_count, 10).Interior.Color = RGB(255, 0, 0)
                ElseIf yearly_change >= 0 Then
                    ws.Cells(ticker_count, 10).Interior.Color = RGB(0, 255, 0)
                End If
                total_stock_volume = 0
                For j = ticker_initial_index To (ticker_final_index + 1)
                    total_stock_volume = total_stock_volume + Cells(j, 7)
                Next j
                ws.Cells(ticker_count, 12) = total_stock_volume
                ticker_count = ticker_count + 1
            End If
        Next i
    
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "Greatest % increase"
    ws.Cells(3, 15) = "Greatest % decrease"
    ws.Cells(4, 15) = "Greatest total volume"
    
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
        For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
            If ws.Cells(i, 11) > greatest_increase Then
                greatest_increase = ws.Cells(i, 11)
                greatest_increase_index = i
            End If
            If ws.Cells(i, 11) < greatest_decrease Then
                greatest_decrease = ws.Cells(i, 11)
                greatest_decrease_index = i
            End If
            If ws.Cells(i, 12) > greatest_volume Then
                greatest_volume = ws.Cells(i, 12)
                greatest_volume_index = i
            End If
        Next i
    
    ws.Cells(2, 16) = ws.Cells(greatest_increase_index, 9)
    ws.Cells(2, 17) = ws.Cells(greatest_increase_index, 11)
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ws.Cells(3, 16) = ws.Cells(greatest_decrease_index, 9)
    ws.Cells(3, 17) = ws.Cells(greatest_decrease_index, 11)
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Cells(4, 16) = ws.Cells(greatest_volume_index, 9)
    ws.Cells(4, 17) = ws.Cells(greatest_volume_index, 12)

Next

End Sub

