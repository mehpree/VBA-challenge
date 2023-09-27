Attribute VB_Name = "Module1"
Sub TickerAnalysis()

' Set and define variables 
Dim i As Long
Dim ws As Worksheet
Dim LastRow As Long

Dim ticker As String
Dim tickervolume As Double
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percentage_change As Double
Dim stock_vol As Double

Dim tickersummary As Integer
Dim max_increase_ticker As String
Dim max_decrease_ticker As String
Dim max_volume_ticker As String
Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Double

' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Summary Table Details
        ws.Range("I:L").Clear
        ws.Range("P:P").Clear 
        ws.Range("Q:Q").NumberFormat = "0.00%"  
        ws.Range("R:R").NumberFormat = "0.00%"  

        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"

        ' Determine the last row for the current worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Initialize variables
        tickersummary = 2
        ticker = " "
        tickervolume = 0
        open_price = 0
        max_increase = 0
        max_decrease = 0
        max_volume = 0

        ' Loop through rows in the current worksheet
        For i = 2 To LastRow
            ' Check if the current row is for a new ticker symbol
            If ws.Cells(i, 1).Value <> ticker Then
                ' Update summary for the previous ticker, but only if it's not the first row
                If i > 2 Then
                    ' Calculate yearly_change and percentage_change for the previous ticker
                    If open_price <> 0 Then
                        yearly_change = close_price - open_price
                        If open_price <> 0 Then
                            percentage_change = yearly_change / open_price
                        Else
                            percentage_change = 0
                        End If
                    Else
                        yearly_change = 0
                        percentage_change = 0
                    End If
                    stock_vol = tickervolume

                    ' Update summary columns
                    ws.Cells(tickersummary, "I").Value = ticker
                    ws.Cells(tickersummary, "J").Value = yearly_change
                    ws.Cells(tickersummary, "K").Value = percentage_change
                    ws.Cells(tickersummary, "L").Value = stock_vol
                    ' Format percentage change as percentage
                    ws.Cells(tickersummary, "K").NumberFormat = "0.00%"
                    ' Move to the next summary row
                    tickersummary = tickersummary + 1
                End If

                ' Update variables for the new ticker
                ticker = ws.Cells(i, 1).Value
                open_price = ws.Cells(i, 3).Value ' Set the new open_price
                tickervolume = 0
            End If

            ' Update variables for the current ticker
            tickervolume = tickervolume + ws.Cells(i, 7).Value
            close_price = ws.Cells(i, 6).Value

            ' Update max increase, max decrease, and max volume
            If percentage_change > max_increase Then
                max_increase = percentage_change
                max_increase_ticker = ticker
            ElseIf percentage_change < max_decrease Then
                max_decrease = percentage_change
                max_decrease_ticker = ticker
            End If

            If stock_vol > max_volume Then
                max_volume = stock_vol
                max_volume_ticker = ticker
            End If

            ' Handle the last row
            If i = LastRow Then
                ' Calculate yearly_change and percentage_change for the last ticker
                If open_price <> 0 Then
                    yearly_change = close_price - open_price
                    If open_price <> 0 Then
                        percentage_change = yearly_change / open_price
                    Else
                        percentage_change = 0
                    End If
                Else
                    yearly_change = 0
                    percentage_change = 0
                End If
                stock_vol = tickervolume

            ' Update summary for the last ticker
                ws.Cells(tickersummary, "I").Value = ticker
                ws.Cells(tickersummary, "J").Value = yearly_change
                ws.Cells(tickersummary, "K").Value = percentage_change
                ws.Cells(tickersummary, "L").Value = stock_vol
                ws.Cells(tickersummary, "K").NumberFormat = "0.00%"
            End If
        Next i

        ' Update greatest increase, decrease, and volume values in the summary
        ws.Cells(2, 17).Value = max_increase_ticker
        ws.Cells(2, 18).Value = max_increase
        ws.Cells(3, 17).Value = max_decrease_ticker
        ws.Cells(3, 18).Value = max_decrease
        ws.Cells(4, 17).Value = max_volume_ticker
        ws.Cells(4, 18).Value = max_volume

        ' Apply conditional formatting based on positive/negative change
        For j = 2 To tickersummary - 1
            If ws.Cells(j, 11).Value > 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 10 ' Green
            ElseIf ws.Cells(j, 11).Value < 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 3 ' Red
            End If
        Next j
    Next ws

End Sub

