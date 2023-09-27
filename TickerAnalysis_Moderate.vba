Attribute VB_Name = "Module1"
Sub TickerAnalysis_Easy()
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

' Initialize tickersummary
    Dim tickersummary As Long 

' Initialize j
    Dim j As Long

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Determine the last row for the current worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Initialize variables
        ticker = " "
        tickervolume = 0
        open_price = 0
        tickersummary = 2 ' Initialize tickersummary to the starting row for summary

        ' Add headers for the summary columns
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percentage Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"

        ' Loop through rows in the current worksheet
        For i = 2 To LastRow
            ' Check if the current row is for a new ticker symbol
            If ws.Cells(i, 1).Value <> ticker Then
                ' Update summary for the previous ticker, but only if it's not the first row
                If i > 2 Then
                    ' Calculate yearly_change and percentage_change for the previous ticker
                    yearly_change = close_price - open_price
                    If open_price <> 0 Then
                        percentage_change = yearly_change / open_price
                    Else
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

            ' Handle the last row
            If i = LastRow Then
                ' Calculate yearly_change and percentage_change for the last ticker
                yearly_change = close_price - open_price
                If open_price <> 0 Then
                    percentage_change = yearly_change / open_price
                Else
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
