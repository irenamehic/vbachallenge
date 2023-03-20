Attribute VB_Name = "Module1"
Sub stock_ticker()
    'Defining our variables'
    Dim ws As Worksheet
    Dim tickername As String
    Dim tickervolume As Double
    tickervolume = 0
    
    'Setting up the summary table'
    Dim summary_ticker_row As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    'Loop through the rows by ticker names'
    For Each ws In ThisWorkbook.Worksheets
        summary_ticker_row = 2
        'Summary table headers'
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                tickername = ws.Cells(i, 1).Value
                tickervolume = tickervolume + ws.Cells(i, 7).Value
                
                'Print ticker name in the table'
                ws.Range("I" & summary_ticker_row).Value = tickername
                
                'Print volume for each ticker in the table'
                ws.Range("L" & summary_ticker_row).Value = tickervolume
                
                close_price = ws.Cells(i, 6).Value
                
                'Find the yearly change for stock price'
                open_price = ws.Cells(i, 3).Value
                yearly_change = (close_price - open_price)
                ws.Range("J" & summary_ticker_row).Value = yearly_change
                
                If (open_price = 0) Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If
                
                ws.Range("K" & summary_ticker_row).Value = percent_change
                ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
                
                'reset row counter'
                summary_ticker_row = summary_ticker_row + 1
                tickervolume = 0
                open_price = ws.Cells(i + 1, 3)
            Else
                tickervolume = tickervolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        'Conditional formatting to highlight positive change in green and negative change in red'
        lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow_summary_table
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    Next ws

End Sub


