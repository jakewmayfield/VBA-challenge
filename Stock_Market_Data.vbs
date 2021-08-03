Attribute VB_Name = "Module1"
Sub tickercounter()

'Make loop for all the worksheets
    For Each ws In Worksheets

        'Make variables for ticker, total volume, row for summary table
        Dim ticker As String
        Dim totalvolume As Double
        totalvolume = 0
        Dim summary_table_row As Integer
        summary_table_row = 2
        
        'make variable for open price, close price, year change, and percent change
        Dim open_price As Double
        open_price = ws.Cells(2, 3).Value
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'Make column titles for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Determine total rows in first column for loop
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Make conditional loop to grab data by comparing current cell in loop to next
        For i = 2 To lastrow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Grab ticker & volume
              ticker = ws.Cells(i, 1).Value
              totalvolume = totalvolume + ws.Cells(i, 7).Value

              'Print ticker & volume in summary table
              ws.Range("I" & summary_table_row).Value = ticker
              ws.Range("L" & summary_table_row).Value = totalvolume

              'Grab close price for yearly change and print in summary table
              close_price = ws.Cells(i, 6).Value
              yearly_change = (close_price - open_price)
              ws.Range("J" & summary_table_row).Value = yearly_change

              'Get percent change & DON'T DIVIDE BY 0 OR YOU'LL BREAK THINGS
                If open_price = 0 Then
                    percent_change = 0
                
                Else
                    percent_change = yearly_change / open_price
                
                End If

              'Print percent change in summary table
              ws.Range("K" & summary_table_row).Value = percent_change
              ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
   
              'Increment table row counter and reset total volume and opening price
              summary_table_row = summary_table_row + 1
              totalvolume = 0
              open_price = ws.Cells(i + 1, 3)
            
            Else
               'add up total volume
              totalvolume = totalvolume + ws.Cells(i, 7).Value
            End If
        
        Next i

    'Set up conditional formatting for positive change and negative change (green & red)
    
    'Get last row of summary table for loop
    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Make loop for formatting
        For i = 2 To lastrow_summary_table
            
            If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 10
            
            End If
        
        Next i

    'Make column titles for greatest overall summary table

        ws.Cells(2, 17).Value = "Ticker"
        ws.Cells(2, 18).Value = "Value"
        ws.Cells(3, 16).Value = "Greatest % Increase"
        ws.Range("P:P").ColumnWidth = 18.75
        ws.Cells(4, 16).Value = "Greatest % Decrease"
        ws.Cells(5, 16).Value = "Greatest Total Volume"
        ws.Range("R:R").ColumnWidth = 10.25
        
      

    'Figure out max and min for percent change and max for total volume to determine print data
    
        For i = 2 To lastrow_summary_table
        
            'Find max percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max _
            (ws.Range("K2:K" & lastrow_summary_table)) Then
            ws.Cells(3, 17).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 18).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 18).NumberFormat = "0.00%"

            'Find min percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min _
            (ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(4, 17).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 18).Value = ws.Cells(i, 11).Value
                ws.Cells(4, 18).NumberFormat = "0.00%"
            
            'Find highest total volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max _
            (ws.Range("L2:L" & lastrow_summary_table)) Then
            ws.Cells(5, 17).Value = ws.Cells(i, 9).Value
            ws.Cells(5, 18).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
End Sub


