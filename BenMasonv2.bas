Attribute VB_Name = "Module2"
Sub tickersolver()

'PURPOSE: Determine how many seconds it took for code to completely run
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer

'Loop through all the sheets
    For Each ws In Worksheets

        'Set a variable for holding the ticker name
        Dim ticker_name As String
    
        'Set a variable for holding a total count
        Dim ticker_volume As Double
        ticker_volume = 0

       'Set ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Set variable for holding the open price
        Dim open_price As Double
        
        'Set initial open_price. Other opening prices will be determined in the conditional loop.
        open_price = ws.Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'Print the Summary Table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Count the number of rows in the first column.
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows by the ticker names
        For i = 2 To lastrow

            'Searches through rows to match value and count
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Set the ticker name
              ticker_name = ws.Cells(i, 1).Value

              'Add the volume of trade
              ticker_volume = ticker_volume + ws.Cells(i, 7).Value

              'Print the ticker name in the summary table
              ws.Range("I" & summary_ticker_row).Value = ticker_name

              'Print the trade volume for each ticker in the summary table
              ws.Range("L" & summary_ticker_row).Value = ticker_volume

              'Collect closing price
              close_price = ws.Cells(i, 6).Value

              'Calculate yearly change
               yearly_change = (close_price - open_price)
              
              'Print the yearly change for each ticker in the summary table
              ws.Range("J" & summary_ticker_row).Value = yearly_change

              'Check condition for calculating percent change
                If open_price = 0 Then
                    percent_change = 0
                
                Else
                    percent_change = yearly_change / open_price
                
                End If

              'Print the yearly change for each ticker in the summary table
              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter. Add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              ticker_volume = 0

              'Reset the opening price
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              ticker_volume = ticker_volume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    'Conditional formatting
    'First find the last row of the summary table

    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
        For i = 2 To lastrow_summary_table
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i

    'Highlight the stock price changes

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    'Determine the max and min values for ticker

        For i = 2 To lastrow_summary_table
        
            'Find the maximum change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Find the minimum change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
'Determine how many seconds code took to run
SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
        
End Sub
