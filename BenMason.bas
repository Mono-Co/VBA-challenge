Attribute VB_Name = "Module1"
Sub Total_Stock_Table()

'PURPOSE: Determine how many seconds it took for code to completely run
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
  StartTime = Timer

'For loop through each worksheet
For Each CurrentWs In Worksheets

'Print the Row Name for new table
CurrentWs.Range("J1").Value = "Ticker"
CurrentWs.Range("K1").Value = "Yearly Change"
CurrentWs.Range("L1").Value = "Percent Change"
CurrentWs.Range("M1").Value = "Total Stock Volume"
CurrentWs.Range("O2").Value = "Greatest % Increase"
CurrentWs.Range("O3").Value = "Greatest % Decrease"
CurrentWs.Range("O4").Value = "Greatest Total Volume"
CurrentWs.Range("P1").Value = "Ticker"
CurrentWs.Range("Q1").Value = "Value"

'Set the lastrow
lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

' Set an initial variable for holding the Stock Ticker name
Dim Ticker As String

' Set an initial variable for holding the total Stock Total
Dim Total_Stock_Volume As Double
Stock_Total = 0

'Set an initial variable for holding the Stock Yearly Open
Dim Total_Stock_Open As Double
Stock_Open = 0

'Set an initial variable for holding the Stock Yearly Close
Dim Total_Stock_Close As Double
Stock_Close = 0

'Set an initial variable forholding the Yearly Change
Dim Total_Stock_Yearly_Change As Double
Stock_Yearly_Change = 0

'Set an Inital variable for holding the Yearly % Change
Dim Total_Stock_Yearly_Percent_Change As Double
Stock_Yearly_Percent_Change = 0

' Keep track of the location for each Stock Ticker name in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 1

' Loop through all Stock Tickers
For i = 2 To lastrow

'Search for the Rows to find the same Ticker code
If CurrentWs.Cells(i + 1, 1).Value = CurrentWs.Cells(i, 1).Value Then

'Write the Stock Trade Volume Colum
Total_Stock_Volume = Total_Stock_Volume + CurrentWs.Cells(i, 7).Value

Else

' Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

' Set the Ticker name
Ticker = CurrentWs.Cells(i, 1).Value

' Set the Open Price
Stock_Open = CurrentWs.Cells(2, 3).Value

'Set the Close Price
Stock_Close = CurrentWs.Cells(i, 6).Value

'Find the Yearly Price Change & Yearly Percentage
Stock_Yearly_Change = Stock_Close - Stock_Open

'Find the Yearly Change in Percentage
Total_Stock_Yearly_Percent_Change = (Stock_Yearly_Change / Stock_Open) * 100

' Print the Ticker in the Summary Table
CurrentWs.Range("J" & Summary_Table_Row).Value = Ticker

' Print the Yearly Change to the Summary Table
CurrentWs.Range("K" & Summary_Table_Row).Value = Stock_Yearly_Change

' Print the Yearly Percentage Change to the Summary Table
CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Stock_Yearly_Percent_Change

' Print the Stock Volume to the Summary Table
CurrentWs.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume

' Reset the Stock Total
Total_Stock_Volume = 0

End If

'Format the Cells
If (Stock_Yearly_Change >= 0) Then

CurrentWs.Range("K" & Summary_Table_Row + 1).Interior.ColorIndex = 4

ElseIf (Stock_Yearly_Change <= 0) Then

CurrentWs.Range("K" & Summary_Table_Row + 1).Interior.ColorIndex = 3
                
End If
       
Next i

'Bonus Works

'Create variable to store Greatest Percentage Decrease value
Dim Greatest_Stock_Yearly_Percent_Deccrease As Single

'Print the Greatest Percentage Increase value
CurrentWs.Range("Q2") = WorksheetFunction.Max(CurrentWs.Range("L:L"))

'Print the Greatest Percentage Decrease value
CurrentWs.Range("Q3") = WorksheetFunction.Min(CurrentWs.Range("L:L"))

'Create variable to store Greatest Percentage Decrease value
Dim Greatest_Stock_Yearly_Percent_Volume As Single

'Print the Greatest Percentage Decrease value
CurrentWs.Range("Q4") = WorksheetFunction.Max(CurrentWs.Range("M:M"))


 ' Create variables to hold Greatest Stocks
    Dim GreatestPercentageIncrease As Double
    Dim GreatestPercentageDecrease As Double
    Dim GreatestTotalVolume As Double

    ' Establish the Greatest Stock numbers
    Greatest_Percentage_Increase = CurrentWs.Cells(2, 17).Value
    Greatest_Percentage_Decrease = CurrentWs.Cells(3, 17).Value
    Greatest_Total_Volume = CurrentWs.Cells(4, 17).Value

    ' Loop through each of the Stock Summary
    For i = 1 To lastrow

        ' Check if the lotto number matches the Greatest Percentage Increase.
        If CurrentWs.Cells(i, 12).Value = Greatest_Percentage_Increase Then

           ' Retrieve the values associated with the Greatest Percentage Incefrase.
            CurrentWs.Cells(2, 16).Value = CurrentWs.Cells(i, 10).Value
                     
            ' Check if the  number matches the Greatest Percentage Decrease.
             ElseIf CurrentWs.Cells(i, 12).Value = Greatest_Percentage_Decrease Then
            
            ' Retrieve the values associated with the Greatest Percentage Decrease.
            CurrentWs.Cells(3, 16).Value = CurrentWs.Cells(i, 10).Value
            
        ' Check if the number matches the Greatest Volume.
        ElseIf CurrentWs.Cells(i, 13).Value = Greatest_Total_Volume Then
           
            ' Retrieve the values associated with the Greatest Volume.
            CurrentWs.Cells(4, 16).Value = CurrentWs.Cells(i, 10).Value
       
        ' Ends this series of IF/ELSE conditionals
        End If

     Next i
             
Next CurrentWs

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub

