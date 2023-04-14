Attribute VB_Name = "Module8"
Sub StocksFinalAllWorksheets()

Dim x As Integer
Dim ws_num As Integer
Dim ws As Worksheet
ws_num = ThisWorkbook.Worksheets.Count

'loop through all worksheets in the workbook
For x = 1 To ws_num

 'find the sheets that need to run code
 If Sheets(x).Name = "2018" Or Sheets(x).Name = "2019" Or Sheets(x).Name = "2020" Then

 'after finding the right sheet, run the code
 ThisWorkbook.Worksheets(x).Activate

'________________________________________________________________

'Part I: Create summary table aggregating stock data

'Determine variables
Dim Ticker As String
Dim Ticker_Total As Double
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'Set an initial variable for holding the total per stock
Ticker_Total = 0
open_price = Cells(2, 3).Value

'Print titles on summary table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Keep track of the location for each stock symbol in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Loop through all stocks
For i = 2 To 753001

 'Check if we are still within the same stock, if it is not...
 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
    'Set the stock symbol
    Ticker = Cells(i, 1).Value
    
    'Add to the Stock Total
    Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
    'Print the Stock smybol in the Summary Table
    Range("I" & Summary_Table_Row).Value = Ticker
    
    'Print the Stock Total to the Summary Table
    Range("L" & Summary_Table_Row).Value = Ticker_Total
      
    'set close price
    close_price = Cells(i, 6).Value
    
    'Calculate the yearly change
    yearly_change = close_price - open_price
    
    'Calculate the percent change
    percent_change = yearly_change / open_price
    
    'Print the yearly change in the summary table
    Range("J" & Summary_Table_Row).Value = yearly_change
    
    'print the percent change in the summary table
    Range("K" & Summary_Table_Row).Value = percent_change
    
    'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    
    'Reset the Stock Total
    Ticker_Total = 0
    open_price = Cells(1 + i, 3).Value
       
 'If the cell immediateely following a row is the same stock symbol
    Else
    
     'Add to the Stock Total
     Ticker_Total = Ticker_Total + Cells(i, 7).Value
  
    End If

 Next i
 
'Change format of percent change column to %
Range("K2:K30001").NumberFormat = "0.00%"

'hightlight positive and negative values in percent range

  'Loop through values in percent_range column
  For j = 2 To 3001

   'If values are negative then make them red
   If Cells(j, 10).Value < 0 Then
   Cells(j, 10).Interior.ColorIndex = 3

   'Else make them green
   Else
   Cells(j, 10).Interior.ColorIndex = 4

   End If

  Next j

'____________________________________________________________________

'Part 2: Create a table to show greatest increase, greatest decrease, and greatest volume

'Determine Variables
Dim PercentRng As Range
Dim VolumeRng As Range
Dim LastRow As Double

'Set location of Last Row
LastRow = Range("I" & Rows.Count).End(xlUp).row

'Set ranges to search
Set PercentRng = Range("K2:K" & LastRow)
Set VolumeRng = Range("L2:L" & LastRow)

'Calculate greatest percent increase
max_increase = Application.WorksheetFunction.Max(PercentRng)

'Calculate greatest percent decrease
min_increase = Application.WorksheetFunction.Min(PercentRng)

'Calculate greatest total volume
max_volume = Application.WorksheetFunction.Max(VolumeRng)

'Set row names of greatest summary table
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Set column names of greatest summary table
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Print greatest values in greatest summary table
Cells(2, 17).Value = max_increase
Cells(3, 17).Value = min_increase
Cells(4, 17).Value = max_volume

'Change format to percent
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"

'Fill ticker column
Dim y, z, a As Integer
Dim Ticker2 As String

'Loop through all stocks
For y = 2 To 3001

 'Check match for greatest increase
 If Cells(y, 11).Value = max_increase Then
  
    'Set the stock symbol
    Ticker2 = Cells(y, 9).Value
     
    'Print the Stock smybol in the Summary Table
    Cells(2, 16).Value = Ticker2
    
    Else
     
    End If

 Next y

For z = 2 To 3001

 'Check match for greatest decrease
 If Cells(z, 11).Value = min_increase Then
  
    'Set the stock symbol
    Ticker2 = Cells(z, 9).Value
     
    'Print the Stock smybol in the Summary Table
    Cells(3, 16).Value = Ticker2
    
    Else
     
    End If

 Next z

For a = 2 To 3001

 'Check match for greatest volume
 If Cells(a, 12).Value = max_volume Then
  
    'Set the stock symbol
    Ticker2 = Cells(a, 9).Value
     
    'Print the Stock smybol in the Summary Table
    Cells(4, 16).Value = Ticker2
    
    Else
     
    End If

 Next a

'_____________________________________________________________________

End If

Next

End Sub
