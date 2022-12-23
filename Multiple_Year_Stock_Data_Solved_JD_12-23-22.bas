Attribute VB_Name = "Module1"

Sub multiple_year_stock_data():

Dim ticker As String
Dim ticker_summary_row As Integer
Dim Open_A As Double
Dim Close_A As Double
Dim vol As LongLong
Dim ws As Worksheet

For Each ws In Sheets

ws.Activate

Open_A = ws.Cells(2, 3).Value
ticker_summary_row = 2
vol = 0

ws.Range("I1") = "ticker"
ws.Range("J1") = "Total Stock volume"
ws.Range("K1") = "Yearly Change"
ws.Range("L1") = "Percent Change"


LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row

For i = 2 To LastRow

'Cells that are not Equal to each other
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Set ticker value
     ticker = ws.Cells(i, 1).Value
    
    'Add stock volume
    vol = vol + ws.Cells(i, 7).Value
    
    'Add close value
    Close_A = ws.Cells(i, 6).Value
    
    'Print Difference between first closing and first opening
    ws.Range("K" & ticker_summary_row).Value = Close_A - Open_A
    
    'Format percentage for stock change
    ws.Cells(ticker_summary_row, 12).NumberFormat = "0.00%"
      
    'Expression to determine Stock change
    ws.Cells(ticker_summary_row, 12) = (Close_A - Open_A) / Open_A
    
    'Set Open_A value and iterate
    Open_A = ws.Cells(i + 1, 3).Value
    
    'Conditional formatting for positive and negative stock change
    If ws.Cells(ticker_summary_row, 11).Value < 0 Then
    ws.Cells(ticker_summary_row, 11).Interior.ColorIndex = 3

Else
    
    ws.Cells(ticker_summary_row, 11).Interior.ColorIndex = 4
    
End If

    'Print ticker to the summary table
    ws.Range("I" & ticker_summary_row).Value = ticker

    'Print stock volume to the summary row
    ws.Range("J" & ticker_summary_row).Value = vol
    
 ticker_summary_row = ticker_summary_row + 1

    'Reset vol total
    vol = 0
    

'Cells that are equal to each other
Else

vol = vol + ws.Cells(i, 7).Value

End If

Next i

Next ws

End Sub

Sub Increase_Decrease_Stock():

Dim ticker As String
Dim Great_Per_Increase As Double
Dim Low_Per As Double
Dim LastRow As Double
Dim i As Integer
Dim High_Stock_Volume As LongLong
Dim ws As Worksheet

'Iterate through active sheets in workbook
For Each ws In Sheets

ws.Activate

'Set headers for printed data
ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"
ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Total Volume"

Low_Per = 0

LastRow = ws.Range("L" & Rows.Count).End(xlUp).Row

For i = 2 To LastRow

'Condition to find greatest percent increase
If ws.Range("L" & i).Value > Great_Per_Increase Then
    Great_Per_Increase = ws.Range("L" & i).Value
    ticker = ws.Range("I" & i).Value
    
    ws.Range("P2") = Great_Per_Increase
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("O2") = ticker
    
End If
    
'Condition to find Greatest % decrease
If ws.Range("L" & i).Value < Low_Per Then
    Low_Per = ws.Range("L" & i).Value
    ticker = ws.Range("I" & i).Value
    
    ws.Range("P3") = Low_Per
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("O3") = ticker
    
    End If
    
 'Condition to find Greatest total volume
If ws.Range("J" & i).Value > High_Stock_Volume Then
    High_Stock_Volume = ws.Range("J" & i).Value
    ticker = ws.Range("I" & i).Value
    
    ws.Range("P4") = High_Stock_Volume
    ws.Range("O4") = ticker
    
    'For some reason beyond my comprehension at this time, this code will not print out the "Great_Per_Increase" or "Greatest Total Volume" in Sheet 3 (2020).
    'There are no issues with the print out in Sheets 1 or 2
    'My belief is there is something wrong either in how vba is referencing the active sheets or something incorrect in my conditionals.
    
    End If
    
Next i

Next ws
    
End Sub

