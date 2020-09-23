
Sub Activate()

'Activate all worksheet so the VBA scripts all the data
For Each ws In Worksheets
    ws.Activate
    Call Calculate 
Next ws

End Sub

Sub Calculate()

'set variables
Dim current_ticker As String
Dim next_ticker As String
Dim total As Double
Dim total_row As Double
Dim open_price As Double
Dim open_price2 As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double


'Create headers/title for the Calculated data column
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'VBA Calc Scripting, assign values to variables
total_row = Cells(Rows.Count, "A").End(xlUp).Row
start_row = 2

'set the inital prices as the first value seen on ws 
open_price = Cells(start_row, 3).Value

'create a loop to check if the ticker name changes for the wholesheet
For current_row = 2 To total_row

'+1 ensure that next row is read
current_ticker = Cells(current_row, 1).Value
next_ticker = Cells(current_row + 1, 1).Value

'Add the total Volume
total = total + Cells(current_row, 7).Value

If current_ticker <> next_ticker Then
    Cells(start_row, 9).Value = current_ticker
    Cells(start_row, 12).Value = total
    close_price = Cells(current_row, 6).Value
    yearly_change = close_price - open_price
    Cells(start_row, 10).Value = yearly_change

'calculate the percent_change also troubleshoot division by zero error
    If open_price = 0 Then
        percent_change = 0
    Else
        percent_change = (yearly_change / open_price)
        Cells(start_row, 11).Value = percent_change
        'format to have 2 decimal point and show value as percentage
        Cells(start_row, 11).NumberFormat = "0.00%"
    End If
    
'condition formating formatting for yearly_changes : green for positive and red for negative.
If yearly_change > 0 Then
    Cells(start_row, 10).Interior.ColorIndex = 4
ElseIf yearly_change <= 0 Then
Cells(start_row, 10).Interior.ColorIndex = 3
End If
'to start on next Row
start_row = start_row + 1
'reset the total
total = 0
'to ensure open price is changed to the value of next_ticker instead of resetting to preset value
open_price = Cells(current_row + 1, 3).Value

End If

Next current_row


'For Challenges part of HW
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Declare Variables
Dim per_lastrow As Double
Dim max_per As Double
Dim min_per As Double
Dim vol_lastrow As Double
Dim max_vol As Double

'set the variables to zero
max_per = 0
min_per = 0
'identifying last percentage row and last vol row
per_lastrow = Cells(Rows.Count, 11).End(xlUp).Row
vol_lastrow = Cells(Rows.Count, 12).End(xlUp).Row

'loop through the row using if statement to get the max and min
For currentrow2 = 2 To per_lastrow

'the loop goest through each row and record the value if larger than the previous and vice versa for min
If max_per < Cells(currentrow2, 11).Value Then
    max_per = Cells(currentrow2, 11).Value
    Range("Q2").Value = max_per
    Range("Q2").NumberFormat = "0.00%"
    Range("P2").Value = Cells(currentrow2, 9).Value
    
ElseIf min_per > Cells(currentrow2, 11).Value Then
    min_per = Cells(currentrow2, 11).Value
    Range("Q3").Value = min_per
    Range("Q3").NumberFormat = "0.00%"
    Range("P3").Value = Cells(currentrow2, 9).Value
End If

  Next currentrow2
  
'create a parameter for the length of the row for volume that the loop needs to run through
For currentrow3 = 2 To vol_lastrow

'loops through each row to find the largest value in the row. if next value is greater max_vol gets overwritten till complete.
If max_vol < Cells(currentrow3, 12).Value Then
    max_vol = Cells(currentrow3, 12).Value
    Range("Q4").Value = max_vol
    Range("P4").Value = Cells(currentrow3, 9).Value
    
    End If
    
    Next currentrow3

'autofit all the columns so data can be properly displayed
Columns("A:Q").AutoFit




End Sub



















