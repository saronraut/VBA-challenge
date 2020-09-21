Attribute VB_Name = "Module3"
Sub Activate()

'Activate all worksheet so the VBA scripts all the data
For Each ws In Worksheets
    ws.Activate
    Call Calculate
Next ws

End Sub

Sub Calculate()

'Part one is set variables
Dim current_ticker As String
Dim next_ticker As String
Dim total As Double
Dim total_row As Double
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'Create headers/title for the Calculated data
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'VBA Calc Scripting, assign values to variavles
total_row = Cells(Rows.Count, "A").End(xlUp).Row
start_row = 2

'set the inital prices as the first value seen on ws, then loop will be used to identify the first value after change in tickername
open_price = Cells(start_row, 3).Value

'create a loop to check if the ticker name changes for the wholesheet
For Current_row = 2 To total_row

current_ticker = Cells(Current_row, 1).Value
next_ticker = Cells(Current_row + 1, 1).Value

'Add the total Volume
total = total + Cells(Current_row, 7).Value


If current_ticker <> next_ticker Then
    Cells(start_row, 9).Value = current_ticker
    Cells(start_row, 12).Value = total
    close_price = Cells(Current_row, 6).Value
    yearly_change = close_price - open_price
    Cells(start_row, 10).Value = yearly_change

'calculate the percent_change also troubleshoot division by zero
    If open_price <> 0 Then
        percent_change = (yearly_change / open_price) * 100
        Cells(start_row, 11).Value = percent_change
    Else
        MsgBox ("error dividing by zero")
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
open_price = Cells(Current_row + 1, 3).Value


End If

Next Current_row




End Sub















