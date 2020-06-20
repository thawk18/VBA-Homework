Attribute VB_Name = "Module1"
Sub ticker():

'variables definition
Dim lastrow As Long
Dim i As Long
Dim j As Long
Dim ticker As Long
Dim YearlyChange As Double
Dim Percentchange As Double
Dim Workingdays As Integer
Dim Sheet As Worksheet

         
For Each Sheet In Worksheets

'Writing up titles
Sheet.Cells(1, 9).Value = "Ticker"
Sheet.Cells(1, 10).Value = "Yearly Change"
Sheet.Cells(1, 11).Value = "Percentage Change"

' how many rows the data source has?
lastrow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row

i = 2

For i = 2 To lastrow

' Compare the current cell with the cell above
' if current cell is different from the cell above then there is a new ticker
If Sheet.Cells(i, 1).Value <> Sheet.Cells(i - 1, 1).Value Then

' Get the next line to fill ticker
tickerrow = Sheet.Cells(Rows.Count, 9).End(xlUp).Row + 1


'place new ticker
Sheet.Cells(tickerrow, 9).Value = Sheet.Cells(i, 1).Value

ElseIf Sheet.Cells(i, 1).Value = Sheet.Cells(2, 9).Value Then

'Get number of extra working days after first day (nr of data source lines per ticker -1)
'this way macro works for any given year
Workingdays = Workingdays + 1

End If


Next i
   
' get the last row of total tickers
nrtickers = Sheet.Cells(Rows.Count, 9).End(xlUp).Row

j = 2
' for each ticker
For j = 2 To nrtickers + 1

i = 2

i = Workingdays + 2

'run again through data source starting on the working days, +2 to adjust to line
For i = Workingdays + 2 To lastrow + Workingdays

' if Column I as the same ticker of column A...
If Sheet.Cells(i, 1).Value = Sheet.Cells(j, 9).Value Then

'calculate yearly change, percentage change and total volume
YearlyChange = Sheet.Cells(i, 6) - Sheet.Cells(i - Workingdays, 3)
Percentchange = YearlyChange / Sheet.Cells(i - Workingdays, 3)
TotalVolume = Sheet.Cells(i - Workingdays, 7) + TotalVolume

Else
End If

Next i

' Place all results in the table
Sheet.Cells(j, 10).Value = YearlyChange
Sheet.Cells(j, 11).Value = Percentchange
Sheet.Cells(j - 1, 12).Value = TotalVolume

' Remove last cells as they don't need a value
Sheet.Cells(nrtickers + 1, 10).Value = ""
Sheet.Cells(nrtickers + 1, 11).Value = ""

'Total Volume to reset for next sheet
TotalVolume = 0

Next j

Sheet.Cells(1, 12).Value = "Total Stock Volume"

'Reset variables for next sheet
Workingdays = 0
tickerrow = 0
lastrow = 0
nrtickers = 0

Next

End Sub
