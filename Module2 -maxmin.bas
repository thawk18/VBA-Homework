Attribute VB_Name = "Module2"
Sub MaxMin()



' Define Variables
Dim Sheet As Worksheet
Dim i As Long
Dim firstRow As Integer
Dim columnNumber As Integer
Dim max As Double
Dim min As Double
Dim lastticker As Integer
Dim ticker As String

         
For Each Sheet In Worksheets

Sheet.Cells(1, 16) = "Ticker"
Sheet.Cells(1, 17) = "Value"
Sheet.Cells(2, 15) = "Greateast % Increase"
Sheet.Cells(3, 15) = "Greateast % Decrease"
Sheet.Cells(4, 15) = "Greateast Total Volume"

' how many different tickers we have?
lastticker = Sheet.Cells(Rows.Count, 9).End(xlUp).Row




'GREATEST INCREASE
'Start in first row
firstRow = 2


For i = firstRow To lastticker
'finding out maximum value for greatest increase
    If Sheet.Cells(i, 11) > max Then max = Sheet.Cells(i, 11)
'finding ticket that correspond to greatest increase
    If max = Sheet.Cells(i, 11) Then ticker = Sheet.Cells(i, 9)
    
Next

'writing down greatest increase
Sheet.Cells(2, 17) = max
    
'writing down ticker with greatest increase
Sheet.Cells(2, 16) = ticker





'GREATEST DECREASE
'initiate first row
firstRow = 2

For i = firstRow To lastticker
'finding out minimum value for greatest decrease
    If Sheet.Cells(i, 11) < min Then min = Sheet.Cells(i, 11)
'finding ticket that correspond to greatest decrease
    If min = Sheet.Cells(i, 11) Then ticker = Sheet.Cells(i, 9)
   
Next

'writing down greatest decrease
Sheet.Cells(3, 17) = min

'writing down ticker with greatest decrease
Sheet.Cells(3, 16) = ticker





'GREATEST TOTAL VOLUME
'initiate first row
firstRow = 2


For i = firstRow To lastticker
'finding out maximum Total Volume
    If Sheet.Cells(i, 12) > max Then max = Sheet.Cells(i, 12)
'finding ticker with greatest total volume
    If max = Sheet.Cells(i, 12) Then ticker = Sheet.Cells(i, 9)

Next


'writing down ticker with greatest Total Volume
Sheet.Cells(4, 16) = ticker

'writing down maximum total volume
Sheet.Cells(4, 17) = max

'reset variables
lastticker = 0
max = 0
min = 0

Next


End Sub

