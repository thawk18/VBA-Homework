Attribute VB_Name = "Module3"
Sub format()

'Define variables
Dim i As Integer
Dim ntickers As Integer

For Each Sheet In Worksheets

'Limit number of cells with tickers
nrtickers = Sheet.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To nrtickers

'Conditional formating
If Sheet.Cells(i, 10).Value >= 0 Then

Sheet.Cells(i, 10).Interior.ColorIndex = 4

Else

Sheet.Cells(i, 10).Interior.ColorIndex = 3

End If

'Format cells to correct units
Sheet.Cells(i, 10).NumberFormat = "$0.00"
Sheet.Cells(i, 11).NumberFormat = "0.00%"
Sheet.Cells(i, 12).NumberFormat = "$0"

Next i

'Format cells to correct units
Sheet.Cells(2, 17).NumberFormat = "0.00%"
Sheet.Cells(3, 17).NumberFormat = "0.00%"
Sheet.Cells(4, 17).NumberFormat = "$0"

'reset variables for next sheet
ntickers = 0
i = 2

Next


End Sub

