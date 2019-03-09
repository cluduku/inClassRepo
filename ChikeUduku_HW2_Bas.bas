Attribute VB_Name = "Stock_Operations"

'Creator: Chike Uduku
'Created: 03/08/2019
'Desc: This is the main sub routine that drives sequence of operation
'Revisions:1.0
Sub Main()

Dim lastUsedRow As Long 'contains row index of last used row for a given worksheet
Dim startTickerIndex As Long 'contains the row index for the start of a new ticker
Dim ws As Worksheet ' contains thee current worksheet object
Dim count As Long
Dim i As Long 'variable used to iterate over a given worksheet

'Let's loop through the worksheets
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate 'set current worksheet as active worksheet
    Call NameColumns(ws) 'write column headers for the respective worksheet
    
    startTickerIndex = 2 'initialize startTickerIndex to starting row index of first stock ticker
    count = 2 'initialize start of row index where processed stock data will be displayed
    lastUsedRow = ws.Cells(Rows.count, 1).End(xlUp).Row 'find last used row for stock data
    
    'We can now loop through the active worksheet to process stock  tickers
    For i = 2 To (lastUsedRow - 1)
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then 'detect the last row of current stock ticker
        
            'Now that we have range of current stock ticker, let's process this ticker
            Call ProcessStock(ws, startTickerIndex, i, count)
            
            'After current stock ticker is processed, update StartTickerIndex to reflect
            'start of next ticker to be processed
            startTickerIndex = i + 1
            
            'also update count so that we write on the next line after for the given
            'columns after next stock ticker is processed
            count = count + 1
        End If
        
        If (i = (lastUsedRow - 1) And (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value)) Then
            Call ProcessStock(ws, startTickerIndex, (i + 1), count)
        End If
        
    Next i
    
    'Now that we are done processing stocks, let's find our max and min % increase as well as
    'our stock with greatest total volume
    Call FindLimits(ws)
Next
               
End Sub

'Creator: Chike Uduku
'Created: 03/08/2019
'Desc: This sub routine names the headers for columns containing processed stock data
'Revision:1.0
Sub NameColumns(myWs As Worksheet)
myWs.Range("I1").Value = "Ticker"
myWs.Range("J1").Value = "Yearly Change"
myWs.Range("K1").Value = "Percent Change"
myWs.Range("L1").Value = "Stock Volume"
myWs.Range("O1").Value = "Ticker"
myWs.Range("P1").Value = "Value"
myWs.Range("N2").Value = "Greatest % Increase"
myWs.Range("N3").Value = "Greatest % Decrease"
myWs.Range("N4").Value = "Greatest total volume"
myWs.Range("$K:$K").NumberFormat = "0.00%"
myWs.Range("P2").NumberFormat = "0.00%"
myWs.Range("P3").NumberFormat = "0.00%"
End Sub

'Creator:Chike Uduku
'Created: 03/09/2019
'Desc: This sub routine processes a given stock ticker and displays the results on the sheet
'Revisions:1.0
Sub ProcessStock(myWs As Worksheet, startIndex As Long, endIndex As Long, writeIndex As Long)
'Dim sumRange As Range

'write the ticker
myWs.Cells(writeIndex, 9).Value = myWs.Cells(startIndex, 1)

'Solve for yearly change
myWs.Cells(writeIndex, 10).Value = myWs.Cells(endIndex, 6) - myWs.Cells(startIndex, 3)

'Now that we have yearly change values, let's format the cell color based on that value
If (myWs.Cells(writeIndex, 10).Value > 0) Then
    myWs.Cells(writeIndex, 10).Interior.ColorIndex = 4 'green
ElseIf (myWs.Cells(writeIndex, 10).Value < 0) Then
    myWs.Cells(writeIndex, 10).Interior.ColorIndex = 3 'red
End If

'Let's solve for percent change
If (myWs.Cells(startIndex, 3) <> 0) Then
    myWs.Cells(writeIndex, 11).Value = myWs.Cells(writeIndex, 10).Value / myWs.Cells(startIndex, 3)
End If

'Let's solve for total stock volume
'sumRange =
myWs.Cells(writeIndex, 12).Value = Application.WorksheetFunction.Sum(myWs.Range(myWs.Cells(startIndex, 7), myWs.Cells(endIndex, 7)))

End Sub

'Creator: Chike Uduku
'Created: 03/09/2019
'Desc: This function looks at processed stock tickers to find max and min % increase as well as
        'stock with greatest  total volume
'Revisions: 1.0
Sub FindLimits(myWs As Worksheet)

Dim endRow As Long 'contains last row for processed stock tickers on a given worksheet
Dim j As Long 'iteration variable

endRow = myWs.Cells(Rows.count, 12).End(xlUp).Row 'find last used row for processed data
myWs.Cells(2, 16).Value = 0 'initialize greatest % increase
myWs.Cells(3, 16).Value = 1000000 'initialize greatest % decrease
myWs.Cells(4, 16).Value = 0 'initialize greatest total volume

For j = 2 To endRow
    'Find max % increase stock
    If (myWs.Cells(j, 11).Value > myWs.Cells(2, 16).Value) Then
        myWs.Cells(2, 16).Value = myWs.Cells(j, 11).Value
        myWs.Cells(2, 15).Value = myWs.Cells(j, 9).Value
    End If
    'Find max % decrease stock
    If (myWs.Cells(j, 11).Value < myWs.Cells(3, 16).Value) Then
        myWs.Cells(3, 16).Value = myWs.Cells(j, 11).Value
        myWs.Cells(3, 15).Value = myWs.Cells(j, 9).Value
    End If
    'find greatest tota volume stock
    If (myWs.Cells(j, 12).Value > myWs.Cells(4, 16).Value) Then
        myWs.Cells(4, 16).Value = myWs.Cells(j, 12).Value
        myWs.Cells(4, 15).Value = myWs.Cells(j, 9).Value
    End If
Next j

End Sub

