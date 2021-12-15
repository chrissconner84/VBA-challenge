Attribute VB_Name = "Module1"
Sub CountST()

'Dim the variables

Dim StockSym As String
Dim STRow As Integer
Dim StockTotal As Double
Dim OSP As Double
Dim CSP As Double
Dim OSP_Counter As Long
Dim OSP_Amount As Double

'Set initial header names

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Columns("I:L").EntireColumn.AutoFit

'Set initial variables values

STRow = 2
StockTotal = 0
OSP = 0
CSP = 0
te = ws.Cells(Rows.Count, 1).End(xlUp).Row
OSP_Counter = 2
    
'Loop through all rows

    For t = 2 To te
 
    If ws.Cells(t + 1, 1).Value <> ws.Cells(t, 1).Value Then
     
    StockSym = ws.Cells(t, 1).Value
    
    OSP_Amount = ws.Cells((t + 2) - OSP_Counter, 3).Value

    StockTotal = StockTotal + ws.Cells(t, 7).Value

    ws.Range("I" & STRow).Value = StockSym
    ws.Range("J" & STRow).Value = CSP - OSP_Amount

             
      ' Reset the variables
      StockTotal = 0
      OSP = 0
      'CSP = 0
      OSP_Counter = 2

    ' If the cell immediately following a row is the ticker
    Else
    'increment the ticker counter
    OSP_Counter = OSP_Counter + 1
      
      ' Add to the Stock volume total
      StockTotal = StockTotal + ws.Cells(t, 7).Value
     
     'get closing stock price from last record in set
     CSP = ws.Cells(t + 1, 6).Value

    End If

    Next t
    
 
    


End Sub



Sub CleanSheets():

For Each ws In Worksheets

ws.Range("I:P").Delete

Next ws

End Sub


