Attribute VB_Name = "Module1"
Sub CountST()

For Each ws In Worksheets

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
'Set colors based on %
    If ws.Range("J" & STRow).Value < 0 Then
        ws.Range("J" & STRow).Interior.ColorIndex = 3
    ElseIf ws.Range("J" & STRow).Value > 0 Then
        ws.Range("J" & STRow).Interior.ColorIndex = 4
    End If
    ws.Range("L" & STRow).Value = StockTotal
    If OSP_Amount = 0 Then
    ws.Range("K" & STRow).Value = 0
    Else
    ws.Range("K" & STRow).Value = (CSP / OSP_Amount) - 1
    ws.Range("K2:K" & STRow).NumberFormat = "0.00%"
    End If
    
    STRow = STRow + 1
             
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
    
    ' Run the bonus here
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"


Dim Biggest, Smallest As Double, Most As Double, Iticker As String, BiggestRow, SmallestRow, MostRow As Integer
    'Find Largest % Change
    Biggest = WorksheetFunction.Max(ws.Range("K2:K" & STRow))
    BiggestRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & STRow)), ws.Range("K2:K" & STRow), 0)
    BiggestRow = BiggestRow + 1
    ws.Range("P2") = Biggest
    Iticker = ws.Range("I" & BiggestRow).Value
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("O2").Value = Iticker
    
    'Find Smallest % Change
    Smallest = WorksheetFunction.Min(ws.Range("K2:K" & STRow))
    SmallestRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & STRow)), ws.Range("K2:K" & STRow), 0)
    SmallestRow = SmallestRow + 1
    ws.Range("P3") = Smallest
    Iticker = ws.Range("I" & SmallestRow).Value
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("O3").Value = Iticker
    
    'Find Largest Stock Volume
    Most = WorksheetFunction.Max(ws.Range("L2:L" & STRow))
    MostRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & STRow)), ws.Range("L2:L" & STRow), 0)
    MostRow = MostRow + 1
    ws.Range("P4") = Most
    Iticker = ws.Range("I" & MostRow).Value
    ws.Range("O4").Value = Iticker
    ws.Columns("N:P").EntireColumn.AutoFit
    
Next ws

End Sub



Sub CleanSheets():

For Each ws In Worksheets

ws.Range("I:P").Delete

Next ws

End Sub


