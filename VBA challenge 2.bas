Attribute VB_Name = "Module1"
Sub stocks():

'variables
    Dim ticker As String
    Dim i As Long
    Dim j As Integer
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim total_volume As Double
    Dim rowcount As Long
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    'intial values
    j = 2
    Change = 0
    counter = 0
    Start = 0
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'what i want to find
    ws.Range("I1").Value = "ticker"
    ws.Range("J1").Value = "yearly_change"
    ws.Range("K1").Value = "pecentage_change" & "%"
    ws.Range("L1").Value = "total_volume"
    ws.Range("P2").Value = "greatest percent increase"
    ws.Range("P3").Value = "greatest percent decrease"
    ws.Range("P4").Value = "greatest total value"
    ws.Range("P1").Value = "value" & "%"
    ws.Range("O2:O4").Value = "ticker"
    ws.Range("N2").Value = "greatest %"
    ws.Range("N3").Value = "least %"
    ws.Range("N4").Value = "greatest total volume"
    
    'loop code
    For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
    ws.Cells(j, 9) = ticker
    
    'total volume
    counter = 0
    total_volume = counter + ws.Cells(i, 7).Value
    ws.Cells(j, 12) = total_volume
    
    'yearly change
    opening_price = ws.Cells(i + 1, 3).Value
    closing_price = ws.Cells(i, 6).Value
    yearly_change = ws.Cells(i, 6) - ws.Cells(i + 1, 3)
    ws.Cells(j, 10) = yearly_change
    
    'percentage change in years
    percentage_change = ((yearly_change / ws.Cells(i, 3)) * 100)
    ws.Cells(j, 11).Value = percentage_change & "%"
    j = j + 1
    
    End If
    
    If ws.Cells(j, 10).Value < 0 Then
    ws.Cells(j, 10).Interior.ColorIndex = 3
    ElseIf Cells(j, 10).Value > 0 Then
    ws.Cells(j, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(j, 10).Interior.ColorIndex = 0
    End If
    
    If ws.Cells(j, 11).Value < 0 Then
    ws.Cells(j, 11).Interior.ColorIndex = 3
    ElseIf Cells(j, 11).Value > 0 Then
    ws.Cells(j, 11).Interior.ColorIndex = 4
    Else
    ws.Cells(j, 11).Interior.ColorIndex = 0
    End If
    
    'max and min of percentage
    ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
    ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) * 100
    ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
    
    'match ticker to max and min by setting variables
    ticker_max = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    ticker_min = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    ticker_volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
    
    'match ticker to max and min
    ws.Range("O2") = ws.Cells(ticker_max, 9)
    ws.Range("O3") = ws.Cells(ticker_min, 9)
    ws.Range("O4") = ws.Cells(ticker_volume, 9)
    
    Next i
    
    Next ws
    
    End Sub
