Attribute VB_Name = "Module1"
Sub StockData()
    
Dim ws As Worksheet
Dim Summary_Table_Row As Long
Dim ticker As String
Dim lastrow As Double
Dim openyear As Single
Dim closeyear As Single
Dim volume As Double

For Each ws In Sheets
        
lastrow = WorksheetFunction.CountA(ActiveSheet.Columns(1))
volume = 0
Summary_Table_Row = 2
        
'Define headers
    
    ws.[I1] = "Ticker"
    ws.[J1] = "Yearly Change"
    ws.[K1] = "Percent Change"
    ws.[L1] = "Volume"
    ws.[P1] = "Ticker"
    ws.[Q1] = "Value"
    ws.[O2] = "Greatest % Increase"
    ws.[O3] = "Greatest % Decrease"
    ws.[O4] = "Greatest Total Volume"

'Tickers & Volume

For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ticker = ws.Cells(i, 1).Value
 volume = volume + ws.Cells(i, 7).Value
 
 ws.Range("I" & Summary_Table_Row) = ticker
 ws.Range("L" & Summary_Table_Row).Value = volume
 
 Summary_Table_Row = Summary_Table_Row + 1
 
Else:

volume = volume + ws.Cells(i, 7).Value

End If
Next i

            
'Loop through all through rows to assign

Summary_Table_Row = 2
For i = 2 To lastrow
            
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    closeyear = ws.Cells(i, 6).Value
ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    openyear = ws.Cells(i, 3).Value
End If
            
If openyear > 0 And closeyear > 0 Then
 Change = closeyear - openyear
 percentchange = Change / openyear
 ws.Cells(Summary_Table_Row, 10).Value = Change
 ws.Cells(Summary_Table_Row, 11).Value = FormatPercent(percentchange)
    closeyear = 0
    openyear = 0

Summary_Table_Row = Summary_Table_Row + 1

End If
Next i
        
'min and max values
    MaxPercent = WorksheetFunction.Max(ActiveSheet.Columns("k"))
    MinPercent = WorksheetFunction.Min(ActiveSheet.Columns("k"))
    maxvolume = WorksheetFunction.Max(ActiveSheet.Columns("l"))
        
    ws.Range("Q2").Value = FormatPercent(MaxPercent)
    ws.Range("Q3").Value = FormatPercent(MinPercent)
    ws.Range("Q4").Value = maxvolume
        
        
'apply tickers to cells
For i = 2 To lastrow
    
If MaxPercent = ws.Cells(i, 11).Value Then
 ws.Range("P2").Value = ws.Cells(i, 9).Value
ElseIf MinPercent = ws.Cells(i, 11).Value Then
    ws.Range("P3").Value = ws.Cells(i, 9).Value
ElseIf maxvolume = ws.Cells(i, 12).Value Then
    ws.Range("P4").Value = ws.Cells(i, 9).Value
End If
Next i
        
'applying conditionals

For i = 2 To lastrow

If IsEmpty(Cells(i, 10).Value) Then Exit For
If ws.Cells(i, 10).Value > 0 Then
   ws.Cells(i, 10).Interior.ColorIndex = 4
Else
  ws.Cells(i, 10).Interior.ColorIndex = 3

End If
Next i


Next ws
End Sub


