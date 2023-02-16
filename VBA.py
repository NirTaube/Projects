Sub Stocks()

For Each ws In Worksheets

'Define Ranges
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

'Set variable type
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    Dim StockOpen As Double
    Dim StockClose As Double
    Dim lastrow As Double
    
'Locate Last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Value to volume
Volume = 0

'Next defining Summary Table Row
Dim Summary_Table_Row As Double
Summary_Table_Row = 2

'Now Making the loop w/ a conditional
For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        Ticker = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value
        
'Place Ticker and Volume
          ws.Range("I" & Summary_Table_Row).Value = Ticker
          ws.Range("L" & Summary_Table_Row).Value = Volume
          
'place stock close
        StockClose = ws.Cells(i, 6)
       
        If StockOpen = 0 Then
            YearlyChange = 0
            PercentChange = 0
        Else:
            YearlyChange = StockClose - StockOpen
            PercentChange = (StockClose - StockOpen) / StockOpen
        End If

    'Label all the rage values
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
            ws.Range("K" & Summary_Table_Row).Style = "Percent"
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
'Summary table row was set to value of 2, and we want to add 1 to that everytime as we continue
            Summary_Table_Row = Summary_Table_Row + 1

    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
         StockOpen = ws.Cells(i, 3)

'Keeping track of volume
    Else: Volume = Volume + ws.Cells(i, 7).Value

    End If

'Loop with the Value w/ Color, Greater than 0 is green, Else Red
    Next i


For R = 2 To lastrow

    If ws.Range("J" & R).Value > 0 Then
        ws.Range("J" & R).Interior.ColorIndex = 4

    ElseIf ws.Range("J" & R).Value < 0 Then
        ws.Range("J" & R).Interior.ColorIndex = 3
        
    End If


    Next R
    
'Ranges
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Lables
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Set value types (increase "inc" / Decrease "Dec"/ Volume "vol")
Dim GreatestInc As Double
Dim GreatestDec As Double
Dim GreatestVol As Double

'set all to zero
GreatestInc = 0
GreatestDec = 0
GreatestVol = 0


'Make loops for Increase, Decrease, Volume.
'-----Increase----
For A = 2 To lastrow

    If ws.Cells(A, 11).Value > GreatestInc Then
    
        GreatestInc = ws.Cells(A, 11).Value
        ws.Range("Q2").Value = GreatestInc
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = ws.Cells(A, 9).Value
        
    End If

    Next A
'-----Decrease----
For B = 2 To lastrow
    
    If ws.Cells(B, 11).Value < GreatestDec Then
    
        GreatestDec = ws.Cells(B, 11).Value
        ws.Range("Q3").Value = GreatestDec
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = ws.Cells(B, 9).Value
        
    End If
    
   Next B
'----- Volume----
For C = 2 To lastrow
    
    If ws.Cells(C, 12).Value > GreatestVol Then
    
        GreatestVol = ws.Cells(C, 12).Value
        ws.Range("Q4").Value = GreatestVol
        ws.Range("P4").Value = ws.Cells(C, 9).Value
        
    End If
  
    Next C
 
ws.Columns("A:Q").AutoFit
    
Next ws

End Sub

