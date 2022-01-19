Sub WallStreetVBA()

For Each ws In Worksheets

    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
Volume = 0
    
    Dim StockOpen As Double
    Dim StockClose As Double
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
   ' Had to Google '
    Dim LastRow As Double
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim index As Double
index = 2

For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        Ticker = ws.Cells(i, 1).Value
        Volume = ws.Cells(i, 7).Value + Volume

        ws.Range("I" & index).Value = Ticker
        ws.Range("L" & index).Value = Volume

        Volume = 0

        StockClose = ws.Cells(i, 6)
       
        If StockOpen = 0 Then
            YearlyChange = 0
            PercentChange = 0
        Else:
            YearlyChange = StockClose - StockOpen
            PercentChange = (StockClose - StockOpen) / StockOpen
        
        End If

            ws.Range("J" & index).Value = YearlyChange
            ws.Range("K" & index).Value = PercentChange

            index = index + 1

    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
         StockOpen = ws.Cells(i, 3)

    Else: Volume = ws.Cells(i, 7).Value + Volume

    End If

    Next i

For k = 2 To LastRow

    If ws.Range("J" & k).Value > 0 Then
        ws.Range("J" & k).Interior.ColorIndex = 4

    ElseIf ws.Range("J" & k).Value < 0 Then
        ws.Range("J" & k).Interior.ColorIndex = 3
        
    End If

    Next k

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

For a = 2 To LastRow

    If ws.Cells(a, 11).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(a, 11).Value
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("P2").Value = ws.Cells(a, 9).Value
    End If

    Next a

For b = 2 To LastRow
    
    If ws.Cells(b, 11).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(b, 11).Value
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("P3").Value = ws.Cells(b, 9).Value
    End If
    
   Next b

For c = 2 To LastRow
    
    If ws.Cells(c, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(c, 12).Value
        ws.Range("Q4").Value = GreatestVolume
        ws.Range("P4").Value = ws.Cells(c, 9).Value
    End If
  
    Next c
    
Next ws

End Sub