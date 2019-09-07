
Sub stockAnalysis()

'Dim ws As Worksheet - Loop through all worksheets
Dim ws As Worksheet
For Each ws In Worksheets
ws.Select

'Column and Cell Texts
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Dim volumeTotal As Double
Dim YearlyChange As Double
Dim PercentChange As Single
Dim SummaryTable As Double
Dim Start As Long

volumeTotal = 0
YearlyChange = 0
PercentChange = 0
Start = 2
SummaryTable = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            volumeTotal = volumeTotal + Cells(i, 7).Value
            If volumeTotal = 0 Then
            Range("I" & SummaryTable).Value = Cells(i, 1).Value
            Range("J" & SummaryTable).Value = 0
            Range("K" & SummaryTable).Value = "%" & 0
            Range("L" & SummaryTable).Value = 0
            Else
                If Cells(Start, 3) = 0 Then
                    For findValue = Start To i
                        If Cells(findValue, 3).Value <> 0 Then
                            Start = findValue
                            Exit For
                        End If
                    Next findValue
                End If
                
                YearlyChange = (Cells(i, 6) - Cells(Start, 3))
                PercentChange = Round((YearlyChange / Cells(Start, 3) * 100), 2)
            
                Start = i + 1
                Range("I" & SummaryTable).Value = Cells(i, 1).Value
                Range("J" & SummaryTable).Value = Round(YearlyChange, 2)
                Range("K" & SummaryTable).Value = "%" & PercentChange
                Range("L" & SummaryTable).Value = volumeTotal
            
                    If Range("J" & SummaryTable).Value > 0 Then
                        Range("J" & SummaryTable).Interior.ColorIndex = 4
                    Else
                        Range("J" & SummaryTable).Interior.ColorIndex = 3
                    End If

            End If

            SummaryTable = SummaryTable + 1
            volumeTotal = 0
            YearlyChange = 0
            PercentChange = 0

        Else

            volumeTotal = volumeTotal + Cells(i, 7).Value

        End If
        
            
    Next i
    
    'Column and Cell Texts
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    

    Dim Max As Single
    Dim Min As Single
    
    Max = WorksheetFunction.Max(Range("K:K").Value)
    Min = WorksheetFunction.Min(Range("K:K").Value)
    GreatestTotalVolume = WorksheetFunction.Max(Range("L:L").Value)
    
    For j = 2 To LastRow
    If Cells(j, 11).Value = Max Then
    Range("P2").Value = Cells(j, 9).Value
    Range("Q2").Value = "%" & (Max) * 100
    
    ElseIf Cells(j, 11).Value = Min Then
    Range("P3").Value = Cells(j, 9).Value
    Range("Q3").Value = "%" & (Min) * 100
    
    ElseIf Cells(j, 12).Value = GreatestTotalVolume Then
    Range("P4").Value = Cells(j, 9).Value
    Range("Q4").Value = GreatestTotalVolume
    
    End If
    
    Next j

Next ws

End Sub


