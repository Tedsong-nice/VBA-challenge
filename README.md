# VBA-challenge
Sub StockTicker()
    Dim ws As Worksheet

    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        Dim Lastrow As Long
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        Dim Ticker As String
        Dim QuartlyChange As Double
        Dim PercentageChange As Double
        Dim TotalVolume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim CurrentQuarter As String
        Dim PreviousQuarter As String
        Dim OutputRow As Long
        OutputRow = 2

        ws.Cells(1, 8).Value = "Quarter"
        Dim i As Long
        For i = 2 To Lastrow
            ws.Cells(i, 8).Value = Year(ws.Cells(i, 2).Value) & "Q" & Application.WorksheetFunction.RoundUp(Month(ws.Cells(i, 2).Value) / 3, 0)
        Next i

        ws.Sort.SortFields.Clear
        ws.Sort.SortFields.Add Key:=Range("H2:H" & Lastrow), Order:=xlAscending
        ws.Sort.SortFields.Add Key:=Range("A2:A" & Lastrow), Order:=xlAscending
        With ws.Sort
            .SetRange Range("A1:H" & Lastrow)
            .Header = xlYes
            .Apply
        End With
        PreviousQuarter = ""
        TotalVolume = 0

        Dim MaxIncrease As Double
        Dim MaxDecrease As Double
        Dim MaxVolume As Double
        Dim MaxIncreaseTicker As String
        Dim MaxDecreaseTicker As String
        Dim MaxVolumeTicker As String

        MaxIncrease = -1
        MaxDecrease = 1
        MaxVolume = 0

        For i = 2 To Lastrow
            Ticker = ws.Cells(i, 1).Value
            CurrentQuarter = ws.Cells(i, 8).Value

            If CurrentQuarter <> PreviousQuarter Or Ticker <> ws.Cells(i - 1, 1).Value Then
                If i > 2 Then
                    QuartlyChange = ClosePrice - OpenPrice
                    PercentageChange = (QuartlyChange / OpenPrice)
                    ws.Cells(OutputRow, 9).Value = ws.Cells(i - 1, 1).Value
                    ws.Cells(OutputRow, 10).Value = QuartlyChange
                    ws.Cells(OutputRow, 11).Value = PercentageChange
                    ws.Cells(OutputRow, 12).Value = TotalVolume
                    If QuartlyChange > 0 Then
                        ws.Cells(OutputRow, 10).Interior.Color = RGB(0, 255, 0)
                    ElseIf QuartlyChange < 0 Then
                        ws.Cells(OutputRow, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Cells(OutputRow, 10).Interior.Color = RGB(255, 255, 255)
                    End If
    
                    ws.Cells(OutputRow, 11).NumberFormat = "0.00%"

                    OutputRow = OutputRow + 1
                    If PercentageChange > MaxIncrease Then
                        MaxIncrease = PercentageChange
                        MaxIncreaseTicker = ws.Cells(i - 1, 1).Value
                    End If
                    If PercentageChange < MaxDecrease Then
                        MaxDecrease = PercentageChange
                        MaxDecreaseTicker = ws.Cells(i - 1, 1).Value
                    End If
                    If TotalVolume > MaxVolume Then
                        MaxVolume = TotalVolume
                        MaxVolumeTicker = ws.Cells(i - 1, 1).Value
                    End If
                End If
                OpenPrice = ws.Cells(i, 3).Value
                TotalVolume = ws.Cells(i, 7).Value
            Else
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
            ClosePrice = ws.Cells(i, 6).Value
            PreviousQuarter = CurrentQuarter
        Next i
        QuartlyChange = ClosePrice - OpenPrice
        PercentageChange = (QuartlyChange / OpenPrice) * 100
        ws.Cells(OutputRow, 9).Value = ws.Cells(Lastrow, 1).Value
        ws.Cells(OutputRow, 10).Value = QuartlyChange
        ws.Cells(OutputRow, 11).Value = PercentageChange
        ws.Cells(OutputRow, 12).Value = TotalVolume
        If QuartlyChange > 0 Then
            ws.Cells(OutputRow, 10).Interior.Color = RGB(0, 255, 0) 
        ElseIf QuartlyChange < 0 Then
            ws.Cells(OutputRow, 10).Interior.Color = RGB(255, 0, 0) 
        Else
            ws.Cells(OutputRow, 10).Interior.Color = RGB(255, 255, 255) 
        End If
        ws.Cells(OutputRow, 11).NumberFormat = "0.00%"
        If PercentageChange > MaxIncrease Then
            MaxIncrease = PercentageChange
            MaxIncreaseTicker = ws.Cells(Lastrow, 1).Value
        End If
        If PercentageChange < MaxDecrease Then
            MaxDecrease = PercentageChange
            MaxDecreaseTicker = ws.Cells(Lastrow, 1).Value
        End If
        If TotalVolume > MaxVolume Then
            MaxVolume = TotalVolume
            MaxVolumeTicker = ws.Cells(Lastrow, 1).Value
        End If

        
        ws.Range("H1:H" & Lastrow).ClearContents
        ws.Cells(2, 16).Value = MaxIncreaseTicker
        ws.Cells(2, 17).Value = MaxIncrease
        ws.Cells(3, 16).Value = MaxDecreaseTicker
        ws.Cells(3, 17).Value = MaxDecrease
        ws.Cells(4, 16).Value = MaxVolumeTicker
        ws.Cells(4, 17).Value = MaxVolume

    Next ws
End Sub
