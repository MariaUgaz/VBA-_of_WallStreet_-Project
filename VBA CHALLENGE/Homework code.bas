Sub homework():

        For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        Dim Ticker As String
        Dim TotalTicker As Double
        TotalTicker = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim YearO As Double
        Dim YearC As Double
        Dim YearCh As Double
        Dim PreAmount As Long
        PreAmount = 2
        Dim PercentChange As Double


        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            TotalTicker = TotalTicker + ws.Cells(i, 7).Value

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                Ticker = ws.Cells(i, 1).Value

                ws.Range("I" & SummaryTableRow).Value = Ticker

                ws.Range("L" & SummaryTableRow).Value = TotalTicker

                TotalTicker = 0

                YearO = ws.Range("C" & PreAmount)
                YearC = ws.Range("F" & i)
                YearCh = YearC - YearO
                ws.Range("J" & SummaryTableRow).Value = YearCh

              
                If YearO = 0 Then
                    PercentChange = 0
                Else
                    YearO = ws.Range("C" & PreAmount)
                    PercentChange = YearCh / YearO
                End If

                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange


                If ws.Range("J" & SummaryTableRow).Value <= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                End If
            

                SummaryTableRow = SummaryTableRow + 1
                PreAmount = i + 1
                End If
            Next i


            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        

            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value

                elseIf ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value

                elseIf ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i

            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            

        ws.Columns("I:Q").AutoFit

    Next ws

End Sub
