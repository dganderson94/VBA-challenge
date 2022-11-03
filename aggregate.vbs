Sub Aggregate()
    'Declare variables
        'opening price
        Dim openPrice As Double
        'total stock volume
        Dim totalVol As Double
        'stock count
        Dim stockCount As Integer
        'greatest % increase value & ticker
        Dim greatInc As Double
        Dim greatIncTic As String
        'greatest %decrease value & ticker
        Dim greatDec As Double
        Dim greatDecTic As String
        'greatest total volume value & ticker
        Dim greatTotVol As Double
        Dim greatTotVolTic As String
    'Loop through worksheets
    For j = 1 To Worksheets.Count
        Worksheets(j).Activate
        'Set variables to default values
            'opening price should be set to first opening price
            openPrice = Cells(2, 3).Value
            totalVol = 0
            stockCount = 0
            greatInc = 0
            greatDec = 0
            greatTotVol = 0
        'Set headers
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
        'For loop that runs to the spreadsheet length
        For i = 2 To Range("A1").End(xlDown).Row
            'add to running total stock volume
            totalVol = totalVol + Cells(i, 7).Value
            'if the next row has a different ticker symbol
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                'increment stock count
                stockCount = stockCount + 1
                'print ticker symbol
                Cells(stockCount + 1, 9).Value = Cells(i, 1).Value
                'calculate and print yearly change
                Cells(stockCount + 1, 10).Value = Cells(i, 6).Value - openPrice
                'format yearly change to relevant color
                If Cells(stockCount + 1, 10).Value > 0 Then
                    Cells(stockCount + 1, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf Cells(stockCount + 1, 10).Value < 0 Then
                    Cells(stockCount + 1, 10).Interior.Color = RGB(255, 0, 0)
                Else
                    Cells(stockCount + 1, 10).Interior.Color = RGB(255, 255, 255)
                End If
                'calculate and print percent change
                Cells(stockCount + 1, 11).Value = Cells(stockCount + 1, 10).Value / openPrice
                Cells(stockCount + 1, 11).NumberFormat = "0.00%"
                'print total stock volume
                Cells(stockCount + 1, 12).Value = totalVol
                'check row against greatest rows
                If Cells(stockCount + 1, 11).Value > greatInc Then
                    greatInc = Cells(stockCount + 1, 11).Value
                    greatIncTic = Cells(stockCount + 1, 9).Value
                End If
                If Cells(stockCount + 1, 11).Value < greatDec Then
                    greatDec = Cells(stockCount + 1, 11).Value
                    greatDecTic = Cells(stockCount + 1, 9).Value
                End If
                If Cells(stockCount + 1, 12).Value > greatTotVol Then
                    greatTotVol = Cells(stockCount + 1, 12).Value
                    greatTotVolTic = Cells(stockCount + 1, 9).Value
                End If
                'set opening price variable to next opening price
                openPrice = Cells(i + 1, 3).Value
                'set total stock volume to zero
                totalVol = 0
            End If
        Next i
        'Fill out greatest stocks
            Cells(2, 16).Value = greatIncTic
            Cells(2, 17).Value = greatInc
            Cells(3, 16).Value = greatDecTic
            Cells(3, 17).Value = greatDec
            Cells(4, 16).Value = greatTotVolTic
            Cells(4, 17).Value = greatTotVol
        'Set column widths
            Range("I1:L1").EntireColumn.AutoFit
            Range("O1:Q1").EntireColumn.AutoFit
        'Set percent format
            Range("Q2:Q3").NumberFormat = "0.00%"
    Next j
End Sub
