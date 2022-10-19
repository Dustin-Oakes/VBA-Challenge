Attribute VB_Name = "Module11"
Sub MarketScript()

    'Loop through each worksheet
    '''''
    Dim wrkSheet As Worksheet
    
    For Each wrkSheet In ThisWorkbook.Worksheets
        wrkSheet.Activate

        'Dim variables
        '''''
        Dim readRow As Double
        Dim writeRow As Integer
        Dim tickerValue As String
        Dim startValue As Double
        Dim endValue As Double
        Dim totalVolume As Double
        Dim rowCount As Double
        Dim changePercent As Double
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestTotalVolume As Double
        Dim greatestIncreaseTicker As String
        Dim greatestDecreaseTicker As String
        Dim greatestTotalVolumeTicker As String

        'Initialize Variables
        '''''
        writeRow = 2
        tickerValue = Cells(2, 1)
        startValue = Cells(2, 3)
        totalVolume = 0
        rowCount = WorksheetFunction.CountA(Range("A:A"))
        greatestIncrease = 0
        greatestDecrease = 0
        greatestTotalVolume = 0

        'Main For loop to read and write info
        '''''
        For readRow = 2 To rowCount + 1

            If Cells(readRow, 1) = tickerValue Then
                totalVolume = totalVolume + Cells(readRow, 7)
   
            Else

                'Write old vlaues
                '''''
                endValue = Cells((readRow - 1), 6)
                changePercent = ((endValue - startValue) / startValue)
        
                Cells(writeRow, 9) = tickerValue
                Cells(writeRow, 10) = endValue - startValue
                Cells(writeRow, 11) = changePercent
                Cells(writeRow, 12) = totalVolume
        
                'Check for new Min/Max
                '''''
                If changePercent > greatestIncrease Then
                    greatestIncrease = changePercent
                    greatestIncreaseTicker = tickerValue
                End If
        
                If changePercent < greatestDecrease Then
                    greatestDecrease = changePercent
                    greatestDecreaseTicker = tickerValue
                End If
        
                If totalVolume > greatestTotalVolume Then
                    greatestTotalVolume = totalVolume
                    greatestTotalVolumeTicker = tickerValue
                End If
        
                'Format Column 10 based on Value
                '''''
                If Cells(writeRow, 10) > 0 Then
                    Range("J" & writeRow).Interior.Color = RGB(0, 255, 0)
            
                ElseIf Cells(writeRow, 10) < 0 Then
                    Range("J" & writeRow).Interior.Color = RGB(255, 0, 0)
            
                End If
        
                'Set new values
                '''''
                startValue = Cells(readRow, 3)
                tickerValue = Cells(readRow, 1)
                totalVolume = Cells(readRow, 7)
        
                'Iterate writeRow
                '''''
                writeRow = writeRow + 1

            End If

        Next readRow

        'Create Headers
        '''''
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"
        Cells(2, 15) = "Greatest % Increase"
        Cells(3, 15) = "Greatest % Decrease"
        Cells(4, 15) = "Greatest Total Volume"

        'Write Min/Max
        '''''
        Cells(2, 16) = greatestIncreaseTicker
        Cells(2, 17) = greatestIncrease
        Cells(3, 16) = greatestDecreaseTicker
        Cells(3, 17) = greatestDecrease
        Cells(4, 16) = greatestTotalVolumeTicker
        Cells(4, 17) = greatestTotalVolume

        'Sheet formatting
        '''''
        Columns("J:J").Select
        Selection.NumberFormat = "0.00"
        Columns("K:K").Select
        Selection.NumberFormat = "0.00%"
        Columns("L:L").Select
        Selection.NumberFormat = "0"
        Range("Q2:Q3").Select
        Selection.NumberFormat = "0.00%"
        Range("Q4").Select
        Selection.NumberFormat = "0"
        Cells.Select
        Cells.EntireColumn.AutoFit

    Next wrkSheet

End Sub
