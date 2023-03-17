

Sub stockSummary()



    Dim tickerSymbol As String

    Dim stockVolume As Double

    Dim openPrice As Long

    Dim closePrice As Double

    Dim yearlyChange As Double

    Dim percentage As Double

    Dim summaryTableRow As Integer

    Dim greatestIncrease As Double

    Dim greatestDecrease As Double

    Dim greatestVolume As Double

    'Row starts here, this is why it starts at 2 becausec 1 is the header
    summaryTableRow = 2

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Start Price"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    openPrice = Cells(2, "C").Value
    Cells(2, 15).Value = "Greatest Percent Increase"
    Cells(3, 15).Value = "Greatest Percent Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"


    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            tickerSymbol = Cells(i, 1).Value

            stockVolume = stockVolume + Cells(i, 7).Value

            'print ticker name
            Range("I" & summaryTableRow).Value = tickerSymbol

            'print total value
            Range("M" & summaryTableRow).Value = stockVolume

            tickerSymbol = 0

            closePrice = Cells(i, 6).Value

            yearlyChange = closePrice - openPrice

        If openPrice = 0 Then

            percentage = NA

        Else

            percentage = yearlyChange / openPrice

        End If

            'print yearly change
            Range("K" & summaryTableRow).Value = yearlyChange

            Range("L" & summaryTableRow).Value = FormatPercent(percentage)

            openPrice = Cells(i + 1, 3).Value

            stockVolume = 0

            summaryTableRow = summaryTableRow + 1

        Else

            stockVolume = stockVolume + Cells(i, 7).Value


    End If

    Next i

        greatestIncrease = WorksheetFunction.Max(Range("L:L"))
        greatestDecrease = WorksheetFunction.Min(Range("L:L"))
        greatestVolume = WorksheetFunction.Max(Range("M:M"))

        Range("Q2").Value = FormatPercent(greatestIncrease)
        Range("Q3").Value = FormatPercent(greatestDecrease)
        Range("Q4").Value = greatestVolume

For i = 2 To lastRow

    'Auto-coloring the yearly chance column accordingly
    If Cells(i, 11).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
    Else
        Cells(i, 11).Interior.ColorIndex = 3
    End If


    If greatestIncrease = Cells(i, 12).Value Then
        Range("P2").Value = Cells(i, 9).Value
    End If

    If greatestDecrease = Cells(i, 12).Value Then
        Range("P3").Value = Cells(i, 9).Value
    End If

    If greatestVolume = Cells(i, 13).Value Then
        Range("P4").Value = Cells(i, 9).Value
    End If

Next i


End Sub




