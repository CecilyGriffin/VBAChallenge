Sub ABC()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"


Dim tickerRange As Long
tickerRange = Range("A2").End(xlDown).Row

Dim rowCount As Long
rowCount = 2

Dim tickerString As String
tickerString = ""


Dim year As Long
Dim monthDay As Long



For i = 1 To tickerRange

'Populate column i with distinct ticker's
If (Cells(i + 1, 1).Value <> tickerString) Then
    tickerString = Cells(i + 1, 1).Value
    Cells(rowCount, 9).Value = tickerString
    rowCount = rowCount + 1
End If
    
Next i


Dim tickerLoop As Long
tickerLoop = Range("I2").End(xlDown).Row

Dim ticker As String 'Current column I ticker value
Dim begYearOpen As Double
Dim endYearClose As Double
Dim volumne As Long



For i = 1 To tickerLoop
    ticker = Cells(i + 1, 9).Value
    If (ticker = "") Then
        GoTo Continue
    End If


    Dim tickerCount As Long 'Current row in Column A
    Dim endCount As Long
    Dim totalStockVolume As Double
    totalStockVolume = 0
    tickerCount = 1
        For a = 1 To tickerRange
            If (Cells(a + 1, 1).Value = ticker) Then
                If (tickerCount = 1) Then
                    begYearOpen = Cells(a + 1, 3).Value

                End If
                tickerCount = tickerCount + 1
                endCount = a
               totalStockVolume = Cells(a + 1, 7).Value + totalStockVolume
            End If

        Next a

        Cells(i + 1, 12).Value = totalStockVolume

        endYearClose = Cells(endCount + 1, 6).Value

        Dim percentChange As String

        If (begYearOpen = endYearClose) Then
            percentChange = 0
        ElseIf (begYearOpen = 0) Then
            percentChange = 100
        Else
            percentChange = (endYearClose - begYearOpen) / begYearOpen
        End If

        Dim yearlyChange As Double

        yearlyChange = endYearClose - begYearOpen

        If (yearlyChange < 0) Then
        Cells(i + 1, 10).Interior.ColorIndex = 3
        ElseIf (yearlyChange > 0) Then
        Cells(i + 1, 10).Interior.ColorIndex = 4
        End If

        Cells(i + 1, 10).Value = endYearClose - begYearOpen

        Cells(i + 1, 11).Value = FormatPercent(percentChange)

Continue:
Next i


Dim percentRange As Range
Set percentRange = Range(Cells(2, 11), Cells(tickerLoop, 11))
Dim volumeRange As Range
Set volumeRange = Range(Cells(2, 12), Cells(tickerLoop, 12))

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Cells(2, 15).Value = "Greatest % Increase"

Dim percentTest As Double
percentTest = Application.WorksheetFunction.Max(percentRange)

Cells(2, 17).Value = Format(Application.WorksheetFunction.Max(percentRange), "Percent")

Cells(3, 15).Value = "Greatest % Decrease"
Cells(3, 17).Value = Format(Application.WorksheetFunction.Min(percentRange), "Percent")

Cells(4, 15).Value = "Greatest Total Volume"
Cells(4, 17).Value = Format(Application.WorksheetFunction.Max(volumeRange), "General Number")


For i = 1 To tickerLoop

If (Format(Cells(i + 1, 11).Value, "Percent") = Format(Application.WorksheetFunction.Max(percentRange), "Percent")) Then
    Cells(2, 16).Value = Cells(i + 1, 9).Value
ElseIf (Format(Cells(i + 1, 11).Value, "Percent") = Format(Application.WorksheetFunction.Min(percentRange), "Percent")) Then
    Cells(3, 16).Value = Cells(i + 1, 9).Value
ElseIf (Cells(i + 1, 12).Value = Application.WorksheetFunction.Max(volumeRange)) Then
    Cells(4, 16).Value = Cells(i + 1, 9).Value
End If


Next i


End Sub





