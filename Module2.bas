Attribute VB_Name = "Module2"
Sub MultipleQuarterStockData()
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim TickCount As Long
    Dim PerChange As Double
    Dim TotalVolume As Double
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim GreatIncr As Double
    Dim GreatDecr As Double
    Dim GreatVol As Double
    Dim GreatIncrTicker As String
    Dim GreatDecrTicker As String
    Dim GreatVolTicker As String

    ' Loops through each worksheet in the workbook
    For Each ws In Worksheets

        ' Creating the column headers required
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        TickCount = 2
        j = 2

        i = 2
        ' Starting at the first data row

        ' Looping until the date column is empty
        Do While ws.Cells(i, 2).Value <> ""
            If i = j Then
                ' Setting up the opening price and initialize total volume for the quarter
                OpeningPrice = ws.Cells(i, 3).Value
                TotalVolume = ws.Cells(i, 7).Value
            Else
                ' Collecting the total volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If

            ' Quarterly or ticker changes
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Or _
               (Year(ws.Cells(i + 1, 2).Value) <> Year(ws.Cells(i, 2).Value)) Or _
               (DatePart("q", ws.Cells(i + 1, 2).Value) <> DatePart("q", ws.Cells(i, 2).Value)) Then

                ' Setting the closing price
                ClosingPrice = ws.Cells(i, 6).Value

                ' Calculating quarterly change
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickCount, 10).Value = ClosingPrice - OpeningPrice

                ' Conditional formatting for Quarterly Change
                If ws.Cells(TickCount, 10).Value < 0 Then
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3 ' Using Red colour
                ElseIf ws.Cells(TickCount, 10).Value > 0 Then
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4 ' Using Green colour
                Else
                    ws.Cells(TickCount, 10).Interior.ColorIndex = xlNone ' No colour for 0
                End If

                ' Calculating and writing percentage change
                If OpeningPrice <> 0 Then
                    PerChange = (ClosingPrice - OpeningPrice) / OpeningPrice
                Else
                    PerChange = 0
                End If
                ws.Cells(TickCount, 11).Value = Format(PerChange, "0.00%")

                ' Calculating and writing total volume for the quarter
                ws.Cells(TickCount, 12).Value = TotalVolume

                ' Reseting for next quarter or ticker
                TickCount = TickCount + 1
                j = i + 1
            End If

            i = i + 1 ' Moving to the next row
        Loop

        ' Handling the greatest values
        Dim LastRowI As Long
        LastRowI = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

        ' Initializing variables for greatest calculations
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        GreatIncrTicker = ws.Cells(2, 9).Value
        GreatDecrTicker = ws.Cells(2, 9).Value
        GreatVolTicker = ws.Cells(2, 9).Value

        ' Calculating the greatest values
        For i = 2 To LastRowI
            If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                GreatVolTicker = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                GreatIncrTicker = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                GreatDecrTicker = ws.Cells(i, 9).Value
            End If
        Next i

        ' Displaying Output the greatest values
        ws.Cells(2, 16).Value = GreatIncrTicker
        ws.Cells(3, 16).Value = GreatDecrTicker
        ws.Cells(4, 16).Value = GreatVolTicker

        ws.Cells(2, 17).Value = Format(GreatIncr, "0.00%")
        ws.Cells(3, 17).Value = Format(GreatDecr, "0.00%")
        ws.Cells(4, 17).Value = GreatVol

        ' Autofitting the columns
        ws.Columns("A:Z").AutoFit
    Next ws
End Sub


