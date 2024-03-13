Attribute VB_Name = "Module1"
Sub ModuleChallenge()
    ' Loop through worksheets
    For Each ws In Worksheets
        ' define variables
        Dim i, tickerrow As Long
        Dim op, cp, vol, yearly As Double
        Dim lastrow As Long
        ' setting start for row,open price and volume
        tickerrow = 2
        op = ws.Cells(2, 3)
        vol = 0
        ' Last row in sheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        ' Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        For i = 2 To lastrow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Getting Ticker Names and printing
                ws.Cells(tickerrow, 9).Value = ws.Cells(i, 1).Value
                ' Getting close price,yearly change, and percent change
                cp = ws.Cells(i, 6).Value
                yearly = cp - op
                If yearly < 0 Then
                    ws.Cells(tickerrow, 10).Interior.ColorIndex = 3
                    ws.Cells(tickerrow, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(tickerrow, 10).Interior.ColorIndex = 4
                    ws.Cells(tickerrow, 11).Interior.ColorIndex = 4
                End If
                ws.Cells(tickerrow, 10).Value = yearly
                ws.Cells(tickerrow, 11).Value = (yearly / op)
                ' Make Percentage format (formula from Stack overflow)
                ws.Cells(tickerrow, 11).NumberFormat = "0.00%"
                ' Adding stock volume and printing
                vol = vol + ws.Cells(i, 7).Value
                ws.Cells(tickerrow, 12).Value = vol
                ' Iterating through rows
                tickerrow = tickerrow + 1
                ' Setting open price to opening of next stock
                op = ws.Cells(i + 1, 3)
                ' Reset vol to 0 when summing new stock
                vol = 0
            Else
                ' Need else to sum between same stock
                vol = vol + ws.Cells(i, 7).Value
            End If
        Next i
        Dim max, min, total As Double
        ' Headers for Greater Functionality
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ' set max min total to 0
        max = 0
        min = 0
        total = 0
        For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
            ' Max Percent Change
            If ws.Cells(i, 11) > max Then
                max = Cells(i, 11)
                ws.Cells(2, 16).Value = max
                ws.Cells(2, 16).NumberFormat = "0.00%"
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
            ' Min Percent Change
            ElseIf ws.Cells(i, 11) < min Then
                min = ws.Cells(i, 11)
                ws.Cells(3, 16).Value = min
                ws.Cells(3, 16).NumberFormat = "0.00%"
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
            ' Total Volume
            ElseIf ws.Cells(i, 12) > total Then
                total = Cells(i, 12)
                ws.Cells(4, 16).Value = total
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
            End If
        Next i
    Next ws
End Sub

