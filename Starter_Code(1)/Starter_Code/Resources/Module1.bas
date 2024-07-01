Attribute VB_Name = "Module1"
Sub forEachWs()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Call Stocks(ws)
    Next
End Sub

Sub Stocks(ws As Worksheet)
    With ws
        .Cells(1, 9).Value = "Ticker"
        .Cells(1, 10).Value = "Quarterly Change"
        .Cells(1, 11).Value = "Percent Change"
        .Cells(1, 12).Value = "Total Stock Volume"

        Dim currentTicker As String
        Dim startPrice As Double
        Dim endPrice As Double
        Dim volume As Double
        Dim TickerID As Long
        Dim g As Long
        Dim i As Long
        
        g = 2
        i = 2
        currentTicker = .Cells(2, 1).Value
        startPrice = .Cells(2, 3).Value
        volume = 0

        Do While Not IsEmpty(.Cells(i, 1).Value)
            If .Cells(i, 1).Value <> currentTicker Then
                .Cells(g, 9).Value = currentTicker
                .Cells(g, 10).Value = endPrice - startPrice
                .Cells(g, 11).Value = ((endPrice - startPrice) / startPrice)
                .Cells(g, 12).Value = volume
                g = g + 1
                
                currentTicker = .Cells(i, 1).Value
                startPrice = .Cells(i, 3).Value
                volume = 0
            End If
            
            endPrice = .Cells(i, 6).Value
            volume = volume + .Cells(i, 7).Value
            
            i = i + 1
        Loop
        
        .Cells(g, 9).Value = currentTicker
        .Cells(g, 10).Value = endPrice - startPrice
        .Cells(g, 11).Value = ((endPrice - startPrice) / startPrice) * 100
        .Cells(g, 12).Value = volume
        
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"
        .Cells(2, 15).Value = "Greatest % Increase"
        .Cells(3, 15).Value = "Greatest % Increase"
        .Cells(4, 15).Value = "Greatest Total Volume"
        
        Dim maxInc As Double
        Dim maxDec As Double
        Dim maxVol As Double
        
        Dim tickerA As String
        Dim tickerB As String
        Dim tickerC As String
        
        maxInc = .Cells(2, 11).Value
        maxDec = .Cells(2, 11).Value
        maxVol = .Cells(2, 12).Value
        
        i = 2
        Do While Not IsEmpty(.Cells(i, 9).Value)
            If .Cells(i, 11).Value > maxInc Then
                maxInc = .Cells(i, 11).Value
                tickerA = .Cells(i, 9).Value
            ElseIf .Cells(i, 11).Value < maxDec Then
                maxDec = .Cells(i, 11).Value
                tickerB = .Cells(i, 9).Value
            End If
            If .Cells(i, 12).Value > maxVol Then
                maxVol = .Cells(i, 12).Value
                tickerC = .Cells(i, 9).Value
            End If
            i = i + 1
        Loop
        .Cells(2, 16).Value = tickerA
        .Cells(3, 16).Value = tickerB
        .Cells(4, 16).Value = tickerC
        .Cells(2, 17).Value = maxInc
        .Cells(3, 17).Value = maxDec
        .Cells(4, 17).Value = maxVol
    End With
End Sub

