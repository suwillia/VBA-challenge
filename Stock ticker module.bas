Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim stockvol As Double, openprice As Double, closeprice As Double, qchange As Double
    Dim Ticker As String
    Dim row As Long, trow As Long, lastrow As Long
    Dim max_vol As Double, maxval As Double, minval As Double
    Dim maxTicker As String, minTicker As String, volTicker As String
    Dim totaltsv As Double
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' Loop through each worksheet and process data separately
    For Each ws In Worksheets
        ' Add headers to the summary section on each worksheet
        With ws
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Quarterly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
            
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
        End With

        ' Initialize variables
        trow = 2
        stockvol = 0
        maxval = -1
        minval = 1
        max_vol = 0

        ' Calculate the last row of data on the current worksheet
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        openprice = ws.Cells(2, 3).Value

        ' Loop through each row of the worksheet to process tickers
        For row = 2 To lastrow
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                ' Store the ticker and calculate quarterly change
                Ticker = ws.Cells(row, 1).Value
                closeprice = ws.Cells(row, 6).Value
                qchange = closeprice - openprice

                ' Write results to the summary table
                ws.Cells(trow, 9).Value = Ticker
                ws.Cells(trow, 10).Value = qchange

                ' Format the quarterly change cell
                If qchange > 0 Then
                    ws.Cells(trow, 10).Interior.ColorIndex = 4 ' Green
                ElseIf qchange < 0 Then
                    ws.Cells(trow, 10).Interior.ColorIndex = 3 ' Red
                Else
                    ws.Cells(trow, 10).Interior.ColorIndex = 0 ' no Change
                End If

                ' Calculate and write percent change
                Dim percentChange As Double
                If openprice <> 0 Then
                    percentChange = qchange / openprice
                    ws.Cells(trow, 11).Value = percentChange
                    ws.Cells(trow, 11).NumberFormat = "0.00%"
                Else
                    ws.Cells(trow, 11).Value = 0
                End If

                ' Calculate total stock volume for this ticker and format cells
                               
                stockvol = stockvol + ws.Cells(row, 7).Value
                    If stockvol = 0 Then
                        ws.Cells(trow, 12).Value = 0
                        ws.Cells(trow, 11).Value = 0
                    Else
                    ws.Cells(trow, 12).Value = stockvol
                    ws.Cells(trow, 12).NumberFormat = "#,###"
                    End If
                ' Update summary values for max/min changes and volume
                If percentChange > maxval Then
                    maxval = percentChange
                    maxTicker = Ticker
                End If
                If percentChange < minval Then
                    minval = percentChange
                    minTicker = Ticker
                End If
                If stockvol > max_vol Then
                    max_vol = stockvol
                    volTicker = Ticker
                End If

                ' Move to the next result row and reset variables
                trow = trow + 1
                stockvol = 0
                openprice = ws.Cells(row + 1, 3).Value
                Else
                ' If Stock volume the same we add accumulate
                stockvol = stockvol + ws.Cells(row, 7).Value
            End If
        Next row

        ' Write the individual summary for the current sheet
        With ws
            .Range("Q2").Value = maxval
            .Range("Q2").NumberFormat = "0.00%"
            .Range("P2").Value = maxTicker

            .Range("Q3").Value = minval
            .Range("Q3").NumberFormat = "0.00%"
            .Range("P3").Value = minTicker

            .Range("Q4").Value = max_vol
            .Range("Q4").NumberFormat = "#,###"
            .Range("P4").Value = volTicker

            ' Autofit columns
            .Range("I1:Q4").Columns.AutoFit
        End With
    Next ws
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
    
End Sub

