Attribute VB_Name = "Module1"
Sub StockData()
        
      '  On Error Resume Next
        
        Dim ws As Worksheet
                
        For Each ws In Worksheets
        
                ' Declare variables
                Dim ticker As String
                Dim yearOpen As Double
                Dim yearClose As Double
                Dim yearlyChange As Double
                Dim totalStockVolume As Double
                        totalStockVolume = 0
                Dim percentChange As Double
                
                'Keep track of each ticker location
                Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2
                
                'Create headers for output data
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total Stock Volume"
                
                
                ' Find the last row of the data
                lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                ' Loop through each row, ignore headers
                For i = 2 To lastRow
                
                        If (yearOpen = 0) Then
                                yearOpen = ws.Cells(i, 3)
                        End If
                        
                        ' Check for new ticker
                        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        
                                'Calculate yearly change
                                yearClose = ws.Cells(i, 6).Value
                                yearlyChange = yearClose - yearOpen
                                
                                'Calculate % change
                                If (yearOpen = 0) Then
                                        percentChange = 0
                                Else
                                        percentChange = yearlyChange / yearOpen
                                End If
                                
                                
                                'Fill the Summary Table
                                
                                ticker = ws.Cells(i, 1).Value
                                
                                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
                                
                                ws.Range("I" & Summary_Table_Row).Value = ticker
                                
                                ws.Range("J" & Summary_Table_Row).Value = yearlyChange
                                
                                        'Format Table
                                        If (yearlyChange > 0) Then
                                                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 35
                                        ElseIf (yearlyChange < 0) Then
                                                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 22
                                        Else
                                                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 37
                                        End If
                                
                                
                                ws.Range("K" & Summary_Table_Row).Value = percentChange
                                
                                        'Format Table
                                        ws.Cells(Summary_Table_Row, 11).Value = Format(percentChange, "Percent")
                                        If (percentChange > 0) Then
                                                ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 35
                                        ElseIf (percentChange < 0) Then
                                                ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 22
                                        Else
                                                ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 37
                                        End If
                                        
                                ws.Range("L" & Summary_Table_Row).Value = totalStockVolume

                                ' Add one to the summary table row
                                Summary_Table_Row = Summary_Table_Row + 1
      
                                ' Reset the Brand Total
                                totalStockVolume = 0
                                
                                ' Reset yearOpen
                                yearOpen = 0
                        Else
                                totalStockVolume = totalStockVolume + Cells(i, 7).Value
                        End If
                                       
                Next i
                
                '---------------------------------------------
                'BONUS
                '---------------------------------------------
                
                'Create headers for output data
                ws.Range("N2").Value = "Greatest Percent Increase"
                ws.Range("N3").Value = "Greatest Percent Decrease"
                ws.Range("N4").Value = "Greatest Total Volume"
                ws.Range("O1").Value = "Ticker"
                ws.Range("P1").Value = "Value"
                
                'Last row of the summary table
                lastRowSummary = ws.Cells(Rows.Count, "I").End(xlUp).Row
                
                'Greatest Percent Increase
                Dim gpi As Double
                Dim gpiTicker As String
                gpi = ws.Cells(2, 11).Value
                gpiTicker = ws.Cells(2, 9).Value
                'Greatest Percent Decrease
                Dim gpd As Double
                Dim gpdTicker As String
                gpd = ws.Cells(2, 11).Value
                gpdTicker = ws.Cells(2, 9).Value
                'Greatest Total Volume
                Dim gtv As Double
                Dim gtvTicker As String
                gtv = ws.Cells(2, 12).Value
                gtvTicker = ws.Cells(2, 9).Value
                
                'Loop through Summary Table, ignoring headers
                For i = 2 To lastRowSummary
                        'Find GPI
                        If (ws.Cells(i, 11).Value > gpi) Then
                                gpi = ws.Cells(i, 11).Value
                                gpiTicker = ws.Cells(i, 9).Value
                        End If
                        'Find GPD
                        If (ws.Cells(i, 11).Value < gpd) Then
                                gpd = ws.Cells(i, 11).Value
                                gpdTicker = ws.Cells(i, 9).Value
                        End If
                        'Find GTV
                        If (ws.Cells(i, 12).Value > gtv) Then
                                gtv = ws.Cells(i, 12).Value
                                gtvTicker = ws.Cells(i, 9).Value
                        End If
                Next i
                
                'Create bonus summary table
                ws.Range("O2").Value = Format(gpiTicker, "Percent")
                ws.Range("P2").Value = Format(gpi, "Percent")
                ws.Range("O3").Value = Format(gpdTicker, "Percent")
                ws.Range("P3").Value = Format(gpd, "Percent")
                ws.Range("O4").Value = gtvTicker
                ws.Range("P4").Value = gtv

        Next ws
    
End Sub
