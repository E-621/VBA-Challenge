Attribute VB_Name = "Module1"
Sub VBA_Challenge():

    
    'Set Variables
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Single
    Dim TotalVolume As LongLong
    Dim SummaryTableRow As Integer
    Dim LastRow As Long
    Dim Volume As Long
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    
    'Set Worksheet Loop
    For Each ws In Worksheets
        'Build Headers and Format
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Columns("J:K").HorizontalAlignment = xlCenter
        ws.Columns("I:L").AutoFit
        ws.Range("I1:L1").Font.Bold = True
        
            'Set inital data values
            SummaryTableRow = 2
            YearlyOpen = ws.Cells(2, 3).Value
            YearlyClose = 0
            YearlyChange = 0
            Volume = 0
            TotalVolume = 0
    
   
                    'Set worksheet Loop
                    For i = 2 To LastRow
                   
                        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            'Calculate Ticker and Volume Data
                            Ticker = ws.Cells(i, 1).Value
                            Volume = ws.Cells(i, 7).Value
                            TotalVolume = TotalVolume + Volume
                            ws.Range("I" & SummaryTableRow).Value = Ticker
                            ws.Range("L" & SummaryTableRow).Value = TotalVolume
                                        'Calculate Yearly Change and % Change Data
                                        YearlyClose = ws.Cells(i, 6).Value
                                        YearlyChange = YearlyClose - YearlyOpen
                                        ws.Range("J" & SummaryTableRow).Value = YearlyChange
                                        
                                        If YearlyOpen <> 0 Then
                                            PercentChange = (YearlyChange / YearlyOpen)
                                            ws.Range("K" & SummaryTableRow).Value = PercentChange
                                        Else
                                            ws.Range("K" & SummaryTableRow).Value = 0
                                        End If
                                            YearlyOpen = ws.Cells(i + 1, 3).Value
                                            SummaryTableRow = SummaryTableRow + 1
                                            TotalVolume = 0
                           
                                        'Set Conditioanl Formatting for J & K Rows
                                        If YearlyChange > 0 Then
                                            ws.Range("J" & SummaryTableRow - 1).Interior.ColorIndex = 4
                                        Else
                                            ws.Range("J" & SummaryTableRow - 1).Interior.ColorIndex = 3
                                        End If
                       
                                     Else
                                         Volume = ws.Cells(i, 7).Value
                                         TotalVolume = TotalVolume + Volume
                         End If
                 Next i
        Next ws
End Sub
