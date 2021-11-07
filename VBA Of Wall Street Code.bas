Attribute VB_Name = "Module1"
Sub VBA_of_Wall_Street()
    
    'Insert Columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Define values
    Dim TotalVolume As Double
    Dim Ticker As String
    Dim TickerCounter As Double
    Dim TickerOpenCloseCounter As Double
    Dim YearlyOpen As Double
    Dim YearlyEnd As Double
    
    TotalVolume = 0   'Volume starting value

    TickerCounter = 2           'Row of ticker counter
    TickerOpenCloseCounter = 2   'Row of open and close counter
    
    'Loop
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Ticker = Cells(i, 1).Value
        TotalVolume = TotalVolume + Cells(i, 7).Value
        YearlyOpen = Cells(TickerOpenCloseCounter, 3)
        YearlyEnd = Cells(i, 6)
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(TickerCounter, 9).Value = Ticker
            Cells(TickerCounter, 10).Value = YearlyEnd - YearlyOpen
            
            'If Else to accomodate for division by 0
            If YearlyOpen = 0 Then
                Cells(TickerCounter, 11).Value = Null
            Else
                Cells(TickerCounter, 11).Value = (YearlyEnd - YearlyOpen) / YearlyOpen
            End If
        
        Cells(TickerCounter, 11).NumberFormat = "0.00%"
        Cells(TickerCounter, 12).Value = TotalVolume
        TickerCounter = TickerCounter + 1
        TickerOpenCloseCounter = i + 1
        TotalVolume = 0
        End If
    Next i
End Sub

