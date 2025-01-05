Sub GetInfo()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets

        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        Dim CurrentTicker As String
        CurrentTicker = ws.Range("A2").Value
    
        Dim CurerntOpen As Double
        CurrentOpen = ws.Range("C2").Value
    
        Dim CurrentClose As Double
        Dim CurrentQC As Double
        Dim CurrentQCP As Double
    
        Dim CurrentVolume As LongLong
        CurrentVolume = 0
    
        Dim StockCounter As Integer
        StockCounter = 1
    
        Dim GPI As Double
        GPI = 0
        Dim GPITicker As String
        
        Dim GPD As Double
        GPD = 0
        Dim GPDTicker As String
        
        Dim GTV As LongLong
        GTV = 0
        Dim GTVTicker As String
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        Dim i As Long
        
        For i = 2 To LastRow
        
            CurrentVolume = CurrentVolume + ws.Cells(i, 7).Value
            
            If (ws.Cells(i + 1, 1) <> CurrentTicker) Then
                
                ws.Cells(StockCounter + 1, 9).Value = CurrentTicker
                
                CurrentClose = ws.Cells(i, 6).Value
                CurrentQC = CurrentClose - CurrentOpen
                ws.Cells(StockCounter + 1, 10).NumberFormat = "0.00"
                ws.Cells(StockCounter + 1, 10).Value = CurrentQC
                If (CurrentQC > 0) Then
                    ws.Cells(StockCounter + 1, 10).Interior.ColorIndex = 4
                ElseIf (CurrentQC < 0) Then
                    ws.Cells(StockCounter + 1, 10).Interior.ColorIndex = 3
                End If
                
                CurrentQCP = CurrentQC / CurrentOpen
                ws.Cells(StockCounter + 1, 11).NumberFormat = "0.00%"
                ws.Cells(StockCounter + 1, 11).Value = CurrentQCP
                
                ws.Cells(StockCounter + 1, 12).Value = CurrentVolume
                
                If (CurrentQCP >= GPI) Then
                    GPI = CurrentQCP
                    GPITicker = CurrentTicker
                End If
                    
                If (CurrentQCP <= GPD) Then
                    GPD = CurrentQCP
                    GPDTicker = CurrentTicker
                End If
                
                If (CurrentVolume >= GTV) Then
                    GTV = CurrentVolume
                    GTVTicker = CurrentTicker
                End If
                
                StockCounter = StockCounter + 1
                CurrentTicker = ws.Cells(i + 1, 1).Value
                CurrentOpen = ws.Cells(i + 1, 3).Value
                CurrentClose = 0
                CurrentQC = 0
                CurrentQCP = 0
                CurrentVolume = 0
                
            End If
            
        Next i
        
        ws.Cells(2, 16).Value = GPITicker
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 17).Value = GPI
        
        ws.Cells(3, 16).Value = GPDTicker
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = GPD
        
        ws.Cells(4, 16).Value = GTVTicker
        ws.Cells(4, 17).Value = GTV
        
    
        ws.Columns("I:Q").AutoFit
    
    Next ws

End Sub