Sub GetInfo()

    ' create variable for current worksheet in the loop
    Dim ws As Worksheet
    
    ' for each worksheet...
    For Each ws In ThisWorkbook.Worksheets

        ' create variable for the last row of the worksheet
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' create columns for output summary information for each ticker and insert their headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        ' create variable for current ticker
        Dim CurrentTicker As String
        ' assign CurrentTicker to the first ticker in the worksheet
        CurrentTicker = ws.Range("A2").Value
    
        ' create variable for open price that we are going to use to calc quarterly change
        Dim CurrentOpen As Double
        ' assign CurrentOpen to the first open price in the worksheet
        CurrentOpen = ws.Range("C2").Value
    
        ' create variables...
        ' for close price that will be used to calc quarterly change
        Dim CurrentClose As Double
        ' for quarterly change of current ticker
        Dim CurrentQC As Double
        ' for percent change of current ticker
        Dim CurrentQCP As Double
    
        ' create variable for total stock volume of current ticker and assign it to 0
        Dim CurrentVolume As LongLong
        CurrentVolume = 0
    
        ' create variable for counting of stocks in the worksheet and assign it to 1
        Dim StockCounter As Integer
        StockCounter = 1
    
        ' create variable for greatest % increase and assign it to 0
        Dim GPI As Double
        GPI = 0
        ' create variable for ticker that has GPI
        Dim GPITicker As String
        
        ' create variable for greatest % decrease and assign it to 0
        Dim GPD As Double
        GPD = 0
        ' create variable for ticker that has GPD
        Dim GPDTicker As String
        
        ' create variable for greatest total volume and assign it to 0
        Dim GTV As LongLong
        GTV = 0
        ' create variable for ticker that has GTV
        Dim GTVTicker As String
        
        ' create minitable for GPI, GPD and GTV, and insert headers
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' create variable for iteraions in loop
        Dim i As Long
        
        ' for every row in current worksheet...
        For i = 2 To LastRow
        
            ' reading and storing values from the row (as stated in requirements for the homework) *
            Dim CurrentRowTicker As String
            CurrentRowTicker = ws.Cells(i, 1).Value
            
            Dim CurrentRowOpen As Double
            CurrentRowOpen = ws.Cells(i, 3).Value
            
            Dim CurrentRowClose As Double
            CurrentRowClose = ws.Cells(i, 6).Value
            
            Dim CurrentRowVolume As LongLong
            CurrentRowVolume = ws.Cells(i, 7).Value
            
            ' add this row stock volume to current total stock volume *
            CurrentVolume = CurrentVolume + CurrentRowVolume
            'CurrentVolume = CurrentVolume + ws.Cells(i, 7).Value **
            
            ' if next row has a different ticker...
            If (ws.Cells(i + 1, 1) <> CurrentTicker) Then
                
                ' add current ticker to output summary column
                ws.Cells(StockCounter + 1, 9).Value = CurrentTicker
                
                ' make current row close price the one we will use for counting quarterly change *
                CurrentClose = CurrentRowClose
                'CurrentClose = ws.Cells(i, 6).Value **
                
                ' count quarterly change
                CurrentQC = CurrentClose - CurrentOpen
                
                ' format the cell for quarterly change and insert quarterly change in it
                ws.Cells(StockCounter + 1, 10).NumberFormat = "0.00"
                ws.Cells(StockCounter + 1, 10).Value = CurrentQC
                
                ' depending on the value of quarterly change (<0, >0 or =0), color the cell in red or green color or leave it uncolored
                If (CurrentQC > 0) Then
                    ws.Cells(StockCounter + 1, 10).Interior.ColorIndex = 4
                ElseIf (CurrentQC < 0) Then
                    ws.Cells(StockCounter + 1, 10).Interior.ColorIndex = 3
                End If
                
                ' count percent change
                CurrentQCP = CurrentQC / CurrentOpen
                ' format the cellfor this percent change and insert the value in it
                ws.Cells(StockCounter + 1, 11).NumberFormat = "0.00%"
                ws.Cells(StockCounter + 1, 11).Value = CurrentQCP
                
                ' insert CurrentVolume in th cell for the total stock volume
                ws.Cells(StockCounter + 1, 12).Value = CurrentVolume
                
                ' check if current QCP is GPI or GPD and if yes, assign needed variable and its ticker string to CurrentQCP and CurrentTicker
                If (CurrentQCP >= GPI) Then
                    GPI = CurrentQCP
                    GPITicker = CurrentTicker
                End If
                    
                If (CurrentQCP <= GPD) Then
                    GPD = CurrentQCP
                    GPDTicker = CurrentTicker
                End If
                
                ' check if current total stock volume is GTV and if yes, assign needed variable and its ticker string to CurrentVolume and CurrentTicker
                If (CurrentVolume >= GTV) Then
                    GTV = CurrentVolume
                    GTVTicker = CurrentTicker
                End If
                
                ' increment stock counter by 1
                StockCounter = StockCounter + 1
                ' assign CurrentTicker to next row ticker
                CurrentTicker = ws.Cells(i + 1, 1).Value
                ' assign CurrentOpen to next row open price
                CurrentOpen = ws.Cells(i + 1, 3).Value
                ' assign other variables to 0
                CurrentClose = 0
                CurrentQC = 0
                CurrentQCP = 0
                CurrentVolume = 0
                
            End If
            
        Next i
        
        ' format needed cells and insert GPI, GPD and GTV info
        ws.Cells(2, 16).Value = GPITicker
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 17).Value = GPI
        
        ws.Cells(3, 16).Value = GPDTicker
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = GPD
        
        ws.Cells(4, 16).Value = GTVTicker
        ws.Cells(4, 17).Value = GTV
        
        ' achieve the best fit by changing the width of the columns
        ws.Columns("I:Q").AutoFit
    
    Next ws

End Sub

' i find reading and storing values for close and open price for each row wasteful and useless. in order to get the summary
' info needed in the homework, we only need to know the open price for the first row and close price for the last row
' (with the same ticker). according to my testing, storing these values for each row increases the running time of the
' program by approximately 10 sec (and the more data is in the workbook, the greater this difference will be). since the
' homework requirements indicate the storage of these values, i left the lines of code necessary for this (they are located
' under the comments ending with *). however, i think that instead of using these lines it's more rational to use the lines
' i left as comments ending with ** (with no storing of open and close price for each row). more information about this can
' be found in the readme.md file.