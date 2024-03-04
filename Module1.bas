Attribute VB_Name = "Module1"

Public Sub StockScript()
    
    Dim Current As Worksheet
    For Each Current In Worksheets
    
        Dim TickerName As String
        
    'x can't be integer because causes overflow in big spreadsheet - too many rows
        
        Dim x As Long
        
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim ChangePrice As Double
        Dim ChangePercent As Double
        Dim TotalStock As Double
        
        Dim tickerrow As Integer
        
        
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim IncreaseTicker As String
        Dim DecreaseTicker As String
        Dim GreatestVolume As Double
        Dim GreatestVolumeName As String
    
    'to reset everything for next worksheet:
        ChangePercent = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        OpenPrice = 0
        ClosePrice = 0
        GreatestVolume = 0
        GreatestVolumeName = ""
        
        
        
    'adds header to report section
        Current.Range("I1").Value = "Ticker"
        Current.Range("J1").Value = "Yearly Change"
        Current.Range("K1").Value = "Percent Change"
        Current.Range("L1").Value = "Total Stock Volume"
        Current.Range("O2").Value = "Greatest % Increase"
        Current.Range("O3").Value = "Greatest % Decrease"
        Current.Range("O4").Value = "Greatest Total Volume"
        Current.Range("P1").Value = "Ticker"
        Current.Range("Q1").Value = "Value"
        
    'makes report section look nicer
        Current.Range("I1:L1").Columns.AutoFit
            
    'initial value for OpenPrice before loop begins
        OpenPrice = Current.Cells(2, 3).Value
    
    'starting row for report section to put first ticker
        tickerrow = 2
            
    'start of loop - checks if next row has same ticker name
    'Current.Cells(Rows.Count, 1).End(xlUp).Row reports value of last row # with data
        For x = 2 To Current.Cells(Rows.Count, 1).End(xlUp).Row
        If Current.Cells(x + 1, 1).Value <> Current.Cells(x, 1).Value Then
        
        TickerName = Current.Cells(x, 1).Value
        ClosePrice = Current.Cells(x, 6).Value
        ChangePrice = ClosePrice - OpenPrice
        
        
        ChangePercent = (ChangePrice / OpenPrice)
        
    'if next row has different ticker, adds stock volume of closing date
        
        TotalStock = TotalStock + Current.Cells(x, 7).Value
    
    'since next ticker is different, ready to report to report section:
    
        Current.Range("I" & tickerrow).Value = TickerName
        Current.Range("J" & tickerrow).Value = ChangePrice
        Current.Range("K" & tickerrow).Value = Format(ChangePercent, "0.00%")
        Current.Range("L" & tickerrow).Value = TotalStock
        
    'conditional color formatting for Yearly Change column
        
        If (ChangePrice > 0) Then
            Current.Range("J" & tickerrow).Interior.ColorIndex = 4
        ElseIf (ChangePrice <= 0) Then
            Current.Range("J" & tickerrow).Interior.ColorIndex = 3
        End If
        
    
    'moves report section down one line for new ticker symbol
    
        tickerrow = tickerrow + 1
        
    'sets new open price with next row
    
        OpenPrice = Current.Cells(x + 1, 3).Value

    'checks to see if Percent Change is greatest increase or decrease on list
    'stores information in variables
    
        If (ChangePercent > GreatestIncrease) Then
            GreatestIncrease = ChangePercent
            IncreaseTicker = TickerName
            
        ElseIf (ChangePercent < GreatestDecrease) Then
            GreatestDecrease = ChangePercent
            DecreaseTicker = TickerName
            
        End If
            
    'checks to see if Total Stock is the greatest so far down the list
            
        If (TotalStock > GreatestVolume) Then
                GreatestVolume = TotalStock
                GreatestVolumeName = TickerName
        End If
     
    'resets Total Stock otherwise next ticker total stock volume messes up
     
        TotalStock = 0

        Else
        
    'if next ticker is the same as current ticker, this just adds to the total stock volume
     
        
        TotalStock = TotalStock + Current.Cells(x, 7).Value
        
        End If
        
        
        
        Next x
        
    'after ticker loop is finished, reports Ticker and Value of Greatest % and Volume
    
        Current.Range("Q2").Value = Format(GreatestIncrease, "0.00%")
        Current.Range("Q3").Value = Format(GreatestDecrease, "0.00%")
        Current.Range("P2").Value = IncreaseTicker
        Current.Range("P3").Value = DecreaseTicker
        Current.Range("Q4").Value = GreatestVolume
        Current.Range("P4").Value = GreatestVolumeName
        Current.Range("O1").ColumnWidth = 20
        Current.Range("P1:Q4").ColumnWidth = 8
    'next worksheet
    
    Next
        
End Sub


