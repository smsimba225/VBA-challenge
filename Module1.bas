Attribute VB_Name = "Module1"

Public Sub Script1()
    
    Dim Current As Worksheet
    For Each Current In Worksheets
    
        Dim TickerName As String
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
        
    
        
        Current.Range("I1").Value = "Ticker"
        Current.Range("J1").Value = "Yearly Change"
        Current.Range("K1").Value = "Percent Change"
        Current.Range("L1").Value = "Total Stock Volume"
        Current.Range("O2").Value = "Greatest % Increase"
        Current.Range("O3").Value = "Greatest % Decrease"
        Current.Range("O4").Value = "Greatest Total Volume"
        Current.Range("P1").Value = "Ticker"
        Current.Range("Q1").Value = "Value"
        
    
        
        Current.Range("I1:L1").Columns.AutoFit
            
            
        OpenPrice = Current.Cells(2, 3).Value
        tickerrow = 2
            
            
        For x = 2 To Current.Cells(Rows.Count, 1).End(xlUp).Row
        If Current.Cells(x + 1, 1).Value <> Current.Cells(x, 1).Value Then
        
        TickerName = Current.Cells(x, 1).Value
        ClosePrice = Current.Cells(x, 6).Value
        ChangePrice = ClosePrice - OpenPrice
        
        
        ChangePercent = (ChangePrice / OpenPrice)
        
        
        
        TotalStock = TotalStock + Current.Cells(x, 7).Value
        
        Current.Range("I" & tickerrow).Value = TickerName
        Current.Range("J" & tickerrow).Value = ChangePrice
        Current.Range("K" & tickerrow).Value = Format(ChangePercent, "0.00%")
        Current.Range("L" & tickerrow).Value = TotalStock
        
        If (ChangePrice > 0) Then
            Current.Range("J" & tickerrow).Interior.ColorIndex = 4
        ElseIf (ChangePrice <= 0) Then
            Current.Range("J" & tickerrow).Interior.ColorIndex = 3
        End If
        

        
        tickerrow = tickerrow + 1
        

        OpenPrice = Current.Cells(x + 1, 3).Value

        
        If (ChangePercent > GreatestIncrease) Then
            GreatestIncrease = ChangePercent
            IncreaseTicker = TickerName
            
        ElseIf (ChangePercent < GreatestDecrease) Then
            GreatestDecrease = ChangePercent
            DecreaseTicker = TickerName
            
        End If
            
            
        If (TotalStock > GreatestVolume) Then
                GreatestVolume = TotalStock
                GreatestVolumeName = TickerName
        End If
            
            
        ChangePercent = 0
        TotalStock = 0
        
        Else
        
        TotalStock = TotalStock + Current.Cells(x, 7).Value
        
        End If
        
        Next x
        

        Current.Range("Q2").Value = Format(GreatestIncrease, "0.00%")
        Current.Range("Q3").Value = Format(GreatestDecrease, "0.00%")
        Current.Range("P2").Value = IncreaseTicker
        Current.Range("P3").Value = DecreaseTicker
        Current.Range("Q4").Value = GreatestVolume
        Current.Range("P4").Value = GreatestVolumeName
        Current.Range("O1:Q4").Columns.AutoFit
        Current.Range("Q4").ColumnWidth = 8
        
    Next
        
End Sub

