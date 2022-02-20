Attribute VB_Name = "Module1"
Sub StockMarketAnalysis():
   
   ' Loop Through All Worksheets
   For Each WS In Worksheets
       
    ' Set Column Headers
       WS.Range("I1").Value = "Ticker"
       WS.Range("J1").Value = "Yearly Change"
       WS.Range("K1").Value = "Percent Change"
       WS.Range("L1").Value = "Total Stock Volume"
       WS.Range("N2").Value = "Greatest % increase"
       WS.Range("N3").Value = "Greatest % decrease"
       WS.Range("N4").Value = "Greatest Total Volume"
       WS.Range("P1").Value = "Ticker"
       WS.Range("Q1").Value = "Value"
       
       
    ' Declare and set the variables
       Dim TickerName As String
       Dim LastRow As Long
       Dim LastRowValue As Long
       Dim TotalTickerVolume As Double
       TotalTickerVolume = 0
       Dim SummaryTableRow As Long
       SummaryTableRow = 2
       Dim YearlyOpen As Double
       Dim YearlyClose As Double
       Dim YearlyChange As Double
       Dim PreviousAmount As Long
       PreviousAmount = 2
       Dim PercentChange As Double
       Dim GreatestIncrease As Double
       GreatestIncrease = 0
       Dim GreatestDecrease As Double
       GreatestDecrease = 0
       Dim GreatestTotalVolume As Double
       GreatestTotalVolume = 0
       
       ' Determine the Last Row
       LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
       For i = 2 To LastRow
        
        ' Add To Ticker Total Volume
        TotalTickerVolume = TotalTickerVolume + WS.Cells(i, 7).Value
        
        ' Check If We Are Still in the The Same Ticker Name or not
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        
        ' Set Ticker Name
          TickerName = WS.Cells(i, 1).Value
        
        ' Print Name of the ticker In The Summary Table
        WS.Range("I" & SummaryTableRow).Value = TickerName
               
        ' Print Total Amount of ticker To The Summary Table
        WS.Range("L" & SummaryTableRow).Value = TotalTickerVolume
        
        ' Reset The Ticker
        TotalTickerVolume = 0
               
        ' Set Yearly Open, Yearly Close and Yearly Change Name
        YearlyOpen = WS.Range("C" & PreviousAmount)
        YearlyClose = WS.Range("F" & i)
        YearlyChange = YearlyClose - YearlyOpen
        WS.Range("J" & SummaryTableRow).Value = YearlyChange
               If YearlyOpen = 0 Then
                   PercentChange = 0
               Else
                   YearlyOpen = WS.Range("C" & PreviousAmount)
                   PercentChange = YearlyChange / YearlyOpen
               End If
               '
               WS.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
               WS.Range("K" & SummaryTableRow).Value = PercentChange
               '
               If WS.Range("J" & SummaryTableRow).Value >= 0 Then
                   WS.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
               Else
                   WS.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
               End If
               SummaryTableRow = SummaryTableRow + 1
               PreviousAmount = i + 1
               End If
           Next i
 
       ' Return stock with Greatest % Increase, Greatest % Decrease and Greatest Total Volume
        

       ' Find last Row
         LastRow = WS.Cells(Rows.Count, 11).End(xlUp).Row + 1
        
       ' Start Loop
         For i = 2 To LastRow
        
        If WS.Range("K" & i).Value > WS.Range("Q2").Value Then
           WS.Range("Q2").Value = WS.Range("K" & i).Value
           WS.Range("P2").Value = WS.Range("I" & i).Value
        
        End If

        If WS.Range("K" & i).Value < WS.Range("Q3").Value Then
           WS.Range("Q3").Value = WS.Range("K" & i).Value
           WS.Range("P3").Value = WS.Range("I" & i).Value
        End If

        If WS.Range("L" & i).Value > WS.Range("Q4").Value Then
           WS.Range("Q4").Value = WS.Range("L" & i).Value
           WS.Range("P4").Value = WS.Range("I" & i).Value
        End If

        Next i
        
      ' Include % Symbol And Set To Two Decimal Places
       WS.Range("Q2").NumberFormat = "0.00%"
       WS.Range("Q3").NumberFormat = "0.00%"
            
      ' Autofit to display data
       WS.Columns("I:Q").AutoFit

      Next WS

End Sub

