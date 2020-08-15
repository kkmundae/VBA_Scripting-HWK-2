Attribute VB_Name = "Module1"
Sub Stock_Analysis()

    'Loop through worksheets
    For Each ws In Worksheets
        
        'Assign new Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
     'Declare Variables
        Dim Ticker As String
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
        Dim LastRow As Long
        Dim TickerVol As Double
        Dim Summary_Table_Row As Long
        
    'Set Variables
        TickerVol = 0
        Summary_Table_Row = 2
        PreviousAmount = 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVolume = 0
    
    'Find last row of data (for each it may be different cuz of number of entries
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
            ' Add To Ticker Total
            TickerVol = TickerVol + ws.Cells(i, 7).Value
            
            ' Check If We Are Still Within The Same Ticker Name
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Set Ticker, Total Ticker Volume and Print in cell
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = TickerVol
                TotalTicker = 0
            
                'Set Yearly Open, Close and Change
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            
                'Calculate the Percent Change for each Ticker
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                    
                End If
            
                'Format cells to add in Percent Symbol
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
            
                'Colour (+ve green and -ve red)
                If ws.Range("K" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            PreviousAmount = i + 1
            
            End If
        
        Next i
        
        'Calculate the Greatest Percent Increase, the greatest percent decrease and the greatest total volume
        For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
        ' Format cells to add in Percent Symbol
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
        ' Format Table Columns To Auto Fit
        ws.Columns("I:Q").AutoFit
    
    Next ws
  
End Sub
            
            
            
            
            
            
            
    
    
     
     



