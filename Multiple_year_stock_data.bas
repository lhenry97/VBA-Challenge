Attribute VB_Name = "Module1"
Sub Stocks()

'Set inital variable to hold ticker symbol
Dim Ticker_Symbol As String

'Set ws as worksheet
Dim ws As Worksheet

'Set opening price of quarter variables
Dim OpeningPrice As Double

'Set closing price of quarter variables
Dim ClosingPrice As Double

'Set Quarterly Change Variable
Dim QuarterlyChange As Double

'Set Percentage Change Variable
Dim PercentageChange As Double

'Set Total Stock Volume Variable
Dim TotalStockVol As Double

'Set variable for Greatest % increase
Dim GreatestIncrease As Double
Dim GreatestIncreaseName As String
    
'Set variable for Greatest % decrease
Dim GreatestDecrease As Double
Dim GreatestDecreaseName As String

'Set variable for Greatest Total Stock volume
Dim GreatestStockVol As Double
Dim GreatestStockVolName As String


For Each ws In Worksheets
    
    'Keep track of location of summary table row
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Add headers in summary table
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Quarterly Change"
    ws.Range("L1").Value = "Percentage Change"
    ws.Range("M1").Value = "Total Stock Volume"
    
    'Counts the number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Obtain first value of opening price
    OpeningPrice = ws.Cells(2, 3).Value
    
    'Initalise volume amount
    TotalStockVol = 0

    'For loop for parsing data
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            
            'Assign ticker symbol to varaiable
            Ticker_Symbol = ws.Cells(i, 1).Value
            
            'Print ticker symbol in summary table
            ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol
              
            'Assign closing price value
            ClosingPrice = ws.Cells(i, 6).Value
            
            'Find Change price over quarter
            QuarterlyChange = ClosingPrice - OpeningPrice
            
            'Print price in summary table
            ws.Range("K" & Summary_Table_Row).NumberFormat = "$0.00"
            ws.Range("K" & Summary_Table_Row).Value = QuarterlyChange
            
            'Apply conditonal formatting based on value of quarterly change
            If (QuarterlyChange > 0) Then
                    
                    ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                
            ElseIf (QuarterlyChange < 0) Then
                
                    ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            End If
    
            'Find Percentage change over quarter
            PercentageChange = QuarterlyChange / OpeningPrice
           
            'Set cell formatting to percentage
            ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
            
            'Print Percentage Change
            ws.Range("L" & Summary_Table_Row).Value = PercentageChange
            
            'Apply conditional formatting based on value of percentage change
            If (PercentageChange > 0) Then
                
                ws.Range("L" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                
            ElseIf (PercentageChange < 0) Then
                
                ws.Range("L" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            
            End If
                    
            'Add final value to stock volume
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
            
            'Print Total Stock Volume value
            ws.Range("M" & Summary_Table_Row).Value = TotalStockVol
            
            'Re-initalise total stock volume to zero for next ticker
            TotalStockVol = 0
            
            'Assign opening price for next ticker
            OpeningPrice = ws.Cells(i + 1, 3).Value
            
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
        
        Else
        
            'Add to total stock volume
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
            
        
        End If
          
        
    Next i
        
    ws.Columns("J:M").EntireColumn.AutoFit
    
    'Initialising variable lastrow2 to identify last row of parsed data
    lastrow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'Initalise values of variables
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestStockVol = 0

   
    'For loop for identiying % increases
    For k = 2 To lastrow2
       'If statement to identify higher % increase
       If (ws.Cells(k, 12).Value > GreatestIncrease) Then
       
       GreatestIncrease = ws.Cells(k, 12).Value
       GreatestIncreaseName = ws.Cells(k, 10).Value
       End If
       
       'If statement to identify higher % decrease
       If (ws.Cells(k, 12).Value < GreatestDecrease) Then
       
       GreatestDecrease = ws.Cells(k, 12).Value
       GreatestDecreaseName = ws.Cells(k, 10).Value
       
       End If
       
       'If statement to identify greatest stock volume
       If (ws.Cells(k, 13).Value > GreatestStockVol) Then
       
       GreatestStockVol = ws.Cells(k, 13).Value
       GreatestStockVolName = ws.Cells(k, 10).Value
       
       End If
            
       
    Next k
    
    'Add headers table
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    
    'Format for percentage
    ws.Range("R2:R3").NumberFormat = "0.00%"
    
    'Add values into summary table
    ws.Range("Q2").Value = GreatestIncreaseName
    ws.Range("R2").Value = GreatestIncrease
    
    ws.Range("Q3").Value = GreatestDecreaseName
    ws.Range("R3").Value = GreatestDecrease
    
    ws.Range("Q4").Value = GreatestStockVolName
    ws.Range("R4").Value = GreatestStockVol
    
    'Autofit columns
    ws.Columns("P:R").EntireColumn.AutoFit

Next ws

End Sub
