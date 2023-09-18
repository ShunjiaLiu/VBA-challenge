Attribute VB_Name = "stockanalysis"
Sub StockAnalysis()

    Dim ws As Worksheet
    Dim LastRow As Long
    
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    
    Dim SummaryRow As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    
    For Each ws In ThisWorkbook.Worksheets
        
        
        SummaryRow = 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        GreatestIncreaseTicker = ""
        GreatestDecreaseTicker = ""
        GreatestVolumeTicker = ""
        
        
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
     
        ws.Range("I" & SummaryRow).Value = "Ticker"
        ws.Range("J" & SummaryRow).Value = "Yearly_Change"
        ws.Range("K" & SummaryRow).Value = "Percent_Change"
        ws.Range("L" & SummaryRow).Value = "Total_Volume"
                
        OpenPrice = ws.Cells(2, 3).Value
        Total_Volume = 0
     
        For i = 2 To LastRow
      
            Ticker = ws.Cells(i, 1).Value
     
            
            
            If Ticker <> ws.Cells(i + 1, 1).Value Then
                ClosePrice = ws.Cells(i, 6).Value
                
                
                Yearly_Change = ClosePrice - OpenPrice
                
                If OpenPrice <> 0 Then
                    Percent_Change = (Yearly_Change / OpenPrice)
                Else
                    Percent_Change = 0
                End If
                OpenPrice = ws.Cells(i + 1, 3).Value
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                ws.Range("I" & SummaryRow).Value = Ticker
                ws.Range("J" & SummaryRow).Value = Yearly_Change
                ws.Range("K" & SummaryRow).Value = Percent_Change
                ws.Range("L" & SummaryRow).Value = Total_Volume
                
                
                If Yearly_Change > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                ElseIf Yearly_Change < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                End If
                
              
                If Percent_Change > GreatestIncrease Then
                    GreatestIncrease = Percent_Change
                    GreatestIncreaseTicker = Ticker
                ElseIf Percent_Change < GreatestDecrease Then
                    GreatestDecrease = Percent_Change
                    GreatestDecreaseTicker = Ticker
                End If
                
                If Total_Volume > GreatestVolume Then
                    GreatestVolume = Total_Volume
                    GreatestVolumeTicker = Ticker
                End If
                
                
                Total_Volume = 0
                Yearly_Change = 0
                Percent_Change = 0
                SummaryRow = SummaryRow + 1
                
            Else

                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = GreatestIncreaseTicker
        ws.Cells(2, 17).Value = GreatestIncrease & "%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = GreatestDecreaseTicker
        ws.Cells(3, 17).Value = GreatestDecrease & "%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = GreatestVolumeTicker
        ws.Cells(4, 17).Value = GreatestVolume
        
    Next ws

End Sub


