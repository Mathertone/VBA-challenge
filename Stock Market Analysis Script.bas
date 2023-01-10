Attribute VB_Name = "Module1"
Sub VBA_Challenge()
For Each ws In Worksheets

    Dim Ticker As String
    Dim Ticker_Greatest_Increased As String
    Dim Ticker_Greatest_Decreased As String
    Dim Ticker_Greatest_Volume As String
    
    Dim Greatest_Increased As Double
    Greatest_Increased = 0
    Dim Greatest_Decreased As Double
    Greatest_Decreased = 0
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0

    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    
    Dim Summary_Table_Row As Integer
    Dim Days As Long
    
    Summary_Table_Row = 2
    Yearly_Open = 2
    Yearly_Close = 0
    Yearly_Change = 0
    Total_Stock_Volume = 0
    Days = 0
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    SumRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To LastRow
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                    Ticker = ws.Cells(i, 1).Value
                    
                    Open_Price = ws.Cells(i - Days, 3).Value
                    
                    Close_Price = ws.Cells(i, 6).Value
                    
                    Yearly_Change = Close_Price - Open_Price
                    
                    If Open_Price = 0 Then
                        Percent_Change = 0
                    Else
                        Percent_Change = (Close_Price - Open_Price) / Open_Price
                    End If
                    
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    
                    ws.Range("I" & Summary_Table_Row).Value = Ticker
                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                    ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                    
                    Summary_Table_Row = Summary_Table_Row + 1
                    Total_Stock_Volume = 0
                    Days = 0
                
                Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                Days = Days + 1
                
                End If
                
        Next i
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For j = 2 To LastRow
                If ws.Cells(j, 10).Value >= 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
            Next j
            
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For k = 2 To LastRow
       
            If ws.Cells(k, 11).Value > Greatest_Increased Then
                Greatest_Increased = ws.Cells(k, 11).Value
                Ticker_Greatest_Increased = ws.Cells(k, 9).Value
            End If
        
            If ws.Cells(k, 11).Value < Greatest_Decreased Then
                Greatest_Decreased = ws.Cells(k, 11).Value
                Ticker_Greatest_Decreased = ws.Cells(k, 9).Value
            End If
         
            If ws.Cells(k, 12).Value > Greatest_Total_Volume Then
                Greatest_Total_Volume = ws.Cells(k, 12).Value
                Ticker_Greatest_Volume = ws.Cells(k, 9).Value
            End If
            
        Next k
        
        Summary_Table_Row = 2
        Total_Stock_Volume = 0
        Yearly_Change = 0
        Percent_Change = 0
        
        ws.Range("P2").Value = Ticker_Greatest_Increased
        ws.Range("P3").Value = Ticker_Greatest_Decreased
        ws.Range("P4").Value = Ticker_Greatest_Volume
        ws.Range("Q2").Value = Greatest_Increased
        ws.Range("Q3").Value = Greatest_Decreased
        ws.Range("Q4").Value = Greatest_Total_Volume

        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

        ws.Columns("A:Q").AutoFit
        
    Next ws

End Sub
