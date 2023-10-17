Attribute VB_Name = "Module1"
Sub Challenge2()
    
    Dim ws As Worksheet
    
    'Variable for last row
    Dim Last_Row As Long
    Dim i As Long
    
    'Set variables
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim Summary_Table_Row As Long
    
    'set variables & values for greatest percent changes and volume
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume As Double
    Dim Top_Ticker As String
    Dim Worst_Ticker As String
    Dim Top_Volume_Ticker As String
    
    'Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
        'set value of variables
        Last_Row = ws.Cells(ws.Rows.Count, 1).Row
        For i = Last_Row To 1 Step -1
            If Not IsEmpty(ws.Cells(i, 1).Value) Then
                Last_Row = i
            Exit For
            End If
        Next i
        Ticker = ws.Cells(2, 1).Value
        Opening_Price = ws.Cells(2, 3).Value
        Total_Volume = 0
        
        'Create a location for the summary table
        Summary_Table_Row = 2
        
        'Add Headers to Summary Table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To Last_Row
        
            ' Check if a stock is within a given year i.e. ticker, if it is...
            If ws.Cells(i, 1).Value <> Ticker Then
                
                'find the yearly change & percent change
                Closing_Price = ws.Cells(i - 1, 6).Value
                Yearly_Change = Closing_Price - Opening_Price
                If Opening_Price <> 0 Then
                    Precent_Change = (Yearly_Change / Opening_Price) * 100
                Else
                    Precent_Change = 0
                End If
                
                'Print the Ticker Symbol in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
            
                'Print the Yearly Change in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                    
                'Print the Total Stock Volume in the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Volume
                
                'Print Precent Change in Summary Table
                ws.Range("k" & Summary_Table_Row).Value = Precent_Change
            
                'Add another row to the summary table
                Summary_Table_Row = Summary_Table_Row + 1
            
                'Reset Variables
                Ticker = ws.Cells(i, 1).Value
                Opening_Price = ws.Cells(i, 3).Value
                Total_Stock = 0
            End If
            
            
            'add to the Total Stock Volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
        Next i
        
        'set conditional formats
        
        For i = 2 To Last_Row
            If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
            If ws.Cells(i, 11).Value > 0 Then
                    ws.Cells(i, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 11).Value < 0 Then
                    ws.Cells(i, 11).Interior.ColorIndex = 3
            End If
        Next i
            
        Closing_Price = ws.Cells(Last_Row, 6).Value
        Yearly_Change = Closing_Price - Opening_Price
        
        If Opening_Price <> 0 Then
            Percent_Change = (Yearly_Change / Opening_Price) * 100
        Else
            Percent_Change = 0
        End If
        'Print to Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        ws.Range("L" & Summary_Table_Row).Value = Total_Volume
        ws.Range("k" & Summary_Table_Row).Value = Precent_Change
        
        'set values for greatest % and total volume variables
        Greatest_Percent_Decrease = Application.WorksheetFunction.Min(ws.Range("k2:k" & Last_Row))
        Greatest_Percent_Increase = Application.WorksheetFunction.Max(ws.Range("k2:k" & Last_Row))
        Greatest_Total_Volume = Application.WorksheetFunction.Max(ws.Range("l2:l" & Last_Row))
        
        'print greatest values & tickers to second summary table
        ws.Cells(2, 16).Value = Top_Ticker
        ws.Cells(3, 16).Value = Worst_Ticker
        ws.Cells(4, 16).Value = Top_Volume_Ticker
        ws.Cells(2, 17).Value = Greatest_Percent_Increase
        ws.Cells(3, 17).Value = Greatest_Percent_Decrease
        ws.Cells(4, 17).Value = Greatest_Total_Volume
        
        'name rows
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'name headers
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
    Next ws
End Sub

