Attribute VB_Name = "Module1"
Sub Stock_price():

    For Each ws In Worksheets
    
        worksheetName = ws.Name
        
        Dim Ticker As String
        Dim Open_Price As Double
        Dim Yearly_Change As Double
        Dim Stock_Volume As Double
        Dim GreatIn As Double
        Dim GreatDe As Double
        Dim GreatTV As Double
           
        'Stock Volume Starting Point
        Stock_Volume = 0
        Open_Price = ws.Cells(2, 3).Value
        Closing_Price = 0
        
        'Summary Table Starting Point
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Last Row Variable
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        'Creating Headers for Combined Data
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
          
        'Loop combining Ticker & Stock Volume
        For i = 2 To LastRow
        
            'First Comparison
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Ticker Title to be Transfered
                Ticker = ws.Cells(i, 1).Value
            
            'Closing Price for the Year
                Closing_Price = ws.Cells(i, 6).Value
            
            'Final Total for Stock Volume
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            
            'Yearly Price Change
                Yearly_Change = Closing_Price - Open_Price
                
            'Annual Ending Percent Change
                Percent_Change = Yearly_Change / Open_Price
            
            'Ticker Title Being Transfered to Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            'Yearly Change Title Being Transfered to Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            'Percent Change Title Being Transfered to Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            'Stock Volume Total Being Transfered to Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
            
            'Moving Down to Insert Next Row of Stock Info
                Summary_Table_Row = Summary_Table_Row + 1
            
            'Stock Volume Total Reset
                Stock_Volume = 0
            
            'Open Price Reset
                Open_Price = ws.Cells(i + 1, 3)
            Else
            
            'Collection of Stock Volume Totals
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                        
            End If
            
          Next i
        
        'Conditional Formatting **Need to fix the highlighting
        LastRow_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To LastRow_Summary_Table
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        
        Next i
        
        'Bonus Work
        'ws.Range("Q2") = Application.WorksheetFunction.Max(Range("K2:LastRow_Summary_Table"))
        
        'ws.Range("Q3") = Application.WorksheetFunction.Min(Range("K3:LastRow_Summary_Table"))
        
        'ws.Range("Q4") = Application.WorksheetFunction.Max(Range("L4:LastRow_Summary_Table"))
        
        'Cell Size Formating
        ws.Columns("A:Q").AutoFit
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("J").Style = "Currency"
    
    Next ws
    
MsgBox ("Task Complete")

End Sub
