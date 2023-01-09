Sub Stock_Price_Analysis()
    
    Dim ws As Integer
    Dim WS_Count As Integer
    
    'loop through all worksheets in a workbook,so same VBS script will run in each tab
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For ws = 1 To WS_Count
    
    'show each tab's name to double check
    MsgBox ActiveWorkbook.Worksheets(ws).Name

'-----------------------------------------------------
    
    'assign title for summary table
    Worksheets(ws).Cells(1, "I").Value = "Ticker"
    Worksheets(ws).Cells(1, "J").Value = "Yearly Change"
    Worksheets(ws).Cells(1, "K").Value = "Percentage Change"
    Worksheets(ws).Cells(1, "L").Value = "Total Stock Volume"
    'show every title completely
    Worksheets(ws).Columns("A:Q").AutoFit
    
'-----------------------------------------------------
    LR = Worksheets(ws).Cells(Rows.Count, "A").End(xlUp).Row
    'or lastrow = Worksheets(ws).Cells(Rows, Count, 1)
    MsgBox LR
    
    Total = 0
    First_Open_Price = 2
    Summary_Position = 2
    
        'loop through all tickers and categorize by ID
        For i = 2 To LR
            If Worksheets(ws).Cells(i, "A").Value <> Worksheets(ws).Cells(i + 1, "A").Value Then
                'calculate column A to G
                Total = Total + Worksheets(ws).Cells(i, "G").Value
                'has to be the first open price for each ticker
                Open_Price = Worksheets(ws).Cells(First_Open_Price, "C").Value
                
                Close_Price = Worksheets(ws).Cells(i, "F").Value
                
                yearly_change = Close_Price - Open_Price
                
                Percentage_Change = yearly_change / Open_Price * 100
                
                
                'populate result
                Worksheets(ws).Cells(Summary_Position, "I").Value = Worksheets(ws).Cells(i, "A").Value
                Worksheets(ws).Cells(Summary_Position, "J").Value = yearly_change
                Worksheets(ws).Cells(Summary_Position, "K").Value = "%" & Percentage_Change
                Worksheets(ws).Cells(Summary_Position, "L").Value = Total
                
                
                'shade cells based on conditions
                If yearly_change > 0 Then
                    Worksheets(ws).Cells(Summary_Position, "J").Interior.ColorIndex = 4
                Else
                    Worksheets(ws).Cells(Summary_Position, "J").Interior.ColorIndex = 3
                End If
                
                
                'has to start from 0 again for each ticker!!!
                Total = 0
                'smaller/inner loop
                First_Open_Price = i + 1
                'bigger/external loop
                Summary_Position = Summary_Position + 1
                
                
                
            Else
                Total = Total + Worksheets(ws).Cells(i, "G").Value
                
            
            End If
        
        
        Next i
    
    Next ws
    
End Sub

Sub Stock_Price_Fluctuation()

    Dim ws As Integer
    Dim WS_Count As Integer
    
    'loop through all worksheets in a workbook,so same VBS script will run in each tab
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For ws = 1 To WS_Count
    
    'assign title for comparison table
    Worksheets(ws).Cells(1, "P").Value = "Ticker"
    Worksheets(ws).Cells(1, "Q").Value = "Value"
    Worksheets(ws).Cells(2, "O").Value = "Greatest % Increase"
    Worksheets(ws).Cells(3, "O").Value = "Greatest % Decrease"
    Worksheets(ws).Cells(4, "O").Value = "Greatest Total Volume"
    
    LR2 = Worksheets(ws).Cells(Rows.Count, "I").End(xlUp).Row
        
        For x = 2 To LR2
            
            'extract corresponding values
            Greatest_Increase = Application.WorksheetFunction.Max(Worksheets(ws).Range("K:K"))
            Greatest_Decrease = Application.WorksheetFunction.Min(Worksheets(ws).Range("K:K"))
            Top_Vol = Application.WorksheetFunction.Max(Worksheets(ws).Range("L:L"))
            
            
            'apply retrieved value to comparison table
            If Worksheets(ws).Cells(x, "K").Value = Greatest_Increase Then
                Worksheets(ws).Cells(2, "P").Value = Worksheets(ws).Cells(x, "I").Value
                Worksheets(ws).Cells(2, "Q").Value = "%" & Worksheets(ws).Cells(x, "K").Value
            
            ElseIf Worksheets(ws).Cells(x, "K").Value = Greatest_Decrease Then
                Worksheets(ws).Cells(3, "P").Value = Worksheets(ws).Cells(x, "I").Value
                Worksheets(ws).Cells(3, "Q").Value = "%" & Worksheets(ws).Cells(x, "K").Value
        
    
            ElseIf Worksheets(ws).Cells(x, "L").Value = Top_Vol Then
                Worksheets(ws).Cells(4, "P").Value = Worksheets(ws).Cells(x, "I").Value
                Worksheets(ws).Cells(4, "Q").Value = Worksheets(ws).Cells(x, "L").Value
    
            End If
            
        
        Next x
        
    Next ws
    
End Sub
