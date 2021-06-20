Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information.
    'The ticker symbol
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub VBA_Challenge()

    ' Dims
    Dim ticker As String
    Dim yearly_change As Double
        yearly_change = 0 ' set initial yearly change as 0
        
    Dim percent_change As Double
        percent_change = 0 ' set initial percent change as 0
        
    Dim total_stock_vol As Long
        total_stock_vol = 0 ' set initial total stock vol as 0
        
    Dim new_table_row As Long
        new_table_row = 2 ' new table data should start on row 2
    
    Dim lastrow As Long
    
    ' Headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Perecent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Bounus Headings
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    ' Find last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loops
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value ' Set ticker
            total_stock_vol = Cells(i, 6).Value + total_stock_vol ' Add total stock volume
            
            Range("I" & new_table_row).Value = ticker ' Print ticker
            Range("L" & new_table_row).Value = total_stock_volume ' Print total stock volume
            
            new_table_row = new_table_row + 1 ' Add one to the new table
            
            total_stock_vol = 0 ' reset total stock volume
    
        End If
        
    Next i
    
    ' Autofit
    Range("I1:Q1").EntireColumn.AutoFit

End Sub
