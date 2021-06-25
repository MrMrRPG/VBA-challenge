Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information.
    'The ticker symbol
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.
'
'Bounus:
'   1. Your solotuion will also be able to return the stock with the "Greatest % increase",
    ' "Greatest % decrease" and "Greatest total volume".
'   2. Make the appropriate adjustments to your VBA script that will allow it to run on
    ' every worksheet, i.e., every year, jsut by running the VBA script once.

Sub vba_challenge()

    Dim ws As Worksheet

    For Each ws In Worksheets ' Code run on each worksheet - not working; copies
                                ' current worksheet info to the others but not
                                ' corresponding worksheet to its own.

        ' Dims
        Dim ticker As String
        
        Dim yearly_change As Double
            yearly_change = 0 ' set initial yearly change as 0
            
        Dim percent_change As Double
            percent_change = 0 ' set initial percent change as 0
            
        Dim total_stock_vol As Double
            total_stock_vol = 0 ' set initial total stock vol as 0
            
        Dim new_table_row As Double
            new_table_row = 2 ' new table data should start on row 2
            
        Dim opening_price As Double
        
        Dim closing_price As Double
        
        Dim open_close As Double
            open_close = 2 ' Defined in the first ticker's open row
        
        Dim lastrow_raw As Long
        
        Dim lastrow_new As Long
        
        
        ' Headings
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Perecent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Bounus Headings
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Find last row
        lastrow_raw = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop for creating basic table
        For i = 2 To lastrow_raw
                
            ' Ticker and Total Stoack Volume
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value ' Set ticker
                total_stock_vol = ws.Cells(i, 7).Value + total_stock_vol ' Add total stock volume
                opening_price = ws.Cells(open_close, 3).Value ' define opening_price
                closing_price = ws.Cells(i, 6).Value ' define closing_price
                yearly_change = closing_price - opening_price ' define yearly_change
                
'               ' Embedded conditions for percentage change in case opening_price = 0
                If opening_price = 0 Then
                    ws.Cells(new_table_row, 11).Value = Null
                Else
                    percentage_change = yearly_change / opening_price ' define percentage_change

                End If
                
                ws.Range("I" & new_table_row).Value = ticker ' Print ticker
                ws.Range("L" & new_table_row).Value = total_stock_vol ' Print total stock volume
                ws.Range("J" & new_table_row).Value = yearly_change ' Print yearly change
                ws.Range("K" & new_table_row).Value = percentage_change
                
                new_table_row = new_table_row + 1 ' Add one to the new table
                
                total_stock_vol = 0 ' reset total stock volume
                open_close = i + 1 ' sets line status from row 2 to row 1
                yearly_change = 0 ' reset yearly change
            
        
            Else
                total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value ' add to the total stock vol
        
        
            End If
            
        Next i
         
        lastrow_new = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For j = 2 To lastrow_new
        
            ' Loop for coloration on yearly change
            If ws.Cells(j, 10).Value < 0 Then ' If the yearly change is a negative number
                ws.Cells(j, 10).Interior.ColorIndex = 3 ' Then set it to red
                
            ElseIf ws.Cells(j, 10).Value >= 0 Then ' If the yearly change is a negative number
                ws.Cells(j, 10).Interior.ColorIndex = 4 ' Then green
                
            End If
        
        ' Bounus Section ----------------------------------------------------------
        Dim max_percent As Double
            max_percent = WorksheetFunction.Max(Range("K" & 2 & ":" & "K" & j))
            ws.Range("Q2").Value = max_percent
            ws.Range("P2").Value = Application.Index(Range("I" & 2 & ":" & "I" & j), Application.Match(max_percent, Range("K" & 2 & ":" & "K" & j), 0))
            
        Dim min_percent As Double
            min_percent = WorksheetFunction.Min(Range("K" & 2 & ":" & "K" & j))
            ws.Range("Q3").Value = min_percent
            ws.Range("P3").Value = Application.Index(Range("I" & 2 & ":" & "I" & j), Application.Match(min_percent, Range("K" & 2 & ":" & "K" & j), 0))
            
        Dim max_total As Double
            max_total = WorksheetFunction.Max(Range("L" & 2 & ":" & "L" & j))
            ws.Range("Q4").Value = max_total
            ws.Range("P4").Value = Application.Index(Range("I" & 2 & ":" & "I" & j), Application.Match(max_total, Range("L" & 2 & ":" & "L" & j), 0))
        ' --------------------------------------------------------------------------
        
        Next j
         
        ' Autofit, because why not lol
        ws.Range("I1:Q1").EntireColumn.AutoFit
        
        ' Foramt percentage_change to %
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws
    
End Sub

Sub clear()

    Dim ws As Worksheet
    
    For Each ws In Worksheets

        ws.Range("I:Q").clear
    
    Next ws
    
End Sub
