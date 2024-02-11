Attribute VB_Name = "Module1"
'Module 2 Assignment - VBA Challenege
'-----------------------------------
'This code is split into 3 parts:
'Part 1: Looping through the main table to create the summary table
'Part 2: Labelling and Formtting of summary table
'Part 3: Calculation of greatest values


Sub ticker_summary()

For Each ws In Worksheets
    
    '---------------------------------
    'Part 1 :Creating of Summary Table
    '---------------------------------
    
    'Definition of Variables
    Dim counter As Integer
    Dim total_row, maxrow, minrow, maxvolrow As Long
    Dim ticker_symbol As String
    Dim total_stock As Variant
    Dim opening_year_price, closing_year_price, yearly_change, percentage_change  As Double
    Dim greatest_increase, greatest_decrease, greatest_total_volume As Double
        
    
    'Initial assignment of values to variables
    counter = 2
    total_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker_symbol = ws.Cells(2, 1).Value
    opening_year_price = ws.Cells(2, 3).Value
    
    'Looping through ticker symbol to create summary
    For x = 2 To total_row
                   
        If ws.Cells(x, 1) = ticker_symbol Then
            
            'Aggregates the volume of stock when the ticker symbol is the same
            total_stock = total_stock + ws.Cells(x, 7).Value
           
        Else
            'This would be at the change point when the ticker symbol does not match.
            'So outputs the calculated information to the summary table
            
            'Extraction of Ticker Symbol
            ws.Cells(counter, 9).Value = ticker_symbol
            ticker_symbol = ws.Cells(x, 1).Value
            
            'Output of calculated total stock
            ws.Cells(counter, 12).Value = total_stock
            total_stock = ws.Cells(x, 7).Value
            
            'Calculation and output of yearly change
            closing_year_price = ws.Cells(x - 1, 6).Value
            yearly_change = closing_year_price - opening_year_price
            ws.Cells(counter, 10).Value = yearly_change
            
            'Calculation and output of percentage change
            percentage_change = yearly_change / opening_year_price
            ws.Cells(counter, 11).Value = percentage_change
            
            'Reassigment of values for next run of loop
            opening_year_price = ws.Cells(x, 3).Value
            counter = counter + 1
        End If
         
    Next x
    
    '-------------------------------
    'Part 2: Labelling and formatting
    '-------------------------------
    
    'Column Labels for summary table
    ws.Range("i1") = "Ticker"
    ws.Range("p1") = "Ticker"
    ws.Range("j1") = "Yearly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("l1") = "Total Stock Volume"
    ws.Range("q1") = "Value"
    ws.Range("o2") = "Greatest % Increase"
    ws.Range("o3") = "Greatest % Decrease"
    ws.Range("o4") = "Greatest Total Volume"
    
    'Auto adjust column width for worksheet
    ws.Range("i:q").Columns.AutoFit
    
    'Formatting of yearly change column to 2 decimal places
    ws.Range("j:j").NumberFormat = "0.00"
    
    'Formatting to 2 decimal places and %
    ws.Range("k:k").NumberFormat = "0.00%"
    ws.Range("q2:q3").NumberFormat = "0.00%"
    
    'Colour Formatting of Yearly Change Column
    For x = 2 To counter - 1
        
        If ws.Cells(x, 10).Value < 0 Then
            'Negative values are red
            ws.Cells(x, 10).Interior.ColorIndex = 3
        Else
            'Positive values are green
            ws.Cells(x, 10).Interior.ColorIndex = 4
        End If
    
    Next x
    
    '----------------------------------------------
    'Part 3: Calculation & Output of Greatest Values
    '----------------------------------------------
    
    'Loops through summary table to find values
    For x = 2 To counter
    
        'Finding of greatest % increase
        If ws.Cells(x, 11).Value > greatest_increase Then
            greatest_increase = ws.Cells(x, 11).Value
            maxrow = x
        End If
        
        'Finding of greatest % decrease
        If ws.Cells(x, 11).Value < greatest_decrease Then
            greatest_decrease = ws.Cells(x, 11).Value
            minrow = x
        End If
        
        'Finding of greatest total volume
        If ws.Cells(x, 12).Value > greatest_total_volume Then
            greatest_total_volume = ws.Cells(x, 12).Value
            maxvolrow = x
        End If
        
    Next x
    
    'Output of greatest % increase - ticker symbol & value
    ws.Range("q2").Value = ws.Cells(maxrow, 11).Value
    ws.Range("p2").Value = ws.Cells(maxrow, 9).Value
    
    'Output of greatest % decrease - ticker symbol & value
    ws.Range("q3").Value = ws.Cells(minrow, 11).Value
    ws.Range("p3").Value = ws.Cells(minrow, 9).Value
       
    'Output of greatest total volume - ticker symbol & value
    ws.Range("q4").Value = ws.Cells(maxvolrow, 12).Value
    ws.Range("p4").Value = ws.Cells(maxvolrow, 9).Value

Next ws

End Sub

