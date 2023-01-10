Sub stockanalysis()
    
    'Dimensions for workbook
    
Dim total As Double
Dim rowindex As Long
Dim change As Double
Dim columnindex As Integer
Dim start As Long
Dim rowcount As Long
Dim percentChange As Double
Dim days As Integer
Dim dailychange As Single
Dim averagechange As Double
Dim ws As Worksheet

For Each ws In Worksheets

    total = 0
    change = 0
    start = 2
    outputrow = 2
    
    'set title row'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Get the row number of the last row with data
    rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For rowindex = 2 To rowcount
        total = total + ws.Cells(rowindex, 7).Value
        'if ticker changes give results
        If ws.Cells(rowindex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
   
    
              change = (ws.Cells(rowindex, 6) - ws.Cells(start, 3))
              percentChange = change / ws.Cells(start, 3)
              
              start = rowindex + 1
              
              ws.Range("i" & outputrow) = ws.Cells(rowindex, 1).Value
              ws.Range("J" & outputrow) = change
              ws.Range("J" & outputrow).NumberFormat = "0.00"
              ws.Range("K" & outputrow).Value = percentChange
              ws.Range("K" & outputrow).NumberFormat = "0.00%"
              ws.Range("L" & outputrow).Value = total
              
              
              outputrow = outputrow + 1
              total = 0
        
            
            
        
        End If
    
    Next rowindex
    'autosize
    ws.Range("I:Q").Columns.AutoFit
    ws.Range("I:Q").HorizontalAlignment = xlHAlignCenter

Next ws


End Sub

