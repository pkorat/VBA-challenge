Attribute VB_Name = "Module1"
Sub Stocks():

' loop through each worksheet
For Each ws In Worksheets
    
    ' initalize the ticker as a string
    Dim ticker As String
    
    ' initalize the lastrow as a integer
    Dim lastrow As Double
    
    ' initalize the beginning row
    Dim starter_row As Integer
    
    ' initalize the total_vol
    Dim total_vol As Double
    
    ' initalize open_price as double
    Dim open_price As Double
    
    ' set the starter_row to start at row 2 of any column
    starter_row = 2
    
    ' set lastrow equal to the last row
    lastrow = ws.Range("A1").End(xlDown).Row
    
    ' set total_vol as 0
    total_vol = 0
    
    ' set the headers for the ticker(i1), yearly change(j1), percent change(k1), and total stock volume(l1)
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' create for loop from 2 to lastrow
    For i = 2 To lastrow
    
        ' add to total_vol for each ticker
        total_vol = total_vol + ws.Cells(i, 7).Value
        
        ' if statement that checks for changes in the ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
            
            ' set equal to the ticker
            ticker = ws.Cells(i, 1).Value
            
            ' place into column I
            ws.Range("I" & starter_row).Value = ticker
            
            ' place into column J
            ws.Range("J" & starter_row).Value = ws.Cells(i, 6).Value - open_price
            
            ' place into column K
            ws.Range("K" & starter_row).Value = (ws.Cells(i, 6).Value - open_price) / open_price * 100
            
            ' place into column L
            ws.Range("L" & starter_row).Value = total_vol
            
            ' return total_vol to 0 at the end of the ticker
            total_vol = 0
            
            ' increment the starter_row by one for each true
            starter_row = starter_row + 1
            
            ' create another if statement that tracks beginning of year
        ElseIf Right(ws.Cells(i, 2).Value, 4) = 101 And ws.Cells(i, 3).Value <> 0 Then
        
            ' set open value
            open_price = ws.Cells(i, 3).Value
            
        End If
        
            ' check to see if yearly change is negative
        If ws.Cells(i, 10).Value < 0 Then
            
            ' set background to red
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
            
            ' check to see if yearly change is positive
        If ws.Cells(i, 10).Value > 0 Then
        
            ' set background to green
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        End If
    Next i
Next ws
        
End Sub
