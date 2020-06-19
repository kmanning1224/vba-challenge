Sub stock_ticker()


'functions
Dim ticker As String
Dim vol As Variant
    vol = 0
Dim yr_open As Variant
    yr_open = 0
Dim yr_close As Variant
    yr_close = 0
Dim yr_change As Double
    yr_change = 0
Dim yr_percent_change As Double
Dim ws As Worksheet


'create table
Summary_Table_Row = 2


For Each WS In ActiveWorkbook.Worksheets
'for looping ticker
    For i = 2 To ws.UsedRange.Rows.Count
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'remember ws cells sets the area its reading
        ticker = ws.Cells(i, 1).Value
        yr_open = yr_open + ws.Cells(i, 3).Value
        yr_close = yr_close + ws.Cells(i, 6).Value
        vol = vol + ws.Cells(i, 7).Value 
        
        'math to create info wanted
        yr_change = ws.Cells(i, 3).Value - ws.Cells(i, 6).Value
           
        
            If yr_percent_change = yr_open  And yr_close = 0 Then
        yr_percent_change = 0
            
            ElseIf yr_open <> 0 And yr_close = 0 Then yr_percent_change = 1
            
            Else: yr_percent_change = (yr_close - yr_open) / yr_open * 100
            End If


        'rows
        Range("j" & Summary_Table_Row).Value = ticker
        Range("k" & Summary_Table_Row).Value = yr_change
        Range("l" & Summary_Table_Row).Value = yr_percent_change
        Range("m" & Summary_Table_Row).Value = vol
        

        ' make percentage for percent change row
        Range("k2", Range("k2").End(xlDown)).NumberFormat = "0.00"
        Range("l2", Range("l2").End(xlDown)).NumberFormat = "0.00%"
        Range("m2", Range("m2").End(xlDown)).NumberFormat = "general"
        
    'titles
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Volume"
    'because im weird and like things centered
    Range("J1", Range("j1").End(xlDown)).HorizontalAlignment = xlCenter
    Range("K1", Range("K1").End(xlDown)).HorizontalAlignment = xlCenter
    Range("L1", Range("L1").End(xlDown)).HorizontalAlignment = xlCenter
    Range("M1", Range("M1").End(xlDown)).HorizontalAlignment = xlCenter
        
        'make sure table works right
        Summary_Table_Row = Summary_Table_Row + 1
        
        'reset vol
        vol = 0
        
        Else
        vol = vol + ws.Cells(i, 7).Value 
        
        
        End If
    Next i
    'pretty colors
    For Each cel In Range("k2", Range("k2").End(xlDown))

        If ((cel.Value) < 0) Then
        cel.Interior.ColorIndex = 3
        
        ElseIf ((cel.Value) > 0) Then
        cel.Interior.ColorIndex = 4
        End If
    Next cel
    

Next ws
End Sub

