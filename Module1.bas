Attribute VB_Name = "Module1"
Sub YearlyChangeTest()

'Loop through all sheets
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    'Heading creator
    ws.Range("i1").value = "Ticker"
    ws.Range("j1").value = "Yearly Change"
    ws.Range("k1").value = "Percent Change"
    ws.Range("l1").value = "Total Stock Volume"
    ws.Range("i1:l1").EntireColumn.AutoFit

    'Set an initial variable for Ticker
    Dim ticker_name As String

    'Set an initial variable for holding the total
    Dim ticker_open As Double
    ticker_open = 0
    Dim ticker_close As Double
    ticker_close = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim percent_change As Double
    percent_change = 0
    Dim ticker_volume As Double
    ticker_volume = 0
    Dim count As Integer
    count = 1

        'Check location of each ticker in total chart
        Dim ticker_location As Integer
        ticker_location = 2

        'Loop through all the tickers
        Dim i As Long
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
        For i = 2 To lastrow
    
        'Check if tickers match in ColumnA
        If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
    
            'Set the ticker name
            ticker_name = ws.Cells(i, 1).value
            
            'Adding volume
            ticker_volume = ticker_volume + ws.Cells(i, 7).value
        
            'Check the first open ticker and last closing ticker
            ticker_open = ws.Cells((i - (count - 1)), 3)
            ticker_close = ws.Cells(i, 6).value
            Yearly_Change = (ticker_close) - (ticker_open)
            percent_change = Yearly_Change / ticker_open
            
        
            'Print the ticker name
            ws.Range("i" & ticker_location).value = ticker_name
        
            'Print the ticker totals
            ws.Range("j" & ticker_location).value = Yearly_Change
                'Format colors for Yearly change
                If ws.Range("j" & ticker_location).value = 0 Then
                ws.Range("j" & ticker_location).Interior.ColorIndex = 6
                ElseIf ws.Range("j" & ticker_location).value > 0 Then
                ws.Range("j" & ticker_location).Interior.ColorIndex = 4
                ElseIf ws.Range("j" & ticker_location).value < 0 Then
                ws.Range("j" & ticker_location).Interior.ColorIndex = 3
                End If
                
            ws.Range("k" & ticker_location).value = percent_change
            ws.Range("l" & ticker_location).value = ticker_volume
        
            'Add 1 to ticker row count
            ticker_location = ticker_location + 1
        
            'Reset totals
            ticker_open = 0
            ticker_close = 0
            ticker_volume = 0
            Yearly_Change = 0
            percent_change = 0
            count = 1
        
    'If next ticker match we still count stock
    Else
        ticker_volume = ticker_volume + ws.Cells(i, 7).value
        count = count + 1
       

    End If

Next i

    'Fixing Percentage Change %
    ws.Range("k1", "k" & lastrow).NumberFormat = "0.00%"
    'reseting ticker location for next sheet
    ticker_location = 2
Next ws


End Sub


