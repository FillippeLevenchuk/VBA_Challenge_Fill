Sub YearlyChangeTest()

'Loop through all sheets
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    'Heading creator
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    ws.Range("i1:l1").EntireColumn.AutoFit

    'Set an initial variable for Ticker
    Dim ticker_name As String

    'Set an initial variable for holding the total
    Dim ticker_total_open As Double
    ticker_total_open = 0
    Dim ticker_total_close As Double
    ticker_total_close = 0
    Dim Yearly_Change As Double
    Dim ticker_volume As Double
    ticker_volume = 0

        'Check location of each ticker in total chart
        Dim ticker_location As Integer
        ticker_location = 2

        'Loop through all the tickers
        Dim i As Integer
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


Sub max_min()

'creating variables
Dim ws As Worksheet
Dim biggest_increase As String
Dim biggest_decrease As String
Dim biggest_volume
Dim ticker As String
Dim value As String

For Each ws In ThisWorkbook.Worksheets

Dim range1 As Range
Dim range2 As Range
Dim lastrow As Long
lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
Set range1 = ws.Range(ws.Cells(2, 11), ws.Cells(lastrow, 11))
Set range2 = ws.Range(ws.Cells(2, 12), ws.Cells(lastrow, 12))

    'Heading creator
    ws.Range("o2").value = "Greatest Increase"
    ws.Range("o3").value = "Greatest_%_Decrease"
    ws.Range("o4").value = "Greatest_Total_Volume"
    ws.Range("p1").value = "Ticker"
    ws.Range("q1").value = "Value"
    ws.Range("o1:q1").EntireColumn.AutoFit
    
            
            biggest_increase = ws.Application.WorksheetFunction.Max(range1)
            biggest_decrease = ws.Application.WorksheetFunction.Min(range1)
            biggest_volume = ws.Application.WorksheetFunction.Max(range2)
            
            'Loop i
            Dim i As Long
            For i = 2 To lastrow
            
            'If function for biggest increase
            If ws.Cells(i, 11).value = biggest_increase Then
            ws.Range("p2").value = ws.Cells(i, 9).value
            ws.Range("q2").value = ws.Cells(i, 11).value
            End If
            
            'If function for biggest decrease
            If ws.Cells(i, 11).value = biggest_decrease Then
            ws.Range("p3").value = ws.Cells(i, 9).value
            ws.Range("q3").value = ws.Cells(i, 11).value
            End If
            
            'If function for biggest increase
            If ws.Cells(i, 12).value = biggest_volume Then
            ws.Range("p4").value = ws.Cells(i, 9).value
            ws.Range("q4").value = ws.Cells(i, 12).value
            End If
            
        Next i
    ws.Range("o1:q1").EntireColumn.AutoFit
    ws.Range("q2").NumberFormat = "0.00%"
    ws.Range("q3").NumberFormat = "0.00%"
Next ws
End Sub

