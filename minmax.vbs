Attribute VB_Name = "Module2"
Sub max_min()

'creating variables
Dim ws As Worksheet
Dim biggest_increase As String
Dim biggest_decrease As String
Dim biggest_volume

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

