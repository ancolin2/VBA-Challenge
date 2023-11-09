Attribute VB_Name = "Module1"
Sub Stocks()
'Loop through worksheets
For Each ws In Worksheets

'Establish Columns and Row titles
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Determine last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Establish Variables
Dim Ticker As Integer
Dim running_volume As Double
Dim opn As Double
Dim cls As Double

Ticker = 2

'Make a loop
For j = 2 To LastRow
    
    'Yearly Change
    If ws.Cells(j, 1).Value <> ws.Cells(j - 1, 1).Value Then
    opn = ws.Cells(j, 3).Value
    End If
    
    'Total Stock Volume
    running_volume = running_volume + ws.Cells(j, 7).Value

    'Have a conditional for <>
    If (ws.Cells(j, 1).Value <> ws.Cells(j + 1, 1).Value) Then
        'add values for tickers,yearly change, percent change, and total stock volume
        cls = Cells(j, 6).Value
        ws.Cells(Ticker, 9).Value = ws.Cells(j, 1).Value
        ws.Cells(Ticker, 12).Value = running_volume
        ws.Cells(Ticker, 10).Value = cls - opn
        running_volume = 0
        ws.Cells(Ticker, 11).Value = (cls - opn) / opn
            'Change color for increase vs decrease
            If ws.Cells(Ticker, 10).Value > 0 Then
            ws.Cells(Ticker, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(Ticker, 10).Interior.ColorIndex = 3
            End If
        Ticker = Ticker + 1
    
    End If
    
Next j

'Greatest Increase
LastNewRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
For p = 3 To LastNewRow
    If ws.Cells(p, 11).Value > ws.Cells(2, 17).Value Then
        ws.Cells(2, 17).Value = ws.Cells(p, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(p, 9).Value
    End If
    Next p

'Greatest decrease
For d = 2 To LastNewRow
    If ws.Cells(d, 11).Value < ws.Cells(3, 17).Value Then
        ws.Cells(3, 17).Value = ws.Cells(d, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(d, 9).Value
    End If
    Next d
'greatest total volume
    For V = 2 To LastNewRow
    If ws.Cells(V, 12).Value > ws.Cells(4, 17) Then
        ws.Cells(4, 17).Value = ws.Cells(V, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(V, 9).Value
        End If
    Next V
        
'Making percentages

For i = 2 To LastNewRow
ws.Cells(i, 11).NumberFormat = "0.00%"

Next i
ws.Range("Q2,Q3").NumberFormat = "0.00%"
Next ws
End Sub

