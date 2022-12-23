Sub Summarize_Data()

'go through each worksheet
For Each ws In Worksheets

Dim lRow As Long
Dim tickRow As Integer
Dim ticker As String
Dim stockVol As LongLong
Dim Edate, Ldate As Double

'get last row in data
lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'start totals line
tickRow = 2
'start stock volume
stockVol = 0
'get first opening price of the worksheet
Edate = ws.Cells(2, 3).Value
    
    'go through each row of data
    For i = 2 To lRow
    
        'check if ticker is unique
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'get ticker
            ticker = ws.Cells(i, 1).Value
            'get last stock volume
            stockVol = stockVol + ws.Cells(i, 7).Value
            'get closing price
            Ldate = ws.Cells(i, 6).Value
            'print ticker
            ws.Range("I" & tickRow).Value = ticker
            'print stock volume
            ws.Range("L" & tickRow).Value = stockVol
            'print yearly change and format color
            ws.Range("J" & tickRow).Value = (Ldate - Edate)
                If ws.Range("J" & tickRow).Value >= 0 Then
                    ws.Range("J" & tickRow).Interior.ColorIndex = 4
                ElseIf ws.Range("J" & tickRow).Value < 0 Then
                    ws.Range("J" & tickRow).Interior.ColorIndex = 3
                End If
            'print percent change
            ws.Range("K" & tickRow).Value = ((Ldate / Edate) - 1)
            'add the next line for the totals
            tickRow = tickRow + 1
            'reset stock volume
            stockVol = 0
            'get the new opening price
            Edate = ws.Cells(i + 1, 3).Value
            
        Else
            'summing the stock volume
            stockVol = stockVol + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    'changing the percent column format
    For i = 2 To tickRow
        ws.Cells(i, 11).Style = "percent"
        ws.Cells(i, 11).NumberFormat = "0.00%"
    Next i
    
    'Adding Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Add Greatest Section
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"

    ws.Cells(1, 16).Value = "Ticker"
    ws.Range("P2:P4").Value = ws.Cells(2, 9).Value
    ws.Cells(1, 17).Value = "Value"
    ws.Range("Q2:Q4").Value = ws.Cells(2, 11).Value
    ws.Range("Q2:Q3").Style = "percent"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Find Greatest Values
    Dim inc, dec As Double
    Dim vol As LongLong
    inc = ws.Range("K2").Value
    dec = ws.Range("K2").Value
    vol = ws.Range("L2").Value
    
    'Compare all rows to find Values
    For i = 2 To tickRow
        'Determine if i has a larger increase percentage
        If inc < ws.Cells(i, 11).Value Then
            inc = ws.Cells(i, 11).Value
            ws.Range("P2").Value = ws.Cells(i, 9).Value
            ws.Range("Q2").Value = ws.Cells(i, 11).Value
        Else: End If
        'Determine if i has a larger decrease percentage
        If dec > ws.Cells(i, 11).Value Then
            dec = ws.Cells(i, 11).Value
            ws.Range("P3").Value = ws.Cells(i, 9).Value
            ws.Range("Q3").Value = ws.Cells(i, 11).Value
        Else: End If
        'Determine if i has a larger increase in volume
        If vol < ws.Cells(i, 12).Value Then
            vol = ws.Cells(i, 12).Value
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            ws.Range("Q4").Value = ws.Cells(i, 12).Value
        Else: End If
        
    Next i
    
    ws.Columns("I:Q").AutoFit
    
Next ws

End Sub
