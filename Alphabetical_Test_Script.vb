Sub Alphabetical_Test()

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
            ws.Range("L" & tickRow).Value = ticker
            'print stock volume
            ws.Range("O" & tickRow).Value = stockVol
            'print yearly change and format color
            ws.Range("M" & tickRow).Value = (Ldate - Edate)
                If ws.Range("M" & tickRow).Value >= 0 Then
                    ws.Range("M" & tickRow).Interior.ColorIndex = 4
                ElseIf ws.Range("M" & tickRow).Value < 0 Then
                    ws.Range("M" & tickRow).Interior.ColorIndex = 3
                End If
            'print percent change
            ws.Range("N" & tickRow).Value = ((Ldate / Edate) - 1)
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
        ws.Cells(i, 14).Style = "percent"
        ws.Cells(i, 14).NumberFormat = "0.00%"
    Next i
    
    ws.Range("L1:O1").Value = Worksheets("A").Range("L1:O1").Value
    ws.Columns("L:O").AutoFit
    
    'Add Greatest Section

    ws.Cells(2, 19).Value = "Greatest % increase"
    ws.Cells(3, 19).Value = "Greatest % decrease"
    ws.Cells(4, 19).Value = "Greatest total volume"

    ws.Cells(1, 20).Value = "Ticker"
    ws.Range("T2:T4").Value = ws.Cells(2, 12).Value
    ws.Cells(1, 21).Value = "Value"
    ws.Range("U2:U4").Value = ws.Cells(2, 14).Value
    ws.Range("U2:U3").NumberFormat = "0.00%"
    
    'Find Greatest Values
    
    Dim inc, dec As Double
    Dim vol As LongLong
    inc = ws.Range("N2").Value
    dec = ws.Range("N2").Value
    vol = ws.Range("O2").Value
    
    'Compare all rows to find Values
    For i = 2 To tickRow
        'Determine if i has a larger increase percentage
        If inc < ws.Cells(i, 14).Value Then
            inc = ws.Cells(i, 14).Value
            ws.Range("T2").Value = ws.Cells(i, 12).Value
            ws.Range("U2").Value = ws.Cells(i, 14).Value
        Else: End If
        'Determine if i has a larger decrease percentage
        If dec > ws.Cells(i, 14).Value Then
            dec = ws.Cells(i, 14).Value
            ws.Range("T3").Value = ws.Cells(i, 12).Value
            ws.Range("U3").Value = ws.Cells(i, 14).Value
        Else: End If
        'Determine if i has a larger increase in volume
        If vol < ws.Cells(i, 15).Value Then
            vol = ws.Cells(i, 15).Value
            ws.Range("T4").Value = ws.Cells(i, 12).Value
            ws.Range("U4").Value = ws.Cells(i, 15).Value
        Else: End If
        
    Next i
    
    ws.Columns("S:U").AutoFit
    
Next ws

End Sub
