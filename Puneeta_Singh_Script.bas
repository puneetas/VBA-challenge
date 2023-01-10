Attribute VB_Name = "Module1"
Sub challenge_2()
    For Each ws In Worksheets

'creating summary table headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

'summary table declarations
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As LongLong

Summary_Table_Row = 2

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

total_volume = 0

        
        ' Loop through all stocks and pull required info
        For i = 1 To lastRow
        
            If i <> 1 Then
            
                'total volume calculation
                total_volume = total_volume + Cells(i, 7).Value
                
            End If
                    
            ' Check if we are still within the same tickers, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                If i <> 1 Then
                
                    'TICKER
                    ' Set the Ticker name
                    Ticker_Name = ws.Cells(i, 1).Value
                    
                    'grab closing value
                    closing_value = ws.Cells(i, 6).Value
                    
                    'yearly change
                    yearly_change = closing_value - opening_value
                    
                    'percent change
                    percent_change = (yearly_change / opening_value)
                    
                    ' Print ticker name in the Summary Table
                    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                    
                    'print yearly change in summary table
                    ws.Range("J" & Summary_Table_Row).Value = yearly_change
                    
                    If yearly_change > 0 Then
                    'set colour to green
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        
                        ElseIf yearly_change < 0 Then
                        'set colour to red
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        
                    End If
                    
                    ' Print percent change in the Summary Table
                    ws.Range("K" & Summary_Table_Row).Value = percent_change
                
                    ' Print total volume in the Summary Table
                    ws.Range("L" & Summary_Table_Row).Value = total_volume
                    
                    ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    total_volume = 0
                
                End If
                
                'grab opening value
                opening_value = ws.Cells(i + 1, 3).Value
                
                
            End If
        
         
            
        Next i
        
'setting up secondary table
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'greatest increase percent change
    max_value = WorksheetFunction.Max(ws.Range("K:K").Value)
    row_index = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(max_value, ws.Range("K:K"), 0))

'greatest increase in secondary summary table
    ws.Range("P2").Value = max_value
    ws.Range("O2").Value = row_index
    
'greatest decrease percent change
    min_value = WorksheetFunction.Min(ws.Range("K:K").Value)
    row_index_min = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(min_value, ws.Range("K:K"), 0))

'greatest decrease in secondary summary table
    ws.Range("P3").Value = min_value
    ws.Range("O3").Value = row_index_min

'greatest Total Volume
    volume_value = WorksheetFunction.Max(ws.Range("L:L").Value)
    row_index_volume = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(volume_value, ws.Range("L:L"), 0))
    
'greatest total volume in secondary summary table
    ws.Range("P4").Value = volume_value
    ws.Range("O4").Value = row_index_volume

'percent change formatting
ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("P:P").NumberFormat = "0.00%"
ws.Range("P4").NumberFormat = "0"


    Next ws
    
End Sub
