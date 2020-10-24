Attribute VB_Name = "Module1"
Sub Summary():

'Declare variables for ticker summary
Dim lastrow As Long
Dim total As Double
Dim summary_row As Integer
Dim openstock As Double
Dim closedstock As Double
Dim ws As Worksheet

'Declare Bonus/Challenge Variables
Dim nrow As Long

For Each ws In Worksheets
    
    'Label headers
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume Stock"
    ws.Range("P1:Q1").Value = Array("Ticker", "Value")

    'Sort data first by date then by ticker name
    ws.Columns("A:G").Sort key1:=ws.Range("B1"), order1:=xlAscending, Header:=xlYes
    ws.Columns("A:G").Sort key1:=ws.Range("A1"), order2:=xlAscending, Header:=xlYes

    '**MAIN SUMMARY**
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set variable initial value
    summary_row = 2
    total = 0
    openstock = ws.Cells(2, 3).Value
    
    'Loops through all the Worksheets in file
    For i = 2 To lastrow

        'Sums the stock volume
        total = total + ws.Cells(i, 7).Value
    
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        
                'Ticker Name
                ws.Cells(summary_row, 9).Value = ws.Cells(i, 1).Value
        
                 'Total Stock Volume
                ws.Cells(summary_row, 12).Value = total
        
                closedstock = ws.Cells(i, 6).Value
        
                'Calculate Yearly Change
                ws.Cells(summary_row, 10).Value = closedstock - openstock
            
                    'Conditional Format Yearly Change with Colour Fill
                    If closedstock > openstock Then
             
                        ws.Cells(summary_row, 10).Interior.ColorIndex = 4
             
                    Else
             
                        ws.Cells(summary_row, 10).Interior.ColorIndex = 3
            
                    End If
        
                    'Calculate and Format Percentage Change (check for when openstock = 0)
                    If openstock = 0 Then
                
                         ws.Cells(summary_row, 11) = "N/A"
            
                    Else
            
                        ws.Cells(summary_row, 11).Value = ws.Cells(summary_row, 10) / openstock
            
                    End If
        
                Cells(summary_row, 11).NumberFormat = "0.00%"
        
                'Next summary row line
                summary_row = summary_row + 1
        
                'Reset total and introduce new openstock for new ticker
                total = 0
                openstock = ws.Cells(i + 1, 3).Value
        
            End If

        Next i
    
    '**BONUS/CHALLENGE SUMMARY**
    nrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    For k = 2 To nrow
    
        'Determine Greatest % Increase using Max function on Percent Change column
        If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K1:K" & nrow)) Then
        
            ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(k, 11).Value

        'Determine Greatest % Decrease using Min function on Percent Change column
        ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K1:K" & nrow)) Then
        
            ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
        
        End If
    
        'Determine Greatest Total Volume using Max function on Total Stock Volume
        If ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L1:L" & nrow)) Then
        
            ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
    
        End If
    
    Next k

Next ws


End Sub
