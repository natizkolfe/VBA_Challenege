# VBA_Challenege
Sub stocks()

'   Set variables
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim change_price As Double
Dim percent_change As Double
Dim volume As LongLong

'
For Each ws In Worksheets

    '   Set table headers
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Change_year"
    ws.Range("L1").Value = "percent_Change"
    ws.Range("M1").Value = "Total Volume"

    
    open_price = ws.Cells(2, 3)
    volume = 0

   
    Dim summary As Integer
    summary = 2

    
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

     
        For i = 2 To last_row
    
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                
                ticker = ws.Cells(i, 1).Value
                close_price = ws.Cells(i, 6).Value
    
                
                change_price = close_price - open_price
                
                
                If open_price <> 0 Then
                    percent_change = (change_price / open_price) * 100
                Else
                    percent_change = 0
                End If
                
                
                volume = volume + ws.Cells(i, 7).Value
                
                
                ws.Range("J" & summary).Value = ticker
                ws.Range("K" & summary).Value = change_price
                
                
                    If change_price < 0 Then
                        ws.Range("K" & summary).Interior.ColorIndex = 3
                    Else
                        ws.Range("K" & summary).Interior.ColorIndex = 4
                    End If
                
                ws.Range("L" & summary).Value = Round(percent_change, 2) & "%"
                
                ws.Range("M" & summary).Value = volume
                
              
                summary = summary + 1
                open_price = ws.Cells(i + 1, 3).Value
                volume = 0
            Else
                volume = volume + ws.Cells(i, 7).Value
            End If
        Next i
Next ws
End Sub
