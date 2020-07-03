Attribute VB_Name = "Module1"
Sub startVBA()

    Dim ws As Worksheet
    Dim yearly_change As Single
    Dim percent_change As Double
    Dim total_stock As Double
    Dim opening_price As Long
    Dim LastRow As Long
    Dim i As Long
    Dim j As Long

    For Each ws In Worksheets
        'variables
        j = 0
        total_stock = 0
        yearly_change = 0
        opening_price = 2
 
        'Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "yearly_change"
        ws.Cells(1, 11).Value = "percent_change"
        ws.Cells(1, 12).Value = "Total_stock_volume"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To LastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    total_stock = total_stock + ws.Cells(i, 7).Value

                    If total_stock = 0 Then
                        ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                        ws.Range("J" & 2 + j).Value = 0
                        ws.Range("K" & 2 + j).Value = "%" & 0
                        ws.Range("L" & 2 + j).Value = 0
    
                Else
                    If ws.Cells(opening_price, 3) = 0 Then
                        For find_value = opening_price To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                    opening_price = find_value
                                    Exit For
                            End If
                        Next find_value
                    End If
                    
                    yearly_change = Round((ws.Cells(i, 6) - ws.Cells(opening_price, 3)), 2)
                    percent_change = Round((yearly_change / ws.Cells(opening_price, 3) * 100), 2)

                    opening_price = i + 1
                    
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Round(yearly_change, 2)
                    ws.Range("K" & 2 + j).Value = "%" & percent_change
                    ws.Range("L" & 2 + j).Value = total_stock
                    
                    'conditions for color format
                    Select Case yearly_change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    End Select
                End If
                   
                'reset variables for new ticker
                yearly_change = 0
                total_stock = 0
                j = j + 1
            Else
                total_stock = total_stock + ws.Cells(i, 7).Value
           End If
                            
        Next i
    Next ws
    MsgBox ("Done!")
End Sub

