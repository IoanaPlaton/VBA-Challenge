Attribute VB_Name = "Module1"

'Create a script that loops through all the stocks for one year and outputs the following information:

'The ticker symbol

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The total stock volume of the stock

Sub Stock_Analysis()

    'Set Dimentions
    Dim Total As Double
    Dim Yearly_Change As Double
    Dim Start As Long
    Dim Last_Row As Long
    Dim Percent_Change As Double
    Dim Daily_Change As Integer
    Dim Average_Change As Double
    
    For Each ws In Worksheets
        j = 0
        Total = 0
        Yearly_Change = 0
        Daily_Change = 0
        Percent_Change = 0
        Start = 2

        'set column titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Determine the Last Row
        Last_Row = ws.Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To Last_Row
    
            'if we are with the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Add Total Stock Volume
                Total = Total + ws.Cells(i, 7).Value
             
                If Total = 0 Then
                    'print the results
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
          
                Else
                    If ws.Cells(Start, 3) = 0 Then
                        For find_value = Start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                Start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                
                    Yearly_Change = (ws.Cells(i, 6) - ws.Cells(Start, 3))
                    Percent_Change = Yearly_Change / ws.Cells(Start, 3)
                
                    Start = i + 1
                
                    ws.Range("I" & 2 + j) = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j) = Yearly_Change
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).Value = Percent_Change
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + j).Value = Total
        
                    Select Case Yearly_Change
                       Case Is > 0
                           ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                       Case Is < 0
                           ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                       Case Else
                           ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                        
                End If
        
                Total = 0
                Yearly_Change = 0
                j = j + 1
                Daily_Change = 0
                Percent_Change = 0
                                          
                      
            Else
            'if ticker is the same add results
            Total = Total + ws.Cells(i, 7).Value
        
            End If
        
        Next i
    
        'add in max and min
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & Last_Row)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & Last_Row)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & Last_Row))
        
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Last_Row)), ws.Range("K2:K" & Last_Row), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Last_Row)), ws.Range("K2:K" & Last_Row), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Last_Row)), ws.Range("L2:L" & Last_Row), 0)
        
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
    
        
    Next ws

End Sub
