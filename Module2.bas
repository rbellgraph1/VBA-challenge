Attribute VB_Name = "Module2"
Sub Multiple_year_stock_data_indvidual_Form2():

    Dim ws As Worksheet
    For Each ws In Worksheets
        'lastRow3 = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row
        'lastRowYear = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
        
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Value"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        'Next w2
        'Dim w2 As Worksheet
        'For Each w2 in Worksheets
        
        Dim Ticker As String
        
        Dim Total As Double
        Total = 0
        Dim open_price As Double
'        open_price = 0
        Dim close_price As Double
'        close_price = 0

        Dim PercentChange As Double
        Dim MaxPercentincrease As Double
        Dim MaxPercentincreaseTicker As String
        Dim MaxPercentDecrease As Double
        Dim MaxPercentDecreaseTicker As String
        Dim MaxTotalVolume As Double
        Dim MaxTotalVolumeTicker As String
        Dim Yearly_Change As Double
                
        Dim Summary_Table As Integer
        Summary_Table = 2
        
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 11).End(xlUp).row
        
        Dim LastRow2 As Long
        LastRow2 = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ws.Range("L2:L" & LastRow).NumberFormat = ".00%"
        ws.Range("A1:S5").Columns.AutoFit
'        open_price = Cells(2, 3).Value
        ws.Range("R2").NumberFormat = ".00%"
        ws.Range("R3").NumberFormat = ".00%"
        
       Dim i As Long
        For i = 2 To LastRow2
        '  first time i am checking to see the ticker
        
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    open_price = ws.Cells(i, 3).Value
                     
            End If
           
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                close_price = ws.Cells(i, 6).Value
                Total = Total + ws.Cells(i, 7).Value
                Yearly_Change = close_price - open_price
                PercentChange = Yearly_Change / open_price
                ws.Range("J" & Summary_Table).Value = Ticker
                ws.Range("K" & Summary_Table).Value = Yearly_Change
                ws.Range("L" & Summary_Table).Value = PercentChange
                ws.Range("M" & Summary_Table).Value = Total
                
                Select Case Yearly_Change
                Case Is > 0
                'you are coloring the range of the row
                ws.Range("K" & Summary_Table).Interior.ColorIndex = 4
                Case Is < 0
                'you are coloring the range of the row
                ws.Range("K" & Summary_Table).Interior.ColorIndex = 3
                Case Else
                'you are coloring the range of the row
                ws.Range("K" & Summary_Table).Interior.ColorIndex = 6
                End Select
                
                'Reset the Total
                Summary_Table = Summary_Table + 1
                Total = 0
                
                Else
                'add to the Total
                Total = Total + ws.Cells(i, 7).Value
                
                
                End If
                
                
                                             
                
                If PercentChange > MaxPercentincrease Then 'this give me greatest increase
                   MaxPercentincrease = PercentChange
                                
                   MaxPercentincreaseTicker = Ticker   ' ticker value for greatest table
                End If
                
                If PercentChange < MaxPercentDecrease Then
                   MaxPercentDecrease = PercentChange

                   MaxPercentDecreaseTicker = Ticker
                End If
                
                If Total > MaxTotalVolume Then
                   MaxTotalVolume = Total

                   MaxTotalVolumeTicker = Ticker
                                 
                End If
                
                
'                '       Reset the Total
'                        Total = 0
'                        open_price = 0
'                        close_price = 0
            
'            Elseif Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
'                    open_price = Cells(i, 3).Value
'                    Total = Total + Cells(i, 7).Value
                    'if the cell immediately following a row is the same brand ...
'            Else
'                'add to the Total
'                Total = Total + Cells(i, 7).Value
                
                
'            End if
            
        
        
            
        'Summary_table section
        Next i
        ws.Range("R2") = MaxPercentincrease
        ws.Range("Q2") = MaxPercentincreaseTicker
        ws.Range("R3") = MaxPercentDecrease
        ws.Range("Q3") = MaxPercentDecreaseTicker
        ws.Range("R4") = MaxTotalVolume
        ws.Range("Q4") = MaxTotalVolumeTicker
        ws.Range("A1:S5").Columns.AutoFit
    Next ws
    
End Sub



