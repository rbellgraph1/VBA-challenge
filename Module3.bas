Attribute VB_Name = "Module3"


'working sheet that is set to only invidual sheet code now

Sub Multiple_year_stock_data_indvidual_Form2():

    Dim ws As Worksheet
    For Each ws In Worksheets
        'lastRow3 = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row
        'lastRowYear = cells(Rows.Count, "A").End(xlUp).Row - 1
        
        Cells(1, 10).Value = "Ticker"
        Cells(1, 11).Value = "Yearly Change"
        Cells(1, 12).Value = "Percent Change"
        Cells(1, 13).Value = "Total Stock Value"
        Cells(1, 17).Value = "Ticker"
        Cells(1, 18).Value = "Value"
        Cells(2, 16).Value = "Greatest % increase"
        Cells(3, 16).Value = "Greatest % Decrease"
        Cells(4, 16).Value = "Greatest Total Volume"
        
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
        LastRow = Cells(Rows.Count, 11).End(xlUp).row
        
        Dim LastRow2 As Long
        LastRow2 = Cells(Rows.Count, 1).End(xlUp).row
        
        Range("L2:L" & LastRow).NumberFormat = ".00%"
        Range("A1:S5").Columns.AutoFit
'        open_price = Cells(2, 3).Value
        Range("R2").NumberFormat = ".00%"
        Range("R3").NumberFormat = ".00%"
        
       Dim i As Long
        For i = 2 To LastRow2
        '  first time i am checking to see the ticker
        
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    open_price = Cells(i, 3).Value
                    
                     
            End If
           
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                Ticker = Cells(i, 1).Value
                Total = Total + Cells(i, 7).Value
                close_price = Cells(i, 6).Value
                
                Yearly_Change = close_price - open_price
                PercentChange = Yearly_Change / open_price
                Range("J" & Summary_Table).Value = Ticker
                Range("K" & Summary_Table).Value = Yearly_Change
                Range("L" & Summary_Table).Value = PercentChange
                Range("M" & Summary_Table).Value = Total
                                
                Select Case Yearly_Change
                Case Is > 0
                'you are coloring the range of the row
                Range("K" & Summary_Table).Interior.ColorIndex = 4
                Case Is < 0
                'you are coloring the range of the row
                Range("K" & Summary_Table).Interior.ColorIndex = 3
                Case Else
                'you are coloring the range of the row
                Range("K" & Summary_Table).Interior.ColorIndex = 6
                End Select
                
                'Reset the Total
                Summary_Table = Summary_Table + 1
                Total = 0
                
                
                Else
                'add to the Total
                Total = Total + Cells(i, 7).Value
                
                
                
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
        Range("R2") = MaxPercentincrease
        Range("Q2") = MaxPercentincreaseTicker
        Range("R3") = MaxPercentDecrease
        Range("Q3") = MaxPercentDecreaseTicker
        Range("R4") = MaxTotalVolume
        Range("Q4") = MaxTotalVolumeTicker
        Range("A1:S5").Columns.AutoFit
    Next ws
    
End Sub




