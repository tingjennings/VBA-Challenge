Attribute VB_Name = "Module1"
Sub tickertotaler_moderate()


'define everything
Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer
Dim Max As Double
Dim Min As Double



'this prevents my overflow error
On Error Resume Next

'run through each worksheet
For Each ws In ThisWorkbook.Worksheets

    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"


    'setup integers for loop
    Summary_Table_Row = 2
    vol = 0
    year_open = 0
    year_close = 0
    
    'loop
    For i = 2 To ws.UsedRange.Rows.Count
        
        If i = 2 Then
            year_open = ws.Cells(i, 3).Value
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'find all the values
            ticker = ws.Cells(i, 1).Value
            year_close = ws.Cells(i, 6).Value
            
            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_open
            
            'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            If yearly_change >= 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.Color = vbGreen
            Else
                ws.Cells(Summary_Table_Row, 10).Interior.Color = vbRed
            End If
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            vol = vol + ws.Cells(i, 7).Value
            ws.Cells(Summary_Table_Row, 12).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1
        
            ' Reset For Next Symbol
            vol = 0
            year_open = ws.Cells(i + 1, 3).Value
        Else
            vol = vol + ws.Cells(i, 7).Value
        End If
        
        
    'finish loop
     Next i
     
     
    Max = 0

    For i = 2 To ws.UsedRange.Rows.Count
          
       With ws.Cells(i, 11)
       If .Value > Max And .Value < 65000 Then
              Max = .Value
              tick = .Offset(0, -2).Value
       End If
       End With
        
    Next i

    ws.Cells(2, 17).Value = Max
    ws.Cells(2, 16).Value = tick
    
    Min = 0

      For i = 2 To ws.UsedRange.Rows.Count
            
         With ws.Cells(i, 11)
         If .Value < Min And .Value < 65000 Then
                Min = .Value
                tick = .Offset(0, -2).Value
          End If
          End With
    Next i
        ws.Cells(3, 17).Value = Min
        ws.Cells(3, 16).Value = tick
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
    
        ws.Columns("K").NumberFormat = "0.00%"
   
   
    GreatestTotalVolume = 0

      For i = 2 To ws.UsedRange.Rows.Count
            
         With ws.Cells(i, 12)
         If .Value > GreatestTotalVolume Then
                GreatestTotalVolume = .Value
                tick = .Offset(0, -3).Value
          End If
          End With
    Next i
        ws.Cells(4, 17).Value = GreatestTotalVolume
        ws.Cells(4, 16).Value = tick
       
    
'move to next worksheet
Next ws

    MsgBox "DONE"

End Sub
