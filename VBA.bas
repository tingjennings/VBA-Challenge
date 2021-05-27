Attribute VB_Name = "Module2"


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

'this prevents my overflow error
On Error Resume Next

'run through each worksheet
For Each ws In ThisWorkbook.Worksheets
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'setup integers for loop
    Summary_Table_Row = 2
    vol = 0
    year_open = 0
    year_close = 0
    
    'loop
    For I = 2 To ws.UsedRange.Rows.Count
        
        If I = 2 Then
            year_open = ws.Cells(I, 3).Value
        ElseIf ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
            'find all the values
            ticker = ws.Cells(I, 1).Value
            year_close = ws.Cells(I, 6).Value
            
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
            vol = vol + ws.Cells(I, 7).Value
            ws.Cells(Summary_Table_Row, 12).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1
        
            ' Reset For Next Symbol
            vol = 0
            year_open = ws.Cells(I + 1, 3).Value
        Else
            vol = vol + ws.Cells(I, 7).Value
        End If
        
        
    'finish loop
    Next I
    
    ws.Columns("K").NumberFormat = "0.00%"
   
'move to next worksheet
Next ws

    MsgBox "DONE"

End Sub
