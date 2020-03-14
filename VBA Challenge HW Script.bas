Attribute VB_Name = "Module1"
Sub tickertotaler()


Dim ws As Worksheet
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer


For Each ws In ThisWorkbook.Worksheets
  
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    Summary_Table_Row = 2
    year_open = ws.Cells(2, 3).Value


        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
         For i = 2 To lastrow
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             

            ticker = ws.Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value

            
            year_close = ws.Cells(i, 6).Value

            yearly_change = year_close - year_open
            If (year_open <> 0) Then
            
            percent_change = (year_close - year_open) / year_open
            Else
            percentage_change = 0
            End If
            
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = vol
            If percent_change >= 0 Then
            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
            Else
            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
            End If
            
            Summary_Table_Row = Summary_Table_Row + 1

             vol = 0
             year_open = ws.Cells(i + 1, 3).Value
        Else: vol = vol + ws.Cells(i, 7).Value
        
        End If

    Next i
    
'ws.Columns("K").NumberFormat = "0.00%"

    'Dim rg As Range
    'Dim g As Long
    'Dim c As Long
    'Dim color_cell As Range
    
    
    'c = rg.Cells.Count
    
    'For g = 1 To c
    'Set color_cell = rg(g)
    'Select Case color_cell
    '    Case Is >= 0
    ''        With color_cell
     ''           .Interior.Color = vbGreen
      '      End With
       ' Case Is < 0
       '     With color_cell
       '         .Interior.Color = vbRed
       '     End With
       'End Select
   ' Next g

Next ws

MsgBox ("Bun")
End Sub



