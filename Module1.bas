Attribute VB_Name = "Module1"
Sub Stock_market()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Double
Dim open_date As Double
open_date = 2

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    Summary_Table_Row = 2

        For I = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
            ticker = ws.Cells(I, 1).Value
            year_open = ws.Cells(open_date, 3).Value
            year_close = ws.Cells(I, 6).Value

            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_open

            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
               
                ws.Cells(Summary_Table_Row, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(open_date, 7), ws.Cells(I, 7)))
                Summary_Table_Row = Summary_Table_Row + 1
                open_date = I + 1
                          
          End If

    Next I
    
ws.Columns("K").NumberFormat = "0.00%"

    Dim rg As Range
    Dim g As Long
    Dim RowCount As Long
    Dim color_cell As Range
    
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    Next ws


End Sub
