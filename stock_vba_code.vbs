
Sub Stocks()

For Each ws in Worksheets

Dim WorksheetName As String
WorksheetName = ws.Name

Dim stock_total as Double
stock_total = 0

Dim open_price as Double
Dim close_price as Double
Dim yearly_change as Double
Dim percent_change as Double

Dim summary_row as Integer
summary_row = 2

'Sort data

ws.Columns("A:G").Sort key1:=ws.Range("B2"), _
    order1:=xlAscending, Header:=xlYes
ws.Columns("A:G").Sort key1:=ws.Range("A2"), _
    order1:=xlAscending, Header:=xlYes

'Assign names for output columns

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

Last_Row =  ws.Cells(Rows.Count, "A").End(xlUp).Row

'Loop for Ticker and Stock Volume

For row = 2 to Last_Row

    IF ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then

        stock_total = stock_total + ws.Cells(row, 7).Value

        ws.Range("I" & summary_row).Value = ws.Cells(row, 1).Value
        ws.Range("L" & summary_row).Value = stock_total
        
        summary_row = summary_row + 1
        stock_total = 0

    ELSE
        stock_total = stock_total + ws.Cells(row, 7).Value
 
    End IF 

Next row  

' Set open price and re-set summary row

open_price = ws.Cells(2, 3).Value
summary_row = 2

' Loop for yearly change and percent change

For row = 2 to Last_Row

    IF ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
        close_price = ws.Cells(row, 6).Value

        yearly_change = close_price - open_price
        
        IF open_price <> 0 THEN
        percent_change = (yearly_change/ open_price) * 100
        open_price = ws.Cells(row + 1, 3).Value

        ELSE
        percent_change = 0
        open_price = ws.Cells(row + 1, 3).Value

        End IF
        
        ws.Range("J" & summary_row).Value = yearly_change
        ws.Range("K" & summary_row).Value = (Cstr(percent_change) & "%")

        summary_row = summary_row + 1

    End If  

Next row    

' Loop for conditional formatting

Last_Row_2 =  ws.Cells(Rows.Count, "J").End(xlUp).Row
ws.Columns("I:L").EntireColumn.AutoFit

For row = 2 to Last_Row_2

    IF ws.Cells(row, 10) < 0 THEN
    ws.Cells(row,10).Interior.ColorIndex = 3

    ELSE
    ws.Cells(row,10).Interior.ColorIndex = 4

    End IF

Next row  

Next ws

End Sub 