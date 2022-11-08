Attribute VB_Name = "Module1"
Sub Stock_Analysis()

Dim w As Worksheet
For Each ws In Worksheets

Dim Stock_Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Total_Volume As Double
Dim Row_Counter As Long

Yearly_Change_Row = 2
Percent_Change_Row = 2
Opening_Price_Row = 2
Total_Volume_Row = 2
Row_Counter = 2
last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To last_row

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    Stock_Ticker = ws.Cells(i, 1).Value
    Opening_Price = ws.Cells(Opening_Price_Row, 3).Value
    Closing_Price = ws.Cells(i, 6).Value
    Yearly_Change = Closing_Price - Opening_Price
    Percent_Change = Yearly_Change / Opening_Price
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
 
    
    ws.Range("I" & Row_Counter).Value = Stock_Ticker
    ws.Range("J" & Row_Counter).Value = Yearly_Change
    ws.Range("K" & Row_Counter).Value = Percent_Change
    ws.Range("L" & Total_Volume_Row).Value = Total_Volume
    Row_Counter = Row_Counter + 1
    Total_Volume = 0

    Else:
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        Total_Volume_Row = Total_Volume_Row + 1
        ws.Range("L" & Total_Volume_Row).Value = Total_Volume
    
    End If
    

Next i

Next ws

End Sub
