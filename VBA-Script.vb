Option Explicit
Sub Stock_loops()
    
    Dim Stock_Name As String
    Dim Yearly_Change As Double
    Dim Summary_Table_Row As Integer
    Dim Change_Frac As Double
    Dim Total_Stock_Volume As LongLong
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim input_row_num As Long
    Dim Last_Data_Row As Long
    
     'creat tital for each new column
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    Range("o2").Value = "Greatest % Increase"
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "Greatest Total Volume"
    Range("p1").Value = "Ticker"
    Range("q1").Value = "Value"
     
    
     Summary_Table_Row = 2
     Total_Stock_Volume = 0
     Opening_Price = Cells(2, 3).Value
     
     'determine the last row
     Last_Data_Row = Cells(Rows.Count, 1).End(xlUp).Row
     
     'loop thrugh row
     For input_row_num = 2 To Last_Data_Row
        Stock_Name = Cells(input_row_num, 1).Value
        Total_Stock_Volume = Total_Stock_Volume + Cells(input_row_num, 7).Value
        'last row of current stock
        If Cells(input_row_num + 1, 1).Value <> Stock_Name Then
            'input
            Closing_Price = Cells(input_row_num, 6).Value
            'calculation
            Yearly_Change = Closing_Price - Opening_Price
            Change_Frac = Yearly_Change / Opening_Price
            
            
            'output
            Range("i" & Summary_Table_Row).Value = Stock_Name
            Range("j" & Summary_Table_Row).Value = Yearly_Change
            Range("k" & Summary_Table_Row).Value = FormatPercent(Change_Frac)
            Range("l" & Summary_Table_Row).Value = Total_Stock_Volume
            'perpare for next stock
            Summary_Table_Row = Summary_Table_Row + 1
            Opening_Price = Cells((input_row_num + 1), 3).Value
            Total_Stock_Volume = 0
        End If

    Next input_row_num
  
    
   
    MsgBox "done"
End Sub

