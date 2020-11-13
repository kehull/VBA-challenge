

Sub stockprocess()

' Setup
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 19).Value = "Ticker"
Cells(1, 20).Value = "Value"
Columns("A:T").AutoFit


' Print Ticker and Total Stock Volume to Summary Table

' Variables
Dim Ticker As String
Dim Stock_Total As Double
Stock_Total = 0
Dim Stummary_Table_Row As Integer
Summary_Table_Row = 2
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row


' Loop through stocks
For i = 2 To LastRow

' Determine if next ticker matches the previous ticker or not
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ' If ticker does not match:
        ' Set ticker value to new ticker
            Ticker = Cells(i, 1).Value
        ' Add to stock total volume
            Stock_Total = Stock_Total + Cells(i, 7).Value
        ' Print ticker to summary table
            Range("I" & Summary_Table_Row).Value = Ticker
      ' Print the stock total volume to summary table
            Range("L" & Summary_Table_Row).Value = Stock_Total
      ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      ' Reset the stock total
            Stock_Total = 0
      ' Set first opening price of a ticker as a variable
            Dim Opening_Price As Double
            Opening_Price = Cells(i + 1, 3).Value

    ' If the ticker does match:
    Else

      ' Add to the stock total
            Stock_Total = Stock_Total + Cells(i, 3).Value
      ' Set closing price of ticker as a variable- this version will keep setting every new closing price
      ' as the last price and will stop updating the value when the ticker cells no longer match
            Dim Closing_Price As Double
            Closing_Price = Cells(i, 6).Value
        ' Set variable so i don't have to type out this reference cell a million times
        Dim Yearly_Change As Range
        Set Yearly_Change = Cells(Summary_Table_Row, 10)
        ' Calculate yearly change
            Yearly_Change.Value = (Closing_Price - Opening_Price)
            ' Conditional formatting turns cell green if value is positive, red if negative
                If Yearly_Change.Value < 0 Then
                    Yearly_Change.Interior.Color = vbRed
                Else
                    Yearly_Change.Interior.Color = vbGreen
                End If
            ' Calculate percent change
                
                'Cells(Summary_Table_Row, 11).Value = ((Yearly_Change.Value / 4) * 100)
    End If
Next i
End Sub

