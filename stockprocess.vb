Sub stockprocess()

'Complete this process on each worksheet in the dataset
Dim WS_Count As Integer
Dim x As Integer
WS_Count = ActiveWorkbook.Worksheets.Count
Dim ws As Worksheet

For Each ws In Worksheets

' Begin the loop.
For x = 1 To WS_Count

' Variables
Dim Ticker As String
Dim Stock_Total As Double
    Stock_Total = 0
Dim Stummary_Table_Row As Integer
    Summary_Table_Row = 2
Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim Opening_Price_Checker As String
    Opening_Price_Checker = "nope"
    
' Loop through stocks
For I = 2 To LastRow
    ' Determine if next ticker value matches the previous ticker value or not
If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
    'If ticker does not match:
    ' Set ticker value to new ticker
    Ticker = ws.Cells(I, 1).Value
    ' Add to stock total volume
    Stock_Total = Stock_Total + ws.Cells(I, 7).Value
    ' Print ticker to summary table
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    ' Print the stock total volume to summary table
    ws.Range("L" & Summary_Table_Row).Value = Stock_Total
    ' Reset the stock total
    Stock_Total = 0
    ' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1


    ' If the ticker does match:
    Else
        ' Set first opening price of a ticker as a variable
        ' The first time this code runs on a sheet, Opening_Price_Checker
        ' already = "nope". Because of this, the code will pull the first
        ' value in the row where the loop starts. This section then changes
        ' Opening_Price_Checker to "yup", which causes the statement to evaluate
        ' as false in every subsequent iteration, preventing it from pulling
        ' any more values from the opening price column.
        If Opening_Price_Checker = "nope" Then
            Opening_Price_Checker = "yup"
            Opening_Price = ws.Cells(I, 3).Value
        Else
        End If
        'Set last closing price of a ticker as a variable
        ' Because Opening_Price_Checker now = "yup", after retrieving
        ' the first opening price, the code exits that part of the statement
        ' and enters this section, which checks for the final appearance of
        ' a particular ticker value and then retrieves the closing price
        ' value from its row.
        If ws.Cells(I + 1, 6).Value <> ws.Cells(I, 6).Value Then
            Closing_Price = ws.Cells(I + 1, 6).Value
        Else
        End If
        ' Because Opening_Price_Checker has global scope, once the code has
        ' run all the way through and starts the next iteration, Opening_Price_Checker
        ' will = "nope" again, allowing the code to retrieve the opening price
        ' for each ticker.
   
    
    ' Add to the stock total
    Stock_Total = Stock_Total + ws.Cells(I, 7).Value
    ' Set variable so i don't have to type out this reference cell a million times
    Dim Yearly_Change As Double
    ' Calculate yearly change
    Yearly_Change = Closing_Price - Opening_Price
    
    ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
    ' Conditional formatting turns cell green if value is positive, red if negative
        If Yearly_Change < 0 Then
            ws.Cells(Summary_Table_Row, 10).Interior.Color = vbRed
        Else
            ws.Cells(Summary_Table_Row, 10).Interior.Color = vbGreen
        End If
    ' Calculate percent change
    ws.Cells(Summary_Table_Row, 11).Value = (Yearly_Change / Opening_Price)
    ws.Cells(Summary_Table_Row, 11).NumberFormat = "#,##0.00" + "%"
    End If
Next I

' set variables for finding the greatest increase, greatest decrease, and greatest total volume
' use the min/max functions to solve each variable
' print the solved variables to the appropriate cells
Dim Greatest_Increase As Double
    Greatest_Increase = Application.WorksheetFunction.Max(ws.Columns("J"))
    ws.Range("Q2").Value = Greatest_Increase
Dim Greatest_Decrease As Double
    Greatest_Decrease = Application.WorksheetFunction.Min(ws.Columns("J"))
    ws.Range("Q3").Value = Greatest_Decrease
Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = Application.WorksheetFunction.Max(ws.Columns("L"))
    ws.Range("Q4").Value = Greatest_Total_Volume
    
' Set up additional cells.
' Because the data is laid out consistantly across all sheets, we can assume
' the exact coordinates of these cells will be the same on every sheet.
' This is not alway the case in every project, but I'm taking advantage
' of it here.
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

' this part isn't necessary to the assignment, it just makes the tables easier to look at
ws.Columns("A:T").AutoFit

' Go to the next worksheet
Next x
Next ws
End Sub
