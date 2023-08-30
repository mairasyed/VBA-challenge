Attribute VB_Name = "Module1"
Sub Stockcounter():

    For Each ws In Worksheets
    
    'Declaring variables:
    Dim ticker_Symbol As String
    Dim yearly_Change As Double
    Dim percentage_Change As Double
    Dim total_Stock As Double
    Dim greatest_Increase As Double
    Dim greatest_Decrease As Double
    Dim greatest_Total_volume As Double
    Dim open_Value As Double
    Dim close_Value As Double
    Dim output_Table As Integer


'Defining worksheets name

WorksheetName = ws.Name
    
    
    'Initialize the variables
    
    greatest_Increase = 0
    
    greatest_Decrease = 0
    
    greatest_Total_volume = 0
    
    output_Table = 2
    
    opening_Price = ws.Cells(2, 3).Value
    closing_Price = ws.Cells(2, 6).Value
    
    'Define last row
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Define output table headers
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    
    total_Stock = 0
    open_Value = ws.Cells(2, 3).Value
    
    For i = 2 To Lastrow
    
    total_Stock = total_Stock + ws.Cells(i, 7).Value
    
    'Check ticker value changes
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ws.Cells(output_Table, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(output_Table, 12).Value = total_Stock
    
    'Explain the yearly change
    yearly_Change = ws.Cells(i, 6).Value - open_Value
    ws.Cells(output_Table, 10) = yearly_Change
    ws.Cells(output_Table, 10).NumberFormat = "$#,##0.00"
    
   'To find out the percentage change
   
    percentage_Change = (ws.Cells(i, 6).Value - open_Value) / open_Value
    ws.Cells(output_Table, 11) = percentage_Change
    ws.Cells(output_Table, 11).NumberFormat = "0.00%"
   
    
'To find out the greatest % increase
    
    If ws.Cells(output_Table, 11).Value > greatest_Increase Then
    
    ws.Cells(2, 17).Value = ws.Cells(output_Table, 11).Value
    
    greatest_Increase = ws.Cells(output_Table, 11).Value
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ws.Cells(2, 16).Value = ws.Cells(output_Table, 9).Value
    
    End If
    

'To find out the greatest % decrease
    
    If ws.Cells(output_Table, 11).Value < greatest_Decrease Then
    
    ws.Cells(3, 17).Value = ws.Cells(output_Table, 11).Value
    
     greatest_Decrease = ws.Cells(output_Table, 11).Value
     
     ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Cells(3, 16).Value = ws.Cells(output_Table, 9).Value
    

End If


'To find out the greatest total volume

 If ws.Cells(output_Table, 12).Value > greatest_Total_volume Then
    
    ws.Cells(4, 17).Value = ws.Cells(output_Table, 12).Value
    
   greatest_Total_volume = ws.Cells(output_Table, 12).Value
    
    ws.Cells(4, 16).Value = ws.Cells(output_Table, 9).Value

End If


'Conditional formatting (colour coding) for Yearly Change

If ws.Cells(output_Table, 10).Value >= 0 Then

ws.Cells(output_Table, 10).Interior.ColorIndex = 4

Else: ws.Cells(output_Table, 10).Interior.ColorIndex = 3

End If


'Conditional formatting (colour coding) for Percentage Change

If ws.Cells(output_Table, 11).Value >= 0 Then

ws.Cells(output_Table, 11).Interior.ColorIndex = 4

Else: ws.Cells(output_Table, 11).Interior.ColorIndex = 3

End If



    output_Table = output_Table + 1
    total_Stock = 0
    open_Value = ws.Cells(i + 1, 3).Value
    close_Value = ws.Cells(i + 1, 6).Value

    
    End If
    
    ws.Cells(4, 17).Value = Format(greatest_Total_volume, "Scientific")
    
       Next i
    
    
    Worksheets(WorksheetName).Columns("A:Z").AutoFit
    
    
    Next ws
    
End Sub

