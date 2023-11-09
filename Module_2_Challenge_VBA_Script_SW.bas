Attribute VB_Name = "Module2Challenge"
Sub Module_2_Challenge()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim Ticker As String
    Dim Open_Price As Double
        Open_Price = 0
    Dim Close_Price As Double
        Close_Price = 0
    Dim Yearly_Change As Double
        Yearly_Change = Close_Price - Open_Price
    Dim Percent_Change As Double
        Percent_Change = 0
    Dim Stock_Volume As Double
        Stock_Volume = 0
        
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
    Dim LastRow As Double
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            Open_Price = Open_Price + ws.Cells(i, 3).Value
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            Close_Price = ws.Cells(i, 6).Value
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            Yearly_Change = Close_Price - Open_Price
            Percent_Change = (Close_Price - Open_Price) / (Open_Price)
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).Style = "Percent"
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
            Summary_Table_Row = Summary_Table_Row + 1
            Open_Price = 0
            Close_Price = 0
            Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0
        Else
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        End If
    Next i

Next ws

For Each ws In Worksheets
    
    Dim TableLastRow As Integer
        TableLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Dim Greatest_Increase As Double
        Greatest_Increase = 0
    Dim Greatest_Decrease As Double
        Greatest_Decrease = 0
    Dim Greatest_Volume As Double
        Greatest_Volume = 0
    
    For t = 2 To TableLastRow
        If ws.Cells(t, 9).Value = ws.Cells(2, 9).Value Then
            Greatest_Increase = Greatest_Increase + ws.Cells(t, 11).Value
        ElseIf ws.Cells(t, 11).Value > Greatest_Increase Then
            Greatest_Increase = ws.Cells(t, 11).Value
        Else
        End If
    
        If ws.Cells(t, 9).Value = ws.Cells(2, 9).Value Then
            Greatest_Decrease = Greatest_Decrease + ws.Cells(t, 11).Value
        ElseIf ws.Cells(t, 11).Value < Greatest_Decrease Then
            Greatest_Decrease = ws.Cells(t, 11).Value
        Else
        End If
    
        If ws.Cells(t, 9).Value = ws.Cells(2, 9).Value Then
            Greatest_Volume = Greatest_Volume + ws.Cells(t, 12).Value
        ElseIf ws.Cells(t, 12).Value > Greatest_Volume Then
            Greatest_Volume = ws.Cells(t, 12).Value
        Else
        End If
    
        ws.Cells(2, 17).Value = Greatest_Increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = Greatest_Decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = Greatest_Volume
    
        If ws.Cells(t, 11).Value = ws.Cells(2, 17).Value Then
            ws.Cells(2, 16).Value = ws.Cells(t, 9)
        End If
        If ws.Cells(t, 11).Value = ws.Cells(3, 17).Value Then
            ws.Cells(3, 16).Value = ws.Cells(t, 9)
        End If
        If ws.Cells(t, 12).Value = ws.Cells(4, 17).Value Then
            ws.Cells(4, 16).Value = ws.Cells(t, 9)
        End If
    Next t

Next ws

End Sub

Sub cleartotal()

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("I2:L150") = ""
    ws.Range("P2:Q4") = ""
    
Next ws
    
End Sub
