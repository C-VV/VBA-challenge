Sub Stock_Ticker()

 'Define everything
Dim Ticker As String
Dim Vol As Double
Dim Year_Open As Double
Dim Year_Close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greattest_Total_Vol As Long

'Define worksheet and run through workbook
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate
 

'Cell Headers and titles

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = " Yearly Change"
ws.Cells(1, 11).Value = " Percent Change"
ws.Cells(1, 12).Value = " Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = " Greatest % Decrease "
ws.Cells(4, 15).Value = " Greatest Total Volume"

Yearly_Change = 0
Percent_Change = 0
Vol = 0
Start = 2
Dim Sum_table As Long
Sum_table = 2

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To Lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            Yearly_Change = (ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value)
             
                If ws.Cells(Start, 3).Value = 0 Then
                Precent_Change = 0
             
            
            Else: Percent_Change = Round((Yearly_Change / ws.Cells(Start, 3).Value * 100), 2)
            End If
            
            
            Vol = Vol + ws.Cells(i, 7).Value
            Start = i + 1
            
       
            ws.Cells(Sum_table, 9).Value = Ticker
            ws.Cells(Sum_table, 10).Value = Yearly_Change
            ws.Cells(Sum_table, 11).Value = Percent_Change & "%"
            
            
            
                If ws.Cells(Sum_table, 10).Value > 0 Then
                ws.Cells(Sum_table, 10).Interior.ColorIndex = 4
                
                ElseIf ws.Cells(Sum_table, 10).Value < 0 Then
                ws.Cells(Sum_table, 10).Interior.ColorIndex = 3
                
                
                End If

            
            ws.Cells(Sum_table, 12).Value = Vol
            
            Sum_table = Sum_table + 1
            Vol = 0
            Yearly_Change = 0
            Else
            Vol = Vol + ws.Cells(i, 7).Value
            End If
 
 
Next

ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)) * 100 & "%"
Greatest_Increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
ws.Range("P2") = ws.Cells(Greatest_Increase + 1, 9)
ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)) * 100 & "%"
Greatest_Decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)), ws.Range("K2:K" & Lastrow), 0)
ws.Range("P3") = ws.Cells(Greatest_Decrease + 1, 9)
ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & Lastrow))
Greatest_Total_Vol = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Lastrow)), ws.Range("L2:L" & Lastrow), 0)

For Each sht In ThisWorkbook.Worksheets
    sht.Cells.EntireColumn.AutoFit
  Next sht


Next ws
End Sub
