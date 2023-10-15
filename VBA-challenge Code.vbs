Attribute VB_Name = "Module1"
Sub Ticker()

'Loop thru all worksheets

For Each ws In Worksheets

'Set LastRow

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
MsgBox (LastRow)

'Create variables for Ticker

Dim Ticker As String
Dim SummaryRow As Integer
Dim i As Long

'Create variable for Yearly Change and set beginning value

Dim Yearly_Change As Double
Dim Year_Open As Double
Dim Year_Close As Double

Yearly_Change = 0

'Create variables for % Change and set beginning value

Dim Percent_Change As Double
Percent_Change = 0

'Create variables for Total_Stock and set beginning value

Dim Volume As Double
Volume = 0
    
'Create Column and Row Headers

ws.Range("I1").EntireColumn.Insert
ws.Cells(1, 9).Value = "Ticker"

ws.Range("J1").EntireColumn.Insert
ws.Cells(1, 10).Value = "Yearly Change"

ws.Range("K1").EntireColumn.Insert
ws.Cells(1, 11).Value = "Percent Change"
ws.Range("K:K").NumberFormat = "0.00%"

ws.Range("L1").EntireColumn.Insert
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Range("L:L").NumberFormat = "0"

ws.Range("A:L").Columns.AutoFit

SummaryRow = 2

'Loop thru to last row on sheet

    For i = 2 To LastRow

        'Find the opening price for each ticker

        If Year_Open = 0 Then
            
            Year_Open = ws.Cells(i, 3).Value
            
        End If
        
        'Find individual ticker data and add to next summary row in appropriate column

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & SummaryRow).Value = Ticker
            
            Year_Close = ws.Cells(i, 6).Value
            Yearly_Change = Year_Close - Year_Open
            ws.Range("J" & SummaryRow).Value = Yearly_Change
            
                If ws.Range("J" & SummaryRow).Value < 0 Then
                    ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
                End If
            
            Percent_Change = Yearly_Change / Year_Open
            ws.Range("K" & SummaryRow).Value = Percent_Change
            
            Volume = Volume + ws.Cells(i, 7).Value
            ws.Range("L" & SummaryRow).Value = Volume
            
            SummaryRow = SummaryRow + 1
            
            Volume = 0
            Year_Open = 0
            
        Else
        
            Volume = Volume + ws.Cells(i, 7).Value
                   
        End If
        
    Next i
    
'Summary

ws.Range("O1").EntireColumn.Insert
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

ws.Range("P1").EntireColumn.Insert
Cells(1, 16).Value = "Ticker"

ws.Range("Q1").EntireColumn.Insert
Cells(1, 17).Value = "Value"
Range("Q2:Q3").NumberFormat = "0.00%"

Range("O:Q").Columns.AutoFit

    For i = 2 To LastRow
    
     If ws.Cells(i, 11).Value > Range("Q2").Value Then
            Range("Q2").Value = ws.Cells(i, 11).Value
            Range("P2").Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11).Value < Range("Q3").Value Then
            Range("Q3").Value = ws.Cells(i, 11).Value
            Range("P3").Value = ws.Cells(i, 9).Value
        End If

        If ws.Cells(i, 12).Value > Range("Q4").Value Then
            Range("Q4").Value = ws.Cells(i, 12).Value
            Range("P4").Value = ws.Cells(i, 9).Value
        End If
        
    Next i
        
'Loop thru next worksheet

Next ws

End Sub


