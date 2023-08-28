VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock()


Dim ws As Worksheet
    
'For Each loop for the worksheets
    For Each ws In ThisWorkbook.Worksheets

'Activate the worksheet
    ws.Activate

'Column Names
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Dimensions

Dim Total As Double
Dim i As Double
Dim RowCount As Double
Dim j As Double
Dim OpeningValue As Double
Dim PercentChange As Double

'Set Opening Value
    OpeningValue = ws.Range("C2")
    max_value = 0
    Min_value = 0
    GreatestTotal = 0

'Create a dynamic range to work across all worksheets
        j = 2
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To RowCount
        Total = Total + ws.Cells(i, 7)

'Calculating the yearly change
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then

        close_value = ws.Cells(i, 6).Value
        yearly_change = close_value - OpeningValue

'Check if opening value is zero to catch div/0 errors
    If OpeningValue = 0 Then
        PercentChange = 0
    Else 'opening val is not zero
        PercentChange = yearly_change / OpeningValue
    End If
        
'Creation of Summary Table
        ws.Range("K" & j).Value = PercentChange
' If % change is > max_value then max_value is %change
    If PercentChange > max_value Then
        max_value = PercentChange
        ws.Range("Q2").Value = PercentChange
        ws.Range("P2").Value = ws.Cells(i, 1).Value
    End If
    
    If PercentChange < Min_value Then
        Min_value = PercentChange
        ws.Range("Q3").Value = PercentChange
        ws.Range("P3").Value = ws.Cells(i, 1).Value
    End If
    If Total > GreatestTotal Then
        GreatestTotal = Total
        ws.Range("Q4").Value = Total
        ws.Range("P4").Value = ws.Cells(i, 1).Value
    End If

' Worksheet Function
' Create Conditional formating for Positive and Negative values
        ws.Range("K" & j).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    If ws.Range("K" & j) >= 0 Then
        ws.Range("K" & j).Interior.ColorIndex = 4
    Else
        ws.Range("K" & j).Interior.ColorIndex = 3
    
    End If
    If ws.Range("J" & j) >= 0 Then
        ws.Range("J" & j).Interior.ColorIndex = 4
    Else
        ws.Range("J" & j).Interior.ColorIndex = 3
    
    End If
'Print Ticker, yearly_change & Total
        ws.Range("I" & j).Value = ws.Cells(i, 1).Value
        ws.Range("L" & j).Value = Total
        ws.Range("J" & j).Value = yearly_change
        j = j + 1
        Total = 0
        OpeningValue = ws.Cells(i + 1, 3).Value

    End If
    Next i

'Alt creation of worksheet func. for Summary Table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"



Next ws

End Sub


