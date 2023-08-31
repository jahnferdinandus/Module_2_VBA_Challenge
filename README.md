# VBA-challenge

The VBA Challenge work is in total of 5 files.
1. alphabetical_testing was used to write and test the code
2. Multiple_year_stock_data was used to run the completed code and find the results
3. Final results can be viewed from the 3 PNG files withing the folder. 
4. The text file contains the code created

# Key Findings

1. RKS has consistently decreased for 2 years in 2018 & 2019
2. The greatest increase in 2018, 2019 & 2020 have been from stocks THB, RYU & YDI respectively.
3. The highest traded stocks were QKN in both 2018 & 2020

# Referencing

Special thanks to my Tutor and TAs within ASK BCS in supporting me to understand, create and fintune my code. The creation of the following codes were supported by my tutors.

 RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To RowCount

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

    If PercentChange > max_value Then
        max_value = PercentChange

        Dim ws As Worksheet
    
'For Each loop for the worksheets
    For Each ws In ThisWorkbook.Worksheets


Websites Refered

https://www.homeandlearn.org/excel_vba_for_loops.html
https://www.automateexcel.com/vba/cycle-and-update-all-worksheets/
https://learn.microsoft.com/en-us/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-excel-worksheet-functions-in-visual-basic


