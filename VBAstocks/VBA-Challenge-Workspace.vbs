Sub VbaChallengeHw()

    Dim yearStartDate As Long
    Dim yearEndDate As Long
    Dim yearPercent As Double
    Dim startingPrice As Double
    Dim endingPrice As Double
    Dim totalVolume As Double
    Dim tickerName As String
    Dim yearDifference As Double
    Dim lastRow As Long
    Dim Summary_Table_Row As Long
    Dim tickerNameCount As Integer

    Summary_Table_Row = 2
    totalVolume = 0
    lastRow = cells(Rows.Count, 1).End(xlUp).Row

' Creating the Headers    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    Range("I1:L1").Borders(xlEdgeBottom).LineStyle = xlContinious
    Range("I1:L1").Borders(xlEdgeBottom).Weight = xlThick
    Range("I1:L1").Font.FontStyle = "Italic"
    'Range("I1:L1").Font.Size = 14

' the for loop to run through the sheet to populate report
    For i = 2 To lastRow
        
        If cells(i + 1, 1).Value <> cells(i, 1).Value Then
            tickerName = cells(i, 1).Value
            startingPrice = cells(i - tickerNameCount, 3).Value
            endingPrice = cells(i, 6).Value
            yearPercent = (((endingPrice - startingPrice) / startingPrice) * 100)
            yearDifference = (endingPrice - startingPrice)
            Range("I" & Summary_Table_Row).Value = tickerName
            Range("J" & Summary_Table_Row).Value = yearDifference
            Range("J" & Summary_Table_Row).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                If yearDifference < 0 Then
                    Range("J" & Summary_Table_Row).Cells.Interior.Color = 192
                ElseIf yearDifference >= 0 Then
                    Range("J" & Summary_Table_Row).Cells.Interior.Color = 4697456
                Else
                End If
            Range("K" & Summary_Table_Row).Value = yearPercent
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            Range("L" & Summary_Table_Row).Value = totalVolume
            Range("L" & Summary_Table_Row).NumberFormat = "#,##0_);[Red](#,##0)"
            Summary_Table_Row = Summary_Table_Row + 1
            'searched from https://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/
            ' Range("J" & Summary_Table_Row).Value = startingPrice
            ' Range("K" & SUmmary_Table_Row).Value = endingPrice
            
            totalVolume = 0
            tickerNameCount = 0
        Else
            totalVolume = totalVolume + cells(i, 7).Value
            tickerNameCount = tickerNameCount + 1

        End If

    Next i

End Sub
' ---- Something tried, but decided to put conditional formatting into for loop.   
' Sub FormattingCells(lastRow)
'     Dim wb As Workbook
'     Dim ws As Worksheet
'     Dim cell As Range
'     Set ws = ActiveSheet
    
'     lastRow = cells(Rows.Count, 1).End(xlUp).Row
'     For Each cell In ws.Range("J" & 2, ":", "J" & lastRow)
'         If cell.Value >= 0 Then
'             cell.Interior.Color = 4697456
'         ElseIf cell.Value < 0 Then
'             cell.Interior.Color = 192
'         Else
'             cell.Interior.Color = 16777215
'         End If

'     Next cell


' End Sub
' ----- For looping across worhsheets
    ' 'Loop obtained from https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook
    ' Dim Current As Worksheet
    ' For Each Current in Worksheets
    ' MsgBox Current.Name
    ' Next
' Sub VBAStocks()
'     Dim tickerValue As String
'     'Dim Brand_Name As String
'     Dim startDate As Long
'     Dim endDate As Long
'     Dim percentChange as Double
'     Dim stockVolume as Long
'     'Dim Brand_Total As Double
'     Dim lastRow As Long
'     Dim Current As Worksheet

'     Set Current = ActiveSheet

'     stockVolume = 0

'     Dim Summary_Table_Row as Integer
'     Summary_Table_Row = 2

'     For i = 2 To lastRow
'         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'             tickerValue = Cells(i, 1).Value
'             stockVolume = stockVolume + Cells(i, 7).Value
'             Range("I" & Summary_Table_Row).Value = tickerValue
'             Range("J" & Summary_Table_Row).Value = stockVolume

'             Summary_Table_Row = Summary_Table_Row + 1
            
'             stockVolume = 0
'         Else
'             stockVolume = stockVolume + Cells(i, 7).Value
'         End If
    
'     Next i 
    




    
'     ' tickerValue = Cells(lastRow,1)
     
'     ' For ticker = 1 to lastRow
'     '     lastRow = Cells(Row.Count, 1).End(xlUp).Row
'     '     MSgBox "lastRow"
'     ' Next
'     ' lastRow = Cells(Row.Count, 1).End(xlUp).Row
    

' '     'for calling down to the last row (shows the last row)
' '     ' * Create a script that will loop through all the stocks for one year for each run and take the following information.

' ' '   * The ticker symbol.

' ' '   * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

' ' '   * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

' ' '   * The total stock volume of the stock.


' ' For j = 4 to 8
' '     For }
' '     =
' ' Next j 


'     ' 'Loop obtained from https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook
'     ' Dim Current As Worksheet
'     ' For Each Current in Worksheets
'     ' MsgBox Current.Name
'     ' Next

' End Sub