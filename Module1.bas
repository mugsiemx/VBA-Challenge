Attribute VB_Name = "Module1"
'Stock Market Analysis Project
'MSU Data Analytics BootCamp Instructions
' Create a script that loops through all the stocks for one year and outputs the following information:
    '- The ticker symbol.
    '- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    '- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    '- The total stock volume of the stock.
'**Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

'Created by Sheila LaRoue on 1/6/2023

Sub StockMarketAnalysis()

'Miscellaneous variables
Dim end_row, next_row, end_col, next_col As Integer
Dim i, j As Integer
Dim ws_index As Long

'Stock information variables
Dim ticker As String
Dim open_price, close_price, yearly_chg, percent_chg As Double
Dim total_stock_vol As LongLong

'Overall greatest change variables
Dim most_ticker As String
Dim most_increase, most_decrease As Double
Dim most_total_vol As LongLong
Dim most_ticker_array As Variant
most_ticker_array = Array("ar0", "ar1", "ar2")
most_increase = 0#
most_decrease = 0#

'Column headers for new section
Dim column_headers() As Variant
column_headers() = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

'Overall greatest change column headers
Dim most_headers() As Variant
most_headers() = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume", "Ticker", "Value")


'Begin looping through the worksheets in the workbook from last to first,
    ' collect the required information per sheet and write the summary results for each ticker in the same sheet,
    ' then write the largest changes to the first worksheet in the workbook
For ws_index = ThisWorkbook.Sheets.Count To 1 Step -1

'Count the rows and columns to set up for a dynamic selection of results placements
Application.ScreenUpdating = False
Worksheets(ws_index).Activate
Range("A1").Select
end_row = Range("A1").End(xlDown).Row
end_col = Range("A1").End(xlToRight).Column

    'For summary individual ticker section, leave one blank column and add column headers, autofit column widths to fit headers
        For j = 0 To 3
            next_col = end_col + 2 + j
            Cells(1, (next_col)).Value = column_headers(j)
            Cells(1, (next_col)).Font.Bold = True
            Columns(next_col).AutoFit
        Next j
    
    'Read each line and calculate the answers as noted in the initial comments section prior to the variable declarations above
        j = 1                       'ticker row number
        For i = 2 To end_row + 1      'individual stock row number
        
            If Cells(i, 1).Value <> ticker Then
                If i <> 2 Then          'skip the initial ticker change from blank
                yearly_chg = close_price - open_price
                           
                    If open_price = 0 Then
                    percent_chg = 0     'avoid division by zero issue
                    Else
                    percent_chg = yearly_chg / open_price
                        'update if ticker changes and %s fit < or > criteria
                        If most_increase <> percent_chg And ticker <> most_ticker_array(0) Then
                            If most_increase < percent_chg Then
                            most_increase = percent_chg
                            most_ticker_array(0) = ticker
                            End If
                        End If
                        If most_decrease <> percent_chg And ticker <> most_ticker_array(1) Then
                            If most_decrease > percent_chg Then
                            most_decrease = percent_chg
                            most_ticker_array(1) = ticker
                            End If
                        End If
                        'update per ticker, total largest stock volume
                        If most_total_vol <> total_stock_vol And ticker <> most_ticker_array(2) Then
                            If most_total_vol < total_stock_vol Then
                            most_total_vol = total_stock_vol
                            most_ticker_array(2) = ticker
                            End If
                        End If
            
                j = j + 1   'next ticker row number
                'Write the individual ticker summary lines in the same worksheet
                Cells(j, 9).Value = ticker
                Cells(j, 10).Value = yearly_chg
                Cells(j, 10).NumberFormat = "0.00"
                    If Cells(j, 10).Value >= 0 Then
                        Cells(j, 10).Interior.ColorIndex = 4    'positive $ change or zero colors cell green
                        Else
                        Cells(j, 10).Interior.ColorIndex = 3        'negative change colors cell red
                    End If
                Cells(j, 11).Value = percent_chg
                Cells(j, 11).NumberFormat = "0.00%"             'format percent change into a percentage
                    If Cells(j, 11).Value >= 0 Then
                        Cells(j, 11).Interior.ColorIndex = 4    'positive percent change or zero colors cell green
                        Else
                        Cells(j, 11).Interior.ColorIndex = 3        'negative change colors cell red
                    End If
                Cells(j, 12).Value = total_stock_vol
                End If
                    
                'reset individual ticker summary values variables
                yearly_chg = 0
                percent_chg = 0
                total_stock_vol = 0
                End If
            'new ticker creates next summary row information
            ticker = Cells(i, 1).Value
            open_price = Cells(i, 3).Value
            total_stock_vol = total_stock_vol + Cells(i, 7).Value
            Else
            'same ticker, continue calculations
            close_price = Cells(i, 6).Value
            total_stock_vol = total_stock_vol + Cells(i, 7).Value
            End If
        Next i  'next individual stock row number
        
'The grand finale!! Write the greatest change section in the same worksheet!!
    next_row = 2                    'start with second row
    next_col = next_col + 3         'leave two blank columns between individual ticker summaries

        For i = 0 To 4              'write the 5 row/column greatest change section labels
            If i < 3 Then           'populate greatest change section row descriptions first
                Cells(next_row, (next_col)).Value = most_headers(i)
                Cells(next_row, (next_col)).Font.Bold = True
                Columns(next_col).AutoFit
                    If i = 0 Then       'greatest % increase
                        Cells(next_row, (next_col + 1)).Value = most_ticker_array(i)
                        Cells(next_row, (next_col + 2)).Value = most_increase
                        Cells(next_row, (next_col + 2)).NumberFormat = "0.00%"
                        Cells(next_row, (next_col + 2)).Interior.ColorIndex = 4    'positive percent change or zero colors cell green
                    ElseIf i = 1 Then   'greatest % decrease
                        Cells(next_row, (next_col + 1)).Value = most_ticker_array(i)
                        Cells(next_row, (next_col + 2)).Value = most_decrease
                        Cells(next_row, (next_col + 2)).NumberFormat = "0.00%"
                        Cells(next_row, (next_col + 2)).Interior.ColorIndex = 3    'negative change colors cell red
                    Else   'i = 2       'greatest total volume
                        Cells(next_row, (next_col + 1)).Value = most_ticker_array(i)
                        Cells(next_row, (next_col + 2)).Value = most_total_vol
                    End If
                next_row = next_row + 1
            Else                     'populate greatest change section column descriptions last
                next_col = next_col + 1
                Cells(1, (next_col)).Value = most_headers(i)
                Cells(1, (next_col)).Font.Bold = True
            End If
        Next i 'write the 5 row/column greatest change section labels
            
        'Reset variables and accumulators
        most_ticker = " "
        most_increase = 0#
        most_decrease = 0#
        most_total_vol = 0
            
 Next ws_index  'continue with reverse worksheet loop
 
Application.ScreenUpdating = True
MsgBox "THANK YOU for grading my project!!"

End Sub
