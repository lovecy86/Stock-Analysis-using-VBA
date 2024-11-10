Option Explicit

'Global variable declaration
Dim wb As Workbook                'declaring workbook as wb
Dim ws As Worksheet                'declaring worksheet as w
Dim i As Double                         'for looping rows/columns
Dim j As Double                         'for looping rows/columns
Dim lastrow As Long                  'for calculating last rows in a column

'This sub procedure filter unique values from column A to column I

Public Sub Consolidated()

Dim rnge As Range                               ' rnge is the range to which the unique values are copied. Here it is column I
Dim criteriaRange As Range                  'criteria range is the range from which unique values are filtered. Here it is Column A
Set wb = ThisWorkbook                       ' setting ThisWorkbook to wb

        For Each ws In ThisWorkbook.Worksheets                                                                                         'for loop for looping worksheets
                     Set rnge = ws.Range("I1")                                                                                                       'the unique values are copied to column I
                     lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row                                                        'finding the last row of column A
                     Set criteriaRange = ws.Range("A1:A" & lastrow)                                                                    'from criteria range i.e column A, the unique values are copied
                    criteriaRange.AdvancedFilter Action:=xlFilterCopy, copytorange:=rnge, Unique:=True          ' using advanced filter function to filter unique values(Ticker symbols)from column A to column I
        
Next ws

'Call Header                                                                                                                                                 ' calling 'Header' function to insert headers in all columns
'Call QuarterlyAndPercentageChange                                                                                                           'calling 'QuarterlyandPercentageChange function' to calculate quarterly change(column J) and percenatge change (column K)
'Call ConditonalFormatting
'Call TotalStockVolumes                                                                                                                              'calling 'TotalStockVolumes' to calculate total stock volume
'Call Calculations                                                                                                                                          'calling 'Calculations' function to calculate Greatest percentage increase, greatest percentage decrease and max stock volume.
End Sub


'Assigning headers for the columns

Public Sub Header()

Set wb = ThisWorkbook

For Each ws In ThisWorkbook.Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quaterly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Columns.AutoFit                                                      ' autofitting the columns
        
Next ws

End Sub

'Calculate quarterly and percenatge change

Public Sub QuarterlyAndPercentageChange()
                                                        
Dim open_price As Double
Dim close_price As Double
Dim PercentageChange As Double
Dim lastrowJ As Long

Set wb = ThisWorkbook

For Each ws In ThisWorkbook.Worksheets

        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row                  'computes last row of column A
        open_price = ws.Cells(2, 3).Value                                                 'stores the open price of ticker AAF. Here it is open_price = 5.02.This is the opening price of AAF
        j = 2                                                                                              'assigning j =2

        For i = 2 To lastrow                                                                         'Using for-next loop on column A

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then            ' compares the value of each row of column A with the subsequent row of column A.
            
                close_price = ws.Cells(i, 6).Value                                         'When the values do not match,the value of  column 'close' of ith row is moved to variable close_price.This gives the closing price at the end of the quarter. For AAF, close_price =5.08                                                                                     'j is incremente.This is because the frst quarterly change value should be stored in  row 2 of column J. With the execution of loop, the quarterly change values is stored in the subsequent rows of column J
                ws.Cells(j, 10).Value = close_price - open_price                   'In column J, the value of quarterly change is moved.
                PercentageChange = ws.Cells(j, 10).Value / open_price         'percentage change = Quarterly Change/ open_price
                ws.Cells(j, 11).Value = PercentageChange                             'in column L the value of percentage change is moved
                open_price = ws.Cells(i + 1, 3).Value                                    ' the value of (i+1)th row of Column C becomes the open_pirce for the next ticker
                j = j + 1                                                                                  ' the row of column J and L are incremented
                
            End If
            
        Next i
        lastrowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row               '
        ws.Range("J1:J" & lastrowJ).NumberFormat = "General"

Next ws

End Sub

' To calculate Greatest Percentage Increase, Gretest Percentage Decrease and Greatest Total Volume

Public Sub Calculations()

'defining variables
            
Dim ColRange As Range   'computes last row of column K
Dim VolRange As Range   'computes last row of column K
Dim TickerValue As Variant   'TickerValue is used as variant data type
Dim lastrow1

Set wb = ThisWorkbook

For Each ws In Worksheets

        
            lastrow = ws.Cells(Rows.Count, "K").End(xlUp).Row           'computes last row of column K
            lastrow1 = ws.Cells(Rows.Count, "L").End(xlUp).Row          'computes last row of column L
            Set ColRange = ws.Range("K1:K" & lastrow)                       'setting ColRange variable to  column K range
            Set VolRange = ws.Range("L1:L" & lastrow1)                      'setting VolRange variable to  column L range
            
            'To compute Greatest Percentage Increase
            
            ws.Range("O2").Value = "Greatest % increase"                                                                                                  'O2 cell value is "Greatest % increase"
            ws.Range("Q2").Value = Application.WorksheetFunction.Max(ColRange)                                                        'using Max function to calculate maximum percentage change and is stored in Q2 cell
            TickerValue = Application.Match(ws.Range("Q2").Value, ws.Range("K1:K" & lastrow), 0)                               'uses match functon to match the values of Q2 cell with the data in column K
            ws.Range("P2").Value = ws.Cells(TickerValue, "K").Offset(0, -2).Value                                                           'once the match is found, the coreesponding ticker symbol is stored in cell P3


            ' To compute Greatest percentage decrease
            
            ws.Range("O3").Value = "Greatest % decrease"                                                                                                          'O3 cell value is "Greatest % decrease"
            ws.Range("Q3").Value = Application.WorksheetFunction.Min(ColRange)                                                                  'using MIN function to calculate minimum percentage change and is stored in Q3 cell
            TickerValue = Application.Match(ws.Range("Q3").Value, ws.Range("K1:K" & lastrow), 0)                                        'uses match functon to match the values of Q3 cell with the data in column K
            ws.Range("P3").Value = ws.Cells(TickerValue, "K").Offset(0, -2).Value                                                                    'once the match is found, the coreesponding ticker symbol is stored in cell P3

            
            'To compute Greatest Total Volume
            
            ws.Range("O4").Value = "Greatest TotalVolume"                                                                                                            'O4 cell value is "Greatest Total Volume"
            ws.Range("Q4").Value = Application.WorksheetFunction.Max(VolRange)                                                                      'using Max function to calculate maximum total volume and is stored in Q4 cell
            TickerValue = Application.Match(ws.Range("Q4").Value, ws.Range("L1:L" & lastrow1), 0)                                            'uses match functon to match the values of Q4 cell with the data in column L
            ws.Range("P4").Value = ws.Cells(TickerValue, "L").Offset(0, -3).Value                                                                          'once the match is found, the coreesponding ticker symbol is stored in cell P3
            
            ws.Columns.AutoFit
            ws.Range("K1:K" & lastrow).NumberFormat = "0.00%"
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
Next ws

End Sub

'To apply conditional formatting on Quarterly change

Public Sub ConditonalFormatting()

Set wb = ThisWorkbook

For Each ws In ThisWorkbook.Worksheets

    lastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row                   'computes last row of column J
        For i = 2 To lastrow                                                                  'loops the rows of column J
                If (Cells(i, 10).Value < 0) Then                                          'if the value of ith row of column J < 0, then red color is filled in the cell else green color is filled
                        ws.Cells(i, 10).Interior.ColorIndex = 9
                Else
                         ws.Cells(i, 10).Interior.ColorIndex = 51
                End If
        Next i
Next ws
                
End Sub

'to compute the Total Stock Volume

Public Sub TotalStockVolumes()

'Variable declaration
Dim sum As Double
Dim lastrowA As Long
Dim lastrowI As Long

For Each ws In Worksheets

        lastrowI = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row                                   'computes last row of column I
        lastrowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row                                  'computes last row of column A

                For j = 2 To lastrowI                                                                                'applying for loop for column I
                sum = 0                                                                                                    'initializing sum to 0

                        For i = 2 To lastrowA                                                                       'applying loop to column A

                                    If ws.Cells(i, 1).Value = ws.Cells(j, 9).Value Then                 'if the ith row of column A = to the jth row of column I
                                                sum = sum + ws.Cells(i, 7).Value                               'then the value of the ith row of column G is added to sum
                                                ws.Cells(j, 12).Value = sum                                        'the sum value is stored in jth row of column L
                                     End If

                          Next i

                Next j

Next ws

End Sub





