Sub Homework2_Moderate_Carlos_Castillo()
Dim ticker_name As String
Dim ticker_volume, ticker_min, ticker_max As Double
Dim lastRow As Long
Dim lastColumn As Long
Dim Summary_Table_Row As Integer
worksheet_count = ActiveWorkbook.Worksheets.Count

    For worksheet_index = 1 To worksheet_count
        Worksheets(worksheet_index).Range("J:W").Clear
    Next worksheet_index
    MsgBox ("All clear")

    For worksheet_index = 1 To worksheet_count
        lastRow = Worksheets(worksheet_index).Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        For worksheet_row_index = 2 To lastRow

            If Worksheets(worksheet_index).Cells(worksheet_row_index, 1).Value = Worksheets(worksheet_index).Cells(worksheet_row_index + 1, 1).Value Then
            'On Error GoTo ErrorHandler
                ticker_volume = Worksheets(worksheet_index).Cells(worksheet_row_index, 7).Value + ticker_volume

            Else
                ticker_volume = Worksheets(worksheet_index).Cells(worksheet_row_index, 7).Value + ticker_volume
                ticker_name = Worksheets(worksheet_index).Cells(worksheet_row_index, 1).Value
                
                'Print all tickers for Easy solution on column J
                Worksheets(worksheet_index).Range("J1").Value = Range("A1").Value
                Worksheets(worksheet_index).Range("J" & Summary_Table_Row).Value = ticker_name

                'Print all tickers' volumes for Easy solution
                Worksheets(worksheet_index).Range("K1").Value = Range("G1").Value
                Worksheets(worksheet_index).Range("K" & Summary_Table_Row).Value = ticker_volume
                
                'Print all tickers for Moderate solution on column J
                Worksheets(worksheet_index).Range("M1").Value = Range("A1").Value
                Worksheets(worksheet_index).Range("M" & Summary_Table_Row).Value = ticker_name

                'Print all tickers' open values for Moderate solution
                Worksheets(worksheet_index).Range("N2").Value = Worksheets(worksheet_index).Cells(2, 3).Value
                Worksheets(worksheet_index).Range("N1").Value = Range("C1").Value
                ticker_min_value = Worksheets(worksheet_index).Cells(worksheet_row_index + 1, 3).Value
                Worksheets(worksheet_index).Range("N" & Summary_Table_Row + 1).Value = ticker_min_value

                'Print all tickers' close values for Moderate solution on colum O
                Worksheets(worksheet_index).Range("O1").Value = Range("F1").Value
                ticker_max_value = Worksheets(worksheet_index).Cells(worksheet_row_index, 6).Value
                Worksheets(worksheet_index).Range("O" & Summary_Table_Row).Value = ticker_max_value

                'Yearly change for Moderate solution
                Worksheets(worksheet_index).Range("P1").Value = "<yearly change>"
                Worksheets(worksheet_index).Range("P" & Summary_Table_Row).Value = Worksheets(worksheet_index).Range("O" & Summary_Table_Row).Value - Worksheets(worksheet_index).Range("N" & Summary_Table_Row).Value
                    If Worksheets(worksheet_index).Range("P" & Summary_Table_Row).Value > 0 Then
                    Worksheets(worksheet_index).Range("P" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                    Worksheets(worksheet_index).Range("P" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If

                'Percent change for Moderate solution
                Worksheets(worksheet_index).Range("Q1").Value = "<percent change>"
                'Worksheets(worksheet_index).Range("Q" & Summary_Table_Row).Value = (Worksheets(worksheet_index).Range("O" & Summary_Table_Row).Value / Worksheets(worksheet_index).Range("N" & Summary_Table_Row).Value) - 1

                If Worksheets(worksheet_index).Range("N" & Summary_Table_Row).Value = 0 Then ' if denominator equals 0 then division by 0 occurs
                Worksheets(worksheet_index).Range("Q" & Summary_Table_Row).Value = "0"
                Else
                Worksheets(worksheet_index).Range("Q" & Summary_Table_Row).Value = (Worksheets(worksheet_index).Range("O" & Summary_Table_Row).Value / Worksheets(worksheet_index).Range("N" & Summary_Table_Row).Value) - 1
                End If

                'Add 1 to variable helping with printing values
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset ticker volume counter for next iteration corresponding to a new ticker
                ticker_volume = 0

            End If
        Next worksheet_row_index
        'Autofits the columns width
        Worksheets(worksheet_index).Range("A:V").Columns.AutoFit
        'Aligns cells to the top
        Worksheets(worksheet_index).Cells.VerticalAlignment = xlTop
        'Aligns cells to the left
        Worksheets(worksheet_index).Cells.HorizontalAlignment = xlLeft
        'Two decimals for yearly change and percent change
        Worksheets(worksheet_index).Columns("P").NumberFormat = "0.00"
        Worksheets(worksheet_index).Columns("Q").Style = "Percent"
        Worksheets(worksheet_index).Columns("Q").NumberFormat = "0.00%"

    Next worksheet_index
        
End Sub
