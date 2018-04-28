Sub Homework2_Easy_Carlos_Castillo()
Dim ticker_name As String
Dim ticker_volume As Double
Dim lastRow As Long
Dim lastColumn As Long
Dim Summary_Table_Row As Integer
worksheet_count = ActiveWorkbook.Worksheets.Count

    For worksheet_index = 1 To worksheet_count
        Worksheets(worksheet_index).Range("J:K").Clear
    Next worksheet_index
    MsgBox ("All clear")

    For worksheet_index = 1 To worksheet_count
        lastRow = Worksheets(worksheet_index).Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        For worksheet_row_index = 2 To lastRow
            If Worksheets(worksheet_index).Cells(worksheet_row_index, 1).Value = Worksheets(worksheet_index).Cells(worksheet_row_index + 1, 1).Value Then
                ticker_volume = Worksheets(worksheet_index).Cells(worksheet_row_index, 7).Value + ticker_volume
            Else
                ticker_volume = Worksheets(worksheet_index).Cells(worksheet_row_index, 7).Value + ticker_volume
                ticker_name = Worksheets(worksheet_index).Cells(worksheet_row_index, 1).Value
                Worksheets(worksheet_index).Range("J1").Value = Range("A1").Value
                Worksheets(worksheet_index).Range("K1").Value = Range("G1").Value
                Worksheets(worksheet_index).Range("J" & Summary_Table_Row).Value = ticker_name
                Worksheets(worksheet_index).Range("K" & Summary_Table_Row).Value = ticker_volume
                Summary_Table_Row = Summary_Table_Row + 1
                ticker_volume = 0
            End If
        Next worksheet_row_index
        'Autofits the columns width
        Worksheets(worksheet_index).Range("A:V").Columns.AutoFit
        'Aligns cells to the top
        Worksheets(worksheet_index).Cells.VerticalAlignment = xlTop
        'Aligns cells to the left
        Worksheets(worksheet_index).Cells.HorizontalAlignment = xlLeft
    Next worksheet_index
End Sub