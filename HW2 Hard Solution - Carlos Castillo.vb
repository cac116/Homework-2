Sub FindMax()
    Dim PercentRnG, VolumeRnG As Range
    Dim Max_Percent_Change, Min_Percent_Change, Max_Volume As Double
    Dim MaxPercentRow, MinPercentRow, MaxVolumeRow, lastRow  As Long
    worksheet_count = ActiveWorkbook.Worksheets.Count

    For worksheet_index = 1 To worksheet_count

        lastRow = Worksheets(worksheet_index).Cells(Rows.Count, "Q").End(xlUp).Row
        
        Set PercentRnG = Worksheets(worksheet_index).Range("Q1" & ":" & "Q" & lastRow)
        Set VolumeRnG = Worksheets(worksheet_index).Range("K1" & ":" & "K" & lastRow)

        Max_Percent_Change = Worksheets(worksheet_index).Application.WorksheetFunction.Max(PercentRnG)
        Min_Percent_Change = Worksheets(worksheet_index).Application.WorksheetFunction.Min(PercentRnG)
        Max_Volume = Worksheets(worksheet_index).Application.WorksheetFunction.Max(VolumeRnG)

        MaxPercentRow = Worksheets(worksheet_index).Application.WorksheetFunction.Match(Max_Percent_Change, PercentRnG, 0)
        MinPercentRow = Worksheets(worksheet_index).Application.WorksheetFunction.Match(Min_Percent_Change, PercentRnG, 0)
        MaxVolumeRow = Worksheets(worksheet_index).Application.WorksheetFunction.Match(Max_Volume, VolumeRnG, 0)

        Worksheets(worksheet_index).Range("T1").Value = "<ticker>"
        Worksheets(worksheet_index).Range("U1").Value = "<value>"

        Worksheets(worksheet_index).Range("S2").Value = "<greatest % increase>"
        Worksheets(worksheet_index).Range("T2").Value = Worksheets(worksheet_index).Range("M" & MaxPercentRow).Value 
        Worksheets(worksheet_index).Range("U2").Value = Worksheets(worksheet_index).Range("Q" & MaxPercentRow).Value
        Worksheets(worksheet_index).Range("U2").Style = "Percent"
        Worksheets(worksheet_index).Range("U2").NumberFormat = "0.00%"

        Worksheets(worksheet_index).Range("S3").Value = "<greatest % decrease>"
        Worksheets(worksheet_index).Range("T3").Value = Worksheets(worksheet_index).Range("M" & MinPercentRow).Value 
        Worksheets(worksheet_index).Range("U3").Value = Worksheets(worksheet_index).Range("Q" & MinPercentRow).Value
        Worksheets(worksheet_index).Range("U3").Style = "Percent"
        Worksheets(worksheet_index).Range("U3").NumberFormat = "0.00%"
        
        Worksheets(worksheet_index).Range("S4").Value = "<greatest total value>"
        Worksheets(worksheet_index).Range("T4").Value = Worksheets(worksheet_index).Range("M" & MaxVolumeRow).Value 
        Worksheets(worksheet_index).Range("U4").Value = Worksheets(worksheet_index).Range("K" & MaxVolumeRow).Value

    Next worksheet_index


End Sub